// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    // Subscribes to the default Outlook calendar folder and mirrors appointment
    // create/update/delete events to a Nextcloud CalDAV calendar.
    //
    // Threading model: Outlook COM events fire on the UI thread. We read all COM
    // properties synchronously on the UI thread (fast — memory only), then dispatch
    // the HTTP request to a thread-pool thread so Outlook is never blocked.
    internal sealed class CalDavCalendarSync : IDisposable
    {
        private readonly CalDavSyncService _syncService = new CalDavSyncService();
        private readonly Dictionary<string, CalDavDeleteTracker> _deleteTrackers =
            new Dictionary<string, CalDavDeleteTracker>(StringComparer.OrdinalIgnoreCase);

        private TalkServiceConfiguration _configuration;
        private string _calendarName;
        private Outlook.MAPIFolder _calendarFolder;
        private Outlook.Items _calendarItems;
        private bool _disposed;

        internal void Attach(Outlook.Application application, TalkServiceConfiguration configuration, string calendarName)
        {
            Detach();
            if (application == null || configuration == null || !configuration.IsComplete())
                return;

            _configuration = configuration;
            _calendarName = string.IsNullOrWhiteSpace(calendarName) ? "personal" : calendarName.Trim();

            try
            {
                _calendarFolder = application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                _calendarItems = _calendarFolder.Items;
                _calendarItems.ItemAdd += OnItemAdd;
                _calendarItems.ItemChange += OnItemChange;
                DiagnosticsLogger.Log(LogCategories.CalDav, "Calendar sync attached (calendar=" + _calendarName + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to attach calendar sync.", ex);
                Detach();
            }
        }

        internal void Detach()
        {
            if (_calendarItems != null)
            {
                try { _calendarItems.ItemAdd -= OnItemAdd; }
                catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to unhook ItemAdd.", ex); }
                try { _calendarItems.ItemChange -= OnItemChange; }
                catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to unhook ItemChange.", ex); }
                ComInteropScope.TryFinalRelease(_calendarItems, LogCategories.CalDav, "Failed to release calendar Items COM object.");
                _calendarItems = null;
            }
            if (_calendarFolder != null)
            {
                ComInteropScope.TryFinalRelease(_calendarFolder, LogCategories.CalDav, "Failed to release calendar folder COM object.");
                _calendarFolder = null;
            }
            foreach (var tracker in _deleteTrackers.Values)
                tracker.Dispose();
            _deleteTrackers.Clear();
            _configuration = null;
            DiagnosticsLogger.Log(LogCategories.CalDav, "Calendar sync detached.");
        }

        private void OnItemAdd(object item)
        {
            var appointment = item as Outlook.AppointmentItem;
            if (appointment != null)
                SyncToCalDav(appointment);
        }

        private void OnItemChange(object item)
        {
            var appointment = item as Outlook.AppointmentItem;
            if (appointment != null)
                SyncToCalDav(appointment);
        }

        private void SyncToCalDav(Outlook.AppointmentItem appointment)
        {
            if (_configuration == null || !_configuration.IsComplete())
                return;

            // Read all COM data on the UI thread — accessing COM from a background thread is not safe.

            // Only sync appointments that have a Nextcloud Talk room attached.
            if (!HasTalkToken(appointment))
                return;

            string entryId = null;
            try { entryId = appointment.EntryID; }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to read appointment EntryID.", ex);
                return;
            }
            if (string.IsNullOrEmpty(entryId))
                return;

            string uid = DeriveUid(entryId);
            string ical = null;
            try
            {
                ical = ICalBuilder.Build(appointment, uid);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to build iCal for appointment (uid=" + uid + ").", ex);
                return;
            }

            // Capture immutable values for the background task — no COM access after this point.
            var config = _configuration;
            var calendarName = _calendarName;
            var syncService = _syncService;

            Task.Run(() =>
            {
                try
                {
                    bool ok = syncService.PutAppointment(config, calendarName, uid, ical);
                    DiagnosticsLogger.Log(LogCategories.CalDav, "Appointment synced (uid=" + uid + ", ok=" + ok + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to PUT appointment (uid=" + uid + ").", ex);
                }
            });

            EnsureDeleteTracker(appointment, entryId, uid);
        }

        private void EnsureDeleteTracker(Outlook.AppointmentItem appointment, string entryId, string uid)
        {
            if (_deleteTrackers.ContainsKey(entryId))
                return;

            var tracker = new CalDavDeleteTracker(
                appointment, uid, _configuration, _calendarName, _syncService, entryId,
                id =>
                {
                    if (id != null)
                        _deleteTrackers.Remove(id);
                });
            _deleteTrackers[entryId] = tracker;
        }

        private static bool HasTalkToken(Outlook.AppointmentItem appointment)
        {
            Outlook.UserProperties props = null;
            try
            {
                props = appointment.UserProperties;
                if (props == null)
                    return false;

                // Check both the iCal property name and the Outlook storage name.
                string[] names = new[] { "X-NCTALK-TOKEN", "NcTalkRoomToken" };
                foreach (string name in names)
                {
                    Outlook.UserProperty prop = null;
                    try
                    {
                        prop = props.Find(name, false);
                        if (prop != null)
                        {
                            string val = prop.Value as string;
                            return !string.IsNullOrWhiteSpace(val);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to read user property '" + name + "'.", ex);
                    }
                    finally
                    {
                        if (prop != null)
                            ComInteropScope.TryRelease(prop, LogCategories.CalDav, "Failed to release UserProperty COM object.");
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to check Talk token on appointment.", ex);
                return false;
            }
            finally
            {
                if (props != null)
                    ComInteropScope.TryRelease(props, LogCategories.CalDav, "Failed to release UserProperties COM object.");
            }
        }

        internal static string DeriveUid(string entryId)
        {
            using (var sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(entryId));
                return BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            Detach();
        }

        // Lightweight per-item subscription that fires a CalDAV DELETE on BeforeDelete.
        private sealed class CalDavDeleteTracker : IDisposable
        {
            private readonly string _uid;
            private readonly string _calendarName;
            private readonly TalkServiceConfiguration _configuration;
            private readonly CalDavSyncService _syncService;
            private readonly string _entryId;
            private readonly Action<string> _onDisposed;
            private Outlook.ItemEvents_10_Event _events;
            private bool _disposed;

            internal CalDavDeleteTracker(
                Outlook.AppointmentItem appointment,
                string uid,
                TalkServiceConfiguration configuration,
                string calendarName,
                CalDavSyncService syncService,
                string entryId,
                Action<string> onDisposed)
            {
                _uid = uid;
                _configuration = configuration;
                _calendarName = calendarName;
                _syncService = syncService;
                _entryId = entryId;
                _onDisposed = onDisposed;
                _events = appointment as Outlook.ItemEvents_10_Event;
                if (_events != null)
                    _events.BeforeDelete += OnBeforeDelete;
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                // Capture all needed values before disposal — uid/config/calendarName are plain .NET objects,
                // no COM access is needed in the background task.
                var config = _configuration;
                var calendarName = _calendarName;
                var uid = _uid;
                var syncService = _syncService;

                // Unhook immediately so we don't fire twice if Outlook retries.
                Dispose();

                if (config == null || !config.IsComplete())
                    return;

                Task.Run(() =>
                {
                    try
                    {
                        bool ok = syncService.DeleteAppointment(config, calendarName, uid);
                        DiagnosticsLogger.Log(LogCategories.CalDav, "Appointment deleted from CalDAV (uid=" + uid + ", ok=" + ok + ").");
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to DELETE appointment (uid=" + uid + ").", ex);
                    }
                });
            }

            public void Dispose()
            {
                if (_disposed)
                    return;
                _disposed = true;
                if (_events != null)
                {
                    try { _events.BeforeDelete -= OnBeforeDelete; }
                    catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to unhook BeforeDelete.", ex); }
                    _events = null;
                }
                if (_onDisposed != null)
                    _onDisposed(_entryId);
            }
        }
    }
}
