// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Folder-level watch on the default calendar so Talk rooms follow appointment edits that never
    // open an inspector: dragging a meeting to a new slot, editing it in the peek/preview pane, or
    // changing attendees from the scheduling view. Those paths never raise AppointmentItem.Write,
    // which is the only trigger AppointmentSubscription has.
    //
    // Appointments that do have an open inspector are left to AppointmentSubscription.OnWrite —
    // that path can persist the attendee baseline as part of the in-flight save. This watcher never
    // writes to the item (a Save() here would mark the user's meeting as needing an update to be
    // resent), so it keeps its attendee baseline in memory instead. The two baselines can drift by
    // one edit, which is harmless: a repeated add answers 409 and a repeated removal answers 404,
    // and both are treated as success.
    public sealed partial class NextcloudTalkAddIn
    {
        private const int CalendarWatchStateLimit = 512;

        private Outlook.MAPIFolder _watchedCalendarFolder;
        private Outlook.Items _watchedCalendarItems;
        private readonly Dictionary<string, CalendarWatchState> _calendarWatchStates =
            new Dictionary<string, CalendarWatchState>(StringComparer.OrdinalIgnoreCase);

        private sealed class CalendarWatchState
        {
            internal string Signature;
            internal string AppliedObjectId;
            internal List<string> AttendeeBaseline;
        }

        private void EnsureTalkCalendarWatcher()
        {
            if (_outlookApplication == null || _watchedCalendarItems != null)
            {
                return;
            }
            try
            {
                _watchedCalendarFolder = _outlookApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                _watchedCalendarItems = _watchedCalendarFolder.Items;
                _watchedCalendarItems.ItemChange += OnTalkCalendarItemChange;
                LogTalk("Calendar watcher attached for Talk appointment changes.");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to attach the Talk calendar watcher.", ex);
                UnhookTalkCalendarWatcher();
            }
        }

        private void UnhookTalkCalendarWatcher()
        {
            if (_watchedCalendarItems != null)
            {
                try
                {
                    _watchedCalendarItems.ItemChange -= OnTalkCalendarItemChange;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to unhook the Talk calendar watcher.", ex);
                }

                ComInteropScope.TryFinalRelease(_watchedCalendarItems, LogCategories.Talk, "Failed to release watched calendar Items COM object.");
                _watchedCalendarItems = null;
            }
            if (_watchedCalendarFolder != null)
            {
                ComInteropScope.TryFinalRelease(_watchedCalendarFolder, LogCategories.Talk, "Failed to release watched calendar folder COM object.");
                _watchedCalendarFolder = null;
            }

            _calendarWatchStates.Clear();
        }

        private void OnTalkCalendarItemChange(object item)
        {
            var appointment = item as Outlook.AppointmentItem;
            if (appointment == null)
            {
                return;
            }
            try
            {
                OnTalkCalendarItemChangeCore(appointment);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to process a calendar item change.", ex);
            }
        }

        private void OnTalkCalendarItemChangeCore(Outlook.AppointmentItem appointment)
        {
            string roomToken = ResolveRoomTokenForAppointment(appointment);
            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }
            string entryId = GetEntryId(appointment);
            if (string.IsNullOrEmpty(entryId))
            {
                return;
            }

            // An open inspector's Write handler already owns this appointment.
            if (_subscriptionByEntryId.ContainsKey(entryId))
            {
                return;
            }
            if (!IsOrganizer(appointment))
            {
                return;
            }

            string delegateId;
            if (_talkAppointmentController.IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Calendar watcher skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return;
            }
            if (_talkAppointmentController.IsDelegationPending(appointment, out delegateId))
            {
                // Handing over moderation involves leaving the room; that only runs on the
                // inspector path where the outcome can be written back to the appointment.
                LogTalk("Calendar watcher skipped (delegation pending, token=" + roomToken + ").");
                return;
            }

            bool lobbyKnown;
            bool lobbyEnabled;
            bool isEventConversation;
            _talkAppointmentController.ResolveRuntimeRoomTraits(
                appointment,
                roomToken,
                false,
                true,
                out lobbyKnown,
                out lobbyEnabled,
                out isEventConversation);

            // Cheap gate first. Outlook raises ItemChange for plenty of edits this add-in does not
            // care about (categories, reminders, reminder dismissal, sync churn) and usually more
            // than once per edit. The probe reads only locally cached fields, so the expensive part
            // — resolving every attendee to an SMTP address, which can round-trip to Exchange — is
            // reached only when something relevant actually changed.
            string signature = BuildChangeProbe(appointment);

            CalendarWatchState state;
            bool hasState = _calendarWatchStates.TryGetValue(entryId, out state);
            if (hasState && string.Equals(state.Signature, signature, StringComparison.Ordinal))
            {
                return;
            }

            TalkRoomSyncSnapshot snapshot = _talkAppointmentController.BuildSyncSnapshot(
                appointment,
                roomToken,
                lobbyKnown,
                lobbyEnabled,
                isEventConversation,
                false,
                "item_change");
            if (snapshot == null)
            {
                return;
            }

            string objectId = BuildObjectId(snapshot);
            string previousObjectId = hasState && !string.IsNullOrEmpty(state.AppliedObjectId)
                ? state.AppliedObjectId
                : TalkAppointmentController.GetUserPropertyText(appointment, IcalObjectId);
            snapshot.TimingChanged = !string.Equals(previousObjectId ?? string.Empty, objectId ?? string.Empty, StringComparison.Ordinal);

            if (hasState && state.AttendeeBaseline != null)
            {
                // The in-memory baseline is newer than the persisted one whenever this watcher has
                // already synced the appointment during this session.
                snapshot.PreviousAttendeeEmails = state.AttendeeBaseline;
            }
            if (!snapshot.TimingChanged && !snapshot.AttendeesChanged && hasState)
            {
                // Subject/body changes still reach the room, but nothing else has to be recomputed.
                LogTalk("Calendar watcher detected a text-only change (token=" + roomToken + ").");
            }

            if (state == null)
            {
                // Deleted appointments leave their entry behind; drop everything rather than track
                // removals, since a rebuilt entry only costs one redundant (idempotent) sync.
                if (_calendarWatchStates.Count >= CalendarWatchStateLimit)
                {
                    LogTalk("Calendar watcher state cache reset after reaching its limit.");
                    _calendarWatchStates.Clear();
                }

                state = new CalendarWatchState();
                _calendarWatchStates[entryId] = state;
            }
            state.Signature = signature;
            state.AppliedObjectId = objectId;
            state.AttendeeBaseline = new List<string>(snapshot.CurrentAttendeeEmails);

            LogTalk("Calendar watcher queueing room sync (token=" + roomToken
                    + ", timingChanged=" + snapshot.TimingChanged
                    + ", attendeesChanged=" + snapshot.AttendeesChanged + ").");
            QueueRoomSync(snapshot);
        }

        private static string BuildObjectId(TalkRoomSyncSnapshot snapshot)
        {
            long? start = TimeUtilities.ToUnixTimeSeconds(snapshot.Start);
            long? end = TimeUtilities.ToUnixTimeSeconds(snapshot.End);
            if (!start.HasValue || !end.HasValue)
            {
                return null;
            }
            return start.Value.ToString(CultureInfo.InvariantCulture) + "#" + end.Value.ToString(CultureInfo.InvariantCulture);
        }

        // Fingerprint of everything this add-in mirrors to Nextcloud, built exclusively from fields
        // Outlook serves from the local store: no SMTP resolution, no address-book lookups, no COM
        // calls that can reach the Exchange server. Attendees are identified by display name here —
        // enough to notice that the list changed, which is all this gate has to decide.
        private static string BuildChangeProbe(Outlook.AppointmentItem appointment)
        {
            var builder = new StringBuilder();
            builder.Append(SafeProbeRead(() => appointment.Subject));
            builder.Append('|').Append(SafeProbeRead(() => appointment.Start.Ticks.ToString(CultureInfo.InvariantCulture)));
            builder.Append('|').Append(SafeProbeRead(() => appointment.End.Ticks.ToString(CultureInfo.InvariantCulture)));
            builder.Append('|').Append(SafeProbeRead(() => (appointment.Body ?? string.Empty).Length.ToString(CultureInfo.InvariantCulture)));

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                if (recipients != null)
                {
                    var names = new List<string>();
                    int count = recipients.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Recipient recipient = null;
                        try
                        {
                            recipient = recipients[i];
                            if (recipient == null)
                            {
                                continue;
                            }
                            names.Add(SafeProbeRead(() => recipient.Type.ToString(CultureInfo.InvariantCulture))
                                      + ":"
                                      + SafeProbeRead(() => recipient.Name));
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(recipient, LogCategories.Talk, "Failed to release Recipient COM object during change probe.");
                        }
                    }

                    names.Sort(StringComparer.OrdinalIgnoreCase);
                    builder.Append('|').Append(string.Join(";", names.ToArray()));
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to probe appointment recipients.", ex);
            }
            finally
            {
                ComInteropScope.TryRelease(recipients, LogCategories.Talk, "Failed to release Recipients COM object during change probe.");
            }

            return builder.ToString();
        }

        private static string SafeProbeRead(Func<string> read)
        {
            try
            {
                return read() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
