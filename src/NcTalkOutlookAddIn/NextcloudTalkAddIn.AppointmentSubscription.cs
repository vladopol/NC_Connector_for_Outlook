// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        private sealed class AppointmentSubscription : IDisposable
        {
            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.AppointmentItem _appointment;
            private readonly string _key;
            private readonly string _roomToken;
            private readonly string _roomUrl;
            private readonly bool _lobbyEnabled;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly bool _isEventConversation;
            // Last start#end pair pushed to Nextcloud. Tracking the pair (rather than only the
            // start) means a changed end time also refreshes the room's event binding.
            private string _lastSyncObjectId;
            private bool _roomDeleted;
            private bool _disposed;
            private bool _unsavedCloseCleanupPending;
            private int _unsavedCloseCleanupAttempts;
            private System.Windows.Forms.Timer _unsavedCloseCleanupTimer;
            private System.Windows.Forms.Timer _deferredWriteLobbyTimer;
            private int _deferredWriteLobbyAttempts;
            private string _entryId;
            private const int UnsavedCloseCleanupMaxAttempts = 90;
            private const int DeferredWriteLobbyMaxAttempts = 4;

            internal AppointmentSubscription(
                NextcloudTalkAddIn owner,
                Outlook.AppointmentItem appointment,
                string key,
                string roomToken,
                string roomUrl,
                bool lobbyEnabled,
                bool isEventConversation,
                string entryId)
            {
                _owner = owner;
                _appointment = appointment;
                _key = key;
                _roomToken = roomToken;
                _roomUrl = roomUrl;
                _lobbyEnabled = lobbyEnabled;
                _isEventConversation = isEventConversation;
                _lastSyncObjectId = TalkAppointmentController.GetUserPropertyText(appointment, IcalObjectId);
                _entryId = entryId;
                _events = appointment as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.BeforeDelete += OnBeforeDelete;
                    _events.Write += OnWrite;
                    _events.Close += OnClose;
                }

                LogTalk("Subscription registered (token=" + _roomToken + ", lobby=" + _lobbyEnabled + ", event=" + _isEventConversation + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }

            private void OnWrite(ref bool cancel)
            {
                if (_unsavedCloseCleanupPending)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("OnWrite canceled pending unsaved cleanup (token=" + _roomToken + ").");
                }
                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("OnWrite ignored (no token).");
                    return;
                }

                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnWrite ignored (not organizer, token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("OnWrite skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                LogTalk("OnWrite for appointment (token=" + _roomToken + ").");

                bool effectiveLobbyKnown;
                bool effectiveLobbyEnabled;
                bool effectiveIsEventConversation;
                _owner._talkAppointmentController.ResolveRuntimeRoomTraits(_appointment, _roomToken, _lobbyEnabled, _isEventConversation, out effectiveLobbyKnown, out effectiveLobbyEnabled, out effectiveIsEventConversation);
                LogTalk("OnWrite traits resolved (token=" + _roomToken + ", lobbyKnown=" + effectiveLobbyKnown + ", lobby=" + effectiveLobbyEnabled + ", event=" + effectiveIsEventConversation + ").");

                // What was last pushed to Nextcloud. The in-memory value wins: the deferred pass can
                // advance it without writing the appointment back, so the stored property may lag.
                string previousObjectId = _lastSyncObjectId;
                if (string.IsNullOrWhiteSpace(previousObjectId))
                {
                    previousObjectId = TalkAppointmentController.GetUserPropertyText(_appointment, IcalObjectId);
                }

                long currentStartEpoch;
                _owner._talkAppointmentController.PersistCoreIcalProperties(
                    _appointment,
                    _roomToken,
                    _roomUrl,
                    effectiveLobbyEnabled,
                    effectiveIsEventConversation,
                    out currentStartEpoch);

                string currentObjectId = TalkAppointmentController.GetUserPropertyText(_appointment, IcalObjectId);
                bool timingChanged = !string.Equals(previousObjectId ?? string.Empty, currentObjectId ?? string.Empty, StringComparison.Ordinal);

                TalkRoomSyncSnapshot snapshot = _owner._talkAppointmentController.BuildSyncSnapshot(
                    _appointment,
                    _roomToken,
                    effectiveLobbyKnown,
                    effectiveLobbyEnabled,
                    effectiveIsEventConversation,
                    timingChanged,
                    "write");
                if (snapshot == null)
                {
                    LogTalk("OnWrite sync skipped: snapshot unavailable (token=" + _roomToken + ").");
                    _owner.TryDirectCalDavSync(_appointment);
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                // Record the attendee list while the save is still in flight, so the new baseline is
                // persisted by this very save and no follow-up Save() is needed.
                _owner._talkAppointmentController.PersistAttendeeBaseline(_appointment, snapshot.CurrentAttendeeEmails);
                _lastSyncObjectId = currentObjectId;

                string pendingDelegateId;
                bool delegationPending = _owner._talkAppointmentController.IsDelegationPending(_appointment, out pendingDelegateId);
                if (delegationPending)
                {
                    // The room must be fully in sync before ownership is handed over and this user
                    // leaves it, so the reconciliation runs inline on this rare one-off path.
                    LogTalk("OnWrite delegation-pending path (token=" + _roomToken + ", delegate=" + pendingDelegateId + ").");
                    _owner.RunRoomSyncInline(snapshot);
                    _owner._talkAppointmentController.TryApplyDelegation(_appointment, _roomToken);
                }
                else
                {
                    _owner.QueueRoomSync(snapshot);
                }

                if (timingChanged || effectiveLobbyEnabled || !effectiveLobbyKnown)
                {
                    // Outlook occasionally still reports the pre-edit time while Write is running;
                    // the deferred pass re-reads it once the save has settled.
                    ScheduleDeferredWriteLobbyVerification();
                }

                _owner.TryDirectCalDavSync(_appointment);
                _owner.RefreshEntryBinding(this);
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("BeforeDelete ignored (not organizer, token=" + _roomToken + ").");
                    return;
                }

                LogTalk("BeforeDelete -> queue saved-event room deletion (token=" + _roomToken + ").");
                QueueSavedEventRoomDeletion();
            }

            private void OnClose(ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnClose without organizer (token=" + _roomToken + ").");
                    Dispose();
                    return;
                }
                if (!_roomDeleted && _appointment != null && !_appointment.Saved)
                {
                    LogTalk("OnClose unsaved state observed (token=" + _roomToken + ", cancel=" + cancel + ").");
                    if (!cancel)
                    {
                        ScheduleUnsavedCloseCleanup();
                    }
                    return;
                }

                LogTalk("OnClose completed (token=" + _roomToken + ", deleted=" + _roomDeleted + ").");
            }

            private void ScheduleUnsavedCloseCleanup()
            {
                if (_unsavedCloseCleanupPending)
                {
                    return;
                }

                _unsavedCloseCleanupPending = true;
                _unsavedCloseCleanupAttempts = 0;
                _unsavedCloseCleanupTimer = new System.Windows.Forms.Timer();
                _unsavedCloseCleanupTimer.Interval = 1000;
                _unsavedCloseCleanupTimer.Tick += OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Start();
                LogTalk("OnClose unsaved cleanup scheduled (token=" + _roomToken + ").");
            }

            private void ScheduleDeferredWriteLobbyVerification()
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    return;
                }

                _deferredWriteLobbyAttempts = 0;
                if (_deferredWriteLobbyTimer == null)
                {
                    _deferredWriteLobbyTimer = new System.Windows.Forms.Timer();
                    _deferredWriteLobbyTimer.Interval = 750;
                    _deferredWriteLobbyTimer.Tick += OnDeferredWriteLobbyTick;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Start();
                LogTalk("Deferred post-write lobby verification scheduled (token=" + _roomToken + ").");
            }

            private void OnDeferredWriteLobbyTick(object sender, EventArgs e)
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }
                try
                {
                    OnDeferredWriteLobbyTickCore();
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    LogTalk("Deferred post-write lobby tick aborted: appointment COM object is no longer valid (token=" + _roomToken + "): " + ex.Message);
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                }
            }

            private void OnDeferredWriteLobbyTickCore()
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }

                _deferredWriteLobbyAttempts++;

                if (!_owner.IsOrganizer(_appointment))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (not organizer, token=" + _roomToken + ").");
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    return;
                }
                bool effectiveLobbyKnown;
                bool effectiveLobbyEnabled;
                bool effectiveIsEventConversation;
                _owner._talkAppointmentController.ResolveRuntimeRoomTraits(_appointment, _roomToken, _lobbyEnabled, _isEventConversation, out effectiveLobbyKnown, out effectiveLobbyEnabled, out effectiveIsEventConversation);

                // Strictly read-only: Outlook has already saved the item by the time this runs, so
                // writing user properties here would leave it dirty and trigger a save prompt.
                string currentObjectId;
                if (!_owner._talkAppointmentController.TryReadAppointmentObjectId(_appointment, _roomToken, out currentObjectId))
                {
                    if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                    {
                        _deferredWriteLobbyAttempts = 0;
                        StopDeferredWriteLobbyTimer();
                        LogTalk("Deferred post-write timing verification stopped after unavailable appointment start (token=" + _roomToken + ").");
                    }
                    return;
                }
                if (string.Equals(_lastSyncObjectId ?? string.Empty, currentObjectId ?? string.Empty, StringComparison.Ordinal))
                {
                    if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                    {
                        _deferredWriteLobbyAttempts = 0;
                        StopDeferredWriteLobbyTimer();
                        LogTalk("Deferred post-write timing verification completed without detected change (token=" + _roomToken + ", objectId=" + (currentObjectId ?? "n/a") + ").");
                    }
                    return;
                }

                LogTalk("Deferred post-write timing verification detected a change (token=" + _roomToken + ", objectId=" + (currentObjectId ?? "n/a") + ").");
                TalkRoomSyncSnapshot snapshot = _owner._talkAppointmentController.BuildSyncSnapshot(
                    _appointment,
                    _roomToken,
                    effectiveLobbyKnown,
                    effectiveLobbyEnabled,
                    effectiveIsEventConversation,
                    true,
                    "deferred_write");
                _lastSyncObjectId = currentObjectId;
                _deferredWriteLobbyAttempts = 0;
                StopDeferredWriteLobbyTimer();
                _owner.QueueRoomSync(snapshot);
            }

            private void OnUnsavedCloseCleanupTick(object sender, EventArgs e)
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    return;
                }

                _unsavedCloseCleanupAttempts++;
                bool saved;
                bool hasSavedState = TryGetAppointmentSaved(_appointment, out saved);
                bool inspectorOpen = IsAppointmentOpenInAnyInspector(_appointment);
                if (saved)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True).");
                    return;
                }
                if (inspectorOpen)
                {
                    if (_unsavedCloseCleanupAttempts >= UnsavedCloseCleanupMaxAttempts)
                    {
                        _unsavedCloseCleanupPending = false;
                        _unsavedCloseCleanupAttempts = 0;
                        StopUnsavedCloseCleanupTimer();
                        LogTalk("Deferred OnClose cleanup skipped after timeout (token=" + _roomToken + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ", inspectorOpen=True).");
                        return;
                    }
                    if (_unsavedCloseCleanupAttempts == 1 || (_unsavedCloseCleanupAttempts % 5) == 0)
                    {
                        LogTalk("Deferred OnClose cleanup waiting for inspector close (token=" + _roomToken + ", attempt=" + _unsavedCloseCleanupAttempts.ToString(CultureInfo.InvariantCulture) + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ").");
                    }
                    return;
                }

                _unsavedCloseCleanupPending = false;
                _unsavedCloseCleanupAttempts = 0;
                StopUnsavedCloseCleanupTimer();

                if (!saved)
                {
                    LogTalk("Deferred OnClose cleanup deleting room (token=" + _roomToken + ").");
                    EnsureRoomDeleted();
                    return;
                }

                LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True, hasSavedState=" + hasSavedState + ", inspectorOpen=False).");
            }

            private static bool TryGetAppointmentSaved(Outlook.AppointmentItem appointment, out bool saved)
            {
                saved = false;
                if (appointment == null)
                {
                    return false;
                }
                try
                {
                    saved = appointment.Saved;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private void StopUnsavedCloseCleanupTimer()
            {
                if (_unsavedCloseCleanupTimer == null)
                {
                    return;
                }

                _unsavedCloseCleanupTimer.Stop();
                _unsavedCloseCleanupTimer.Tick -= OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Dispose();
                _unsavedCloseCleanupTimer = null;
            }

            private void StopDeferredWriteLobbyTimer()
            {
                if (_deferredWriteLobbyTimer == null)
                {
                    return;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Tick -= OnDeferredWriteLobbyTick;
                _deferredWriteLobbyTimer.Dispose();
                _deferredWriteLobbyTimer = null;
            }

            private bool IsAppointmentOpenInAnyInspector(Outlook.AppointmentItem appointment)
            {
                if (appointment == null || _owner == null || _owner._inspectors == null)
                {
                    return false;
                }
                string appointmentEntryId = null;
                try
                {
                    appointmentEntryId = appointment.EntryID;
                }
                catch
                {
                }
                int inspectorCount = 0;
                try
                {
                    inspectorCount = _owner._inspectors.Count;
                }
                catch
                {
                    return false;
                }
                for (int i = 1; i <= inspectorCount; i++)
                {
                    Outlook.Inspector inspector = null;
                    object currentItem = null;
                    Outlook.AppointmentItem currentAppointment = null;
                    try
                    {
                        inspector = _owner._inspectors[i];
                        if (inspector == null)
                        {
                            continue;
                        }

                        currentItem = inspector.CurrentItem;
                        currentAppointment = currentItem as Outlook.AppointmentItem;
                        if (currentAppointment == null)
                        {
                            continue;
                        }
                        if (currentAppointment == appointment)
                        {
                            return true;
                        }
                        if (!string.IsNullOrWhiteSpace(appointmentEntryId))
                        {
                            string currentEntryId = null;
                            try
                            {
                                currentEntryId = currentAppointment.EntryID;
                            }
                            catch
                            {
                            }
                            if (!string.IsNullOrWhiteSpace(currentEntryId)
                                && string.Equals(currentEntryId, appointmentEntryId, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (currentAppointment != null && !ReferenceEquals(currentAppointment, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentAppointment,
                                LogCategories.Talk,
                                "Failed to release current appointment COM object during unsaved-close lookup.");
                        }
                        if (currentItem != null
                            && !ReferenceEquals(currentItem, currentAppointment)
                            && !ReferenceEquals(currentItem, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentItem,
                                LogCategories.Talk,
                                "Failed to release current item COM object during unsaved-close lookup.");
                        }
                        if (inspector != null)
                        {
                            ComInteropScope.TryRelease(
                                inspector,
                                LogCategories.Talk,
                                "Failed to release inspector COM object during unsaved-close lookup.");
                        }
                    }
                }
                return false;
            }

            private void EnsureRoomDeleted()
            {
                if (_roomDeleted || !_owner.IsOrganizer(_appointment))
                {
                    if (_roomDeleted)
                    {
                        LogTalk("EnsureRoomDeleted: room already deleted (token=" + _roomToken + ").");
                    }
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("EnsureRoomDeleted skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner._talkAppointmentController.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    Dispose();
                    return;
                }
                // The HTTP delete goes to a background thread: this runs from a timer tick on the UI
                // thread, and a stalled server would freeze Outlook. The local metadata is cleared
                // immediately — the appointment is being discarded either way.
                _roomDeleted = true;
                _owner._talkAppointmentController.ClearTalkProperties(_appointment);
                _owner.QueueDiscardedRoomDeletion(_roomToken, _isEventConversation);
                LogTalk("EnsureRoomDeleted queued (token=" + _roomToken + ").");
                Dispose();
            }

            private void QueueSavedEventRoomDeletion()
            {
                if (_roomDeleted || string.IsNullOrWhiteSpace(_roomToken))
                {
                    return;
                }

                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("Saved-event room deletion skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _roomDeleted = true;
                    Dispose();
                    return;
                }

                _roomDeleted = true;
                _owner.QueueSavedEventRoomDeletion(_roomToken, _isEventConversation);
                _owner.TryQueueCalDavDelete(_entryId);
                Dispose();
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    LogTalk("Subscription.Dispose called again (token=" + _roomToken + ").");
                    return;
                }
                if (_events != null)
                {
                    _events.BeforeDelete -= OnBeforeDelete;
                    _events.Write -= OnWrite;
                    _events.Close -= OnClose;
                }

                StopUnsavedCloseCleanupTimer();
                StopDeferredWriteLobbyTimer();

                _owner.UnregisterSubscription(_key, _roomToken, _entryId);
                _disposed = true;
                LogTalk("Subscription.Dispose completed (token=" + _roomToken + ").");
            }

            internal Outlook.AppointmentItem Appointment
            {
                get { return _appointment; }
            }

            internal string EntryId
            {
                get { return _entryId; }
            }

            internal bool IsFor(Outlook.AppointmentItem appointment)
            {
                if (appointment == null)
                {
                    return false;
                }
                if (appointment == _appointment)
                {
                    return true;
                }
                string thisEntryId = _entryId;
                if (string.IsNullOrWhiteSpace(thisEntryId))
                {
                    thisEntryId = GetEntryId(_appointment);
                }
                string otherEntryId = GetEntryId(appointment);
                if (!string.IsNullOrWhiteSpace(thisEntryId)
                    && !string.IsNullOrWhiteSpace(otherEntryId)
                    && string.Equals(thisEntryId, otherEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                return false;
            }

            internal bool MatchesToken(string token)
            {
                return string.Equals(_roomToken, token, StringComparison.OrdinalIgnoreCase);
            }

            internal void UpdateEntryId(string entryId)
            {
                _entryId = entryId;
                LogTalk("Subscription EntryId updated (token=" + _roomToken + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }
        }

    }
}
