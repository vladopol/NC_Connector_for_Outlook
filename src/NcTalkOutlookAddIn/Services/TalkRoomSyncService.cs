// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Reconciles a Talk room with the state of its Outlook appointment: name, description,
    // lobby/event timing and the participant list.
    //
    // Operates purely on a TalkRoomSyncSnapshot — no COM access — so every method here is safe to
    // call from a thread-pool thread. Failures are logged and never surfaced as dialogs: this runs
    // on every appointment save and a modal warning per save would be unusable.
    internal sealed class TalkRoomSyncService
    {
        private const string ActorTypeUsers = "users";
        private const string ActorTypeEmails = "emails";

        // Event conversations derive their name and description from the linked calendar object on
        // some Talk versions and reject direct updates. The first rejection per room is remembered
        // for the rest of the session so later saves don't repeat a request that cannot succeed.
        private static readonly HashSet<string> NameSyncRejectedTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> DescriptionSyncRejectedTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly object RejectionSyncRoot = new object();

        private readonly TalkServiceConfiguration _configuration;
        private readonly IfbAddressBookCache _addressBookCache;
        private readonly int _cacheHours;
        private readonly string _selfUserId;

        internal TalkRoomSyncService(
            TalkServiceConfiguration configuration,
            IfbAddressBookCache addressBookCache,
            int cacheHours,
            string selfUserId)
        {
            _configuration = configuration;
            _addressBookCache = addressBookCache;
            _cacheHours = cacheHours;
            _selfUserId = selfUserId ?? string.Empty;
        }

        // Runs the full reconciliation. Returns false when at least one step failed, so callers can
        // decide whether the recorded attendee baseline should be kept for a later retry.
        internal bool Run(TalkRoomSyncSnapshot snapshot)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.RoomToken))
            {
                return false;
            }
            if (_configuration == null || !_configuration.IsComplete())
            {
                LogTalk("Room sync skipped: Talk configuration incomplete (token=" + snapshot.RoomToken + ").");
                return false;
            }

            string token = snapshot.RoomToken.Trim();
            LogTalk("Room sync started (token=" + token
                    + ", reason=" + (snapshot.Reason ?? "n/a")
                    + ", event=" + snapshot.IsEventConversation
                    + ", timingChanged=" + snapshot.TimingChanged
                    + ", attendeesChanged=" + snapshot.AttendeesChanged + ").");

            TalkService service;
            try
            {
                service = new TalkService(_configuration);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Room sync could not create the Talk service (token=" + token + ").", ex);
                return false;
            }

            bool ok = true;
            ok &= SyncName(service, token, snapshot);
            ok &= SyncDescription(service, token, snapshot);
            ok &= SyncTiming(service, token, snapshot);
            ok &= SyncParticipants(service, token, snapshot);

            LogTalk("Room sync finished (token=" + token + ", ok=" + ok + ").");
            return ok;
        }

        private bool SyncName(TalkService service, string token, TalkRoomSyncSnapshot snapshot)
        {
            string roomName = (snapshot.Subject ?? string.Empty).Trim();
            if (roomName.Length == 0)
            {
                return true;
            }
            if (IsRejected(NameSyncRejectedTokens, token))
            {
                return true;
            }

            try
            {
                service.UpdateRoomName(token, roomName);
                LogTalk("Room name synced (token=" + token + ", length=" + roomName.Length + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsPropertyRejected(ex, snapshot.IsEventConversation))
                {
                    Reject(NameSyncRejectedTokens, token);
                    LogTalk("Room name is managed by the linked event; name sync disabled for this session (token=" + token + ", reason=" + ex.Message + ").");
                    return true;
                }
                if (IsMissingOrForbidden(ex))
                {
                    LogTalk("Room name sync skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Room name could not be synced (token=" + token + "): " + ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while syncing room name (token=" + token + ").", ex);
            }
            return false;
        }

        private bool SyncDescription(TalkService service, string token, TalkRoomSyncSnapshot snapshot)
        {
            string description = (snapshot.Description ?? string.Empty).Trim();
            if (IsRejected(DescriptionSyncRejectedTokens, token))
            {
                return true;
            }

            try
            {
                service.UpdateDescription(token, description, snapshot.IsEventConversation);
                LogTalk("Room description synced (token=" + token + ", length=" + description.Length + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsPropertyRejected(ex, snapshot.IsEventConversation))
                {
                    Reject(DescriptionSyncRejectedTokens, token);
                    LogTalk("Room description is managed by the linked event; description sync disabled for this session (token=" + token + ", reason=" + ex.Message + ").");
                    return true;
                }
                if (IsMissingOrForbidden(ex))
                {
                    LogTalk("Room description sync skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Room description could not be synced (token=" + token + "): " + ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while syncing room description (token=" + token + ").", ex);
            }
            return false;
        }

        private bool SyncTiming(TalkService service, string token, TalkRoomSyncSnapshot snapshot)
        {
            if (!snapshot.TimingChanged)
            {
                return true;
            }
            // Lobby timers are only meaningful when the lobby is on; the event-object binding that
            // carries start/end to Nextcloud Calendar is updated for every event conversation.
            if (!snapshot.LobbyEnabled && snapshot.LobbyKnown && !snapshot.IsEventConversation)
            {
                return true;
            }

            DateTime start = snapshot.Start;
            DateTime end = snapshot.End == DateTime.MinValue ? snapshot.Start : snapshot.End;
            if (start == DateTime.MinValue)
            {
                LogTalk("Timing sync skipped: appointment start unavailable (token=" + token + ").");
                return false;
            }

            try
            {
                service.UpdateLobby(token, start.ToUniversalTime(), end, snapshot.IsEventConversation);
                LogTalk("Room timing synced (token=" + token
                        + ", start=" + start.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)
                        + ", end=" + end.ToString("o", CultureInfo.InvariantCulture) + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsMissingOrForbidden(ex))
                {
                    LogTalk("Timing sync skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Room timing could not be synced (token=" + token + "): " + ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while syncing room timing (token=" + token + ").", ex);
            }
            return false;
        }

        // Adds attendees that appeared in the Outlook invitation and removes those that disappeared
        // from it. Only addresses that this add-in previously recorded as attendees are ever
        // removed, so participants invited directly in Talk survive an Outlook-side edit.
        private bool SyncParticipants(TalkService service, string token, TalkRoomSyncSnapshot snapshot)
        {
            List<string> added = Difference(snapshot.CurrentAttendeeEmails, snapshot.PreviousAttendeeEmails);
            List<string> removed = Difference(snapshot.PreviousAttendeeEmails, snapshot.CurrentAttendeeEmails);
            if (added.Count == 0 && removed.Count == 0)
            {
                return true;
            }

            string selfEmail = null;
            if (_addressBookCache != null && !string.IsNullOrWhiteSpace(_selfUserId))
            {
                try
                {
                    _addressBookCache.TryGetPrimaryEmailForUid(_configuration, _cacheHours, _selfUserId, out selfEmail);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve the organizer's own address for participant sync.", ex);
                }
            }

            bool ok = AddParticipants(service, token, snapshot, added, selfEmail);
            ok &= RemoveParticipants(service, token, removed, selfEmail);
            return ok;
        }

        private bool AddParticipants(
            TalkService service,
            string token,
            TalkRoomSyncSnapshot snapshot,
            List<string> emails,
            string selfEmail)
        {
            if (emails.Count == 0 || (!snapshot.AddUsers && !snapshot.AddGuests))
            {
                return true;
            }

            int userAdds = 0;
            int guestAdds = 0;
            int skipped = 0;
            bool ok = true;

            for (int i = 0; i < emails.Count; i++)
            {
                string email = emails[i];
                if (IsSelf(email, selfEmail))
                {
                    skipped++;
                    continue;
                }

                string uid = ResolveUid(email);
                try
                {
                    if (!string.IsNullOrWhiteSpace(uid))
                    {
                        if (!snapshot.AddUsers)
                        {
                            skipped++;
                            continue;
                        }
                        if (service.AddUserParticipant(token, uid))
                        {
                            userAdds++;
                        }
                        else
                        {
                            ok = false;
                            LogTalk("Participant sync failed while adding Nextcloud user (uid=" + uid + ", token=" + token + ").");
                        }
                        continue;
                    }
                    if (!snapshot.AddGuests)
                    {
                        skipped++;
                        continue;
                    }
                    if (service.AddGuestParticipant(token, email))
                    {
                        guestAdds++;
                    }
                    else
                    {
                        ok = false;
                        LogTalk("Participant sync failed while adding guest (token=" + token + ").");
                    }
                }
                catch (TalkServiceException ex)
                {
                    if (IsMissingOrForbidden(ex))
                    {
                        LogTalk("Participant add skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                        return ok;
                    }

                    ok = false;
                    LogTalk("Participant could not be added (token=" + token + "): " + ex.Message);
                }
                catch (Exception ex)
                {
                    ok = false;
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while adding a participant (token=" + token + ").", ex);
                }
            }

            LogTalk("Participants added (users=" + userAdds + ", guests=" + guestAdds + ", skipped=" + skipped + ", token=" + token + ").");
            return ok;
        }

        private bool RemoveParticipants(TalkService service, string token, List<string> emails, string selfEmail)
        {
            if (emails.Count == 0)
            {
                return true;
            }

            // Build the actor identities to look for before fetching the room roster, so the
            // participant list is only requested once.
            var wantedUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var wantedEmails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < emails.Count; i++)
            {
                string email = emails[i];
                if (IsSelf(email, selfEmail))
                {
                    continue;
                }

                wantedEmails.Add(email);
                string uid = ResolveUid(email);
                if (!string.IsNullOrWhiteSpace(uid) && !IsSelfUid(uid))
                {
                    wantedUsers.Add(uid);
                }
            }
            if (wantedUsers.Count == 0 && wantedEmails.Count == 0)
            {
                return true;
            }

            List<TalkParticipant> participants;
            try
            {
                participants = service.GetParticipants(token);
            }
            catch (TalkServiceException ex)
            {
                if (IsMissingOrForbidden(ex))
                {
                    LogTalk("Participant removal skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Participant removal failed: room roster unavailable (token=" + token + "): " + ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while reading the room roster (token=" + token + ").", ex);
                return false;
            }

            int removals = 0;
            int protectedModerators = 0;
            bool ok = true;

            for (int i = 0; i < participants.Count; i++)
            {
                TalkParticipant participant = participants[i];
                if (participant == null || participant.AttendeeId <= 0)
                {
                    continue;
                }

                bool matches;
                if (string.Equals(participant.ActorType, ActorTypeUsers, StringComparison.OrdinalIgnoreCase))
                {
                    matches = wantedUsers.Contains(participant.ActorId);
                }
                else if (string.Equals(participant.ActorType, ActorTypeEmails, StringComparison.OrdinalIgnoreCase))
                {
                    matches = wantedEmails.Contains(participant.ActorId);
                }
                else
                {
                    matches = false;
                }
                if (!matches)
                {
                    continue;
                }

                // Owners, moderators and the delegated moderator stay: dropping them would leave
                // the room without anyone able to manage it.
                if (participant.IsModerator)
                {
                    protectedModerators++;
                    LogTalk("Participant removal skipped for moderator (actorType=" + participant.ActorType + ", token=" + token + ").");
                    continue;
                }

                try
                {
                    if (service.RemoveParticipant(token, participant.AttendeeId))
                    {
                        removals++;
                    }
                    else
                    {
                        ok = false;
                    }
                }
                catch (TalkServiceException ex)
                {
                    if (IsMissingOrForbidden(ex))
                    {
                        LogTalk("Participant removal skipped after access loss (token=" + token + ", status=" + (int)ex.StatusCode + ").");
                        return ok;
                    }

                    ok = false;
                    LogTalk("Participant could not be removed (token=" + token + "): " + ex.Message);
                }
                catch (Exception ex)
                {
                    ok = false;
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Unexpected error while removing a participant (token=" + token + ").", ex);
                }
            }

            LogTalk("Participants removed (count=" + removals
                    + ", moderatorsKept=" + protectedModerators
                    + ", candidates=" + emails.Count
                    + ", token=" + token + ").");
            return ok;
        }

        private string ResolveUid(string email)
        {
            if (_addressBookCache == null || string.IsNullOrWhiteSpace(email))
            {
                return null;
            }
            try
            {
                string uid;
                if (_addressBookCache.TryGetUid(_configuration, _cacheHours, email, out uid))
                {
                    return uid;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve a Nextcloud user id for participant sync.", ex);
            }
            return null;
        }

        private bool IsSelfUid(string uid)
        {
            return !string.IsNullOrWhiteSpace(_selfUserId)
                   && string.Equals(uid, _selfUserId.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSelf(string email, string selfEmail)
        {
            return !string.IsNullOrWhiteSpace(selfEmail)
                   && string.Equals(email, selfEmail, StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> Difference(List<string> source, List<string> exclude)
        {
            var result = new List<string>();
            if (source == null)
            {
                return result;
            }

            var excludeSet = new HashSet<string>(
                exclude ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < source.Count; i++)
            {
                string value = source[i];
                if (string.IsNullOrWhiteSpace(value) || excludeSet.Contains(value))
                {
                    continue;
                }
                result.Add(value);
            }
            return result;
        }

        private static bool IsRejected(HashSet<string> set, string token)
        {
            lock (RejectionSyncRoot)
            {
                return set.Contains(token);
            }
        }

        private static void Reject(HashSet<string> set, string token)
        {
            lock (RejectionSyncRoot)
            {
                set.Add(token);
            }
        }

        // Talk answers with a BadRequest carrying "event" when a property is owned by the linked
        // calendar object rather than by the room itself. For event conversations any BadRequest is
        // treated the same way: it will not start succeeding later in the session, and retrying it
        // on every appointment save would be pure noise.
        private static bool IsPropertyRejected(TalkServiceException ex, bool isEventConversation)
        {
            if (ex == null || ex.StatusCode != HttpStatusCode.BadRequest)
            {
                return false;
            }
            if (isEventConversation)
            {
                return true;
            }
            return !string.IsNullOrWhiteSpace(ex.Message)
                   && ex.Message.IndexOf("event", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsMissingOrForbidden(TalkServiceException ex)
        {
            return ex != null
                   && (ex.StatusCode == HttpStatusCode.NotFound || ex.StatusCode == HttpStatusCode.Forbidden);
        }

        private static void LogTalk(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Talk, message);
        }
    }
}
