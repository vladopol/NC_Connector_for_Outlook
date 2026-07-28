// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
    // COM-free description of an Outlook appointment's Talk-relevant state.
    //
    // Built on the UI thread from an AppointmentItem (memory-only property reads), then handed to
    // TalkRoomSyncService on a background thread. Nothing in here touches Outlook, so the sync can
    // run without blocking Outlook's message loop.
    internal sealed class TalkRoomSyncSnapshot
    {
        internal TalkRoomSyncSnapshot()
        {
            CurrentAttendeeEmails = new List<string>();
            PreviousAttendeeEmails = new List<string>();
        }

        internal string RoomToken { get; set; }

        internal string EntryId { get; set; }

        // Free-text origin of the sync ("write", "item_change", …). Log noise only.
        internal string Reason { get; set; }

        internal string Subject { get; set; }

        internal string Description { get; set; }

        internal DateTime Start { get; set; }

        internal DateTime End { get; set; }

        internal bool IsEventConversation { get; set; }

        internal bool LobbyEnabled { get; set; }

        internal bool LobbyKnown { get; set; }

        // Set when the appointment's start/end changed since the last successful sync.
        internal bool TimingChanged { get; set; }

        internal bool AddUsers { get; set; }

        internal bool AddGuests { get; set; }

        // SMTP addresses of the appointment's current attendees (organizer excluded by the caller).
        internal List<string> CurrentAttendeeEmails { get; set; }

        // SMTP addresses recorded by the previous sync (X-NCTALK-ATTENDEES). Anything that was in
        // this list but is no longer an attendee is a removal candidate — participants added
        // directly in Talk never appear here and are therefore never touched.
        internal List<string> PreviousAttendeeEmails { get; set; }

        internal bool AttendeesChanged
        {
            get
            {
                return !EmailListsEqual(CurrentAttendeeEmails, PreviousAttendeeEmails);
            }
        }

        internal static bool EmailListsEqual(List<string> left, List<string> right)
        {
            int leftCount = left != null ? left.Count : 0;
            int rightCount = right != null ? right.Count : 0;
            if (leftCount != rightCount)
            {
                return false;
            }
            if (leftCount == 0)
            {
                return true;
            }

            var rightSet = new HashSet<string>(right, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < left.Count; i++)
            {
                if (!rightSet.Contains(left[i]))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
