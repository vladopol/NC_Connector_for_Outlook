// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
        // Describes user input for creating a Talk room.
    internal sealed class TalkRoomRequest
    {
        public string Title { get; set; }

        public string Password { get; set; }

        public bool LobbyEnabled { get; set; }

        public bool SearchVisible { get; set; }

        public TalkRoomType RoomType { get; set; }

        public DateTime? AppointmentStart { get; set; }

        public DateTime? AppointmentEnd { get; set; }

        public string Description { get; set; }

        public string DescriptionLanguage { get; set; }

        public string DescriptionType { get; set; }

        public string InvitationTemplate { get; set; }

        public bool AddUsers { get; set; }

        public bool AddGuests { get; set; }

        // Nextcloud user ids promoted to moderator after the room is created. The organizer stays
        // owner and stays in the room — this is not the ownership handover the older builds did.
        public List<string> ModeratorIds { get; set; }
    }
}
