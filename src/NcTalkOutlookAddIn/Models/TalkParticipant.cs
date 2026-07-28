// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
        // Minimal Talk participant (actorType/actorId + attendeeId) from the Talk API.
    internal sealed class TalkParticipant
    {
        // Talk participantType values (see spreed/lib/Participant.php).
        internal const int ParticipantTypeOwner = 1;
        internal const int ParticipantTypeModerator = 2;
        internal const int ParticipantTypeGuestModerator = 6;

        internal TalkParticipant(string actorType, string actorId, int attendeeId)
            : this(actorType, actorId, attendeeId, 0)
        {
        }

        internal TalkParticipant(string actorType, string actorId, int attendeeId, int participantType)
        {
            ActorType = actorType ?? string.Empty;
            ActorId = actorId ?? string.Empty;
            AttendeeId = attendeeId;
            ParticipantType = participantType;
        }

        internal string ActorType { get; private set; }

        internal string ActorId { get; private set; }

        internal int AttendeeId { get; private set; }

        internal int ParticipantType { get; private set; }

        // Owners and moderators are never removed by attendee-list synchronization: they are
        // either the organizer themselves or a delegated moderator, and losing them would
        // leave the room unmanageable.
        internal bool IsModerator
        {
            get
            {
                return ParticipantType == ParticipantTypeOwner
                       || ParticipantType == ParticipantTypeModerator
                       || ParticipantType == ParticipantTypeGuestModerator;
            }
        }
    }
}
