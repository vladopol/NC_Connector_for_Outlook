// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
        // Simple representation of a Nextcloud user from the system address book.
    // Contains the UID (actorId) and a primary email address.
    internal sealed class NextcloudUser
    {
        internal NextcloudUser(string userId, string email)
            : this(userId, email, null)
        {
        }

        internal NextcloudUser(string userId, string email, string displayName)
        {
            UserId = userId ?? string.Empty;
            Email = email ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
        }

        internal string UserId { get; private set; }

        internal string Email { get; private set; }

        // Human-readable name, when the caller knows one — for meeting attendees this is Outlook's
        // resolved recipient name ("Мамина Алина"). A login is what the account is called, not what
        // the person is called, so it is only shown when nothing better is available.
        internal string DisplayName { get; private set; }

        internal string DisplayLabel
        {
            get
            {
                string name = !string.IsNullOrWhiteSpace(DisplayName) ? DisplayName.Trim() : UserId;
                if (string.IsNullOrEmpty(Email))
                {
                    return name;
                }
                return name + " <" + Email + ">";
            }
        }

        public override string ToString()
        {
            return DisplayLabel;
        }
    }
}
