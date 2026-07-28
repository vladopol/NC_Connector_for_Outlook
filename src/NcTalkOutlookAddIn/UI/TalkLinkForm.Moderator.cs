// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
        // Additional moderators for the Talk room.
    //
    // Candidates are the meeting's own attendees that resolve to Nextcloud accounts, not the whole
    // user directory: offering people who were never invited is noise, and the short list makes a
    // plain checked list the right control — no search, no autocomplete, no dropdown.
    //
    // These are additions. The organizer stays the room owner and stays in the room; nothing here
    // hands moderation over or leaves the conversation.
    internal sealed partial class TalkLinkForm
    {
        private void PopulateModeratorCandidates(List<NextcloudUser> candidates)
        {
            _moderatorListBox.BeginUpdate();
            try
            {
                _moderatorListBox.Items.Clear();
                if (candidates != null)
                {
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < candidates.Count; i++)
                    {
                        NextcloudUser candidate = candidates[i];
                        if (candidate == null || string.IsNullOrWhiteSpace(candidate.UserId))
                        {
                            continue;
                        }
                        if (!seen.Add(candidate.UserId.Trim()))
                        {
                            continue;
                        }

                        _moderatorListBox.Items.Add(candidate);
                    }
                }
            }
            finally
            {
                _moderatorListBox.EndUpdate();
            }

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "Moderator candidates populated from meeting attendees (count=" + _moderatorListBox.Items.Count + ").");
        }

        private List<string> CollectCheckedModeratorIds()
        {
            var result = new List<string>();
            foreach (object item in _moderatorListBox.CheckedItems)
            {
                var user = item as NextcloudUser;
                if (user == null || string.IsNullOrWhiteSpace(user.UserId))
                {
                    continue;
                }

                string id = user.UserId.Trim();
                if (!result.Contains(id))
                {
                    result.Add(id);
                }
            }
            return result;
        }

        private void ClearModeratorSelection()
        {
            for (int i = 0; i < _moderatorListBox.Items.Count; i++)
            {
                _moderatorListBox.SetItemChecked(i, false);
            }
        }

        private void UpdateModeratorHint()
        {
            _moderatorHintLabel.Visible = true;

            // The empty states have three different causes and three different remedies, so they
            // are reported separately rather than as one vague "no moderators available".
            if (_moderatorListBox.Items.Count == 0)
            {
                if (_meetingAttendeeCount == 0)
                {
                    // The usual one: the room is created before anyone has been invited.
                    _moderatorHintLabel.Text = Strings.TalkModeratorHintNoAttendees;
                    return;
                }
                if (!_systemAddressbookAvailable)
                {
                    _moderatorHintLabel.Text = Strings.TalkSystemAddressbookRequiredMessage;
                    return;
                }

                _moderatorHintLabel.Text = Strings.TalkModeratorHintNoCandidates;
                return;
            }

            _moderatorHintLabel.Text = Strings.TalkModeratorHint;
        }
    }
}
