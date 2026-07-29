// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
        // Handles Talk ribbon interactions including authentication gate, wizard orchestration,
    // room replacement flow, and appointment persistence.
    internal sealed class TalkRibbonController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal TalkRibbonController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal async Task OnTalkButtonPressedAsync(IRibbonControl control)
        {            if (_owner == null)
            {
                return;
            }

            AddinSettings settings = _owner.CurrentSettings;
            if (!_owner.SettingsAreComplete())
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: settings are incomplete.");
                MessageBox.Show(
                    Strings.ErrorMissingCredentials,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                _owner.OnSettingsButtonPressed(control);
                return;
            }
            if (!await EnsureAuthenticationValidAsync(control))
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: authentication failed.");
                return;
            }

            Outlook.AppointmentItem appointment = _owner.GetActiveAppointment();            if (appointment == null)
            {
                NextcloudTalkAddIn.LogTalkMessage("Talk link cancelled: no active appointment found.");
                MessageBox.Show(
                    Strings.ErrorNoAppointment,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            var subject = appointment.Subject ?? string.Empty;
            var start = appointment.Start == DateTime.MinValue ? DateTime.Now : appointment.Start;
            var end = appointment.End == DateTime.MinValue ? start.AddHours(1) : appointment.End;
            NextcloudTalkAddIn.LogTalkMessage("Talk link started (subject='" + subject + "', start=" + start.ToString("o") + ", end=" + end.ToString("o") + ").");

            settings = _owner.CurrentSettings ?? new AddinSettings();
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => _owner.FetchBackendPolicyStatus(configuration, "talk_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => _owner.FetchPasswordPolicyForTalkWizard(configuration));
            await Task.WhenAll(policyStatusTask, passwordPolicyTask);
            BackendPolicyStatus policyStatus = policyStatusTask.Result;
            PasswordPolicyInfo passwordPolicy = passwordPolicyTask.Result;

            var addressbookCache = new IfbAddressBookCache(_owner.SettingsStorage != null ? _owner.SettingsStorage.DataDirectory : null);
            NextcloudTalkAddIn.LogTalkMessage("System address book status check requested (context=talk_click, forceRefresh=True).");
            var talkClickAddressbookStatus = await Task.Run(() => addressbookCache.GetSystemAddressbookStatus(
                configuration,
                settings.IfbCacheHours,
                true));
            NextcloudTalkAddIn.LogTalkMessage(
                "System address book status result (context=talk_click, available=" + talkClickAddressbookStatus.Available
                + ", count=" + talkClickAddressbookStatus.Count
                + ", hasError=" + (!string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error)) + ").");
            if (!talkClickAddressbookStatus.Available && !string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error))
            {
                NextcloudTalkAddIn.LogTalkMessage("System address book unavailable on talk click: " + talkClickAddressbookStatus.Error);
            }

            // Moderator candidates are the meeting's own attendees, not the whole directory. The
            // recipient list is read here on the UI thread (COM), then mapped to Nextcloud accounts
            // on a background thread because the lookup can refresh the address-book cache.
            List<string> attendeeEmails = NextcloudTalkAddIn.GetAppointmentAttendeeEmails(appointment);
            Dictionary<string, string> attendeeNames =
                OutlookRecipientResolverController.CollectAppointmentAttendeeNamesByEmail(appointment);
            int cacheHours = settings.IfbCacheHours;
            List<NextcloudUser> moderatorCandidates;
            try
            {
                moderatorCandidates = talkClickAddressbookStatus.Available
                    ? await Task.Run(() => ResolveModeratorCandidates(addressbookCache, configuration, cacheHours, attendeeEmails, attendeeNames, settings.Username))
                    : new List<NextcloudUser>();
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Moderator candidates could not be resolved: " + ex.Message);
                moderatorCandidates = new List<NextcloudUser>();
            }
            NextcloudTalkAddIn.LogTalkMessage(
                "Moderator candidates resolved (attendees=" + attendeeEmails.Count
                + ", withNextcloudAccount=" + moderatorCandidates.Count + ").");

            using (var dialog = new TalkLinkForm(
                settings,
                configuration,
                passwordPolicy,
                policyStatus,
                moderatorCandidates,
                attendeeEmails.Count,
                talkClickAddressbookStatus,
                subject,
                start,
                end))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Talk link dialog cancelled.");
                    return;
                }
                string descriptionLanguage = NextcloudTalkAddIn.ResolveTalkDescriptionLanguage(
                    policyStatus,
                    settings.EventDescriptionLang);
                string descriptionType = NextcloudTalkAddIn.ResolveTalkEventDescriptionType(policyStatus);
                string invitationTemplate = NextcloudTalkAddIn.ResolveTalkInvitationTemplate(policyStatus);
                string appointmentBody = string.Empty;
                try { appointmentBody = appointment.Body ?? string.Empty; }
                catch (Exception ex) { NextcloudTalkAddIn.LogTalkMessage("Failed to read appointment body for room description: " + ex.Message); }
                string initialDescription = TalkDescriptionTemplateController.BuildInitialRoomDescription(appointmentBody);
                var request = new TalkRoomRequest
                {
                    Title = dialog.TalkTitle,
                    Password = dialog.TalkPassword,
                    LobbyEnabled = dialog.LobbyUntilStart,
                    SearchVisible = dialog.SearchVisible,
                    RoomType = dialog.SelectedRoomType,
                    AppointmentStart = start,
                    AppointmentEnd = end,
                    DescriptionLanguage = descriptionLanguage,
                    DescriptionType = descriptionType,
                    InvitationTemplate = invitationTemplate,
                    Description = initialDescription,
                    AddUsers = dialog.AddUsers,
                    AddGuests = dialog.AddGuests,
                    ModeratorIds = dialog.ModeratorIds
                };
                NextcloudTalkAddIn.LogTalkMessage("Room request prepared (title='" + request.Title + "', type=" + request.RoomType + ", lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ", passwordSet=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = TalkAppointmentController.GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalToken);
                if (!string.IsNullOrWhiteSpace(existingToken))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Existing room found (token=" + existingToken + "), replacement requested.");
                    var overwrite = MessageBox.Show(
                        Strings.ConfirmReplaceRoom,
                        Strings.DialogTitle,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (overwrite != DialogResult.Yes)
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Replacement declined, operation ended.");
                        return;
                    }
                    var existingType = TalkAppointmentController.GetRoomType(appointment);
                    bool existingIsEvent = existingType.HasValue && existingType.Value == TalkRoomType.EventConversation;
                    NextcloudTalkAddIn.LogTalkMessage("Attempting to delete existing room (event=" + existingIsEvent + ").");

                    bool existingDeleted;
                    using (new WaitCursorScope())
                    {
                        existingDeleted = await Task.Run(() => _owner.TryDeleteRoom(existingToken, existingIsEvent));
                    }
                    if (!existingDeleted)
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Deleting existing room failed.");
                        return;
                    }
                }

                TalkRoomCreationResult result;
                try
                {
                    using (new WaitCursorScope())
                    {
                        NextcloudTalkAddIn.LogTalkMessage("Sending CreateRoom request to Nextcloud.");
                        var service = _owner.CreateTalkService();
                        // CreateRoom issues several requests in sequence; none of them touch COM,
                        // so the whole chain belongs off the UI thread.
                        result = await Task.Run(() => service.CreateRoom(request));
                    }
                    NextcloudTalkAddIn.LogTalkMessage("Room created successfully (token=" + result.RoomToken + ", URL=" + result.RoomUrl + ", event=" + result.CreatedAsEventConversation + ").");
                }
                catch (TalkServiceException ex)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Talk room could not be created: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoom, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Unexpected error while creating talk room: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                _owner.ApplyRoomToAppointment(appointment, request, result);
                NextcloudTalkAddIn.LogTalkMessage("Room data stored in appointment (EntryID=" + (appointment.EntryID ?? "n/a") + ").");

                MessageBox.Show(
                    string.Format(Strings.InfoRoomCreated, request.Title),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        // The connectivity probe runs on a thread-pool thread: it is the first thing a Talk click
        // does, and on a server that accepts the connection but never answers it would otherwise
        // block Outlook's message loop for the whole timeout.
        private async Task<bool> EnsureAuthenticationValidAsync(IRibbonControl control)
        {
            try
            {
                VerifyConnectionOutcome outcome;
                using (new WaitCursorScope())
                {
                    var service = _owner.CreateTalkService();
                    NextcloudTalkAddIn.LogTalkMessage("Starting credential verification request.");
                    outcome = await Task.Run(() =>
                    {
                        string probeResponse;
                        bool ok = service.VerifyConnection(out probeResponse);
                        return new VerifyConnectionOutcome(ok, probeResponse);
                    });
                }

                if (outcome.Succeeded)
                {
                    _owner.UpdateStoredServerVersion(outcome.Response);
                    NextcloudTalkAddIn.LogTalkMessage("Credentials verified (response=" + (string.IsNullOrEmpty(outcome.Response) ? "OK" : outcome.Response) + ").");
                    return true;
                }
                string message = string.IsNullOrEmpty(outcome.Response)
                    ? Strings.ErrorCredentialsNotVerified
                    : string.Format(CultureInfo.CurrentCulture, Strings.ErrorCredentialsNotVerifiedFormat, outcome.Response);
                NextcloudTalkAddIn.LogTalkMessage("Invalid credentials: " + message);
                return PromptOpenSettings(message, control);
            }
            catch (TalkServiceException ex)
            {
                string message;
                if ((int)ex.StatusCode == 0)
                {
                    message = Strings.ErrorServerUnavailable;
                }
                else if (ex.IsAuthenticationError)
                {
                    message = string.Format(Strings.ErrorAuthenticationRejected, ex.Message);
                }
                else
                {
                    message = string.Format(Strings.ErrorConnectionFailed, ex.Message);
                }

                NextcloudTalkAddIn.LogTalkMessage("Connection check failed: " + message);
                return PromptOpenSettings(message, control);
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error during connection check: " + ex.Message);
                return PromptOpenSettings(string.Format(Strings.ErrorUnknownAuthentication, ex.Message), control);
            }
        }

        // VerifyConnection reports through an out parameter, which cannot cross a Task boundary.
        private sealed class VerifyConnectionOutcome
        {
            internal VerifyConnectionOutcome(bool succeeded, string response)
            {
                Succeeded = succeeded;
                Response = response;
            }

            internal bool Succeeded { get; private set; }

            internal string Response { get; private set; }
        }

        // Maps the meeting's attendee addresses onto Nextcloud accounts. Attendees without an
        // account simply drop out — they can join the room as guests but cannot moderate it.
        private static List<NextcloudUser> ResolveModeratorCandidates(
            IfbAddressBookCache addressBookCache,
            TalkServiceConfiguration configuration,
            int cacheHours,
            List<string> attendeeEmails,
            Dictionary<string, string> attendeeNames,
            string selfUserId)
        {
            var result = new List<NextcloudUser>();
            if (addressBookCache == null || attendeeEmails == null)
            {
                return result;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < attendeeEmails.Count; i++)
            {
                string email = attendeeEmails[i];
                if (string.IsNullOrWhiteSpace(email))
                {
                    continue;
                }
                try
                {
                    string uid;
                    if (!addressBookCache.TryGetUid(configuration, cacheHours, email, out uid)
                        || string.IsNullOrWhiteSpace(uid))
                    {
                        continue;
                    }
                    // The organizer creates the room and is its owner, so they are already a
                    // moderator — the hint says so. Offering them as something to tick would imply
                    // the opposite.
                    if (!string.IsNullOrWhiteSpace(selfUserId)
                        && string.Equals(uid.Trim(), selfUserId.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (!seen.Add(uid.Trim()))
                    {
                        continue;
                    }

                    string displayName;
                    if (attendeeNames == null || !attendeeNames.TryGetValue(email.Trim().ToLowerInvariant(), out displayName))
                    {
                        displayName = null;
                    }
                    result.Add(new NextcloudUser(uid.Trim(), email.Trim(), displayName));
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to map an attendee to a Nextcloud account.", ex);
                }
            }
            return result;
        }

        private bool PromptOpenSettings(string message, IRibbonControl control)
        {
            var result = MessageBox.Show(
                string.Format(Strings.PromptOpenSettings, message),
                Strings.DialogTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                _owner.OnSettingsButtonPressed(control);
            }
            return false;
        }
    }
}

