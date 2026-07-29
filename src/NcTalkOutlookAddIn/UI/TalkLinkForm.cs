// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{

    /// <summary>
    /// Dialog for configuring and creating a Nextcloud Talk room for the active appointment.
    /// </summary>
    internal sealed partial class TalkLinkForm : ScaledForm
    {
        private static readonly string DefaultTitle = Strings.TalkDefaultTitle;
        private const int DefaultMinPasswordLength = 5;
        private const int ModeratorDropdownMaxRows = 6;
        private const int ModeratorDropdownMargin = 4;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

        private readonly Label _roomTypeLabel = new Label();
        private readonly GroupBox _settingsGroup = new GroupBox();
        private readonly Label _eventSupportHintLabel = new Label();
        private readonly ComboBox _roomTypeComboBox = new ComboBox();
        // A room password is the exception rather than the rule, and it has no meaningful "off"
        // state worth displaying — so it is offered as an action rather than a checkbox that sits
        // unticked most of the time. One button in a fixed position handles both directions, with
        // only its caption changing: "Add password" / "Remove password".
        private readonly Button _passwordToggleButton = new Button();
        private bool _passwordEnabled;
        private readonly TextBox _passwordTextBox = new TextBox();
        private readonly Button _passwordGenerateButton = new Button();
        private readonly CheckBox _addUsersCheckBox = new CheckBox();
        private readonly CheckBox _addGuestsCheckBox = new CheckBox();
        private readonly CheckBox _lobbyCheckBox = new CheckBox();
        private readonly CheckBox _searchCheckBox = new CheckBox();
        private readonly GroupBox _moderatorGroup = new GroupBox();
        // Candidates are the meeting's own attendees that map to Nextcloud accounts — not the whole
        // user directory. That keeps the list short enough for a plain checked list and stops the
        // dialog offering people who were never invited.
        private readonly CheckedListBox _moderatorListBox = new CheckedListBox();
        private readonly Panel _moderatorAddressbookWarningPanel = new Panel();
        private readonly Label _moderatorAddressbookWarningTitleLabel = new Label();
        private readonly Label _moderatorAddressbookWarningTextLabel = new Label();
        private readonly LinkLabel _moderatorAddressbookWarningLinkLabel = new LinkLabel();
        private readonly Panel _policyWarningPanel = new Panel();
        private readonly Label _policyWarningTitleLabel = new Label();
        private readonly Label _policyWarningTextLabel = new Label();
        private readonly LinkLabel _policyWarningLinkLabel = new LinkLabel();
        private readonly Label _moderatorHintLabel = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly DisabledControlTooltipHintHelper _disabledTooltipHints;
        private readonly Button _okButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private const int HeaderHeight = 48;
        private readonly bool _eventConversationsSupported;
        private readonly string _serverVersionHint;
        private readonly PasswordPolicyInfo _passwordPolicy;
        private readonly TalkServiceConfiguration _configuration;
        private readonly BackendPolicyStatus _backendPolicyStatus;
        // Meeting attendees that map to Nextcloud accounts — the moderator candidates.
        private readonly List<NextcloudUser> _moderatorCandidates;
        private readonly bool _systemAddressbookAvailable;
        // How many people the meeting actually invites, regardless of Nextcloud accounts. Lets the
        // moderator list explain an empty state properly: "invite someone first" is a different
        // problem from "none of the invitees have an account".
        private readonly int _meetingAttendeeCount;
        private readonly string _systemAddressbookError;
        private readonly Dictionary<string, Image> _avatarCache = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _avatarLoading = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _avatarLock = new object();
        private bool _layoutApplying;

        private string _talkTitle;
        private string _talkPassword;
        private bool _lobbyUntilStart;
        private bool _searchVisible;
        private TalkRoomType _selectedRoomType;
        private bool _addUsers;
        private bool _addGuests;
        private List<string> _moderatorIds = new List<string>();

        internal string TalkTitle
        {
            get { return _talkTitle; }
            private set { _talkTitle = value; }
        }

        internal string TalkPassword
        {
            get { return _talkPassword; }
            private set { _talkPassword = value; }
        }

        internal bool LobbyUntilStart
        {
            get { return _lobbyUntilStart; }
            private set { _lobbyUntilStart = value; }
        }

        internal bool SearchVisible
        {
            get { return _searchVisible; }
            private set { _searchVisible = value; }
        }

        internal TalkRoomType SelectedRoomType
        {
            get { return _selectedRoomType; }
            private set { _selectedRoomType = value; }
        }

        internal bool AddUsers
        {
            get { return _addUsers; }
            private set { _addUsers = value; }
        }

        internal bool AddGuests
        {
            get { return _addGuests; }
            private set { _addGuests = value; }
        }

        // Nextcloud user ids ticked in the dialog. The organizer is the room owner regardless.
        internal List<string> ModeratorIds
        {
            get { return _moderatorIds; }
            private set { _moderatorIds = value ?? new List<string>(); }
        }

        internal TalkLinkForm(
            AddinSettings defaults,
            TalkServiceConfiguration configuration,
            PasswordPolicyInfo passwordPolicy,
            BackendPolicyStatus policyStatus,
            List<NextcloudUser> moderatorCandidates,
            int meetingAttendeeCount,
            IfbAddressBookCache.SystemAddressbookStatus addressbookStatus,
            string appointmentSubject,
            DateTime startTime,
            DateTime endTime)
        {
            _eventConversationsSupported = DetermineEventConversationSupport(defaults, out _serverVersionHint);
            _passwordPolicy = passwordPolicy;
            _configuration = configuration;
            _backendPolicyStatus = policyStatus;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            _moderatorCandidates = moderatorCandidates ?? new List<NextcloudUser>();
            _meetingAttendeeCount = meetingAttendeeCount;
            _systemAddressbookAvailable = addressbookStatus != null && addressbookStatus.Available;
            _systemAddressbookError = addressbookStatus != null ? (addressbookStatus.Error ?? string.Empty) : string.Empty;

            Text = Strings.TalkFormTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ControlBox = true;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 520);
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimumSize = new Size(ScaleLogical(580), ScaleLogical(580));
            Icon = BrandingAssets.GetAppIcon(32);

            BrandedHeader.AttachToParent(_headerPanel, Controls, HeaderHeight);
            InitializeComponents();
            ApplyDefaults(defaults, appointmentSubject);
            ApplyDialogLayout(true);

            UiThemeManager.ApplyToForm(this, _toolTip);
        }

                // Builds all dialog controls (title, password, options, buttons).
        private void InitializeComponents()
        {
            // Room type selector hidden — EventConversation is always used.

            _passwordToggleButton.AutoSize = false;
            _passwordToggleButton.TextAlign = ContentAlignment.MiddleCenter;
            _passwordToggleButton.Click += (s, e) => SetPasswordEnabled(!_passwordEnabled);


            _passwordTextBox.UseSystemPasswordChar = false;

            _passwordGenerateButton.Text = Strings.TalkPasswordGenerate;
            _passwordGenerateButton.AutoSize = false;
            _passwordGenerateButton.TextAlign = ContentAlignment.MiddleCenter;
            _passwordGenerateButton.Click += (s, e) => GeneratePassword();

            _settingsGroup.Text = Strings.TalkSettingsGroup;

            InitializeSettingsOptionCheckBox(_addUsersCheckBox, Strings.TalkAddUsersCheck);
            // AddGuests hidden — external users connect via room link and password.
            InitializeSettingsOptionCheckBox(_lobbyCheckBox, Strings.TalkLobbyCheck);
            InitializeSettingsOptionCheckBox(_searchCheckBox, Strings.TalkSearchCheck);

            _moderatorGroup.Text = Strings.TalkModeratorGroup;

            _moderatorAddressbookWarningPanel.Visible = false;
            _moderatorAddressbookWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _moderatorAddressbookWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _moderatorAddressbookWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };
            _moderatorGroup.Controls.Add(_moderatorAddressbookWarningPanel);

            _moderatorAddressbookWarningTitleLabel.AutoSize = true;
            _moderatorAddressbookWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _moderatorAddressbookWarningTitleLabel.Font = new Font(
                _moderatorAddressbookWarningTitleLabel.Font,
                FontStyle.Bold);
            _moderatorAddressbookWarningTitleLabel.Text = "\u26a0 " + Strings.TalkSystemAddressbookRequiredShort;
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningTitleLabel);

            _moderatorAddressbookWarningTextLabel.AutoSize = true;
            _moderatorAddressbookWarningTextLabel.Text = Strings.TalkSystemAddressbookRequiredMessage;
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningTextLabel);

            _moderatorAddressbookWarningLinkLabel.AutoSize = true;
            _moderatorAddressbookWarningLinkLabel.Text = Strings.TalkSystemAddressbookAdminLinkLabel;
            _moderatorAddressbookWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _moderatorAddressbookWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _moderatorAddressbookWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _moderatorAddressbookWarningLinkLabel.LinkClicked += (s, e) => OpenSystemAddressbookSetupGuide();
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningLinkLabel);

            _moderatorListBox.CheckOnClick = true;
            _moderatorListBox.IntegralHeight = false;
            _moderatorListBox.BorderStyle = BorderStyle.FixedSingle;
            _moderatorGroup.Controls.Add(_moderatorListBox);

            _moderatorHintLabel.Text = Strings.TalkModeratorHint;
            _moderatorHintLabel.ForeColor = Color.DimGray;
            // Labels paint their background; transparent keeps a stray overlap from hiding the list
            // behind it rather than showing both.
            _moderatorHintLabel.BackColor = Color.Transparent;
            _moderatorGroup.Controls.Add(_moderatorHintLabel);

            // Filled here rather than after the constructor finishes: the first layout pass runs at
            // the end of this method, and a pass that sees an empty list positions the hint as
            // though no list existed. Later passes fixed it, which is why the group only lined up
            // once some other control forced a relayout.
            PopulateModeratorCandidates(_moderatorCandidates);
            // Settle the hint text too: it wraps, so its height feeds the group's height and the
            // first pass should already measure the final wording.
            UpdateModeratorHint();

            _eventSupportHintLabel.AutoSize = true;
            _eventSupportHintLabel.MaximumSize = new Size(ScaleLogical(260), 0);
            _eventSupportHintLabel.ForeColor = Color.DimGray;
            _eventSupportHintLabel.Visible = !_eventConversationsSupported;
            if (_eventSupportHintLabel.Visible)
            {
                var versionInfo = string.IsNullOrEmpty(_serverVersionHint) ? Strings.TalkVersionUnknown : _serverVersionHint;
                _eventSupportHintLabel.Text = string.Format(Strings.TalkEventHint, versionInfo);
            }

            _policyWarningPanel.Visible = false;
            _policyWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _policyWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _policyWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };

            _policyWarningTitleLabel.AutoSize = true;
            _policyWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _policyWarningTitleLabel.Font = new Font(
                _policyWarningTitleLabel.Font,
                FontStyle.Bold);
            _policyWarningTitleLabel.Text = "\u26a0 " + Strings.PolicyWarningTitle;
            _policyWarningPanel.Controls.Add(_policyWarningTitleLabel);

            _policyWarningTextLabel.AutoSize = true;
            _policyWarningTextLabel.Text = string.Empty;
            _policyWarningPanel.Controls.Add(_policyWarningTextLabel);

            _policyWarningLinkLabel.AutoSize = true;
            _policyWarningLinkLabel.Text = Strings.PolicyWarningAdminLinkLabel;
            _policyWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _policyWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.LinkClicked += (s, e) =>
                BrowserLauncher.OpenUrl(
                    Strings.PolicyAdminGuideUrl,
                    LogCategories.Talk,
                    "Failed to open policy admin guide URL.");
            _policyWarningPanel.Controls.Add(_policyWarningLinkLabel);

            _okButton.Text = Strings.DialogOk;
            _okButton.AutoSize = false;
            _okButton.DialogResult = DialogResult.OK;
            _okButton.Click += OnOkButtonClick;

            _cancelButton.Text = Strings.DialogCancel;
            _cancelButton.AutoSize = false;
            _cancelButton.DialogResult = DialogResult.Cancel;

            Controls.Add(_eventSupportHintLabel);
            Controls.Add(_passwordToggleButton);
            Controls.Add(_passwordTextBox);
            Controls.Add(_passwordGenerateButton);
            Controls.Add(_settingsGroup);
            Controls.Add(_moderatorGroup);
            Controls.Add(_policyWarningPanel);
            Controls.Add(_okButton);
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;

            _toolTip.AutoPopDelay = 20000;
            _toolTip.InitialDelay = 250;
            _toolTip.ReshowDelay = 150;
            _toolTip.SetToolTip(_addUsersCheckBox, Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_lobbyCheckBox, Strings.TooltipLobby);
            _toolTip.SetToolTip(_searchCheckBox, Strings.TooltipSearchVisible);
            _toolTip.SetToolTip(_moderatorListBox, Strings.TooltipModerator);
            ApplyDialogLayout(false);
        }

        private void InitializeSettingsOptionCheckBox(CheckBox checkBox, string text)
        {            if (checkBox == null)
            {
                return;
            }

            checkBox.Text = text ?? string.Empty;
            checkBox.AutoSize = true;
            _settingsGroup.Controls.Add(checkBox);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (_layoutApplying)
            {
                return;
            }

            ApplyDialogLayout(false);
        }

        private void ApplyDialogLayout(bool ensureClientHeight)
        {
            if (_layoutApplying || IsDisposed || Disposing)
            {
                return;
            }

            _layoutApplying = true;
            try
            {
                int outerPadding = ScaleLogical(18);
                int labelX = ScaleLogical(16);
                int inputX = ScaleLogical(170);
                int verticalGap = ScaleLogical(12);
                int rowGap = ScaleLogical(14);

                int y = _headerPanel.Bottom + ScaleLogical(16);
                int inputWidth = Math.Max(ScaleLogical(180), ClientSize.Width - inputX - ScaleLogical(18));


                _eventSupportHintLabel.Visible = !_eventConversationsSupported;
                if (_eventSupportHintLabel.Visible)
                {
                    _eventSupportHintLabel.Location = new Point(inputX, y - ScaleLogical(2));
                    _eventSupportHintLabel.MaximumSize = new Size(inputWidth, 0);
                    y = _eventSupportHintLabel.Bottom + verticalGap;
                }
                if (_policyWarningPanel.Visible)
                {
                    int warningPadding = ScaleLogical(8);
                    int panelWidth = Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2));
                    int warningTextWidth = Math.Max(ScaleLogical(160), panelWidth - (warningPadding * 2));

                    _policyWarningTitleLabel.Location = new Point(warningPadding, warningPadding);
                    _policyWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningTextTop = _policyWarningTitleLabel.Bottom + ScaleLogical(4);
                    _policyWarningTextLabel.Location = new Point(warningPadding, warningTextTop);
                    _policyWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningLinkTop = _policyWarningTextLabel.Bottom + ScaleLogical(6);
                    _policyWarningLinkLabel.Location = new Point(warningPadding, warningLinkTop);

                    int panelHeight = _policyWarningLinkLabel.Bottom + warningPadding;
                    _policyWarningPanel.SetBounds(outerPadding, y, panelWidth, panelHeight);
                    y = _policyWarningPanel.Bottom + verticalGap;
                }
                else
                {
                    _policyWarningPanel.SetBounds(outerPadding, y, Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2)), 0);
                }

                // One row, always: [toggle] [password] [generate]. The toggle keeps a fixed position
                // so its target does not move under the pointer, and the field grows to the right of
                // it instead of pushing everything below down a line. No caption — the button says
                // what the row is.
                int toggleMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(_passwordToggleButton, out toggleMinWidth);
                int toggleWidth = _passwordToggleButton.Width;
                int toggleHeight = _passwordToggleButton.Height;
                _passwordToggleButton.SetBounds(outerPadding, y, toggleWidth, toggleHeight);

                if (_passwordEnabled)
                {
                    int ignoredGenerateMinWidth;
                    FooterButtonLayoutHelper.ApplyButtonSize(_passwordGenerateButton, out ignoredGenerateMinWidth);
                    int generateButtonWidth = _passwordGenerateButton.Width;
                    int generateButtonHeight = _passwordGenerateButton.Height;
                    int gap = ScaleLogical(8);

                    int passwordRowRight = Math.Max(
                        ScaleLogical(260) + outerPadding,
                        ClientSize.Width - outerPadding);
                    int passwordLeft = _passwordToggleButton.Right + gap;
                    int passwordWidth = Math.Max(
                        ScaleLogical(110),
                        passwordRowRight - passwordLeft - generateButtonWidth - gap);
                    int passwordHeight = _passwordTextBox.PreferredHeight + ScaleLogical(2);

                    // Vertically centre the field against the buttons so the row reads as one line.
                    int passwordTop = y + Math.Max(0, (toggleHeight - passwordHeight) / 2);
                    _passwordTextBox.SetBounds(passwordLeft, passwordTop, passwordWidth, passwordHeight);
                    _passwordGenerateButton.SetBounds(
                        _passwordTextBox.Right + gap,
                        y,
                        generateButtonWidth,
                        generateButtonHeight);
                }

                y = _passwordToggleButton.Bottom + ScaleLogical(16);

                int groupWidth = Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2));

                _settingsGroup.SetBounds(outerPadding, y, groupWidth, ScaleLogical(132));
                int settingsLeft = ScaleLogical(12);
                int settingsTop = ScaleLogical(24);
                int settingsLineGap = ScaleLogical(24);
                _addUsersCheckBox.Location = new Point(settingsLeft, settingsTop);
                _lobbyCheckBox.Location = new Point(settingsLeft, settingsTop + settingsLineGap);
                _searchCheckBox.Location = new Point(settingsLeft, settingsTop + (settingsLineGap * 2));
                int settingsGroupHeight = _searchCheckBox.Bottom + ScaleLogical(14);
                _settingsGroup.Height = Math.Max(ScaleLogical(108), settingsGroupHeight);

                y = _settingsGroup.Bottom + verticalGap;
                _moderatorGroup.SetBounds(outerPadding, y, groupWidth, ScaleLogical(72));
                _moderatorGroup.Height = LayoutModeratorGroupControls(groupWidth);
                var footerButtons = new List<Button> { _okButton, _cancelButton };
                int minClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);
                if (ensureClientHeight && minClientWidth > ClientSize.Width)
                {
                    ClientSize = new Size(minClientWidth, ClientSize.Height);
                }
                int buttonHeight = Math.Max(_okButton.Height, _cancelButton.Height);
                int minClientHeight = _moderatorGroup.Bottom + ScaleLogical(16) + buttonHeight + FooterButtonLayoutHelper.DefaultBottomPadding;
                if (ensureClientHeight && minClientHeight > ClientSize.Height)
                {
                    ClientSize = new Size(ClientSize.Width, minClientHeight);
                }

                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);

            }
            finally
            {
                _layoutApplying = false;
            }
        }

        // The width is passed in rather than read from _moderatorGroup.ClientSize: the group may not
        // have been sized yet when this runs, and reading a default size once produced a narrow list
        // sitting on top of the hint instead of above it.
        private int LayoutModeratorGroupControls(int groupWidth)
        {
            int innerPadding = ScaleLogical(12);
            int contentWidth = Math.Max(ScaleLogical(160), groupWidth - (innerPadding * 3));
            int contentTop = ScaleLogical(22);

            // The list is only shown when there is something to tick; otherwise the hint alone
            // explains what to do, and an empty box would just be a hole in the dialog.
            // SetBounds runs in both cases. Skipping it while hidden left the control at WinForms'
            // default size and position, which then showed through as a small box overlapping the
            // hint as soon as the list became visible.
            bool hasCandidates = _moderatorListBox.Items.Count > 0;
            int rowHeight = _moderatorListBox.ItemHeight > 0 ? _moderatorListBox.ItemHeight : ScaleLogical(17);
            int visibleRows = Math.Min(Math.Max(_moderatorListBox.Items.Count, 1), 5);
            int listHeight = hasCandidates ? (rowHeight * visibleRows) + ScaleLogical(6) : 0;
            _moderatorListBox.SetBounds(innerPadding, contentTop, contentWidth, listHeight);
            _moderatorListBox.Visible = hasCandidates;
            if (hasCandidates)
            {
                contentTop = _moderatorListBox.Bottom + ScaleLogical(8);
            }

            if (_moderatorAddressbookWarningPanel.Visible)
            {
                int panelPadding = ScaleLogical(8);
                int panelWidth = contentWidth;
                int warningTextWidth = Math.Max(ScaleLogical(120), panelWidth - (panelPadding * 2));

                _moderatorAddressbookWarningTitleLabel.Location = new Point(panelPadding, panelPadding);
                _moderatorAddressbookWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningTextTop = _moderatorAddressbookWarningTitleLabel.Bottom + ScaleLogical(4);
                _moderatorAddressbookWarningTextLabel.Location = new Point(panelPadding, warningTextTop);
                _moderatorAddressbookWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningLinkTop = _moderatorAddressbookWarningTextLabel.Bottom + ScaleLogical(6);
                _moderatorAddressbookWarningLinkLabel.Location = new Point(panelPadding, warningLinkTop);

                int panelHeight = _moderatorAddressbookWarningLinkLabel.Bottom + panelPadding;
                _moderatorAddressbookWarningPanel.SetBounds(innerPadding, contentTop, panelWidth, panelHeight);
                contentTop = _moderatorAddressbookWarningPanel.Bottom + ScaleLogical(8);
            }
            // Let the hint wrap to its own height instead of being stretched to fill the group —
            // stretching left a large empty gap between the list and the text. Same AutoSize +
            // MaximumSize pattern the warning labels above use.
            _moderatorHintLabel.AutoSize = true;
            _moderatorHintLabel.MaximumSize = new Size(contentWidth, 0);
            _moderatorHintLabel.Location = new Point(innerPadding, contentTop);
            return _moderatorHintLabel.Bottom + innerPadding;
        }

                // Determines whether event conversations are supported (Nextcloud >= 31).
        private static bool DetermineEventConversationSupport(AddinSettings defaults, out string versionText)
        {
            versionText = string.Empty;            if (defaults == null)
            {
                return true;
            }

            versionText = defaults.LastKnownServerVersion ?? string.Empty;

            Version parsed;
            if (NextcloudVersionHelper.TryParse(versionText, out parsed))
            {
                versionText = parsed.ToString();
                return parsed.Major >= 31;
            }
            return true;
        }

        private void ApplyDefaults(AddinSettings defaults, string appointmentSubject)
        {
            // The room name is the meeting subject; there is no separate editing surface for it,
            // so this only picks a placeholder for a meeting that has no subject yet.
            string titleDefault = string.IsNullOrWhiteSpace(appointmentSubject) ? DefaultTitle : appointmentSubject.Trim();
            bool passwordDefault = defaults == null || defaults.TalkDefaultPasswordEnabled;
            bool addUsersDefault = defaults == null || defaults.TalkDefaultAddUsers;
            bool addGuestsDefault = defaults != null && defaults.TalkDefaultAddGuests;
            bool lobbyDefault = defaults == null || defaults.TalkDefaultLobbyEnabled;
            bool searchDefault = defaults == null || defaults.TalkDefaultSearchVisible;
            TalkRoomType roomTypeDefault = defaults != null ? defaults.TalkDefaultRoomType : TalkRoomType.EventConversation;

            if (PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "talk"))
            {
                bool policyBool;
                string policyString;

                // talk_title is deliberately not applied: the room name follows the meeting
                // subject now, and TalkRoomSyncService would overwrite a pinned title on the very
                // next save anyway. Two owners for one value is worse than none.
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_set_password", out policyBool))
                {
                    passwordDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_users", out policyBool))
                {
                    addUsersDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_guests", out policyBool))
                {
                    addGuestsDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_lobby_active", out policyBool))
                {
                    lobbyDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_show_in_search", out policyBool))
                {
                    searchDefault = policyBool;
                }

                policyString = _backendPolicyStatus.GetPolicyString("talk", "talk_room_type");
                if (!string.IsNullOrWhiteSpace(policyString))
                {
                    roomTypeDefault = string.Equals(policyString.Trim(), "event", StringComparison.OrdinalIgnoreCase)
                        ? TalkRoomType.EventConversation
                        : TalkRoomType.StandardRoom;
                }
            }

            TalkTitle = titleDefault;

            _passwordEnabled = passwordDefault;
            TalkPassword = string.Empty;
            _passwordTextBox.Text = string.Empty;
            if (_passwordEnabled)
            {
                TalkPassword = PasswordGenerationHelper.GenerateRoomPin(_passwordPolicy, DefaultMinPasswordLength);
                _passwordTextBox.Text = TalkPassword;
            }

            _addUsersCheckBox.Checked = addUsersDefault;
            _addGuestsCheckBox.Checked = addGuestsDefault;
            _lobbyCheckBox.Checked = lobbyDefault;
            _searchCheckBox.Checked = searchDefault;

            if (!_eventConversationsSupported && roomTypeDefault == TalkRoomType.EventConversation)
            {
                roomTypeDefault = TalkRoomType.StandardRoom;
            }

            SelectRoomType(roomTypeDefault);

            LobbyUntilStart = _lobbyCheckBox.Checked;
            SearchVisible = _searchCheckBox.Checked;
            AddUsers = _addUsersCheckBox.Checked;
            AddGuests = _addGuestsCheckBox.Checked;
            ModeratorIds = new List<string>();

            ApplyPolicyWarningUi();
            ApplyPolicyLockState();
            ApplySystemAddressbookLockState();
            UpdatePasswordState();
            UpdateModeratorHint();
        }

        private bool IsPolicyLocked(string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked("talk", key);
        }

        private bool IsPolicyGeneratePasswordEnabled()
        {
            if (!PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "talk"))
            {
                return true;
            }
            bool value;
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_generate_password", out value))
            {
                return value;
            }
            return true;
        }

        private void ApplyPolicyWarningUi()
        {
            bool visible = _backendPolicyStatus != null
                           && _backendPolicyStatus.WarningVisible
                           && !string.IsNullOrWhiteSpace(_backendPolicyStatus.WarningMessage);
            _policyWarningPanel.Visible = visible;
            _policyWarningTextLabel.Text = visible ? _backendPolicyStatus.WarningMessage : string.Empty;
        }

        private void ApplyPolicyLockState()
        {
            bool lockRoomType = IsPolicyLocked("talk_room_type");
            bool lockPassword = IsPolicyLocked("talk_set_password");
            bool lockLobby = IsPolicyLocked("talk_lobby_active");
            bool lockSearch = IsPolicyLocked("talk_show_in_search");
            bool lockUsers = IsPolicyLocked("talk_add_users");
            bool lockGuests = IsPolicyLocked("talk_add_guests");

            _roomTypeComboBox.Enabled = !lockRoomType;
            _passwordToggleButton.Enabled = !lockPassword;
            _lobbyCheckBox.Enabled = !lockLobby;
            _searchCheckBox.Enabled = !lockSearch;

            // Address book lock is applied in ApplySystemAddressbookLockState().
            if (!lockUsers)
            {
                _addUsersCheckBox.Enabled = true;
            }
            if (!lockGuests)
            {
                _addGuestsCheckBox.Enabled = true;
            }

            _disabledTooltipHints.Apply(_roomTypeComboBox, lockRoomType ? Strings.PolicyAdminControlledTooltip : _toolTip.GetToolTip(_roomTypeComboBox), lockRoomType, _roomTypeLabel);
            _disabledTooltipHints.Apply(
                _passwordToggleButton,
                lockPassword ? Strings.PolicyAdminControlledTooltip : string.Empty,
                lockPassword,
                _passwordGenerateButton,
                _passwordTextBox);
            _disabledTooltipHints.Apply(_lobbyCheckBox, lockLobby ? Strings.PolicyAdminControlledTooltip : Strings.TooltipLobby, lockLobby);
            _disabledTooltipHints.Apply(_searchCheckBox, lockSearch ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSearchVisible, lockSearch);

            if (lockUsers)
            {
                _disabledTooltipHints.Apply(_addUsersCheckBox, Strings.PolicyAdminControlledTooltip, true);
            }
            if (lockGuests)
            {
                _disabledTooltipHints.Apply(_addGuestsCheckBox, Strings.PolicyAdminControlledTooltip, true);
            }
        }

        private void ApplySystemAddressbookLockState()
        {
            bool lockActive = !_systemAddressbookAvailable;
            bool usersPolicyLocked = IsPolicyLocked("talk_add_users");
            bool guestsPolicyLocked = IsPolicyLocked("talk_add_guests");
            string lockDetail = lockActive
                ? (!string.IsNullOrWhiteSpace(_systemAddressbookError) ? _systemAddressbookError : Strings.TalkSystemAddressbookRequiredMessage)
                : string.Empty;

            if (lockActive)
            {
                _addUsersCheckBox.Checked = false;
                _addGuestsCheckBox.Checked = false;
                ClearModeratorSelection();
            }

            _addUsersCheckBox.Enabled = !lockActive && !usersPolicyLocked;
            _addGuestsCheckBox.Enabled = !lockActive && !guestsPolicyLocked;
            _moderatorListBox.Enabled = !lockActive;

            _moderatorAddressbookWarningPanel.Visible = lockActive;
            _moderatorAddressbookWarningTextLabel.Text = lockActive
                ? Strings.TalkSystemAddressbookRequiredMessage
                : string.Empty;

            _disabledTooltipHints.Apply(
                _addUsersCheckBox,
                lockActive ? Strings.TooltipAddUsersLocked : (usersPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddUsers),
                lockActive || usersPolicyLocked);
            _disabledTooltipHints.Apply(
                _addGuestsCheckBox,
                lockActive ? Strings.TooltipAddGuestsLocked : (guestsPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddGuests),
                lockActive || guestsPolicyLocked);
            _disabledTooltipHints.Apply(_moderatorListBox, lockActive ? Strings.TooltipModeratorLocked : Strings.TooltipModerator, lockActive, _moderatorHintLabel);

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "Talk wizard system address book lock state applied (locked=" + lockActive +
                ", available=" + _systemAddressbookAvailable +
                ", hasError=" + (!string.IsNullOrWhiteSpace(lockDetail)) +
                ", usersPolicyLocked=" + usersPolicyLocked +
                ", guestsPolicyLocked=" + guestsPolicyLocked + ").");
        }

        private static void OpenSystemAddressbookSetupGuide()
        {
            BrowserLauncher.OpenUrl(
                Strings.TalkSystemAddressbookAdminGuideUrl,
                LogCategories.Talk,
                "Failed to open system address book setup guide URL.");
        }

        private void UpdateRoomTypeTooltip()
        {
            if (IsPolicyLocked("talk_room_type"))
            {
                _disabledTooltipHints.Apply(_roomTypeComboBox, Strings.PolicyAdminControlledTooltip, true, _roomTypeLabel);
                return;
            }
            var selected = _roomTypeComboBox.SelectedItem as RoomTypeOption;
            var roomType = selected != null ? selected.Value : TalkRoomType.EventConversation;
            _disabledTooltipHints.Apply(
                _roomTypeComboBox,
                roomType == TalkRoomType.EventConversation ? Strings.TooltipRoomTypeEvent : Strings.TooltipRoomTypeStandard,
                false,
                _roomTypeLabel);
        }

                // Collects user input, validates the password, and exposes the selection to the caller.
        private void OnOkButtonClick(object sender, EventArgs e)
        {

            bool passwordEnabled = _passwordEnabled;
            TalkPassword = passwordEnabled ? _passwordTextBox.Text.Trim() : string.Empty;

            // Mirror the server's password_policy locally, so a password that Nextcloud would
            // reject is caught here with the actual reason instead of failing room creation.
            string policyViolation = passwordEnabled
                ? PasswordGenerationHelper.DescribePolicyViolation(TalkPassword, _passwordPolicy, DefaultMinPasswordLength)
                : null;
            if (!string.IsNullOrEmpty(policyViolation))
            {
                MessageBox.Show(
                    policyViolation,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _passwordTextBox.Focus();
                _passwordTextBox.SelectAll();
                return;
            }

            LobbyUntilStart = _lobbyCheckBox.Checked;
            SearchVisible = _searchCheckBox.Checked;
            AddUsers = _addUsersCheckBox.Checked;
            AddGuests = false;
            SelectedRoomType = TalkRoomType.EventConversation;

            ModeratorIds = CollectCheckedModeratorIds();
        }

        private void UpdatePasswordState()
        {
            bool enabled = _passwordEnabled;
            bool lockPassword = IsPolicyLocked("talk_set_password");
            bool allowGenerate = IsPolicyGeneratePasswordEnabled();

            // When the administrator pins talk_set_password the state cannot be changed at all, so
            // the toggle is disabled rather than hidden — the reason stays visible via its tooltip.
            _passwordToggleButton.Text = enabled ? Strings.TalkPasswordRemove : Strings.TalkPasswordAdd;
            _passwordToggleButton.Enabled = !lockPassword;
            _passwordTextBox.Visible = enabled;
            _passwordGenerateButton.Visible = enabled;

            _passwordTextBox.Enabled = enabled && !lockPassword;
            _passwordGenerateButton.Enabled = enabled && allowGenerate;
            _disabledTooltipHints.Apply(
                _passwordGenerateButton,
                !enabled ? string.Empty : (allowGenerate ? string.Empty : Strings.PolicyAdminControlledTooltip),
                enabled && !allowGenerate);
        }

        // Adding a password fills it in straight away, so the common path is a single click; taking
        // it away clears the field so a stale value cannot be submitted.
        private void SetPasswordEnabled(bool enabled)
        {
            if (_passwordEnabled == enabled)
            {
                return;
            }

            _passwordEnabled = enabled;
            if (enabled)
            {
                if (string.IsNullOrWhiteSpace(_passwordTextBox.Text))
                {
                    _passwordTextBox.Text = PasswordGenerationHelper.GenerateRoomPin(_passwordPolicy, DefaultMinPasswordLength);
                }
            }
            else
            {
                _passwordTextBox.Text = string.Empty;
            }

            UpdatePasswordState();
            ApplyDialogLayout(true);
            if (enabled)
            {
                _passwordTextBox.Focus();
                _passwordTextBox.SelectAll();
            }
            else
            {
                _passwordToggleButton.Focus();
            }
        }

        private void GeneratePassword()
        {
            if (!_passwordEnabled)
            {
                return;
            }
            _passwordTextBox.Text = PasswordGenerationHelper.GenerateRoomPin(_passwordPolicy, DefaultMinPasswordLength);
        }

        private void SelectRoomType(TalkRoomType type)
        {
            foreach (var item in _roomTypeComboBox.Items)
            {
                var option = item as RoomTypeOption;                if (option != null && option.Value == type)
                {
                    _roomTypeComboBox.SelectedItem = option;
                    return;
                }
            }
            if (_roomTypeComboBox.Items.Count > 0)
            {
                _roomTypeComboBox.SelectedIndex = 0;
            }
        }

        private sealed class RoomTypeOption
        {
            internal RoomTypeOption(TalkRoomType value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }

            internal TalkRoomType Value { get; private set; }

            internal string Label { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }

    }
}

