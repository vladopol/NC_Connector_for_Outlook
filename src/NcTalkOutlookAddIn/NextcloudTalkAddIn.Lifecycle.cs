// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Add-in lifecycle and bootstrap/teardown flow.
    public sealed partial class NextcloudTalkAddIn
    {
        // Long enough to clear Outlook's own startup burst, short enough that calendar changes made
        // in the first seconds are not missed by much.
        private const int StartupWiringDelayMs = 5000;

        private System.Windows.Forms.Timer _startupWiringTimer;

                // Outlook calls this method when the add-in is loaded.
        // Stores the Application instance for later actions.
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            _uiSynchronizationContext = SynchronizationContext.Current;
            string outlookProfileName = ResolveCurrentOutlookProfileName();
            _settingsStorage = new NcTalkOutlookAddIn.Settings.SettingsStorage(outlookProfileName);
            _currentSettings = _settingsStorage.Load();
            if (_currentSettings != null && !_currentSettings.TalkDeleteRoomOnEventDelete)
            {
                _currentSettings.TalkDeleteRoomOnEventDelete = true;
                _settingsStorage.Save(_currentSettings);
                LogCore("TalkDeleteRoomOnEventDelete force-enabled and saved.");
            }
            if (_currentSettings != null && _currentSettings.TalkDefaultRoomType != NcTalkOutlookAddIn.Models.TalkRoomType.EventConversation)
            {
                _currentSettings.TalkDefaultRoomType = NcTalkOutlookAddIn.Models.TalkRoomType.EventConversation;
                _settingsStorage.Save(_currentSettings);
                LogCore("TalkDefaultRoomType force-set to EventConversation and saved.");
            }
            if (_currentSettings != null && _currentSettings.TalkDefaultPasswordEnabled)
            {
                // The room dialog adds a password per room on demand, so the stored "always set a
                // password" default no longer has a UI and would silently pre-fill every room.
                _currentSettings.TalkDefaultPasswordEnabled = false;
                _settingsStorage.Save(_currentSettings);
                LogCore("TalkDefaultPasswordEnabled force-disabled and saved.");
            }
            ConfigureDiagnosticsLogger(_currentSettings);
            TryApplyTransportSecurityFromSettings("startup", false);
            TryApplyOfficeUiLanguage();
            MigrateShareNameSeededFromLabel();
            LogCore("Add-in connected (Outlook version=" + (_outlookApplication != null ? _outlookApplication.Version : "unknown") + ").");
            if (!string.IsNullOrWhiteSpace(outlookProfileName))
            {
                LogCore("Using Outlook profile settings: " + outlookProfileName + ".");
            }
            if (_currentSettings != null)
            {
                LogSettings("Settings loaded (AuthMode=" + _currentSettings.AuthMode + ", Debug=" + _currentSettings.DebugLoggingEnabled + ", LogAnonymize=" + _currentSettings.LogAnonymizationEnabled + ").");
            }

            _freeBusyManager = new FreeBusyManager();
            _freeBusyManager.Initialize(_outlookApplication);
            // Explorer.InlineResponse is deliberately never hooked: subscribing to it — regardless of
            // what the handler does — breaks Outlook's rendering of the inline-reply command bar
            // (Send/Discard/PopOut) on some builds (confirmed 16.0.0.14334 / 17932.20842). Bisected across
            // 3.1.0.2-3.1.0.7: ribbon XML and NewInspector are not implicated, only this specific event.
            // Automatic FileLink attachment sharing is unavailable for Reading Pane inline replies as a
            // result; it still works for popped-out compose windows via NewInspector.
            EnsureInspectorHook();

            // Everything below touches the MAPI store (Session.GetDefaultFolder) or the registry and
            // is not needed for Outlook to finish loading. Outlook times OnConnection and disables
            // add-ins that exceed its threshold, so this work is posted back to the UI thread and
            // runs once the message loop is pumping instead of inside the measured window.
            DeferStartupWiring();
        }

        // Posting this to the UI thread runs it at the next message-loop turn — which is exactly
        // when the user is clicking their first mail. The work itself opens the default calendar
        // folder, and on an online-mode Exchange profile (common on terminal servers, where the OST
        // is impractical) every folder open is a round-trip to the server. A short delay keeps it
        // out of the window where the user is waiting on their own first interaction.
        private void DeferStartupWiring()
        {
            if (_startupWiringTimer != null)
            {
                return;
            }
            try
            {
                _startupWiringTimer = new System.Windows.Forms.Timer();
                _startupWiringTimer.Interval = StartupWiringDelayMs;
                _startupWiringTimer.Tick += OnStartupWiringTick;
                _startupWiringTimer.Start();
                return;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to schedule deferred startup wiring.", ex);
                StopStartupWiringTimer();
            }

            RunStartupWiring();
        }

        private void OnStartupWiringTick(object sender, EventArgs e)
        {
            StopStartupWiringTimer();
            RunStartupWiring();
        }

        private void StopStartupWiringTimer()
        {
            if (_startupWiringTimer == null)
            {
                return;
            }
            try
            {
                _startupWiringTimer.Stop();
                _startupWiringTimer.Tick -= OnStartupWiringTick;
                _startupWiringTimer.Dispose();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to dispose the startup wiring timer.", ex);
            }
            finally
            {
                _startupWiringTimer = null;
            }
        }

        private void RunStartupWiring()
        {
            // Shutdown can win the race against the deferred callback.
            if (_outlookApplication == null)
            {
                return;
            }
            try
            {
                // Timed individually: on a slow profile these are the add-in's only expensive
                // startup steps, and the log is the only way to tell them apart from Outlook's own
                // startup work when diagnosing a sluggish launch.
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                EnsureTalkCalendarWatcher();
                long watcherMs = stopwatch.ElapsedMilliseconds;

                ApplyIfbSettings();
                long ifbMs = stopwatch.ElapsedMilliseconds - watcherMs;

                ApplyCalDavSyncSettings();
                long calDavMs = stopwatch.ElapsedMilliseconds - watcherMs - ifbMs;

                LogCore(
                    "Deferred startup wiring completed (calendarWatcher="
                    + watcherMs.ToString(CultureInfo.InvariantCulture)
                    + "ms, freeBusyRegistry="
                    + ifbMs.ToString(CultureInfo.InvariantCulture)
                    + "ms, calDavSync="
                    + calDavMs.ToString(CultureInfo.InvariantCulture)
                    + "ms).");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Deferred startup wiring failed.", ex);
            }
        }

        // Older builds seeded SharingDefaultShareName with the settings field's own caption, so the
        // stored value was a UI label — and it ended up as the real folder name on Nextcloud
        // ("20260729_Share name"). Clear it once so the wizard's localized fallback applies again.
        //
        // Runs after TryApplyOfficeUiLanguage so the caption is compared in the same language it
        // would have been written in. The English literal is checked too, for profiles written
        // before a language change.
        private void MigrateShareNameSeededFromLabel()
        {
            if (_currentSettings == null || string.IsNullOrWhiteSpace(_currentSettings.SharingDefaultShareName))
            {
                return;
            }

            string stored = _currentSettings.SharingDefaultShareName.Trim();
            if (!string.Equals(stored, Strings.SharingDefaultShareNameLabel, StringComparison.Ordinal)
                && !string.Equals(stored, "Share name", StringComparison.Ordinal))
            {
                return;
            }
            try
            {
                _currentSettings.SharingDefaultShareName = string.Empty;
                _settingsStorage.Save(_currentSettings);
                LogCore("SharingDefaultShareName cleared: it held the settings label, not a share name.");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to clear the label-seeded share name.", ex);
            }
        }

        private void TryApplyOfficeUiLanguage()
        {
            try
            {                if (_outlookApplication == null)
                {
                    return;
                }

                LanguageSettings languageSettings = _outlookApplication.LanguageSettings;
                // Guard against missing COM runtime objects.
                if (languageSettings == null)
                {
                    return;
                }
                int lcid = languageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
                CultureInfo culture = CultureInfo.GetCultureInfo(lcid);
                Strings.SetPreferredUiLanguage(culture.Name);
                LogCore("Office UI language detected: " + culture.Name + " (LCID=" + lcid + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to detect Office UI language.", ex);
            }
        }

        private string ResolveCurrentOutlookProfileName()
        {            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            object session = null;
            try
            {
                session = _outlookApplication.Session;                if (session == null)
                {
                    return string.Empty;
                }

                object rawProfileName = session.GetType().InvokeMember(
                    "CurrentProfileName",
                    BindingFlags.GetProperty,
                    null,
                    session,
                    null,
                    CultureInfo.InvariantCulture);

                string profileName = rawProfileName as string;
                return string.IsNullOrWhiteSpace(profileName) ? string.Empty : profileName.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve current Outlook profile name.", ex);
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Core,
                    "Failed to release Outlook session COM object after profile resolution.");
            }
        }

                // Outlook signals that the add-in is unloading.
        // Cleanup hooks follow once resources are held.
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            TearDownAddInState("disconnect", true);
            LogCore("Add-in disconnected (removeMode=" + removeMode + ").");
        }

                // Required IDTExtensibility2 callback.
        // Intentionally no-op because runtime wiring is already complete in OnConnection.
        public void OnAddInsUpdate(ref Array custom)
        {
        }

                // Required IDTExtensibility2 callback.
        // Intentionally no-op because startup work is handled in OnConnection.
        public void OnStartupComplete(ref Array custom)
        {
        }

                // Called when Outlook shuts down; reserved for future cleanup steps.
        public void OnBeginShutdown(ref Array custom)
        {
            TearDownAddInState("shutdown", false);
        }

                // Centralized teardown used by both OnBeginShutdown and OnDisconnection.
        // This path must be idempotent, because Outlook can call both callbacks.
        private void TearDownAddInState(string origin, bool clearOutlookApplication)
        {
            StopStartupWiringTimer();
            UnhookInspector();
            UnhookTalkCalendarWatcher();
            UnhookMailComposeSubscriptions();
            _freeBusyManager = null;

            if (_calDavCalendarSync != null)
            {
                _calDavCalendarSync.Dispose();
                _calDavCalendarSync = null;
            }

            if (clearOutlookApplication)
            {
                _outlookApplication = null;
            }

            _ribbonUi = null;
            _uiSynchronizationContext = null;
            _deferredAppointmentEnsureState.ClearPendingKeys();
        }
    }
}
