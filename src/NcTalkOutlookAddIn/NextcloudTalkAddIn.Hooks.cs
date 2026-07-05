// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Outlook hook lifecycle for inspector/application events and compose-tracking entry points.
    public sealed partial class NextcloudTalkAddIn
    {
        private void EnsureInspectorHook()
        {            if (_outlookApplication == null || _inspectors != null)
            {
                return;
            }
            try
            {
                _inspectors = _outlookApplication.Inspectors;                if (_inspectors != null)
                {
                    _inspectors.NewInspector += OnNewInspector;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Inspectors.NewInspector.", ex);
                _inspectors = null;
            }
        }

        private void UnhookInspector()
        {            if (_inspectors == null)
            {
                return;
            }
            try
            {
                _inspectors.NewInspector -= OnNewInspector;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Inspectors.NewInspector.", ex);
            }
            finally
            {
                ComInteropScope.TryFinalRelease(_inspectors, LogCategories.Core, "Failed to release Inspectors COM object.");
                _inspectors = null;
            }
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {            if (inspector == null)
            {
                return;
            }
            try
            {
                var appointment = inspector.CurrentItem as Outlook.AppointmentItem;                if (appointment != null)
                {
                    EnsureSubscriptionForAppointment(appointment);
                }
                var mail = inspector.CurrentItem as Outlook.MailItem;                if (mail != null && IsMailComposeCandidate(mail, "new_inspector"))
                {
                    string inspectorIdentityKey = ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
                    DeferMailComposeSubscriptionEnsure(mail, inspectorIdentityKey, false, "new_inspector");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewInspector event.", ex);
            }
        }

        // Deferring this off NewInspector's own callback isn't required for correctness (the ribbon bug
        // traced back to Explorer.InlineResponse, not NewInspector — see NextcloudTalkAddIn.Lifecycle.cs),
        // but it's kept as a defensive pattern matching QueueDeferredAppointmentSubscriptionEnsure.
        private void DeferMailComposeSubscriptionEnsure(Outlook.MailItem mail, string inspectorIdentityKey, bool isInlineResponse, string reason)
        {
            SynchronizationContext context = _uiSynchronizationContext;
            if (context == null)
            {
                RunDeferredMailComposeSubscriptionEnsure(mail, inspectorIdentityKey, isInlineResponse, reason);
                return;
            }
            context.Post(_ => RunDeferredMailComposeSubscriptionEnsure(mail, inspectorIdentityKey, isInlineResponse, reason), null);
        }

        private void RunDeferredMailComposeSubscriptionEnsure(Outlook.MailItem mail, string inspectorIdentityKey, bool isInlineResponse, string reason)
        {
            try
            {
                EnsureMailComposeSubscription(mail, inspectorIdentityKey, isInlineResponse);
                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink("Compose subscription ensured (reason=" + (reason ?? "n/a") + ", deferred=True).");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to ensure deferred compose subscription (reason=" + (reason ?? "n/a") + ").", ex);
            }
        }

        private static bool IsMailComposeCandidate(Outlook.MailItem mail, string reason)
        {
            if (mail == null)
            {
                return false;
            }
            try
            {
                if (mail.Sent)
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Mail compose subscription skipped (reason=" + (reason ?? "n/a") + ", sent=True).");
                    }
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to verify mail compose state (reason=" + (reason ?? "n/a") + ").", ex);
                return false;
            }
        }
    }
}

