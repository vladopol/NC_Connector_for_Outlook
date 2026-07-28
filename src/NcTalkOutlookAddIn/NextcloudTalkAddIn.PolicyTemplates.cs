// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn
{
        // Backend policy retrieval and Talk template/language normalization helpers.
    public sealed partial class NextcloudTalkAddIn
    {
        // Blocking fetch, served from BackendPolicyCache when it is still fresh. Never call this
        // from an Outlook event handler — the status endpoint uses a 45 s timeout, and a stall there
        // freezes Outlook's UI thread. Use PeekBackendPolicyStatus instead.
        internal BackendPolicyStatus FetchBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            string cacheKey = BackendPolicyCache.BuildKey(configuration);
            BackendPolicyStatus cached;
            if (BackendPolicyCache.TryGetFresh(cacheKey, out cached))
            {
                LogCore("Backend policy status served from cache (trigger=" + (trigger ?? "n/a") + ", age=" + BackendPolicyCache.DescribeAge(cacheKey) + ").");
                return cached;
            }

            BackendPolicyStatus fetched = FetchBackendPolicyStatusUncached(configuration, trigger);
            BackendPolicyCache.Store(cacheKey, fetched);
            return fetched;
        }

        // Non-blocking counterpart for UI-thread callers: hands back whatever is cached (possibly
        // stale, possibly null) and schedules a background refresh when the value is due.
        internal BackendPolicyStatus PeekBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            if (configuration == null || !configuration.IsComplete())
            {
                return null;
            }

            string cacheKey = BackendPolicyCache.BuildKey(configuration);
            BackendPolicyStatus cached = BackendPolicyCache.Peek(cacheKey);

            if (BackendPolicyCache.TryBeginBackgroundRefresh(cacheKey))
            {
                Task.Run(() =>
                {
                    BackendPolicyStatus fetched = null;
                    try
                    {
                        fetched = FetchBackendPolicyStatusUncached(configuration, trigger);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Background backend policy refresh failed (trigger=" + (trigger ?? "n/a") + ").", ex);
                    }
                    finally
                    {
                        BackendPolicyCache.Store(cacheKey, fetched);
                    }
                });
            }

            return cached;
        }

        private BackendPolicyStatus FetchBackendPolicyStatusUncached(TalkServiceConfiguration configuration, string trigger)
        {
            try
            {
                var service = new BackendPolicyService(configuration);
                BackendPolicyStatus status = service.FetchStatus();
                LogCore(
                    "Backend policy status fetched (trigger=" + (trigger ?? "n/a")
                    + ", active=" + (status != null && status.PolicyActive)
                    + ", share=" + (status != null && status.IsDomainActive("share"))
                    + ", talk=" + (status != null && status.IsDomainActive("talk"))
                    + ", emailSignature=" + (status != null && status.IsDomainActive("email_signature"))
                    + ", warningVisible=" + (status != null && status.WarningVisible)
                    + ", mode=" + (status != null ? status.Mode : "local")
                    + ", reason=" + (status != null ? status.Reason : "n/a")
                    + ").");
                return status;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Backend policy status fetch failed (trigger=" + (trigger ?? "n/a") + ").", ex);
                return null;
            }
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForTalkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogTalk("Password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForFileLinkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogFileLink("Sharing password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        internal static string ResolveTalkDescriptionLanguage(BackendPolicyStatus policyStatus, string fallbackLanguageOverride)
        {
            if (policyStatus != null
                && policyStatus.IsDomainActive("talk")
                && policyStatus.IsLocked("talk", "language_talk_description"))
            {
                string policyLanguageRaw = policyStatus.GetPolicyString("talk", "language_talk_description");
                if (!string.IsNullOrWhiteSpace(policyLanguageRaw))
                {
                    return TalkDescriptionTemplateController.NormalizeTalkDescriptionLanguage(policyLanguageRaw);
                }
            }
            return TalkDescriptionTemplateController.NormalizeTalkDescriptionLanguage(fallbackLanguageOverride);
        }

        internal static string ResolveTalkInvitationTemplate(BackendPolicyStatus policyStatus)
        {
            // Guard against null/inactive backend policy state.
            if (policyStatus == null || !policyStatus.IsDomainActive("talk"))
            {
                return string.Empty;
            }
            return policyStatus.GetPolicyString("talk", "talk_invitation_template");
        }

        internal static string ResolveTalkEventDescriptionType(BackendPolicyStatus policyStatus)
        {
            if (policyStatus != null && policyStatus.IsDomainActive("talk"))
            {
                string policyTypeRaw = policyStatus.GetPolicyString("talk", "event_description_type");
                if (!string.IsNullOrWhiteSpace(policyTypeRaw))
                {
                    return NormalizeTalkEventDescriptionType(policyTypeRaw);
                }
            }
            return "plain_text";
        }

        internal static string NormalizeTalkEventDescriptionType(string descriptionType)
        {
            return string.Equals(descriptionType, "html", StringComparison.OrdinalIgnoreCase)
                ? "html"
                : "plain_text";
        }

    }
}
