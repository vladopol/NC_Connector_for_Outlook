// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared password-policy generation helper used by Talk and Filelink dialogs.
    internal static class PasswordGenerationHelper
    {
        internal static int ResolveMinLength(PasswordPolicyInfo passwordPolicy, int defaultMinLength)
        {
            if (passwordPolicy != null && passwordPolicy.MinLength > 0)
            {
                return passwordPolicy.MinLength;
            }

            return defaultMinLength;
        }

        internal static string GenerateWithServerPolicyFallback(
            TalkServiceConfiguration configuration,
            PasswordPolicyInfo passwordPolicy,
            int minLength,
            string logCategory)
        {
            string generated = null;

            try
            {
                if (configuration != null && passwordPolicy != null && passwordPolicy.HasPolicy)
                {
                    var policyService = new PasswordPolicyService(configuration);
                    generated = policyService.GeneratePassword(passwordPolicy);
                }
            }
            catch (Exception ex)
            {
                generated = null;
                DiagnosticsLogger.LogException(logCategory, "Password generation via server policy failed; falling back to local generator.", ex);
            }

            if (string.IsNullOrWhiteSpace(generated) || generated.Trim().Length < minLength)
            {
                generated = PasswordGenerator.GenerateLocalPassword(minLength);
            }

            return generated;
        }

        internal static string GenerateWithPolicyDefaults(
            TalkServiceConfiguration configuration,
            PasswordPolicyInfo passwordPolicy,
            int defaultMinLength,
            string logCategory)
        {
            int minLength = ResolveMinLength(passwordPolicy, defaultMinLength);
            return GenerateWithServerPolicyFallback(configuration, passwordPolicy, minLength, logCategory);
        }

        // Room PIN generated locally, without calling Nextcloud's server-side generator (which
        // returns long strings with special characters that defeat the point of a PIN).
        //
        // Stays digits-only whenever the server accepts that, and adds the minimum the policy
        // demands when it does not — Nextcloud validates Talk conversation passwords against the
        // same password_policy as account passwords, so a digits-only PIN is refused outright on a
        // server enforcing letters or symbols.
        internal static string GenerateRoomPin(PasswordPolicyInfo passwordPolicy, int defaultMinLength)
        {
            int minLength = ResolveMinLength(passwordPolicy, defaultMinLength);
            if (passwordPolicy == null || passwordPolicy.AllowsDigitsOnly)
            {
                return PasswordGenerator.GenerateNumericLocalPassword(minLength);
            }

            return PasswordGenerator.GeneratePinWithComplexity(
                minLength,
                passwordPolicy.EnforceUpperLowerCase,
                passwordPolicy.EnforceSpecialCharacters);
        }

        // Local pre-check mirroring what the server's password_policy will do, so a password typed
        // by hand fails in the dialog with a clear reason instead of failing room creation later.
        // Returns null when the password is acceptable, or a ready-to-show message when it is not.
        internal static string DescribePolicyViolation(
            string password,
            PasswordPolicyInfo passwordPolicy,
            int defaultMinLength)
        {
            string candidate = (password ?? string.Empty).Trim();
            if (candidate.Length == 0)
            {
                return null;
            }

            int minLength = ResolveMinLength(passwordPolicy, defaultMinLength);
            if (candidate.Length < minLength)
            {
                return string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.TalkPasswordTooShort,
                    minLength);
            }
            if (passwordPolicy == null)
            {
                return null;
            }

            bool hasUpper = false;
            bool hasLower = false;
            bool hasDigit = false;
            bool hasSpecial = false;
            for (int i = 0; i < candidate.Length; i++)
            {
                char c = candidate[i];
                if (char.IsUpper(c)) { hasUpper = true; }
                else if (char.IsLower(c)) { hasLower = true; }
                else if (char.IsDigit(c)) { hasDigit = true; }
                else { hasSpecial = true; }
            }

            if (passwordPolicy.EnforceUpperLowerCase && (!hasUpper || !hasLower))
            {
                return Strings.TalkPasswordNeedsUpperLower;
            }
            if (passwordPolicy.EnforceSpecialCharacters && !hasSpecial)
            {
                return Strings.TalkPasswordNeedsSpecial;
            }
            if (passwordPolicy.EnforceNumericCharacters && !hasDigit)
            {
                return Strings.TalkPasswordNeedsDigit;
            }
            return null;
        }

        internal static bool MeetsMinimumLength(
            string password,
            PasswordPolicyInfo passwordPolicy,
            int defaultMinLength)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                return false;
            }

            int minLength = ResolveMinLength(passwordPolicy, defaultMinLength);
            return password.Trim().Length >= minLength;
        }
    }
}
