// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
        // Representation of Nextcloud password policy capabilities: minimum length, the complexity
    // flags the password_policy app enforces, and the generator endpoint.
    //
    // The complexity flags matter because Nextcloud validates Talk conversation passwords against
    // the same policy it applies to account and share passwords — the password_policy app has no
    // per-app exemption. A digits-only room PIN is rejected outright on a server that enforces
    // letters or symbols, so generation has to know what the server will actually accept.
    internal sealed class PasswordPolicyInfo
    {
        internal PasswordPolicyInfo(bool hasPolicy, int minLength, string generateUrl)
            : this(hasPolicy, minLength, generateUrl, false, false, false)
        {
        }

        internal PasswordPolicyInfo(
            bool hasPolicy,
            int minLength,
            string generateUrl,
            bool enforceUpperLowerCase,
            bool enforceSpecialCharacters,
            bool enforceNumericCharacters)
        {
            HasPolicy = hasPolicy;
            MinLength = minLength;
            GenerateUrl = generateUrl ?? string.Empty;
            EnforceUpperLowerCase = enforceUpperLowerCase;
            EnforceSpecialCharacters = enforceSpecialCharacters;
            EnforceNumericCharacters = enforceNumericCharacters;
        }

        internal bool HasPolicy { get; private set; }

        internal int MinLength { get; private set; }

        internal string GenerateUrl { get; private set; }

        internal bool EnforceUpperLowerCase { get; private set; }

        internal bool EnforceSpecialCharacters { get; private set; }

        internal bool EnforceNumericCharacters { get; private set; }

        // True when digits alone satisfy the server — the case this fork optimizes for: a room
        // password a guest can read aloud or type on a phone keypad.
        internal bool AllowsDigitsOnly
        {
            get { return !EnforceUpperLowerCase && !EnforceSpecialCharacters; }
        }
    }
}
