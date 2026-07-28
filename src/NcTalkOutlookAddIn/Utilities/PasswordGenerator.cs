// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Security.Cryptography;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared local password generator fallback used when backend password generation
    // is unavailable or policy responses are missing.
    internal static class PasswordGenerator
    {
        private const string Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
        private const string NumericAlphabet = "0123456789";
        private const int MinLength = 8;

        internal static string GenerateLocalPassword(int minLength)
        {
            return GenerateFromAlphabet(Alphabet, Math.Max(MinLength, minLength));
        }

        // Digits-only PIN, e.g. for Talk room passwords read aloud or typed by meeting guests.
        internal static string GenerateNumericLocalPassword(int minLength)
        {
            return GenerateFromAlphabet(NumericAlphabet, Math.Max(1, minLength));
        }

        // A room PIN that still satisfies a server enforcing password complexity.
        //
        // Nextcloud runs Talk conversation passwords through the same password_policy validation as
        // account passwords, so on a strict server a digits-only PIN is simply refused. Rather than
        // fall back to a random complex string that nobody can dictate over the phone, this keeps
        // the digit run as the body of the password and appends only the characters the policy
        // actually demands, always in the same place: "5729481!Kp" reads as "five seven two nine
        // four eight one, exclamation mark, capital K, small p".
        //
        // Ambiguous glyphs are excluded so a password read aloud or copied off a screen survives
        // the trip, and the symbol pool is deliberately tiny: "!" and "$" have an unambiguous name
        // in most languages, sit on every keyboard layout without a modifier surprise, and are
        // squarely inside the character set Nextcloud's password_policy counts as special.
        internal static string GeneratePinWithComplexity(
            int minLength,
            bool requireUpperLowerCase,
            bool requireSpecialCharacter)
        {
            const string UpperAlphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ";
            const string LowerAlphabet = "abcdefghijkmnopqrstuvwxyz";
            const string SpecialAlphabet = "!$";
            const int MinDigits = 4;

            int suffixLength = 0;
            if (requireSpecialCharacter)
            {
                suffixLength += 1;
            }
            if (requireUpperLowerCase)
            {
                suffixLength += 2;
            }
            if (suffixLength == 0)
            {
                return GenerateNumericLocalPassword(minLength);
            }

            int digitCount = Math.Max(MinDigits, Math.Max(1, minLength) - suffixLength);
            var builder = new System.Text.StringBuilder();
            builder.Append(GenerateFromAlphabet(NumericAlphabet, digitCount));
            if (requireSpecialCharacter)
            {
                builder.Append(GenerateFromAlphabet(SpecialAlphabet, 1));
            }
            if (requireUpperLowerCase)
            {
                builder.Append(GenerateFromAlphabet(UpperAlphabet, 1));
                builder.Append(GenerateFromAlphabet(LowerAlphabet, 1));
            }
            return builder.ToString();
        }

        private static string GenerateFromAlphabet(string alphabet, int length)
        {
            var chars = new char[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                byte[] data = new byte[4];
                for (int i = 0; i < chars.Length; i++)
                {
                    rng.GetBytes(data);
                    int index = (int)(BitConverter.ToUInt32(data, 0) % alphabet.Length);
                    chars[i] = alphabet[index];
                }
            }
            return new string(chars);
        }
    }
}
