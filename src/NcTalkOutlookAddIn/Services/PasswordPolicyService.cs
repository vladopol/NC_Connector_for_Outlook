// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
        // Fetches password policy information and the generator endpoint from Nextcloud capabilities.
    internal sealed class PasswordPolicyService
    {
        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;
        private static readonly string[] MinLengthKeys = { "minLength", "min_length", "minimumLength", "minimum_length" };
        // The password_policy app has used both singular and plural spellings across versions, and
        // occ shows snake_case for some keys — accept every variant rather than silently reading
        // "no complexity required" from a key name mismatch.
        private static readonly string[] UpperLowerKeys = { "enforceUpperLowerCase", "enforce_upper_lower_case", "enforceUpperLower" };
        private static readonly string[] SpecialCharKeys = { "enforceSpecialCharacters", "enforceSpecialCharacter", "enforce_special_characters", "enforce_special_character" };
        private static readonly string[] NumericCharKeys = { "enforceNumericCharacters", "enforceNumericCharacter", "enforce_numeric_characters", "enforce_numeric_character" };
        private static readonly string[] GenerateUrlKeys = { "api_generate", "apiGenerateUrl", "generateUrl", "generate_url" };

        internal PasswordPolicyService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        internal PasswordPolicyInfo FetchPolicy()
        {
            if (!_configuration.IsComplete())
            {
                return new PasswordPolicyInfo(false, 0, string.Empty);
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "PasswordPolicy.FetchPolicy"))
            {
                string baseUrl = _configuration.GetNormalizedBaseUrl();
                string url = baseUrl + "/ocs/v2.php/cloud/capabilities?format=json";
                DiagnosticsLogger.LogApi("GET " + url);

                IDictionary<string, object> root;
                HttpStatusCode statusCode;
                ExecuteJsonRequest("GET", url, null, out statusCode, out root);

                if (statusCode != HttpStatusCode.OK)
                {
                    DiagnosticsLogger.LogApi("Password policy fetch returned HTTP " + (int)statusCode + ".");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }
                if (root == null)
                {
                    DiagnosticsLogger.LogApi("Password policy fetch returned no parsable JSON payload.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }
                var ocs = NcJson.GetDictionary(root, "ocs");
                var data = NcJson.GetDictionary(ocs, "data");
                var caps = NcJson.GetDictionary(data, "capabilities");                if (caps == null)
                {
                    DiagnosticsLogger.LogApi("Password policy capabilities block missing.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }
                var policy = NcJson.GetDictionary(caps, "password_policy") ?? NcJson.GetDictionary(caps, "passwordPolicy");                if (policy == null)
                {
                    DiagnosticsLogger.LogApi("Password policy block missing.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }
                int minLength = GetFirstPositiveInt(policy, MinLengthKeys);
                if (minLength <= 0)
                {
                    var policies = NcJson.GetDictionary(policy, "policies");
                    var accountPolicy = NcJson.GetDictionary(policies, "account");
                    minLength = GetFirstPositiveInt(accountPolicy, MinLengthKeys);
                }
                var api = NcJson.GetDictionary(policy, "api");
                string generateRaw = GetFirstString(api, "generate", "generateUrl", "generate_url");
                if (string.IsNullOrWhiteSpace(generateRaw))
                {
                    generateRaw = GetFirstString(policy, GenerateUrlKeys);
                }
                string generateUrl = ResolvePolicyUrl(generateRaw, baseUrl);

                // Complexity flags can also sit one level down under policies.account, the same
                // place minLength falls back to.
                var accountPolicies = NcJson.GetDictionary(NcJson.GetDictionary(policy, "policies"), "account");
                bool upperLower = GetFirstBool(policy, UpperLowerKeys) || GetFirstBool(accountPolicies, UpperLowerKeys);
                bool specialChars = GetFirstBool(policy, SpecialCharKeys) || GetFirstBool(accountPolicies, SpecialCharKeys);
                bool numericChars = GetFirstBool(policy, NumericCharKeys) || GetFirstBool(accountPolicies, NumericCharKeys);

                bool hasPolicy = minLength > 0
                                 || !string.IsNullOrWhiteSpace(generateUrl)
                                 || upperLower
                                 || specialChars
                                 || numericChars;

                DiagnosticsLogger.LogApi(
                    "Password policy normalized (hasPolicy=" + hasPolicy
                    + ", minLength=" + minLength
                    + ", upperLower=" + upperLower
                    + ", special=" + specialChars
                    + ", numeric=" + numericChars
                    + ", hasGenerateUrl=" + (!string.IsNullOrWhiteSpace(generateUrl)) + ").");
                return new PasswordPolicyInfo(
                    hasPolicy,
                    minLength,
                    generateUrl ?? string.Empty,
                    upperLower,
                    specialChars,
                    numericChars);
            }
        }

        // Nextcloud serializes these as real booleans, but older payloads and some proxies hand
        // them back as "1"/"true" strings.
        private static bool GetFirstBool(IDictionary<string, object> parent, string[] keys)
        {
            if (parent == null)
            {
                return false;
            }
            for (int i = 0; i < keys.Length; i++)
            {
                object raw;
                if (!parent.TryGetValue(keys[i], out raw) || raw == null)
                {
                    continue;
                }
                if (raw is bool)
                {
                    return (bool)raw;
                }
                if (raw is int)
                {
                    return (int)raw != 0;
                }
                string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }
                bool parsedBool;
                if (bool.TryParse(text.Trim(), out parsedBool))
                {
                    return parsedBool;
                }
                int parsedInt;
                if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedInt))
                {
                    return parsedInt != 0;
                }
            }
            return false;
        }

        internal string GeneratePassword(PasswordPolicyInfo policy)
        {            if (policy == null || !policy.HasPolicy || string.IsNullOrEmpty(policy.GenerateUrl) || !_configuration.IsComplete())
            {
                return null;
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "PasswordPolicy.GeneratePassword"))
            {
                string apiUrl = ResolvePolicyUrl(policy.GenerateUrl, _configuration.GetNormalizedBaseUrl());
                if (string.IsNullOrWhiteSpace(apiUrl))
                {
                    DiagnosticsLogger.LogApi("Password generate skipped: no usable generator URL.");
                    return null;
                }

                DiagnosticsLogger.LogApi("GET " + apiUrl);
                IDictionary<string, object> root;
                HttpStatusCode statusCode;
                ExecuteJsonRequest("GET", apiUrl, null, out statusCode, out root);                if (statusCode != HttpStatusCode.OK || root == null)
                {
                    return null;
                }
                var data = NcJson.GetDictionary(NcJson.GetDictionary(root, "ocs"), "data");
                string password = NcJson.GetString(data, "password");
                if (string.IsNullOrWhiteSpace(password))
                {
                    return null;
                }
                return password.Trim();
            }
        }

        private void ExecuteJsonRequest(string method, string url, string payload, out HttpStatusCode statusCode, out IDictionary<string, object> parsed)
        {
            parsed = null;
            statusCode = 0;

            if (string.IsNullOrEmpty(url))
            {
                return;
            }
            var response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Payload = payload,
                TimeoutMs = NcTimeouts.PolicyMs,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {                if (response.TransportException != null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Api, "Password policy request failed without HTTP response.", response.TransportException);
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.Api, "Password policy request failed without HTTP response.", null);
                }
                return;
            }

            statusCode = response.StatusCode;
            DiagnosticsLogger.LogApi(method + " " + url + " -> " + statusCode);            if (response.JsonParseException != null)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Password policy response parsing failed.", response.JsonParseException);
                DiagnosticsLogger.LogApi("Password policy response sample: " + GetResponseSample(response.ResponseText));
            }

            parsed = response.ParsedJson;
        }

        private static string ResolvePolicyUrl(string rawUrl, string baseUrl)
        {
            string trimmed = rawUrl == null ? string.Empty : rawUrl.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return null;
            }
            try
            {
                Uri absoluteUri;
                if (Uri.TryCreate(trimmed, UriKind.Absolute, out absoluteUri))
                {
                    return absoluteUri.ToString();
                }

                Uri baseUri;
                if (!string.IsNullOrWhiteSpace(baseUrl) && Uri.TryCreate(baseUrl, UriKind.Absolute, out baseUri))
                {
                    Uri resolved = new Uri(baseUri, trimmed);
                    return resolved.ToString();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Password policy URL normalization failed.", ex);
            }
            return null;
        }

        private static string GetResponseSample(string responseText)
        {
            if (string.IsNullOrEmpty(responseText))
            {
                return "<empty>";
            }
            string normalized = responseText.Replace("\r", " ").Replace("\n", " ").Trim();
            return normalized.Length <= 180 ? normalized : normalized.Substring(0, 180) + "...";
        }
        private static int GetFirstPositiveInt(IDictionary<string, object> parent, params string[] keys)
        {            if (parent == null || keys == null)
            {
                return 0;
            }
            foreach (string key in keys)
            {
                int value = NcJson.GetInt(parent, key);
                if (value > 0)
                {
                    return value;
                }
            }
            return 0;
        }

        private static string GetFirstString(IDictionary<string, object> parent, params string[] keys)
        {            if (parent == null || keys == null)
            {
                return null;
            }
            foreach (string key in keys)
            {
                string value = NcJson.GetTrimmedString(parent, key);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            return null;
        }
    }
}

