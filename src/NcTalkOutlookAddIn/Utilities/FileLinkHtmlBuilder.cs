// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Text;
using System.Web;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Utilities
{
        // Builds the HTML block inserted into mail compose windows (download link, password, expiration, permissions).
    // The layout is intentionally kept static to remain stable when embedded in Outlook.
    internal static class FileLinkHtmlBuilder
    {
        internal static string Build(FileLinkResult result, FileLinkRequest request, string languageOverride, BackendPolicyStatus policyStatus = null)
        {
            if (result == null)
            {
                throw new ArgumentNullException("result");
            }
            bool attachmentMode = request != null && request.AttachmentMode;
            bool separatePassword = request != null
                && request.PasswordSeparateEnabled
                && !string.IsNullOrWhiteSpace(result.Password);
            string effectiveLanguage = ResolveEffectiveLanguage(languageOverride, policyStatus);

            string policyTemplate = ResolvePolicyTemplate(policyStatus, false, effectiveLanguage);
            if (!string.IsNullOrWhiteSpace(policyTemplate))
            {
                return RenderPolicyTemplate(
                    policyTemplate,
                    result,
                    request,
                    effectiveLanguage,
                    attachmentMode,
                    separatePassword,
                    passwordOnly: false);
            }
            if (string.Equals(effectiveLanguage, "custom", StringComparison.OrdinalIgnoreCase))
            {
                effectiveLanguage = "default";
            }

            string downloadLabel = Strings.GetInLanguage(effectiveLanguage, "sharing_html_download_label", "Download link");
            string passwordLabel = Strings.GetInLanguage(effectiveLanguage, "sharing_html_password_label", "Password");
            string expireLabel = Strings.GetInLanguage(effectiveLanguage, "sharing_html_expire_label", "Expiration date");
            string passwordSeparateHint = Strings.GetInLanguage(effectiveLanguage, "sharing_html_password_separate_hint", "The password will be sent in a separate email.");

            string downloadUrl = attachmentMode
                ? BuildAttachmentZipDownloadUrl(result.ShareUrl, result.ShareToken)
                : (result.ShareUrl ?? string.Empty);

            var html = new StringBuilder();
            html.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;\">");
            html.Append("<tbody>");
            html.Append("<tr><td valign=\"top\" align=\"left\">")
                .Append(HttpUtility.HtmlEncode(downloadLabel))
                .Append("<br><a href=\"")
                .Append(HttpUtility.HtmlAttributeEncode(downloadUrl))
                .Append("\">")
                .Append(HttpUtility.HtmlEncode(downloadUrl))
                .Append("</a></td></tr>");

            if (!string.IsNullOrEmpty(result.Password) && !separatePassword)
            {
                html.Append("<tr><td valign=\"top\" align=\"left\"><br>")
                    .Append(HttpUtility.HtmlEncode(passwordLabel))
                    .Append(": ")
                    .Append(HttpUtility.HtmlEncode(result.Password))
                    .Append("</td></tr>");
            }
            else if (separatePassword)
            {
                html.Append("<tr><td valign=\"top\" align=\"left\"><br>")
                    .Append(HttpUtility.HtmlEncode(passwordLabel))
                    .Append(": ")
                    .Append(HttpUtility.HtmlEncode(passwordSeparateHint))
                    .Append("</td></tr>");
            }

            if (result.ExpireDate.HasValue)
            {
                html.Append("<tr><td valign=\"top\" align=\"left\"><br>")
                    .Append(HttpUtility.HtmlEncode(expireLabel))
                    .Append(": ")
                    .Append(HttpUtility.HtmlEncode(result.ExpireDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                    .Append("</td></tr>");
            }

            html.Append("</tbody>");
            html.Append("</table>");
            return html.ToString();
        }

        internal static string BuildPasswordOnly(FileLinkResult result, string languageOverride, BackendPolicyStatus policyStatus = null)
        {
            if (result == null)
            {
                throw new ArgumentNullException("result");
            }
            string effectiveLanguage = ResolveEffectiveLanguage(languageOverride, policyStatus);
            string policyTemplate = ResolvePolicyTemplate(policyStatus, true, effectiveLanguage);
            if (!string.IsNullOrWhiteSpace(policyTemplate))
            {
                return RenderPolicyTemplate(
                    policyTemplate,
                    result,
                    request: null,
                    effectiveLanguage: effectiveLanguage,
                    attachmentMode: false,
                    separatePassword: false,
                    passwordOnly: true);
            }
            if (string.Equals(effectiveLanguage, "custom", StringComparison.OrdinalIgnoreCase))
            {
                effectiveLanguage = "default";
            }
            string passwordLabel = Strings.GetInLanguage(effectiveLanguage, "sharing_html_password_label", "Password");

            var html = new StringBuilder();
            html.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;\">");
            html.Append("<tbody>");
            html.Append("<tr><td valign=\"top\" align=\"left\">")
                .Append(HttpUtility.HtmlEncode(passwordLabel))
                .Append(": ")
                .Append(HttpUtility.HtmlEncode(result.Password ?? string.Empty))
                .Append("</td></tr>");
            html.Append("</tbody>");
            html.Append("</table>");
            return html.ToString();
        }

                // Resolve effective HTML block language with policy override support.
        private static string ResolveEffectiveLanguage(string languageOverride, BackendPolicyStatus policyStatus)
        {
            if (policyStatus != null
                && policyStatus.PolicyActive
                && policyStatus.IsLocked("share", "language_share_html_block"))
            {
                string policyLang = policyStatus.GetPolicyString("share", "language_share_html_block");
                if (!string.IsNullOrWhiteSpace(policyLang))
                {
                    return string.Equals(policyLang, "custom", StringComparison.OrdinalIgnoreCase)
                        ? "custom"
                        : Strings.NormalizeLanguageOverride(policyLang);
                }
            }
            string normalized = Strings.NormalizeLanguageOverride(languageOverride);
            if (string.Equals((languageOverride ?? string.Empty).Trim(), "custom", StringComparison.OrdinalIgnoreCase))
            {
                return "custom";
            }
            return normalized;
        }

                // Resolve custom policy template for normal or password-only mode.
        private static string ResolvePolicyTemplate(BackendPolicyStatus policyStatus, bool passwordOnly, string effectiveLanguage)
        {            if (policyStatus == null || !policyStatus.PolicyActive)
            {
                return string.Empty;
            }
            if (!string.Equals(effectiveLanguage, "custom", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }
            string key = passwordOnly ? "share_password_template" : "share_html_block_template";
            string template = policyStatus.GetPolicyString("share", key);
            return template ?? string.Empty;
        }

                // Render one backend-provided custom HTML template.
        private static string RenderPolicyTemplate(
            string template,
            FileLinkResult result,
            FileLinkRequest request,
            string effectiveLanguage,
            bool attachmentMode,
            bool separatePassword,
            bool passwordOnly)
        {            if (string.IsNullOrWhiteSpace(template) || result == null)
            {
                return string.Empty;
            }
            string passwordSeparateHint = Strings.GetInLanguage(effectiveLanguage, "sharing_html_password_separate_hint", "The password will be sent in a separate email.");
            string permissionRead = Strings.GetInLanguage(effectiveLanguage, "sharing_permission_read", "Read");
            string permissionCreate = Strings.GetInLanguage(effectiveLanguage, "sharing_permission_create", "Upload");
            string permissionWrite = Strings.GetInLanguage(effectiveLanguage, "sharing_permission_write", "Modify");
            string permissionDelete = Strings.GetInLanguage(effectiveLanguage, "sharing_permission_delete", "Delete");

            string downloadUrl = attachmentMode
                ? BuildAttachmentZipDownloadUrl(result.ShareUrl, result.ShareToken)
                : (result.ShareUrl ?? string.Empty);

            string passwordValue = result.Password ?? string.Empty;
            if (!passwordOnly && separatePassword)
            {
                passwordValue = passwordSeparateHint;
            }
            string noteValue = string.Empty;            if (request != null && request.NoteEnabled && !string.IsNullOrWhiteSpace(request.Note))
            {
                noteValue = request.Note.Trim();
            }
            string rightsValue = attachmentMode
                ? string.Empty
                : BuildPermissions(result.Permissions, permissionRead, permissionCreate, permissionWrite, permissionDelete);

            string expireValue = result.ExpireDate.HasValue
                ? result.ExpireDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : string.Empty;

            string html = template;
            if (attachmentMode)
            {
                html = StripTemplateRow(html, "RIGHTS");
            }
            html = html.Replace("{URL}", HttpUtility.HtmlEncode(downloadUrl ?? string.Empty));
            html = html.Replace("{PASSWORD}", HttpUtility.HtmlEncode(passwordValue));
            html = html.Replace("{EXPIRATIONDATE}", HttpUtility.HtmlEncode(expireValue));
            html = html.Replace("{RIGHTS}", rightsValue);
            html = html.Replace("{NOTE}", HttpUtility.HtmlEncode(noteValue));

            string sanitized = HtmlTemplateSanitizer.SanitizeShareTemplateHtml(html);
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                throw new InvalidOperationException("Share HTML template sanitized to empty output.");
            }
            return sanitized;
        }

                // Remove one placeholder row from backend-provided HTML templates.
        // This is used to reduce the custom share block for attachment mode.
        private static string StripTemplateRow(string template, string placeholder)
        {
            string token = "{" + (placeholder ?? string.Empty).Trim() + "}";
            if (string.IsNullOrWhiteSpace(token) || string.Equals(token, "{}", StringComparison.Ordinal))
            {
                return template ?? string.Empty;
            }
            string output = template ?? string.Empty;
            int tokenIndex = output.IndexOf(token, StringComparison.Ordinal);
            if (tokenIndex < 0)
            {
                return output;
            }
            int rowStart = LastIndexOfIgnoreCase(output, "<tr", tokenIndex);
            int rowEnd = IndexOfIgnoreCase(output, "</tr>", tokenIndex);
            if (rowStart >= 0 && rowEnd >= 0 && rowEnd >= rowStart)
            {
                output = output.Remove(rowStart, (rowEnd + 5) - rowStart);
            }
            return output.Replace(token, string.Empty);
        }

                // Case-insensitive search for the last occurrence before one absolute index.
        private static int LastIndexOfIgnoreCase(string value, string search, int startIndexExclusive)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(search))
            {
                return -1;
            }
            int maxIndex = Math.Min(startIndexExclusive, value.Length);
            if (maxIndex <= 0)
            {
                return -1;
            }
            return value.LastIndexOf(search, maxIndex - 1, StringComparison.OrdinalIgnoreCase);
        }

                // Case-insensitive forward search.
        private static int IndexOfIgnoreCase(string value, string search, int startIndex)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(search))
            {
                return -1;
            }
            int normalizedStart = Math.Max(0, startIndex);
            if (normalizedStart >= value.Length)
            {
                return -1;
            }
            return value.IndexOf(search, normalizedStart, StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildAttachmentZipDownloadUrl(string shareUrl, string shareToken)
        {
            if (string.IsNullOrWhiteSpace(shareUrl))
            {
                return string.Empty;
            }
            string token = string.IsNullOrWhiteSpace(shareToken) ? string.Empty : shareToken.Trim();
            try
            {
                var shareUri = new Uri(shareUrl, UriKind.Absolute);
                string[] segments = shareUri.AbsolutePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                int shareSegmentIndex = Array.FindLastIndex(
                    segments,
                    segment => string.Equals(segment, "s", StringComparison.OrdinalIgnoreCase));
                if (shareSegmentIndex >= 0 && shareSegmentIndex + 1 < segments.Length)
                {
                    token = segments[shareSegmentIndex + 1];
                }
                if (!string.IsNullOrWhiteSpace(token))
                {
                    return shareUri.GetLeftPart(UriPartial.Authority).TrimEnd('/')
                        + "/s/"
                        + Uri.EscapeDataString(token)
                        + "/download";
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to derive attachment-mode ZIP download URL from share URL.", ex);
            }
            if (string.IsNullOrEmpty(token))
            {
                DiagnosticsLogger.Log(LogCategories.FileLink, "Attachment-mode ZIP download URL fallback used public share URL because no token was available.");
                return shareUrl;
            }
            try
            {
                var shareUri = new Uri(shareUrl, UriKind.Absolute);
                return shareUri.GetLeftPart(UriPartial.Authority).TrimEnd('/')
                    + "/s/"
                    + Uri.EscapeDataString(token)
                    + "/download";
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to build attachment-mode ZIP download URL from token fallback.", ex);
                return shareUrl;
            }
        }

                // Renders the permissions badges (checkmarks / red crosses).
        private static string BuildPermissions(FileLinkPermissionFlags permissions, string readLabel, string createLabel, string writeLabel, string deleteLabel)
        {
            var builder = new StringBuilder();
            builder.Append("<table style=\"border-collapse:collapse;\">");
            builder.Append("<tr>");
            AppendPermissionCell(builder, readLabel, (permissions & FileLinkPermissionFlags.Read) == FileLinkPermissionFlags.Read);
            AppendPermissionCell(builder, createLabel, (permissions & FileLinkPermissionFlags.Create) == FileLinkPermissionFlags.Create);
            AppendPermissionCell(builder, writeLabel, (permissions & FileLinkPermissionFlags.Write) == FileLinkPermissionFlags.Write);
            AppendPermissionCell(builder, deleteLabel, (permissions & FileLinkPermissionFlags.Delete) == FileLinkPermissionFlags.Delete);
            builder.Append("</tr>");
            builder.Append("</table>");
            return builder.ToString();
        }

        private static void AppendPermissionCell(StringBuilder builder, string label, bool enabled)
        {
            builder.Append("<td style=\"padding:0 18px 6px 0;\">");
            builder.Append("<span style=\"display:inline-flex;align-items:center;\">");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<span style=\"display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border:1px solid {0};color:{0};font-size:13px;font-weight:700;\">{1}</span>",
                enabled ? BrandingAssets.BrandBlueHex : "#c62828",
                enabled ? "&#10003;" : "&#10007;");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<span style=\"padding-left:6px;font-weight:600;\">{0}</span>",
                HttpUtility.HtmlEncode(label));
            builder.Append("</span>");
            builder.Append("</td>");
        }
    }
}

