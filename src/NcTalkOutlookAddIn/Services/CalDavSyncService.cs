// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Sends CalDAV PUT/DELETE requests to a Nextcloud calendar.
    internal sealed class CalDavSyncService
    {
        internal bool PutAppointment(TalkServiceConfiguration configuration, string calendarName, string uid, string icalContent)
        {
            string url = BuildIcsUrl(configuration.GetNormalizedBaseUrl(), configuration.Username, calendarName, uid);
            var client = new NcHttpClient(configuration);
            NcHttpResponse response = client.Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = url,
                Payload = icalContent,
                ContentType = "text/calendar; charset=utf-8",
                Accept = "*/*",
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 30000
            });
            if (!response.HasHttpResponse)
            {
                DiagnosticsLogger.Log(LogCategories.CalDav, "PUT failed: no HTTP response (uid=" + uid + ").");
                return false;
            }
            int code = (int)response.StatusCode;
            bool ok = code >= 200 && code < 300;
            if (!ok)
            {
                DiagnosticsLogger.Log(LogCategories.CalDav, "PUT failed: HTTP " + code + " (uid=" + uid + ").");
            }
            return ok;
        }

        internal bool DeleteAppointment(TalkServiceConfiguration configuration, string calendarName, string uid)
        {
            string url = BuildIcsUrl(configuration.GetNormalizedBaseUrl(), configuration.Username, calendarName, uid);
            var client = new NcHttpClient(configuration);
            NcHttpResponse response = client.Send(new NcHttpRequestOptions
            {
                Method = "DELETE",
                Url = url,
                Accept = "*/*",
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 30000
            });
            if (!response.HasHttpResponse)
            {
                DiagnosticsLogger.Log(LogCategories.CalDav, "DELETE failed: no HTTP response (uid=" + uid + ").");
                return false;
            }
            int code = (int)response.StatusCode;
            // 404 means already gone — treat as success
            bool ok = (code >= 200 && code < 300) || code == (int)HttpStatusCode.NotFound;
            if (!ok)
            {
                DiagnosticsLogger.Log(LogCategories.CalDav, "DELETE failed: HTTP " + code + " (uid=" + uid + ").");
            }
            return ok;
        }

        private static string BuildIcsUrl(string baseUrl, string username, string calendarName, string uid)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/calendars/{1}/{2}/{3}.ics",
                baseUrl,
                Uri.EscapeDataString(username ?? string.Empty),
                Uri.EscapeDataString(calendarName ?? "personal"),
                Uri.EscapeDataString(uid));
        }
    }
}
