// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Text;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    // Builds a minimal RFC 5545 iCalendar payload from an Outlook AppointmentItem.
    internal static class ICalBuilder
    {
        internal static string Build(Outlook.AppointmentItem appointment, string uid)
        {
            string subject = SafeRead(() => appointment.Subject) ?? string.Empty;
            string body = SafeRead(() => appointment.Body) ?? string.Empty;
            string location = SafeRead(() => appointment.Location) ?? string.Empty;
            DateTime start = SafeRead(() => appointment.Start);
            DateTime end = SafeRead(() => appointment.End);
            bool allDay = SafeRead(() => appointment.AllDayEvent);

            var sb = new StringBuilder();
            sb.Append("BEGIN:VCALENDAR\r\n");
            sb.Append("VERSION:2.0\r\n");
            sb.Append("PRODID:-//NC Connector for Outlook//Calendar Sync//EN\r\n");
            sb.Append("BEGIN:VEVENT\r\n");
            AppendFolded(sb, "UID", uid);
            AppendFolded(sb, "DTSTAMP", FormatUtcDateTime(DateTime.UtcNow));

            if (allDay)
            {
                AppendFolded(sb, "DTSTART;VALUE=DATE", start.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
                AppendFolded(sb, "DTEND;VALUE=DATE", end.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
            }
            else
            {
                AppendFolded(sb, "DTSTART", FormatUtcDateTime(start.ToUniversalTime()));
                AppendFolded(sb, "DTEND", FormatUtcDateTime(end.ToUniversalTime()));
            }

            if (!string.IsNullOrEmpty(subject))
                AppendFolded(sb, "SUMMARY", EscapeText(subject));
            if (!string.IsNullOrEmpty(location))
                AppendFolded(sb, "LOCATION", EscapeText(location));
            if (!string.IsNullOrWhiteSpace(body))
                AppendFolded(sb, "DESCRIPTION", EscapeText(body.Trim()));

            sb.Append("STATUS:CONFIRMED\r\n");
            sb.Append("END:VEVENT\r\n");
            sb.Append("END:VCALENDAR\r\n");
            return sb.ToString();
        }

        private static void AppendFolded(StringBuilder sb, string name, string value)
        {
            string line = name + ":" + value;
            if (line.Length <= 75)
            {
                sb.Append(line);
                sb.Append("\r\n");
                return;
            }
            sb.Append(line.Substring(0, 75));
            sb.Append("\r\n");
            int pos = 75;
            while (pos < line.Length)
            {
                int len = Math.Min(74, line.Length - pos);
                sb.Append(' ');
                sb.Append(line.Substring(pos, len));
                sb.Append("\r\n");
                pos += len;
            }
        }

        private static string FormatUtcDateTime(DateTime dt)
        {
            return dt.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
        }

        private static string EscapeText(string text)
        {
            return text
                .Replace("\\", "\\\\")
                .Replace("\r\n", "\\n")
                .Replace("\n", "\\n")
                .Replace("\r", "\\n")
                .Replace(",", "\\,")
                .Replace(";", "\\;");
        }

        private static T SafeRead<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to read appointment field for iCal.", ex);
                return default(T);
            }
        }
    }
}
