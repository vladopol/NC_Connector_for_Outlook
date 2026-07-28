// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using NcTalkOutlookAddIn.Controllers;
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
            DateTime lastModified = SafeRead(() => appointment.LastModificationTime);
            bool allDay = SafeRead(() => appointment.AllDayEvent);
            bool cancelled = SafeRead(() => appointment.MeetingStatus) == Outlook.OlMeetingStatus.olMeetingCanceled;

            var sb = new StringBuilder();
            sb.Append("BEGIN:VCALENDAR\r\n");
            sb.Append("VERSION:2.0\r\n");
            sb.Append("PRODID:-//NC Connector for Outlook//Calendar Sync//EN\r\n");
            sb.Append("BEGIN:VEVENT\r\n");
            AppendFolded(sb, "UID", uid);
            AppendFolded(sb, "DTSTAMP", FormatUtcDateTime(DateTime.UtcNow));
            if (lastModified != DateTime.MinValue)
            {
                AppendFolded(sb, "LAST-MODIFIED", FormatUtcDateTime(lastModified.ToUniversalTime()));
            }

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

            AppendParticipants(sb, appointment);

            sb.Append(cancelled ? "STATUS:CANCELLED\r\n" : "STATUS:CONFIRMED\r\n");
            sb.Append("END:VEVENT\r\n");
            sb.Append("END:VCALENDAR\r\n");
            return sb.ToString();
        }

        // Writes ORGANIZER/ATTENDEE so the Nextcloud calendar entry shows who is invited and picks
        // up additions and removals made in Outlook.
        //
        // Every scheduling property carries SCHEDULE-AGENT=CLIENT (RFC 6638): Outlook already sends
        // the meeting invitations, and without this Nextcloud's scheduling plugin would treat the
        // PUT as a new scheduling request and mail every attendee a second, duplicate invitation.
        private static void AppendParticipants(StringBuilder sb, Outlook.AppointmentItem appointment)
        {
            string organizerEmail = null;
            string organizerName = SafeRead(() => appointment.Organizer);

            Outlook.Recipient organizer = null;
            try
            {
                organizer = appointment.GetOrganizer();
                if (organizer != null)
                {
                    organizerEmail = OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(organizer);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to resolve the appointment organizer for iCal.", ex);
            }
            finally
            {
                if (organizer != null)
                    ComInteropScope.TryRelease(organizer, LogCategories.CalDav, "Failed to release organizer Recipient COM object.");
            }

            if (!string.IsNullOrWhiteSpace(organizerEmail))
            {
                AppendFolded(sb, BuildCalendarUserProperty("ORGANIZER", organizerName, null, null), "mailto:" + organizerEmail.Trim());
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                if (recipients == null)
                    return;

                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        if (recipient == null)
                            continue;

                        int type = SafeRead(() => recipient.Type);
                        // 0 = organizer (already written above), 3 = resource (rooms/equipment).
                        if (type == 0 || type == 3)
                            continue;

                        string email = OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient);
                        if (string.IsNullOrWhiteSpace(email))
                            continue;

                        email = email.Trim();
                        if (!string.IsNullOrWhiteSpace(organizerEmail)
                            && string.Equals(email, organizerEmail.Trim(), StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (!seen.Add(email))
                            continue;

                        string name = SafeRead(() => recipient.Name);
                        string role = type == 2 ? "OPT-PARTICIPANT" : "REQ-PARTICIPANT";
                        string partStat = MapResponseStatus(SafeRead(() => recipient.MeetingResponseStatus));
                        AppendFolded(sb, BuildCalendarUserProperty("ATTENDEE", name, role, partStat), "mailto:" + email);
                    }
                    finally
                    {
                        if (recipient != null)
                            ComInteropScope.TryRelease(recipient, LogCategories.CalDav, "Failed to release Recipient COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.CalDav, "Failed to enumerate attendees for iCal.", ex);
            }
            finally
            {
                if (recipients != null)
                    ComInteropScope.TryRelease(recipients, LogCategories.CalDav, "Failed to release Recipients COM object.");
            }
        }

        private static string BuildCalendarUserProperty(string name, string commonName, string role, string partStat)
        {
            var builder = new StringBuilder(name);
            builder.Append(";SCHEDULE-AGENT=CLIENT");
            if (!string.IsNullOrWhiteSpace(role))
            {
                builder.Append(";ROLE=").Append(role);
            }
            if (!string.IsNullOrWhiteSpace(partStat))
            {
                builder.Append(";PARTSTAT=").Append(partStat);
            }
            if (!string.IsNullOrWhiteSpace(commonName))
            {
                builder.Append(";CN=").Append(EscapeParameterValue(commonName));
            }
            return builder.ToString();
        }

        private static string MapResponseStatus(Outlook.OlResponseStatus status)
        {
            switch (status)
            {
                case Outlook.OlResponseStatus.olResponseAccepted:
                    return "ACCEPTED";
                case Outlook.OlResponseStatus.olResponseTentative:
                    return "TENTATIVE";
                case Outlook.OlResponseStatus.olResponseDeclined:
                    return "DECLINED";
                default:
                    return "NEEDS-ACTION";
            }
        }

        // Parameter values are quoted when they contain characters that would otherwise terminate
        // the parameter; the quote character itself has no escape in RFC 5545 and is dropped.
        private static string EscapeParameterValue(string value)
        {
            string sanitized = value
                .Replace("\"", string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            if (sanitized.Length == 0)
            {
                return "\"\"";
            }
            return "\"" + sanitized + "\"";
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
