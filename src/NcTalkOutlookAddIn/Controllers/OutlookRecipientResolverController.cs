// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
        // Centralized Outlook recipient resolution for attendee extraction and SMTP mapping.
    internal static class OutlookRecipientResolverController
    {
        // Resolving an Exchange recipient to its SMTP address goes through AddressEntry.GetExchangeUser()
        // or the MAPI PropertyAccessor, either of which can round-trip to the Exchange server. These
        // resolutions happen on Outlook's UI thread inside event handlers and repeat constantly for
        // the same colleagues, so the mapping (X.500 DN -> SMTP) is memoized for the session.
        private const int SmtpCacheLimit = 2048;
        private static readonly object SmtpCacheSyncRoot = new object();
        private static readonly Dictionary<string, string> SmtpCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static bool TryGetCachedSmtp(string addressKey, out string smtp)
        {
            smtp = null;
            if (string.IsNullOrWhiteSpace(addressKey))
            {
                return false;
            }
            lock (SmtpCacheSyncRoot)
            {
                return SmtpCache.TryGetValue(addressKey, out smtp);
            }
        }

        private static void CacheSmtp(string addressKey, string smtp)
        {
            if (string.IsNullOrWhiteSpace(addressKey) || string.IsNullOrWhiteSpace(smtp))
            {
                return;
            }
            lock (SmtpCacheSyncRoot)
            {
                if (SmtpCache.Count >= SmtpCacheLimit)
                {
                    SmtpCache.Clear();
                }
                SmtpCache[addressKey] = smtp;
            }
        }

        internal static List<string> CollectAppointmentAttendeeEmails(Outlook.AppointmentItem appointment)
        {
            var emails = new List<string>();            if (appointment == null)
            {
                return emails;
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment recipients.", ex);
                recipients = null;
            }
            if (recipients == null)
            {
                return emails;
            }
            try
            {
                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];                        if (recipient == null)
                        {
                            continue;
                        }
                        int type = 0;
                        try
                        {
                            type = recipient.Type;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read recipient.Type.", ex);
                            type = 0;
                        }

                        // 1=Required, 2=Optional, 3=Resource
                        if (type == 3)
                        {
                            continue;
                        }
                        string email = TryResolveRecipientSmtpAddress(recipient);
                        if (string.IsNullOrWhiteSpace(email))
                        {
                            continue;
                        }

                        email = email.Trim().ToLowerInvariant();
                        if (!emails.Contains(email))
                        {
                            emails.Add(email);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(recipient, LogCategories.Talk, "Failed to release Recipient COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to enumerate appointment recipients.", ex);
            }
            finally
            {
                ComInteropScope.TryRelease(recipients, LogCategories.Talk, "Failed to release Recipients COM object.");
            }
            return emails;
        }

        // Same resolution for an AddressEntry, which is what AppointmentItem.GetOrganizer() returns.
        // Shares the session cache, since both paths resolve the same X.500 DNs.
        internal static string TryResolveAddressEntrySmtpAddress(Outlook.AddressEntry entry)
        {
            if (entry == null)
            {
                return null;
            }

            string address = null;
            try
            {
                address = entry.Address;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read AddressEntry.Address.", ex);
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            string addressKey = address;
            string cachedSmtp;
            if (TryGetCachedSmtp(addressKey, out cachedSmtp))
            {
                return cachedSmtp;
            }

            Outlook.ExchangeUser exUser = null;
            try
            {
                exUser = entry.GetExchangeUser();
                if (exUser != null)
                {
                    address = exUser.PrimarySmtpAddress;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve Exchange user from an address entry.", ex);
            }
            finally
            {
                if (exUser != null)
                {
                    ComInteropScope.TryRelease(exUser, LogCategories.Talk, "Failed to release ExchangeUser COM object.");
                }
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                CacheSmtp(addressKey, address);
                return address;
            }
            try
            {
                Outlook.PropertyAccessor accessor = entry.PropertyAccessor;
                if (accessor != null)
                {
                    try
                    {
                        const string SmtpSchema = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        address = accessor.GetProperty(SmtpSchema) as string;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(accessor, LogCategories.Talk, "Failed to release PropertyAccessor COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve SMTP address from an address entry via PropertyAccessor.", ex);
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                CacheSmtp(addressKey, address);
                return address;
            }
            return null;
        }

        internal static string TryResolveRecipientSmtpAddress(Outlook.Recipient recipient)
        {            if (recipient == null)
            {
                return null;
            }
            string address = null;
            try
            {
                address = recipient.Address;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.Address.", ex);
                address = null;
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            // Everything below can hit the Exchange server; serve a previous resolution instead.
            string addressKey = address;
            string cachedSmtp;
            if (TryGetCachedSmtp(addressKey, out cachedSmtp))
            {
                return cachedSmtp;
            }

            Outlook.AddressEntry entry = null;
            try
            {
                entry = recipient.AddressEntry;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.AddressEntry.", ex);
                entry = null;
            }
            if (entry != null)
            {
                try
                {
                    Outlook.ExchangeUser exUser = null;
                    try
                    {
                        exUser = entry.GetExchangeUser();
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve Exchange user from address entry.", ex);
                        exUser = null;
                    }
                    if (exUser != null)
                    {
                        try
                        {
                            address = exUser.PrimarySmtpAddress;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read ExchangeUser.PrimarySmtpAddress.", ex);
                            address = null;
                        }

                        ComInteropScope.TryRelease(exUser, LogCategories.Talk, "Failed to release ExchangeUser COM object.");
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(entry, LogCategories.Talk, "Failed to release AddressEntry COM object.");
                }
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                CacheSmtp(addressKey, address);
                return address;
            }
            try
            {
                Outlook.PropertyAccessor accessor = recipient.PropertyAccessor;                if (accessor != null)
                {
                    try
                    {
                        const string SmtpSchema = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        address = accessor.GetProperty(SmtpSchema) as string;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(accessor, LogCategories.Talk, "Failed to release PropertyAccessor COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve SMTP address via PropertyAccessor.", ex);
            }
            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                CacheSmtp(addressKey, address);
                return address;
            }
            return null;
        }
    }
}

