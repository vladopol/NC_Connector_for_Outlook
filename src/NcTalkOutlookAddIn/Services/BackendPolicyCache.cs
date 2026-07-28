// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Process-wide cache for the backend policy status.
    //
    // The status endpoint answers with a 45 s timeout and the policy rarely changes, yet it used to
    // be re-fetched on every call — including from Outlook event handlers that must return a
    // decision synchronously. A blocking fetch there freezes Outlook for as long as the server takes
    // to answer, which is what makes Outlook flag the add-in as slow and disable it.
    //
    // Callers that can wait use FetchOrRefresh (on a background thread). Callers on the UI thread
    // use PeekWithBackgroundRefresh, which never performs I/O.
    internal static class BackendPolicyCache
    {
        private static readonly TimeSpan FreshFor = TimeSpan.FromMinutes(10);
        private static readonly TimeSpan RetryAfterFailure = TimeSpan.FromMinutes(2);
        private static readonly object SyncRoot = new object();

        private static string _key;
        private static BackendPolicyStatus _status;
        private static DateTime _fetchedUtc = DateTime.MinValue;
        private static DateTime _failedUtc = DateTime.MinValue;
        private static bool _refreshInFlight;

        internal static string BuildKey(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                return string.Empty;
            }
            return (configuration.GetNormalizedBaseUrl() ?? string.Empty) + "|" + (configuration.Username ?? string.Empty);
        }

        // Returns the cached status when it is still fresh. Blocking fetches are the caller's job.
        internal static bool TryGetFresh(string key, out BackendPolicyStatus status)
        {
            lock (SyncRoot)
            {
                status = _status;
                return _status != null
                       && string.Equals(_key, key, StringComparison.Ordinal)
                       && (DateTime.UtcNow - _fetchedUtc) < FreshFor;
            }
        }

        // Returns whatever is cached — even stale — and never performs I/O. Intended for Outlook
        // event handlers, which must answer synchronously.
        internal static BackendPolicyStatus Peek(string key)
        {
            lock (SyncRoot)
            {
                return string.Equals(_key, key, StringComparison.Ordinal) ? _status : null;
            }
        }

        internal static void Store(string key, BackendPolicyStatus status)
        {
            lock (SyncRoot)
            {
                if (status == null)
                {
                    // Remember the failure so a dead server is not retried on every single event.
                    _failedUtc = DateTime.UtcNow;
                    _refreshInFlight = false;
                    return;
                }

                _key = key;
                _status = status;
                _fetchedUtc = DateTime.UtcNow;
                _failedUtc = DateTime.MinValue;
                _refreshInFlight = false;
            }
        }

        // Grants permission to start exactly one background refresh at a time, and only when the
        // cached value is actually due for renewal.
        internal static bool TryBeginBackgroundRefresh(string key)
        {
            lock (SyncRoot)
            {
                if (_refreshInFlight)
                {
                    return false;
                }
                DateTime now = DateTime.UtcNow;
                if (_failedUtc != DateTime.MinValue && (now - _failedUtc) < RetryAfterFailure)
                {
                    return false;
                }
                if (_status != null
                    && string.Equals(_key, key, StringComparison.Ordinal)
                    && (now - _fetchedUtc) < FreshFor)
                {
                    return false;
                }

                _refreshInFlight = true;
                return true;
            }
        }

        internal static void AbandonBackgroundRefresh()
        {
            lock (SyncRoot)
            {
                _refreshInFlight = false;
            }
        }

        internal static void Invalidate()
        {
            lock (SyncRoot)
            {
                _key = null;
                _status = null;
                _fetchedUtc = DateTime.MinValue;
                _failedUtc = DateTime.MinValue;
                _refreshInFlight = false;
                DiagnosticsLogger.Log(LogCategories.Core, "Backend policy cache invalidated.");
            }
        }

        internal static string DescribeAge(string key)
        {
            lock (SyncRoot)
            {
                if (_status == null || !string.Equals(_key, key, StringComparison.Ordinal))
                {
                    return "none";
                }
                return ((int)(DateTime.UtcNow - _fetchedUtc).TotalSeconds).ToString(CultureInfo.InvariantCulture) + "s";
            }
        }
    }
}
