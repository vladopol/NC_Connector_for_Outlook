// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Utilities
{
    // HTTP timeouts, grouped by what the caller is actually waiting for.
    //
    // These are ceilings for a *stalled* server, not budgets for a healthy one — a Nextcloud that
    // answers at all answers these calls in well under a second. The values matter only in the bad
    // case: a server that accepts the TCP connection and then goes quiet (restarting, overloaded,
    // black-holed VPN route, proxy). HttpWebRequest.Timeout covers connect and response together,
    // and the add-in issues these calls in sequence, so the ceilings add up per user action.
    //
    // Several of these calls still run on Outlook's UI thread, where the ceiling is exactly how long
    // Outlook stays frozen. Anything raised here is paid for by the user watching "Not Responding".
    internal static class NcTimeouts
    {
        // Pure connectivity/credential probe against /ocs/v2.php/cloud/capabilities. Runs on the UI
        // thread on every Talk button click and forces a fresh TLS handshake, so it must fail fast:
        // its whole purpose is to answer "is the server reachable right now?".
        internal const int ConnectivityProbeMs = 10000;

        // Optional metadata that gates a dialog opening (backend policy, password policy). Failure
        // is non-fatal — the caller falls back to local defaults — so waiting longer than this buys
        // nothing but a later dialog.
        internal const int PolicyMs = 10000;

        // Control-plane operations the user is actively waiting on, on the UI thread: room create /
        // delete / update, folder existence checks, folder creation. All are cheap server-side.
        internal const int InteractiveMs = 15000;

        // Same class of operation, but issued from a background thread where nobody is watching.
        // Still bounded so a stalled server cannot pin a thread-pool thread.
        internal const int BackgroundMs = 30000;

        // Downloading the Nextcloud user directory. Legitimately larger than a control-plane call
        // in a big organization.
        internal const int DirectoryMs = 30000;

        // Actual file bytes, and the server-side assembly of a chunked upload. These are the only
        // calls whose duration scales with user data, so they keep a generous ceiling.
        internal const int BulkTransferMs = 120000;
    }
}
