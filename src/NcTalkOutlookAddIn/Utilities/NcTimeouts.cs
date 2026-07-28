// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Utilities
{
    // HTTP timeouts, grouped by what the caller is actually waiting for.
    //
    // These are ceilings for a *stalled* server, not budgets for a healthy one. A Nextcloud that is
    // working at all answers every control-plane call here in well under a second; the ceiling is
    // only ever reached when the server accepts the TCP connection and then goes quiet (restarting,
    // overloaded, black-holed VPN route, proxy). They are therefore sized as "how long is it still
    // plausible that an answer is coming", not "how long might this legitimately take".
    //
    // No call governed by these values runs on Outlook's UI thread any more, so a ceiling no longer
    // means frozen Outlook — it means how long the user watches a wait cursor before getting a
    // usable error. That is what keeps them short.
    //
    // HttpWebRequest.Timeout covers connect and response together. It does NOT cover streaming a
    // request body: that is governed by ReadWriteTimeout, a per-operation idle timeout.
    internal static class NcTimeouts
    {
        // Pure connectivity/credential probe against /ocs/v2.php/cloud/capabilities, issued on every
        // Talk button click and forcing a fresh TLS handshake. Its entire job is to answer "is the
        // server reachable right now?", so a slow answer is already a failed answer.
        internal const int ConnectivityProbeMs = 5000;

        // Optional metadata that gates a dialog opening (backend policy, password policy). Failure
        // is non-fatal — the caller falls back to local defaults — so waiting longer only delays the
        // dialog for no benefit.
        internal const int PolicyMs = 5000;

        // Control-plane operations the user is actively waiting on: Talk room create/delete/update,
        // WebDAV folder existence checks, folder creation. Simple database or filesystem work
        // server-side; this is already generous for a healthy server under load.
        internal const int InteractiveMs = 8000;

        // The same class of operation issued from a background thread, where nobody is watching.
        // A little longer to tolerate a busy server, still bounded so a stall cannot pin a
        // thread-pool thread indefinitely.
        internal const int BackgroundMs = 15000;

        // Downloading the Nextcloud user directory — a genuinely larger response in a big
        // organization, and cached afterwards, so it gets more room than a control-plane call.
        internal const int DirectoryMs = 20000;

        // Waiting for the response to a file transfer, and for the server-side assembly of a
        // chunked upload into its final file. This is the one ceiling that must stay large: the
        // work behind it scales with the size of the user's attachment, and assembling a
        // multi-gigabyte file legitimately takes minutes. Shortening this makes nothing faster —
        // it only breaks large uploads.
        internal const int BulkTransferMs = 120000;
    }
}
