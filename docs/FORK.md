# Fork Guide — NC Connector for Outlook (Infosuite deployment)

This document describes how this fork differs from
[nc-connector/NC_Connector_for_Outlook](https://github.com/nc-connector/NC_Connector_for_Outlook)
and why. It is the reference for anyone maintaining, merging upstream into, or reviewing this
codebase.

`README.md` and `docs/DEVELOPMENT.md` describe the product as upstream ships it. Where this file
disagrees with them, this file is correct for this build.

## Contents

- [Relationship to upstream](#relationship-to-upstream)
- [Versioning](#versioning)
- [Deliberate differences from upstream](#deliberate-differences-from-upstream)
- [Appointment to Talk room synchronization](#appointment-to-talk-room-synchronization)
- [CalDAV calendar synchronization](#caldav-calendar-synchronization)
- [UI-thread budget](#ui-thread-budget)
- [Diagnosing a slow Outlook on a restricted terminal server](#diagnosing-a-slow-outlook-on-a-restricted-terminal-server)
- [Invariants](#invariants)
- [Settings force-applied at startup](#settings-force-applied-at-startup)
- [Upstream bugs fixed here](#upstream-bugs-fixed-here)
- [Verification status](#verification-status)

## Relationship to upstream

The fork targets one environment: an Infosuite corporate deployment with Exchange and Nextcloud.
That focus is the justification for most differences below — features that require infrastructure
this deployment does not have are removed rather than hidden, and options that always have the same
answer here are fixed rather than presented.

Remotes:

```
origin    git@github.com:vladopol/NC_Connector_for_Outlook.git
upstream  https://github.com/nc-connector/NC_Connector_for_Outlook.git
```

Merge strategy: create an `upstream-X.Y.Z` branch, resolve, PR into `main`. Files that reliably
conflict: `SettingsStorage.cs` (added fields), `SettingsForm.cs` (removed controls),
`TalkLinkForm.cs`, `TalkRibbonController.cs`, `TalkAppointmentController.cs`, `AssemblyInfo.cs`.

## Versioning

`<upstream_major>.<upstream_minor>.<upstream_patch>.<fork_patch>`

Upstream `3.1.0` → fork `3.1.0.1`, `3.1.0.2`, … When upstream releases `3.1.1`, the fork targets
`3.1.1.1`. The git tag and `AssemblyVersion` always match.

Windows Installer supports only three version fields, so `installer/Product.wxs` receives a
`ProductVersion` where the fork patch is folded into the build field
(`upstream_patch * 100 + fork_patch`). `3.1.0.20` therefore installs as `3.1.20`, and versions stay
monotonically ordered in Add/Remove Programs.

## Deliberate differences from upstream

| Area | State in this fork | Reason |
|---|---|---|
| IFB (Free/Busy local HTTP listener) | **Removed entirely.** `FreeBusyServer.cs` deleted; `IfbEnabled` / `IfbPort` / `IfbDays` settings gone | Exchange serves free/busy natively, and the listener was force-disabled every session anyway. **`IfbAddressBookCache` and `IfbCacheHours` are kept** — they are unrelated to the listener and are actively used by Talk and FileLink for the Nextcloud user lookup. `FreeBusyManager` now holds only a one-time registry restore, in case an older build ever hijacked Outlook's native free/busy paths |
| Email signature | **Removed entirely** — files, settings fields, hidden tab | Requires a server-side Nextcloud component never installed here. The tab was already hidden upstream, but the runtime still performed an uncached synchronous network policy check on every compose, reply and forward — the root cause of a Reading Pane rendering bug |
| Room type selector | Hidden; always `EventConversation` | Outlook meetings map to Event Conversations. Group conversations are created in Talk directly |
| "Add guests" | Hidden; always `false` | External participants join by link and PIN, and Outlook already delivers the invitations. The guest-invite API only adds a second, redundant email |
| Talk room deletion on event delete | Force-enabled, checkbox hidden | Always the correct behaviour for the organizer |
| Talk room password | **Policy-aware PIN**, offered as a per-room action, default off | See [Room password](#room-password) |
| Talk room title field | **Removed from the dialog** | See [Room name](#room-name) |
| Moderators | **Additions**, chosen from the meeting's attendees | See [Moderators](#moderators) |
| Share folder name | Derived from the mail subject | See [Share name](#share-name) |

### Room password

A Talk room password is a PIN for a meeting room, not an account credential: guests read it aloud
or type it on a phone keypad. Generation therefore produces digits — but Nextcloud validates Talk
conversation passwords against the same `password_policy` it applies to account and share
passwords, and the app has **no per-app exemption**. On a server enforcing complexity a digits-only
PIN is refused outright.

`PasswordPolicyService` reads `minLength` **and** the complexity flags
(`enforceUpperLowerCase`, `enforceSpecialCharacters`, `enforceNumericCharacters`), accepting the
singular and snake_case spellings and the `policies.account` nesting — a key-name mismatch must not
read back as "no complexity required".

Generation keeps the digit run as the body and appends only what the policy demands, in a fixed
position: `5729481!Kp`. Ambiguous glyphs (`I`, `l`, `O`, `0`) are excluded, and the symbol pool is
limited to `!` and `$` — unambiguous to dictate, present on every keyboard layout, and inside the
set Nextcloud counts as special. `PasswordGenerationHelper.DescribePolicyViolation` applies the same
rules to a hand-typed password, so the dialog reports the real reason instead of letting room
creation fail against the server.

The password is offered as a button that toggles caption with state ("Add password" / "Remove
password") rather than a permanently displayed checkbox: it is the exception rather than the rule,
and it has no meaningful "off" state worth showing. The corresponding Settings default was removed
and is pinned off; the `talk_set_password` backend policy can still require one, in which case the
toggle is disabled rather than hidden so its tooltip still explains why.

### Room name

The dialog's "Title" field was removed. It defaulted to the appointment subject and, on OK, wrote
itself **back** into `appointment.Subject`. Outlook does not commit a freshly typed subject to the
item until the field loses focus, so pressing the Talk button right after typing handed the dialog
an empty subject; it fell back to its placeholder, and OK replaced the subject the user had just
typed.

The meeting subject is now the single source of the room name, and `TalkRoomSyncService` keeps them
aligned on every save. `ApplyRoomToAppointment` no longer writes `appointment.Subject`. The
`talk_title` backend policy is no longer applied — nothing is left for it to control, and a pinned
title would lose to the same sync. Two owners for one value is worse than none.

### Moderators

The single "Moderator (optional)" field was not what its name suggested: it performed a **handover**.
It promoted the chosen user and then called `LeaveRoom` for the organizer, who was dropped out of
their own meeting. It also wrote `X-NCTALK-DELEGATED`, after which the add-in stopped managing the
appointment entirely — no participant sync, no time or lobby updates, no room deletion.

Moderators are now **additions**: several may be selected, the organizer keeps ownership and stays
in the room, and appointment synchronization keeps working. Candidates are the meeting's own
attendees mapped to Nextcloud accounts, not the whole user directory — the addresses are read on the
UI thread and mapped on a background thread, since the lookup can refresh the address-book cache.
With the list bounded by the invitees, the autocomplete field, dropdown and avatar loading were
removed in favour of a checked list.

The organizer is filtered out of the candidates: they create the room and own it, so they already
are a moderator — the hint says exactly that, and listing them to tick implied the opposite. The
filter matches on the Nextcloud account creating the room, which is the account that becomes owner.

Empty states are reported by cause, because the remedies differ: no attendees invited yet (the
common case when the room is created first), no system address book, or attendees with no Nextcloud
account.

**Migration:** appointments delegated by builds up to 3.1.0.15 keep their old behaviour. Nothing
writes `X-NCTALK-DELEGATED` any more, but `IsDelegatedToOtherUser` is deliberately kept — dropping
it would make the add-in start managing rooms whose organizer handed them over and left.

### Share name

`AddinSettings` seeded `SharingDefaultShareName` with `Strings.SharingDefaultShareNameLabel` — the
settings field's own caption. The value was therefore never empty, the wizard's whitespace check
never fired, its localized fallback was dead code, and the caption flowed into
`BuildShareFolderName`. Every share created with defaults landed on Nextcloud as
`20260729_Share name`.

The default share name now resolves, most specific source first:

1. an administrator's `share_name_template`, or a value deliberately typed in Settings,
2. the subject of the mail the wizard was launched from,
3. the localized fallback.

Subjects are sanitized as a path component and capped at 60 characters, truncated on a word boundary
where one is available. Attachment-mode shares use the same subject-derived base, falling back to
`email_attachment`, and keep their suffix probing for collisions. `RE:` / `FW:` prefixes are **not**
stripped — doing so correctly requires a per-language prefix list, which misfires on mixed-language
correspondence.

## Appointment to Talk room synchronization

Changes to a Talk-linked appointment are reconciled by `TalkRoomSyncService`, which operates on a
COM-free `TalkRoomSyncSnapshot` and therefore runs on a thread-pool thread. Two triggers feed it:

| Trigger | Path | Notes |
|---|---|---|
| Meeting saved in its own window | `AppointmentSubscription.OnWrite` | Builds the snapshot and persists the attendee baseline **during** the Write event, so it rides along with the in-flight save |
| Any other edit — calendar-grid drag, peek pane, scheduling view | `NextcloudTalkAddIn.CalendarWatcher.cs` → `Items.ItemChange` | Skips items with an open inspector (`_subscriptionByEntryId`); **never writes to the item**, so meetings are not silently marked as needing a resend |

What propagates: subject → room name, body → room description, start **and** end → lobby timer and
event-object binding, attendee additions and removals → room participants.

Removals are **baseline-scoped**. `X-NCTALK-ATTENDEES` stores the SMTP addresses seen at the last
sync; only actors derived from that list are ever removed, so participants invited directly in Talk
survive an Outlook-side edit. Owners and moderators are never removed.

The watcher's attendee baseline is in memory only (`_calendarWatchStates`), because persisting it
would require a `Save()`. It can drift one edit from the stored property, which is harmless: a
repeat add answers 409 and a repeat removal answers 404, and both are treated as success.

For event conversations, name and description updates are attempted and a server rejection is cached
per room for the session. These are never surfaced as dialogs — the sync runs on every save.

## CalDAV calendar synchronization

Outlook → Nextcloud only. Enabled per user; calendar name defaults to `personal`. Fields
`CalDavSyncEnabled` and `CalDavCalendarName` live in `AddinSettings` and `SettingsStorage`.

`ICalBuilder` emits `ORGANIZER` and `ATTENDEE` (with role and response status), `LAST-MODIFIED`, and
`STATUS:CANCELLED` for cancelled meetings.

**Every scheduling property carries `SCHEDULE-AGENT=CLIENT` (RFC 6638).** Removing that parameter
would make Nextcloud's scheduling plugin treat the PUT as a new scheduling request and mail every
attendee a second, duplicate invitation on top of the ones Outlook already sends.

On appointment delete, the CalDAV DELETE is explicitly queued in `QueueSavedEventRoomDeletion` —
`CalDavDeleteTracker` is volatile and lost on Outlook restart, so it cannot be relied on.

## UI-thread budget

Outlook times add-in startup and event-handler responsiveness and **auto-disables** add-ins that
exceed its thresholds (`HKCU\...\Outlook\Resiliency\DisabledItems`). It measures startup, shutdown,
folder switch and item open — not arbitrary button clicks — but a frozen click still gets an add-in
disabled by hand, or lands it in `CrashingAddinList` if the user kills Outlook while it hangs.

No add-in code path blocks the UI thread on the network. The Talk ribbon flow and the FileLink
wizard navigation are fully asynchronous
(`NavigateAsync` → `ValidateCurrentStepAsync` → `EnsureShareFolderAvailableAsync`).

`NcTimeouts` names each HTTP ceiling by what the caller is waiting for. These are ceilings for a
*stalled* server, not budgets for a healthy one — a working Nextcloud answers every control-plane
call in well under a second:

| Constant | Value | Applies to |
|---|---|---|
| `ConnectivityProbeMs` | 5 s | `VerifyConnection` on every Talk click, forcing a fresh TLS handshake |
| `PolicyMs` | 5 s | Backend and password policy fetches that gate a dialog and fall back to local defaults |
| `InteractiveMs` | 8 s | Room create/delete/update, WebDAV folder checks, folder creation |
| `BackgroundMs` | 15 s | Share creation and metadata, chunk folder setup and cleanup, CalDAV sync |
| `DirectoryMs` | 20 s | Nextcloud user directory download |
| `BulkTransferMs` | 120 s | File transfer responses and server-side assembly of a chunked upload |

`BulkTransferMs` stays large deliberately: it is the only ceiling whose work scales with the size of
the user's attachment. Shortening it makes nothing faster, it only breaks large uploads. The Login
Flow v2 poll keeps its own 60 s and is annotated — that wait is the user authenticating in a
browser, not a stalled server.

## Diagnosing a slow Outlook on a restricted terminal server

The add-in performs **no network requests at startup**: settings file, registry and MAPI only. Its
only potentially expensive startup step is opening the default calendar folder — twice, once for the
Talk change watcher and once for the CalDAV sync — which on an online-mode Exchange profile is a
round-trip per open. That runs from a one-shot timer 5 s after connection and logs its own timings.

So when Outlook launches slowly, or the first click on a mail stalls for tens of seconds, the add-in
is usually not the thing doing the waiting. **This was investigated in the target environment and
confirmed environmental**: the stall reproduced identically with the add-in disabled, happens on the
first clicks after every launch, and clears on its own. Do not re-investigate it as an add-in
problem. Attribute it before changing anything:

1. **Disable the add-in** (File → Options → Add-ins → COM Add-ins) and repeat the action. If the
   stall persists, it is not this add-in. This is the only decisive test.
2. **Read Outlook's own measurements** at
   `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinPerformanceMetrics` — Outlook
   records per-add-in load and response times there itself.
3. **Enable debug logging** and compare timestamps around the action. `Deferred startup wiring
   completed (...)` reports how long each startup step actually took. Note the logger appends to
   the file synchronously per line, so it is a diagnostic mode, not a permanent one.

Environmental causes worth checking, roughly in order of how often they explain a stall of tens of
seconds that only happens on the first action:

- **Certificate revocation checks (CRL / OCSP).** Validating a certificate fetches a revocation list.
  If outbound access is *dropped* rather than refused, each check waits out its timeout — and the
  result is cached afterwards, which is exactly why the first action is slow and later ones are not.
  Test by turning off "Check for publisher's certificate revocation" and "Check for server
  certificate revocation" in Internet Options → Advanced. (This add-in's MSI and assembly are
  unsigned, so they do not trigger an Authenticode check themselves.)
- **WPAD proxy auto-detection.** "Automatically detect settings" with no WPAD host makes every new
  connection wait for discovery.
- **Online-mode Exchange profile.** Without a cached OST, opening a mail fetches its body over the
  network — the first click after launch competes with Outlook's own startup traffic.
- **Outlook's own startup network work:** AutoDiscover, OAB download, licensing, Connected
  Experiences. On a restricted network these can each hang.

## Invariants

Break these and the symptoms are slow, intermittent and hard to attribute. Each exists because it
was violated once.

**Never subscribe to `ExplorerEvents_10_Event.InlineResponse.** Merely hooking it breaks Outlook's
rendering of the Reading Pane inline-reply command bar (Send / Discard / PopOut) on some builds
(confirmed on 16.0.0.14334 / 17932.20842) — regardless of what the handler does. Proven by bisection
across 3.1.0.2–3.1.0.7: neither the ribbon XML, nor `Inspectors.NewInspector`, nor the handler body
was implicated. Upstream still carries the hook as of 3.2.1. `Inspectors.NewInspector` is fine.

**Never do network I/O in a COM event handler.** Handlers that must answer synchronously
(`BeforeAttachmentAdd`, `Write`, `ItemChange`) may only read caches and local settings.
`BeforeAttachmentAdd` once blocked on an uncached policy request with a 45 s timeout, freezing
Outlook on every attachment added to any mail.

**Never call `FetchBackendPolicyStatus` from the UI thread.** Use `PeekBackendPolicyStatus`, which is
cache-only and refreshes in the background. `FetchBackendPolicyStatus` is cached
(`BackendPolicyCache`, 10 min TTL, 2 min back-off after failure, one in-flight refresh) but still
blocks on a miss.

**Keep `OnConnection` minimal.** Anything touching `Session.GetDefaultFolder`, the registry or the
network belongs in `RunStartupWiring`, posted to the UI thread and run after load completes.

**Resolving recipients to SMTP can reach Exchange.** `AddressEntry.GetExchangeUser()` and the MAPI
`PropertyAccessor` round-trip to the server. `OutlookRecipientResolverController` memoizes
X.500 DN → SMTP for the session; keep that cache.

**Gate `ItemChange` work behind a cheap probe.** `BuildChangeProbe` reads only locally cached fields —
no SMTP resolution — so the expensive snapshot is built only when something relevant changed.

**The deferred post-write pass must stay read-only** (`TryReadAppointmentObjectId`). It runs after
Outlook has saved the item; writing user properties there leaves the appointment dirty and produces
a save prompt.

**Across an `await` the message loop keeps pumping.** Anything awaiting a server call must lock its
controls first (`DisableNavigationDuringServerCall`, restored via `UpdateNavigationState`) and check
`IsDisposed` in the continuation before touching them.

**Never do network work in a form constructor.** The attachment-mode share name used to probe WebDAV
there, so the wizard window did not appear at all until the server answered. It runs from `OnShown`.

**Never write a raw `TimeoutMs` number** — use `NcTimeouts`.

**Track appointment timing as `start#end`** (`X-NCTALK-OBJECTID`), not just the start, or a changed
end time is invisible.

**`TalkService.RemoveParticipant` passes `attendeeId` in the query string** — `NcHttpClient` drops
the body on DELETE requests.

**In WinForms the last `Controls.Add` wins.** Adding a control to a container and then to the form
re-parents it: it keeps being positioned by the container's layout code but is drawn in form
coordinates. This put the moderator list at the top of the dialog while its group box sat empty.

**Layout must not depend on when it runs.** Give every control bounds in all of its states — a
control skipped while hidden keeps WinForms' default size and position and shows through later — and
populate list contents before the first layout pass, not after the constructor returns. Both of
these produced a dialog that only lined up once some unrelated button forced a relayout.

**`Inspector.CurrentItem` returns a new COM reference on every read.** Read it once and release it
unless something retains it. Reading it twice per event and releasing neither leaked two references
for every item opened in a window.

**Never seed a stored setting from a `Strings.*` value.** It has happened twice — the Talk room title
and the share name — and in both cases a UI caption ended up as real data on the server. When adding
a setting, default it to empty and let the consumer apply its own fallback.

**Add new settings in three places:** `AddinSettings.cs` (property + constructor default),
`SettingsStorage.cs` (`AppendElement` in `SaveToXmlFile` **and** `case` in `ApplySettingValue`), and
the UI — or force-apply in `Lifecycle.cs` if no UI is wanted. Missing the storage half means the
value silently resets on every Outlook restart.

## Settings force-applied at startup

Applied in `NextcloudTalkAddIn.Lifecycle.cs → OnConnection` after loading settings, then saved back
to XML, so existing profiles are corrected once rather than keeping values with no UI:

| Setting | Forced to | Why |
|---|---|---|
| `TalkDeleteRoomOnEventDelete` | `true` | Always correct for the organizer |
| `TalkDefaultRoomType` | `EventConversation` | Only room type this fork creates |
| `TalkDefaultPasswordEnabled` | `false` | The dialog offers the password per room; a persistent default would pre-fill every room |
| `SharingDefaultShareName` | cleared **if it equals the settings caption** | One-time repair of the label-seeded value. Runs after `TryApplyOfficeUiLanguage` so the caption is compared in the language it would have been written in; the English literal is checked too |

## Upstream bugs fixed here

Candidates for contributing back:

1. UI thread blocked in `FileLinkLaunchController` and `TalkRibbonController`, and again in
   `BeforeAttachmentAdd` via an uncached policy request with a 45 s timeout.
2. `COMException` crash in `IsOrganizer` and the deferred lobby timer after the appointment COM
   object is invalidated.
3. Unhandled exceptions in `async void` VSTO ribbon handlers crashing Outlook.
4. Duplicate attachment prompt on drag-and-drop (`cancel=true` ignored by Outlook).
5. `ComposeAttachmentPromptForm` not `TopMost`.
6. Subscribing to `ExplorerEvents_10_Event.InlineResponse` breaking the Reading Pane inline-reply
   command bar.
7. `OnNewInspector` reading `Inspector.CurrentItem` twice and releasing neither reference, leaking
   two COM references for every item opened in its own window.
8. `AddinSettings` seeding `SharingDefaultShareName` with the settings field's own caption, so every
   share folder was named after a UI label.
9. The Talk dialog's title field writing itself back into `appointment.Subject`, silently renaming a
   meeting whose subject Outlook had not yet committed.

## Verification status

Builds are produced by GitHub Actions (`.github/workflows/build.yml`, `windows-latest`, MSBuild +
WiX 6.0.2) and uploaded as an artifact named `NCConnectorForOutlook-<ProductVersion>`. The workflow
runs on push to `main` and on manual dispatch; it creates no tags and no releases.

`gh` resolves an ambiguous repository when both `origin` and `upstream` remotes exist — always pass
`-R vladopol/NC_Connector_for_Outlook`, or set a default once with `gh repo set-default`.

**Versions 3.1.0.10 through 3.1.0.25 are verified in Outlook** in the target environment. Confirmed
working against a live Nextcloud and Exchange:

- appointment edits reaching the Talk room, from the meeting window and from the calendar grid;
- creating a Talk room leaving the meeting subject alone;
- the policy-aware room PIN accepted by a server that enforces password complexity;
- additional moderators promoted without the organizer leaving the room;
- share folders named from the mail subject;
- the Talk room dialog laying out correctly on first open.

Three defects were found only by running it, each after a build that compiled cleanly — worth
remembering when judging how much CI proves here:

| Symptom | Cause |
|---|---|
| Moderator list drawn at the top of the dialog, outside its group | A leftover `Controls.Add` re-parented it to the form; the **last** `Controls.Add` wins |
| Hint overlapping the list, correct only after pressing any button | The first layout pass ran before the candidates were added, and the list was only given bounds when non-empty |
| Numeric room PIN rejected by the server | Only `minLength` was read from the password policy; the complexity flags were ignored |
