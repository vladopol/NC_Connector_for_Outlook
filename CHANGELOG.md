# Changelog

All notable changes to **NC Connector for Outlook** will be documented in this file.

This project follows the principles of **Keep a Changelog** and **Semantic Versioning**.

## [3.1.0.25] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### Talk room dialog: organizer removed from the moderator list, hint no longer overlaps it

- **The organizer is no longer offered as a moderator to tick.** They create the room and own it, so
  they already are one — which the hint says. Listing them implied the opposite. Filtered by the
  Nextcloud account creating the room, which is the account that becomes the owner.
- **The hint sits below the list on first display.** The list was only given bounds when it had
  items, so the very first layout pass — which runs before the candidates are added — left it at
  WinForms' default size and position and placed the hint as though no list existed. Any later
  relayout corrected it, which is why the group only lined up after pressing a button.

  The candidates are now filled before the first layout pass, the list is given bounds in both
  states so it can never keep its default ones, and the hint's background is transparent so a stray
  overlap could not hide the list behind it.

Startup timing instrumentation added in 3.1.0.24 is kept: it goes through the diagnostics logger,
which returns immediately when debug logging is off, so it costs nothing in normal operation.

## [3.1.0.24] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### Startup wiring moved off the user's first interaction, and timed

The add-in performs **no network requests at startup** — settings file, registry and MAPI only. Its
one potentially expensive startup step is opening the default calendar folder, which it does twice:
once for the Talk change watcher and once for the CalDAV sync. On an online-mode Exchange profile —
common on terminal servers, where a cached OST is impractical — each folder open and each `Items`
binding is a round-trip to the server.

That work was posted to the UI thread, meaning it ran at the next message-loop turn: exactly while
the user is clicking their first mail.

- The wiring now runs from a one-shot timer 5 s after connection, clear of Outlook's own startup
  burst and of the user's first interactions.
- Each step is timed individually and the elapsed values are logged, so a sluggish launch can be
  attributed from the debug log instead of guessed at.

This does not by itself explain a minute-long stall — see `docs/FORK.md` for what else to check on a
restricted terminal server.

## [3.1.0.23] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### FIXED: COM reference leak on every item opened in its own window

`OnNewInspector` read `Inspector.CurrentItem` **twice** and released neither reference. Every
`CurrentItem` read hands back a separate COM reference, and the handler runs for every item opened
in a window — including received mail, which the add-in does nothing with. Two references leaked per
opened item, held until garbage collection finalized them.

`CurrentItem` is now read once and released unless something retains it (an appointment tracking
subscription, or a compose-mail subscription). Received mail, contacts, tasks and notes — everything
the add-in ignores — no longer leak.

This is an upstream defect, present since the handler was written.

## [3.1.0.22] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### Talk room dialog: alignment, moderator list order and names

- **The password button is aligned with the group boxes below it** instead of sitting in the old
  input column, where it was indented far to the right with nothing above or below to line up with.
- **The moderator hint sits below the list, not underneath it.** `LayoutModeratorGroupControls` read
  its width from `_moderatorGroup.ClientSize`, which could still be the default size when the method
  ran; the list then came out at its minimum width and the hint was positioned as though no list was
  present, so the list — earlier in z-order — painted over it. The width is now passed in by the
  caller, which already knows it.
- **Attendees are listed by name.** `Фамилия Имя <mail@example.org>` instead of
  `login <mail@example.org>` — a Nextcloud login identifies an account, not a person. The name comes
  from Outlook's resolved `Recipient.Name` for the meeting attendee; the login is still shown when
  Outlook has no name for the entry.

## [3.1.0.21] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### FIXED: Talk room dialog layout

- **The moderator list rendered outside its group box, at the top of the dialog.** A leftover
  `Controls.Add(_moderatorListBox)` from the old search dropdown re-parented the control to the form
  after it had been added to the group — the last `Add` wins — so it was positioned in form
  coordinates while the "Additional moderators" box below sat empty. The stray line is gone.
- **The password row is one line:** `[Add/Remove password] [field] [Generate]`. The field now grows
  to the right of the button instead of appearing on a line below it, and is vertically centred
  against the buttons.
- **The "Password" caption is removed.** The button already says what the row is, and the caption
  sat in a column of its own with nothing else in it.
- **The moderator list is hidden entirely when there is nothing to tick**, rather than leaving an
  empty box; the hint alone explains what to do. When shown, it sizes to its content up to five
  rows.
- **The moderator hint is measured instead of stretched.** It used to be sized to fill whatever
  remained of a fixed-height group, which left a large empty gap between the list and the text. The
  group now sizes to its content.

## [3.1.0.20] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### Share folders are named after the mail subject

3.1.0.19 stopped share folders being named after a settings label, but the result was still
generic: every share landed as `20260729_Общий доступ`, and several shares on the same day were
indistinguishable. The mail's own subject is the obvious name, the same way the Talk room takes the
meeting subject.

Default share name, most specific source first:

1. an administrator's `share_name_template`, or a value deliberately typed in Settings,
2. the subject of the mail the wizard was launched from,
3. the localized fallback ("Общий доступ" / "Share").

The field stays editable — this only changes what it starts with. Attachment-mode shares use the
same subject-derived name (falling back to `email_attachment`), keeping their existing suffix
probing for collisions.

Subjects are sanitized for use as a path component and capped at 60 characters, truncated on a word
boundary where one is available. Mail subjects that Outlook has not yet committed to the item — the
subject typed immediately before pressing the button — read back empty, and the fallback applies.

## [3.1.0.19] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### FIXED: share folders were named after a UI label

`AddinSettings` seeded `SharingDefaultShareName` with `Strings.SharingDefaultShareNameLabel` — the
settings field's own caption. The value was therefore never empty, so the FileLink wizard's
whitespace check never fired and its localized fallback ("Общий доступ" / "Share") was dead code.
The caption went straight into `BuildShareFolderName`, and every share created with defaults landed
on Nextcloud as **`20260729_Название общего доступа`**.

Same class of bug as the Talk room title fixed in 3.1.0.17: a UI string leaking into stored data.

- The default is now empty, so the wizard's fallback applies as intended.
- Profiles that already stored the caption are cleared once at startup. The check runs after the
  Office UI language is resolved, so the caption is compared in the language it would have been
  written in, and the English literal is checked as well.

## [3.1.0.18] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### "Set password" removed from Talk defaults

With the room dialog offering a password per room on demand (3.1.0.15), a persistent "always set a
password" default contradicts it: it silently pre-filled a password for every room, which is the
behaviour the on-demand button was introduced to avoid.

- The checkbox is gone from Settings → Talk link.
- `TalkDefaultPasswordEnabled` is force-disabled in `OnConnection` and saved back, so existing
  profiles that stored `true` are corrected once rather than keeping a value with no UI.
- The `talk_set_password` backend policy is untouched — an administrator can still require a
  password, and the room dialog honours it by disabling the toggle.

## [3.1.0.17] - 2026-07-29

Fork patch on upstream 3.1.0.

---

### FIXED: creating a Talk room could silently rename the meeting

The dialog's "Title" field defaulted to the appointment subject and, on OK, wrote itself back into
`appointment.Subject`. Outlook does not commit a freshly typed subject to the item until the field
loses focus, so pressing the Talk button right after typing gave the dialog an **empty** subject —
it fell back to its placeholder ("Встреча" / "Meeting"), and clicking OK replaced the subject the
user had just typed.

The field is removed. The meeting subject is the single source of the room name.

- `ApplyRoomToAppointment` no longer writes `appointment.Subject`. The subject is the source of the
  room name, not a copy of it.
- The field had also become redundant: since 3.1.0.11 `TalkRoomSyncService` pushes the appointment
  subject to the room name on every save, so anything typed into it was overwritten at the first
  save regardless.
- The `talk_title` backend policy is no longer applied — there is nothing left for it to control,
  and a pinned title would be overwritten by the same sync. Two owners for one value is worse than
  none.

A meeting created with no subject at all still names its room "Meeting"; the first save replaces
that with the real subject.

## [3.1.0.16] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### Additional moderators — several, chosen from the meeting's attendees, without leaving the room

The single "Moderator (optional)" field was not what its name suggested: it performed a **handover**.
It promoted the chosen user and then called `LeaveRoom` for the organizer, who was dropped out of
their own meeting. It also wrote `X-NCTALK-DELEGATED`, after which this add-in stopped managing the
appointment entirely — no participant sync, no time or lobby updates, no room deletion.

- **Several moderators instead of one**, and they are **added**, not handed over to. The organizer
  stays the room owner, stays in the room, and appointment synchronization keeps working.
- **Candidates are the meeting's own attendees**, mapped to Nextcloud accounts — not the whole user
  directory. Attendee addresses are read on the UI thread and mapped on a background thread, since
  the lookup can refresh the address-book cache.
- The autocomplete field, dropdown and avatar loading are replaced by a plain checked list. With the
  candidate list bounded by the meeting's own invitees, search was solving a problem that no longer
  exists.
- The group is now labelled **"Additional moderators"**, and the hint states that the organizer is a
  moderator by default. Empty states are reported by cause rather than as one vague message:
  no attendees invited yet ("add attendees first" — the common case when the room is created before
  anyone is invited), no system address book, or attendees that have no Nextcloud account.

**Migration:** appointments delegated by builds up to 3.1.0.15 keep their old behaviour. Nothing
writes `X-NCTALK-DELEGATED` any more, but the reader is deliberately kept — dropping it would make
the add-in start managing rooms whose organizer deliberately handed them over and left.

## [3.1.0.15] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### Room password is an action now, not a permanent checkbox

A room password is the exception rather than the rule, and an unticked checkbox occupying the
dialog permanently communicated nothing. It is replaced by a single button that toggles caption
with the state — "Add password" / "Remove password" — keeping a fixed position so its target does
not move under the pointer when the password row appears or disappears below it.

- Adding a password fills one in immediately, so the common path is a single click; removing it
  clears the field so a stale value cannot be submitted.
- `TalkDefaultPasswordEnabled` now defaults to **off** for fresh installs, matching how the option
  is actually used. Existing profiles keep whatever is stored in their settings XML — untick
  "Set password" in Settings once to adopt the new default.
- When an administrator pins `talk_set_password`, the toggle is disabled rather than hidden, so the
  policy tooltip still explains why the state cannot be changed.

## [3.1.0.14] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### Room PIN now adapts to the server's password policy

Confirmed in the target environment: Nextcloud runs Talk conversation passwords through the same
`password_policy` validation it applies to account and share passwords — the app has no per-app
exemption — so the digits-only PIN introduced in 3.1.0.9 was rejected outright on a server that
enforces complexity.

The cause was that `PasswordPolicyService` read only `minLength` from capabilities and ignored the
complexity flags entirely, which is why the dialog happily generated a 10-digit PIN the server was
always going to refuse.

- **Complexity flags are now read** from `password_policy` capabilities:
  `enforceUpperLowerCase`, `enforceSpecialCharacters`, `enforceNumericCharacters` — including their
  singular/snake_case spellings and the `policies.account` nesting, so a key-name mismatch cannot
  silently read back "no complexity required".
- **Generation stays a PIN wherever the server allows it.** With no complexity enforced, the
  password is digits-only exactly as before. When the server does enforce it, the digit run stays
  the body and only the required characters are appended, always in the same position:
  `5729481!Kp` — "five seven two nine four eight one, exclamation mark, capital K, small p".
  Ambiguous glyphs (`I`, `l`, `O`, `0`) are excluded from the letters, and the symbol pool is just
  `!` and `$`: unambiguous to dictate, present on every keyboard layout, and squarely inside the
  set Nextcloud counts as special.
- **Passwords typed by hand are pre-checked** against the same rules, so the dialog reports the
  actual reason ("the server requires an uppercase and a lowercase letter") instead of letting room
  creation fail against the server.

## [3.1.0.13] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### No add-in code path blocks Outlook's UI thread on the network any more

3.1.0.12 only capped how long the wizards could freeze Outlook. This removes the blocking itself, so
the caps are no longer load-bearing — and with the blocking gone the caps could be cut to values
that match how a working server actually behaves.

**Talk ribbon flow** (`TalkRibbonController`) — the async plumbing already existed end to end, so
each blocking call simply moved off the thread:

- The connectivity probe, the room-replacement delete chain and `CreateRoom` (itself several
  requests in sequence) now run via `Task.Run`. COM access stays on the UI thread.
- The room description push after creation is queued to the background instead of running inline
  inside `ApplyRoomToAppointment` (`QueueRoomDescriptionUpdate`); the appointment body is read on the
  UI thread and handed over as a plain string.

**FileLink wizard** (`FileLinkWizardForm`) — the navigation chain is now asynchronous:

- `Navigate` → `ValidateCurrentStep` → `EnsureShareFolderAvailable` became
  `NavigateAsync` → `ValidateCurrentStepAsync` → `EnsureShareFolderAvailableAsync`, with the Back and
  Next handlers awaiting them. Because the message loop keeps pumping across an await, navigation is
  locked while a server call is in flight so a second click cannot start a parallel transition, and
  every continuation checks `IsDisposed` before touching controls.
- **The attachment-mode share name is no longer resolved in the constructor.** It probes WebDAV for
  a free folder name in a loop, so the wizard window did not appear at all until the server answered.
  It now runs from `OnShown` with the whole probing loop on a thread-pool thread: the window opens
  immediately and the name fills in once resolved.
- `PrepareUpload` (both call sites) moved to `Task.Run`; the share-folder cleanup on cancel/close
  paths became fire-and-forget, since nothing consumes its result and some callers run from
  `FormClosing` where awaiting is not possible.

**Timeouts** could then be cut to what a functioning server actually needs: connectivity probe and
policy fetches 10 s → **5 s**, interactive control-plane calls 15 s → **8 s**, background ones
30 s → **15 s**, user directory 30 s → **20 s**.

File transfers keep their 120 s ceiling, deliberately. That timeout is the wait for the server's
*response* — including assembling a chunked upload into its final file — and that work scales with
the size of the attachment. Shortening it would not make anything faster, only break large uploads.

## [3.1.0.12] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### HTTP timeouts sized to what the caller is waiting for

The Talk wizard and the FileLink wizard still issue some requests synchronously on Outlook's UI
thread, and those requests carried 45-120 s ceilings inherited from upstream. Because the calls run
in sequence, the ceilings add up per user action: a Talk button click against a server that accepts
the connection and then goes quiet could freeze Outlook for roughly 3 minutes, or 6 with an existing
room being replaced (verify 60 s + delete chain 180 s + create chain 120 s).

Timeouts are now named by purpose in `NcTimeouts` and applied accordingly:

- **Connectivity probe** (`VerifyConnection`, runs on the UI thread on every Talk click and forces a
  fresh TLS handshake): 60 s → **10 s**. Its entire job is to answer "is the server reachable?".
- **Policy fetches** that gate a dialog opening and fall back to local defaults on failure: 45-60 s →
  **10 s**.
- **Interactive control-plane calls** — Talk room create/delete/update, WebDAV folder existence
  checks, share folder creation: 60 s → **15 s**. These are cheap server-side; the ceiling only ever
  applies to a stalled server.
- **Background control-plane calls** — share creation/metadata, chunk folder setup and cleanup,
  CalDAV sync: 45-90 s → **30 s**, so a stalled server cannot pin a thread-pool thread either.
- **User directory download**: 60 s → **30 s** (legitimately larger than a control-plane call).
- **Unchanged:** file uploads and the server-side assembly of a chunked upload keep their 120 s
  ceiling — they are the only calls whose duration scales with user data. The Login Flow v2 poll
  keeps 60 s and is annotated: that wait is the user authenticating in a browser, not a stalled
  server.

Worst-case freeze for a Talk click drops from ~3 minutes to ~40 seconds (~1.5 minutes with room
replacement). This bounds the symptom rather than removing it — the remaining fix is to move those
calls off the UI thread, which requires making the wizard navigation asynchronous.

## [3.1.0.11] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### Appointment changes now propagate to the Talk room and the Nextcloud calendar

Editing a meeting that has a Talk room previously synced almost nothing back: for event
conversations — the room type this fork always uses — the subject and the agenda were skipped
outright, attendees were only ever added, a changed **end** time was ignored, and any edit made
without opening the meeting in its own window (dragging it in the calendar grid, editing it in the
peek pane, changing attendees from the scheduling view) reached the Nextcloud calendar but never
the Talk room.

- **Attendee removal.** Removing someone from the Outlook invitation now removes them from the Talk
  room. Only addresses this add-in previously recorded as attendees are eligible — the list is kept
  in a new `X-NCTALK-ATTENDEES` appointment property — so anyone invited directly in Talk is never
  touched. Owners and moderators (including a delegated moderator) are never removed.
- **Subject and agenda.** The room name and description are now updated for event conversations as
  well. If the server answers that the property belongs to the linked calendar object, the attempt
  is remembered per room for the rest of the session and silently skipped from then on — no repeated
  requests and no warning dialogs on every save.
- **End time.** Timing changes are tracked as a `start#end` pair instead of only the start, so
  shortening or extending a meeting refreshes the room's event binding and lobby timer.
- **Edits outside the meeting window.** A folder-level watch on the default calendar picks up
  changes to Talk appointments that never raise `AppointmentItem.Write`. It only reads the item —
  it never saves it, so a meeting is never silently marked as needing an update to be resent — and
  it skips appointments that have an open inspector, which the existing `Write` path already owns.
  Repeat and irrelevant `ItemChange` notifications are filtered by a content signature.
- **Attendees in the Nextcloud calendar.** The CalDAV payload now carries `ORGANIZER` and
  `ATTENDEE` (with role and response status), plus `LAST-MODIFIED` and `STATUS:CANCELLED` for
  cancelled meetings. Every scheduling property carries `SCHEDULE-AGENT=CLIENT` (RFC 6638) so
  Nextcloud does not mail its own duplicate invitations on top of the ones Outlook already sends.

**Threading:** the reconciliation was extracted into `TalkRoomSyncService`, which works from a
COM-free `TalkRoomSyncSnapshot` and runs on a thread-pool thread. Saving a meeting no longer blocks
Outlook's UI thread on Talk HTTP calls — this also resolves the appointment half of known upstream
bug #1. The delegation path still runs inline, because the room must be fully in sync before
moderation is handed over and the organizer leaves the room.

## [3.1.0.10] - 2026-07-28

Fork patch on upstream 3.1.0.

---

### RESOLVED: Outlook freezes and "slow add-in" auto-disable

Outlook times add-in startup and event-handler responsiveness and disables add-ins that exceed its
thresholds. An audit of every path that runs on Outlook's UI thread found blocking network I/O in
places that execute during normal mail and calendar work.

- **Attachment handling blocked the UI thread on an uncached HTTP request (the main offender).**
  `Explorer`/`Inspector` `BeforeAttachmentAdd` must return a decision synchronously, and it called
  the backend policy endpoint via `.GetAwaiter().GetResult()`. That endpoint was **never cached** and
  uses a **45 s timeout**, so *every attachment added to any email* froze Outlook until the Nextcloud
  server answered — up to 45 seconds with the server unreachable (VPN down, server restarting).
  The handler is now I/O-free: it reads the policy from a process cache, falls back to local settings
  when nothing is cached, and triggers a background refresh for next time.
- **Backend policy status is now cached** (`BackendPolicyCache`, 10 min TTL, 2 min back-off after a
  failure, one in-flight refresh at a time). It was previously re-fetched on every single call.
- **Exchange address resolution is memoized.** Mapping a recipient to its SMTP address falls through
  to `AddressEntry.GetExchangeUser()` / `PropertyAccessor`, which can round-trip to Exchange. This
  happens on the UI thread, repeatedly, for the same colleagues; the X.500 DN → SMTP mapping is now
  cached for the session.
- **Startup work moved out of the measured window.** `OnConnection` no longer opens the calendar
  folder, touches the registry or attaches the CalDAV sync inline — that wiring is posted back to the
  UI thread and runs once Outlook's message loop is pumping. Only settings loading and the inspector
  hook remain in the timed path.
- **Room deletion for a discarded appointment** ran a synchronous HTTP DELETE from a WinForms timer
  tick on the UI thread. It is now queued to a background thread.
- **Debug logging** no longer issues a `Directory.CreateDirectory` syscall per written line.

## [3.1.0.9] - 2026-07-05

Fork patch on upstream 3.1.0.

---

### Fork-specific changes

- **Talk room passwords are now simple numeric PINs** — generated and regenerated passwords for Talk rooms are now digits-only (e.g. `48213`), generated locally instead of via Nextcloud's server-side password generator (which could return long strings with special characters). Still respects the server's minimum-length policy if one is configured. FileLink share passwords are unchanged.

## [3.1.0.8] - 2026-07-05

Fork patch on upstream 3.1.0.

---

### RESOLVED: Reading Pane inline-reply losing Send/Discard/PopOut

Root cause confirmed by bisection across 3.1.0.2–3.1.0.7: subscribing to the `ExplorerEvents_10_Event.InlineResponse` event — regardless of what the handler does — breaks Outlook's rendering of the inline-reply command bar on the affected builds (16.0.0.14334 / 17932.20842). Neither the ribbon XML customizations, nor `Inspectors.NewInspector`, nor the handler's contents (network calls, COM property reads, event subscription timing) were responsible; the mere act of hooking `InlineResponse` was enough.

**Fix:** `Explorer.InlineResponse` is no longer hooked at all. Removed the now-dead supporting code: `EnsureApplicationHook`, `EnsureExplorerInlineResponseHooks`, `HookInlineResponseExplorer`, `UnhookExplorerInlineResponseHooks`, `OnNewExplorer`, `OnExplorerInlineResponse`, and the associated `_explorers`/`_explorersEvents`/`_inlineResponseExplorers`/`_inlineResponseExplorerEvents` fields.

**Trade-off:** automatic FileLink attachment-sharing no longer triggers for attachments added during a Reading Pane inline reply. It still works normally for popped-out compose windows (New Mail, Reply/Forward opened in a separate window) via `Inspectors.NewInspector`, and the Settings/FileLink ribbon buttons are unaffected — both were proven innocent during the investigation and are fully restored.

**Investigation trail** (each build ruled out one variable): 3.1.0.2 removed a synchronous network call suspected of blocking the UI thread — no effect. 3.1.0.3 removed a `TabComposeTools/TabMessage` ribbon customization — no effect. 3.1.0.4 deferred event-handler COM work to the next message-loop tick — no effect. 3.1.0.5 disabled all mail/Explorer ribbon customization and event hooks simultaneously — fixed. 3.1.0.6 restored the ribbon while keeping event hooks disabled — stayed fixed, clearing the ribbon. 3.1.0.7 restored `NewInspector`'s mail branch while keeping `Explorer.InlineResponse` disabled — stayed fixed, isolating the cause to `Explorer.InlineResponse` specifically.

## [3.1.0.7] - 2026-07-05 — BISECTION BUILD, not for general deployment

Fork patch on upstream 3.1.0.

---

### 3.1.0.6 confirmed: not the ribbon. Narrowing to Explorer.InlineResponse vs. NewInspector

3.1.0.6 restored both ribbon customizations while keeping `Explorer.InlineResponse` and `NewInspector`'s mail branch disabled — the bug did not reproduce, ruling out the ribbon XML entirely. This build restores `NewInspector`'s mail-compose branch (relevant to popped-out compose windows) while keeping `Explorer.InlineResponse` disabled, to determine whether `NewInspector` is also implicated or whether `Explorer.InlineResponse` alone is the cause.

## [3.1.0.6] - 2026-07-05 — BISECTION BUILD, not for general deployment

Fork patch on upstream 3.1.0.

---

### 3.1.0.5 fixed the bug — bisecting which change actually mattered

3.1.0.5 removed 4 things simultaneously: the Explorer ribbon tab (Settings button), the Mail.Compose ribbon group (FileLink button), the `Explorer.InlineResponse`/`Explorers.NewExplorer` hooks, and `NewInspector`'s mail branch. The bug disappeared, but it's unknown which change(s) were responsible.

This build restores **both ribbon customizations** (Settings + FileLink buttons, still without the `TabComposeTools/TabMessage` contextualTabs group removed back in 3.1.0.3) while keeping the **event hooks disabled** (`EnsureApplicationHook` not called, `NewInspector`'s mail branch skipped). If the bug returns here, the ribbon XML is the cause. If it stays fixed, the event hooks (`Explorer.InlineResponse` / `Inspectors.NewInspector` touching the mail COM object) are the cause instead.

## [3.1.0.5] - 2026-07-05 — DIAGNOSTIC BUILD, not for general deployment

Fork patch on upstream 3.1.0.

---

### Diagnostic isolation test: Reading Pane inline-reply Send/Discard/PopOut bug

Three targeted fixes (3.1.0.2 threading, 3.1.0.3 ribbon TabMessage removal, 3.1.0.4 deferred event handling) made zero observable difference, confirmed via clean debug logs and Programs & Features version checks each time. This build performs the most aggressive isolation possible short of uninstalling the add-in:

- `GetCustomUI` returns `null` for `Microsoft.Outlook.Explorer` and `Microsoft.Outlook.Mail.Compose` (no Settings/FileLink ribbon buttons anywhere for mail). `Microsoft.Outlook.Appointment` (Talk button) is untouched — different, unrelated ribbon context.
- `EnsureApplicationHook` (which wires `Explorer.InlineResponse` / `Explorers.NewExplorer`) is not called at all.
- `Inspectors.NewInspector`'s mail-compose branch is skipped entirely, including the `mail.Sent` COM property read that previously happened synchronously inside the event handler even after the 3.1.0.4 deferral fix.
- Talk (appointment subscriptions, room creation) and CalDAV sync are untouched.
- Added diagnostics: every `GetCustomUI` call now logs the requested `ribbonID`; `OnRibbonLoad` now logs each firing. This will reveal whether Outlook queries a ribbon context we don't even handle, and how ribbon-load timing correlates with `InlineResponse`/`NewInspector` events.

If the bug still reproduces with mail items receiving **zero** addin COM interaction and **zero** mail-related ribbon customization, the cause is not in this add-in's mail-compose/ribbon subsystem at all — likely the mere presence of an `IRibbonExtensibility`-implementing COM add-in interacting with this specific Outlook build (16.0.0.14334 / 17932.20842), which would need a different remediation strategy entirely.

**This build sacrifices FileLink automation and the Settings/FileLink ribbon buttons for mail — diagnostic only, not for general rollout.**

## [3.1.0.4] - 2026-07-05

Fork patch on upstream 3.1.0.

---

### Fix attempt: Reading Pane inline-reply still losing Send/Discard/PopOut after 3.1.0.3

Confirmed via Programs & Features (version/publisher now correctly shown) that 3.1.0.3 was the build actually tested, and the bug still reproduced with the `TabMessage` ribbon customization removed — ruling that out too.

New hypothesis: `Explorer.InlineResponse` and `Inspectors.NewInspector` fire *while Outlook is still constructing* the embedded inline-compose command bar. Both handlers did synchronous COM work (`EnsureMailComposeSubscription`, resolving identity keys, subscribing to `ItemEvents_10`) directly inside the event callback. Re-entrant COM calls into the object model during that specific construction window are a known class of Outlook ribbon-rendering glitch, independent of exceptions or delays — consistent with the clean debug logs from the last two attempts. Deferred this work to the next message-loop iteration via the already-established `_uiSynchronizationContext.Post` pattern (used elsewhere for `QueueDeferredAppointmentSubscriptionEnsure`), so Outlook finishes building the inline-compose UI before the addin touches the mail item's COM interface. Another targeted test, not a confirmed fix.

## [3.1.0.3] - 2026-07-05

Fork patch on upstream 3.1.0.

---

### Fix attempt: Reading Pane inline-reply still losing Send/Discard/PopOut after 3.1.0.2

The threading fix in 3.1.0.2 made no observable difference — a debug log captured during a live repro showed no exceptions, no delays, and normal compose-subscription lifecycle timing (all under ~200ms), ruling out the addin's own runtime code as the cause. The only remaining untested part of the original hypothesis is the ribbon XML: `GetCustomUI("Microsoft.Outlook.Explorer")` added a custom group to the **native `TabMessage` tab under the `TabComposeTools` contextual tab set** — exactly the tab Outlook shows for Reading Pane inline replies. Removed this customization; the FileLink button remains available via the popped-out compose ribbon and the automatic attachment-triggered share flow. This is a targeted test, not a confirmed fix.

### Installer fixes

- **Publisher corrected** — `Manufacturer` in the WiX package was still "Bastian Kleinschmidt" (upstream author, uninvolved in this fork's builds). Changed to "Vladimir Poluliashenko".
- **Fork patch not visible in Add/Remove Programs** — MSI `ProductVersion` only supports 3 numeric fields, so every fork patch collapsed to the same 3-part version (e.g. both 3.1.0.1 and 3.1.0.2 showed as "3.1.0"). The CI workflow now folds the fork-patch digit into the build field (`upstream_patch * 100 + fork_patch`), so each release is distinguishable and correctly ordered in Windows Installer's version comparisons. Not a regression from this fork's recent changes — it affected every previous release too.

## [3.1.0.2] - 2026-07-05

Fork patch on upstream 3.1.0.

---

### Upstream bugs fixed

- **UI thread blocked during compose attachment policy check** — `OnAttachmentEvalTimerTick` (a WinForms `Timer` tick, running on the UI thread) called `ReadAttachmentAutomationSettings`, which fetched `FetchBackendPolicyStatus` synchronously — a blocking HTTP request with a 60s timeout. Split into `ReadAttachmentAutomationSettingsAsync` (awaited via `Task.Run`) for the timer-driven evaluation path; a synchronous wrapper is kept only where `BeforeAttachmentAdd`'s COM contract requires Outlook to receive an immediate decision.

---

### Bug fixed: Reading Pane inline-reply losing Send/Discard/PopOut

Diagnosed root cause: every inline reply started a 250ms debounce timer (`MailComposeSubscription`, Email Signature feature) that fetched `FetchBackendPolicyStatus` — a blocking, uncached HTTP request (60s timeout) — synchronously on the UI thread, right as Outlook finished constructing the embedded inline-compose command bar. Removing the Email Signature feature (below) eliminates both call sites responsible for this (`MailComposeSubscription.Signature.cs` and `ComposeShareLifecycleController.ApplySeparatePasswordBackendSignature`).

---

### Fork-specific changes

- **Email Signature feature removed** — Requires a server-side Nextcloud plugin never installed in this deployment. The settings tab was already hidden, but the runtime still fired a real network policy check on every compose, reply, and forward. Removed entirely: `NextcloudTalkAddIn.MailComposeSubscription.Signature.cs`, `EmailSignaturePolicyService`, `EmailSignaturePlainTextController`, the `EmailSignaturePolicy` model, the hidden Signature settings tab, and the `EmailSignatureOn*` settings fields (old settings XML files with these keys still load fine — unknown keys are ignored).
- **Local IFB Free/Busy HTTP listener removed** — Already force-disabled on every startup since Exchange handles Free/Busy natively for this deployment. Removed the dead `FreeBusyServer` listener and the `IfbEnabled` / `IfbPort` / `IfbDays` settings and hidden settings tab. Kept `IfbAddressBookCache` and `IfbCacheHours` — still actively used by Talk and FileLink for the Nextcloud user picker — and a narrow one-time registry-restore safety net in `FreeBusyManager`, in case an earlier build ever ran with the listener enabled and hijacked Outlook's native free/busy registry paths.

## [3.1.0.1] - 2026-06-01

Fork patch on upstream 3.1.0.

---

### Upstream bugs fixed

These bugs exist in the original upstream codebase and are fixed here.
They are candidates for contributing back to upstream.

- **Outlook UI freezes on FileLink and Talk button** — `FileLinkLaunchController.RunFileLinkWizardForMail` blocked the UI thread with `.GetAwaiter().GetResult()` while waiting for two parallel HTTP requests. `TalkRibbonController` called `GetSystemAddressbookStatus(forceRefresh=true)` and `GetUsers` synchronously on the UI thread. Fixed by making `RunFileLinkWizardForMail` fully `async Task<bool>` and wrapping address book calls in `Task.Run`.
- **Outlook crash on meeting cancellation (`COMException 0x9284010A`)** — The deferred lobby verification timer continued firing after the appointment COM object was invalidated by a delete or cancel. `IsOrganizer` had no `COMException` guard; the timer tick had no outer catch. Added `COMException` handling in `IsOrganizer` (returns `false`) and a try/catch around the tick body that stops the timer cleanly.
- **Unhandled exceptions crash Outlook via async void ribbon handlers** — `OnTalkButtonPressed`, `OnSettingsButtonPressed`, and `OnFileLinkButtonPressed` are `async void` COM callbacks with no `try/catch`. Any unhandled exception propagated through `SynchronizationContext` and crashed Outlook. Added exception guards with diagnostic logging to all three.
- **Duplicate attachment prompt after drag-and-drop** — Outlook DnD sometimes ignores `BeforeAttachmentAdd cancel=true` and adds the file anyway. After the wizard created the Nextcloud link, `OnAttachmentAdd` fired, the evaluation timer triggered, and a second "limit exceeded" dialog appeared. Fixed by setting `_attachmentSuppressed = true` during the wizard in `StartBeforeAddAttachmentShareFlow` and removing the attachment by name if Outlook added it despite `cancel=true`.
- **Attachment prompt dialog hidden behind other windows** — `ComposeAttachmentPromptForm` appeared below Explorer and other app windows because it was only modal relative to Outlook. Timer-triggered dialogs require `TopMost = true` since Outlook may not be the foreground app at the moment they appear.
- **UI thread blocked during attachment flow timer callbacks** — `OnAttachmentEvalTimerTick` and `OnBeforeAddShareTimerTick` were synchronous `void` callbacks. If the share wizard was triggered from a timer tick, it blocked the UI thread for the duration of HTTP prefetch requests. Made both timer callbacks `async void` and propagated `async Task` through `EvaluateAttachmentAutomation`, `StartComposeAttachmentShareFlow`, and `RunQueuedBeforeAddAttachmentShareFlow`.

---

### Fork-specific fixes

These fixes apply only to features or changes introduced in this fork.

- **CalDAV sync settings not persisted** — `CalDavSyncEnabled` and `CalDavCalendarName` were missing from both `SaveToXmlFile` and `ApplySettingValue` in `SettingsStorage`. The checkbox state was lost on every Outlook restart.
- **Nextcloud Calendar event not deleted when meeting is removed** — When the organizer deleted a saved Outlook event, the Talk room was removed but the CalDAV event remained in Nextcloud Calendar. `CalDavDeleteTracker` is ephemeral and lost on Outlook restart; if the appointment was deleted without a prior modification, no tracker existed. Fixed by explicitly queuing a CalDAV DELETE in `QueueSavedEventRoomDeletion` using the EntryID-derived UID, independent of the tracker.
- **`_calDavCalendarSync` zombie state after Detach** — After `Detach()`, `_calDavCalendarSync` was not nulled. Subsequent calls to `TryDirectCalDavSync` reached a detached instance. Fixed by setting `_calDavCalendarSync = null` after `Detach()`.
- **Wrong room type silently returned after hiding room type combo** — After hiding the room type selector in `TalkLinkForm`, `SelectedRoomType` was read from an empty combo which defaulted to `StandardRoom`. Hardcoded `SelectedRoomType = TalkRoomType.EventConversation` directly on OK.
- **Merge artifact: duplicate `BuildPlainText` build error** — The fork's older `BuildPlainText` implementation in `FileLinkHtmlBuilder` was not removed during the upstream 3.1.0 merge, causing `CS0111`. Removed the stale copy and added the missing `using System.Collections.Generic` required by the upstream implementation.

---

### Fork-specific changes

Intentional behavior and UI changes specific to this deployment.

- **IFB disabled and hidden** — The IFB (Free/Busy) endpoint is not used in this deployment. Force-disabled on startup; settings tab hidden.
- **Signature tab hidden** — The email signature feature requires a server-side Nextcloud plugin component not available in this environment. Tab is hidden; the underlying settings remain accessible in code.
- **Room type fixed to "Event Conversation"** — Outlook meetings always map to Nextcloud Talk Event Conversations. Group conversations are created directly in Nextcloud Talk. Room type selector hidden in both settings and create-room dialog; `EventConversation` is force-set on startup.
- **"Add Guests" option hidden** — External participants connect via the room link and optional password. The email-guest invite API is redundant when Outlook already delivers calendar invitations to all attendees. Hidden in both settings and create-room dialog; value is always `false`.
- **Talk room deletion on event delete always enabled** — Deleting a saved Outlook event always removes the linked Talk room on Nextcloud. The per-user opt-in checkbox is hidden; the setting is force-enabled on startup.

## [3.1.0] - 2026-05-12

### Added
- Central backend-managed email signatures for matching Outlook sender identities.
- Backend signatures for HTML/RTF and plain-text compose, including replies and forwards.
- Nextcloud share insertion from inline reply and forward windows.
- Plain-text share insertion that keeps Outlook's reply text intact.
- Chunked Nextcloud WebDAV upload v2 for large shared files.
- Per-file upload speed display in the Outlook sharing wizard.

### Changed
- Outlook WordEditor insertion paths now use shared helpers for inspector and inline compose windows.
- Share, signature, and separate-password mail handling now preserve the sender identity captured from the original compose item.
- Talk room deletion for saved appointments is controlled by an explicit opt-in policy.
- Talk event cleanup stays local to Outlook metadata and no longer depends on generic Talk URL fallbacks.
- Sharing wizard status rendering was adjusted so progress and speed text remain readable.

### Fixed
- Inline reply sends now dispatch queued separate password follow-up mails.
- Separate password follow-up mails now receive the backend signature when the captured sender matches the signature policy.
- Manual fallback drafts for separate password mails include the same backend signature handling as auto-send.
- Talk lobby timing and saved-event cleanup were corrected for edited appointments.

## [3.0.4] - 2026-04-28

### Added
- Configurable local IFB listener port in settings and runtime.

### Changed
- Runtime API logging and JSON serialization helpers were centralized.
- Cleanup and guard-related consolidation was streamlined across the active runtime paths.
- Talk and Sharing wording was updated across all supported locales.

### Fixed
- Removed the dead Talk appointment HTML read fallback path.
- Password notification icon cleanup is now marshaled onto the captured UI context.
- Talk help URL and block-marker detection were hardened.

## [3.0.3] - 2026-04-17

### Added
- Daily runtime log rotation with 7-file retention and 30-day cleanup.
- Log anonymization option with redaction of sensitive runtime diagnostics.
- Backend HTML template hardening with fail-closed sanitizer policy.

### Changed
- Shared helper refactor for URL launch, size formatting, and COM release cleanup.
- Refactor cleanup pass to normalize shell/COM patterns and remove leftover redundancies.
- Runtime HTTP request handling centralized via `NcHttpClient` and `NcJson`.
- Cleanup refactor: removed duplicated helpers and redundant conditions.

### Fixed
- Password-notification icon lifecycle marshaled to the captured UI context.
- Removed ineffective password-selection JavaScript handlers for Outlook HTML rendering.

## [3.0.2] - 2026-04-15

### Added
- New runtime TLS controls in Outlook settings (`Use OS default TLS policy`, `Enable TLS 1.2`, `Enable TLS 1.3`) with immediate test/login-flow usage and persisted profile-specific configuration.

### Changed
- Release line/version references were aligned to `3.0.2` across assembly and installer defaults.
- TLS hint text in advanced settings/docs was tightened and no longer makes assumptions about machine-wide registry changes.
- TLS diagnostics guidance text was updated to be mode-neutral (no hardcoded OS-default assumption) and aligned across all supported locale files.

### Fixed
- Admin-controlled `?` hint glyph placement in advanced settings now avoids overlapping adjacent inputs in tight rows (language override section and neighboring fields).
- Connection diagnostics and login-flow connectivity checks now force fresh HTTP/TLS handshakes (no pooled keep-alive reuse), so runtime TLS mode switches are validated deterministically without requiring an Outlook restart.

## [3.0.1] - 2026-04-14

### Added
- The FileLink wizard now accepts Explorer drag & drop for files/folders across the entire file step (queue, surrounding pane, and action area), not just via explicit add buttons.
- TLS defaults are now hardened at add-in scope via `NcTalkOutlookAddIn.dll.config`, explicitly enabling strong crypto and OS-default TLS negotiation without changing machine-wide .NET registry settings.
- Settings connection diagnostics now classify transport failures (TLS handshake, certificate trust, DNS, proxy/connectivity, timeout) and surface actionable guidance instead of only generic secure-channel errors.

### Changed
- Release line/version references were aligned to `3.0.1` across assembly metadata, installer defaults, readmes, and admin/development docs.
- The new connection diagnostics strings were translated across all supported Outlook locales, so non-English users do not fall back to generic English placeholders.
- Compose attachment automation now evaluates attachments pre-add via `BeforeAttachmentAdd`; threshold/always flows can cancel the host add and open NC share mode before Outlook post-add handling.
- Compose automation subscription coverage was expanded with `ApplicationEvents_11.ItemLoad` so mail items loaded outside `NewInspector` (for example inline compose contexts) are also tracked for attachment automation.
- Repository text-format policy was standardized with new `.editorconfig` and `.gitattributes`; project text files were normalized to consistent CRLF line endings.
- With debug logging enabled, compose pre-add attachment automation now logs detailed candidate/decision/fallback diagnostics (including unresolved-path reasons) to improve support traceability.
- `NextcloudTalkAddIn` was split into dedicated `partial` units for runtime hooks and logging to reduce orchestration density in the main source file.
- COM cleanup now uses shared `ComInteropScope` helpers (including scoped wrappers), and high-risk recipient/explorer release paths were migrated to the centralized implementation.
- Oversized UI forms were modularized with `partial` feature files (`FileLinkWizardForm.DragDrop`, `SettingsForm.Language`, `TalkLinkForm.Moderator`) to improve maintainability and test focus.
- Talk body/template rendering and block sanitation were extracted into `Controllers/TalkDescriptionTemplateController`.
- Appointment attendee e-mail discovery and SMTP recipient resolution were extracted into `Controllers/OutlookRecipientResolverController`.
- Compose subscription lifecycle (get-or-create/remove/dispose-all) now runs through `Controllers/MailComposeSubscriptionRegistryController` instead of direct list/lock management in the add-in root.
- Talk appointment lifecycle logic (`ApplyRoomToAppointment`, runtime room trait resolution, room mutation sync, delegation and participant sync) was extracted into `Controllers/TalkAppointmentController`.
- Compose share cleanup and separate-password dispatch flow (including recipient normalization, auto-send, and manual fallback) was extracted into `Controllers/ComposeShareLifecycleController`.
- Legacy recipient helper forwarders in `NextcloudTalkAddIn` were removed; compose recipient normalization now calls `ComposeShareLifecycleController` helpers directly.
- Remaining recipient CSV/normalization passthrough wrappers were removed from the add-in root to avoid duplicate helper paths.
- `ComposeShareLifecycleController` COM release paths now use centralized `ComInteropScope.TryRelease(...)` for consistent exception-safe cleanup.
- `build.ps1` now supports `-SkipIceValidation` for environments where WiX ICE execution is unavailable (`WIX0217`), while keeping the default validated build path unchanged.
- Large nested runtime subscription classes were moved out of the root file into dedicated partial units:
  - `NextcloudTalkAddIn.MailComposeSubscription.cs`
  - `NextcloudTalkAddIn.AppointmentSubscription.cs`
- Minor UI redundancy cleanup (no behavior change) was applied in `FileLinkWizardForm`, `SettingsForm`, and `TalkLinkForm` (shared helpers for selection validation, resize wiring, and settings-option checkbox setup).

### Fixed
- Pre-add multi-file drag/drop is now debounced and batched into a single wizard launch instead of opening one wizard per file.
- Folder uploads in the FileLink wizard now correctly create required subfolders for mixed file+directory queues; reserved-name tracking no longer suppresses required DAV `MKCOL` calls.
- Upload status/progress UI now flushes buffered per-item progress updates immediately on failure/cancel/finalize paths, preventing stale bars or missing final state labels.
- Admin-controlled `?` hint glyphs now support explicit row anchors, preventing them from drifting into adjacent password and attachment threshold input fields in the FileLink wizard, sharing settings, and Talk password block.
- Pre-add attachment interception now also probes `Attachment.FileName` and `Attachment.DisplayName` for resolvable local paths when `Attachment.PathName` is unavailable, improving early capture reliability for drag/drop scenarios.
- Pre-add attachment candidate materialization now falls back to `Attachment.SaveAsFile(...)` and compares COM-reported size vs. measured file size, reducing false below-threshold decisions for unresolved path scenarios.
- `ApplicationEvents_11.ItemLoad` compose subscription is now limited to active inline-compose contexts to prevent duplicate compose subscriptions and duplicate threshold prompts when inspector-based subscriptions are already active.
- Add-in lifecycle teardown was de-duplicated by centralizing shutdown/disconnect cleanup into a shared idempotent path.

## [3.0.0] - 2026-03-30

### Added
- Optional NC Connector backend policy runtime for Talk and Sharing:
  - backend status endpoint is queried on Talk wizard open, Sharing wizard open, and Settings open/save
  - active valid seats enable backend policy values and `policy_editable` locks
  - paused/invalid seat states show an in-product warning and fall back to local add-in settings
  - central templates are supported for share HTML/password blocks and Talk description text
  - separate password follow-up delivery is explicitly gated behind backend endpoint + active assigned seat
  - backend custom text templates are only activated when the language override is set to `Custom`, otherwise local UI-default text remains active
  - backend attachment-threshold policy now treats `attachments_min_size_mb: null` as an explicit "disabled" state

### Changed
- Release line/version references were aligned to `3.0.0` across assembly metadata, installer defaults, readmes, and admin/development docs.
- Backend policy runtime is now fetched live on the relevant entry points and evaluated by compose attachment automation as well as in Settings/Wizard UI, so locked backend attachment rules stay authoritative without reusing stale cached policy data.
- Backend policy runtime now targets `/apps/ncc_backend_4mc/api/v1/status`; if the backend is unreachable, the license/seat state is no longer usable, or the backend grace window has expired, Outlook falls back to local add-in settings.
- Talk event descriptions now honor backend `event_description_type`; HTML templates are written into the open Outlook appointment editor with stable NC block markers while `Body` stays aligned for room-description sync.
- Locked settings and wizard controls now expose their admin/seat/backend explanation through active hint anchors instead of relying on WinForms tooltips on disabled controls.

### Fixed
- Outlook issue #4 (`HTTP 400` at final share creation) was hardened in the sharing service:
  - create-share now extracts and surfaces OCS server error details instead of only generic WebException text
  - share creation now follows the documented Nextcloud OCS contract more closely: `label` is sent on create, mutable metadata is updated separately through form-encoded OCS update arguments
  - the previous `HTTP 400` retry path that silently dropped optional `label` / `note` metadata was removed
- Runtime exceptions are now always written to `addin-runtime.log`, even when the optional debug log switch is off.
- Backend custom share HTML in attachment mode now removes only the `RIGHTS` row, preventing truncated or fragmentary HTML blocks.
- Compose mail insertion now prefers the `HTMLBody` path for backend HTML blocks, avoiding partial insertion from the old Word-editor fallback path.
- Talk appointment save handling now performs a short deferred post-write lobby verification so start-time changes from externally created appointments are still picked up after Outlook commits the final value.

## [2.3.0] - 2026-02-28

### Added
- Compose attachment automation settings:
  - `Always handle attachments via NC Connector`
  - `Offer upload above X MB`
- Two-action threshold prompt in compose:
  - `Share with NC Connector`
  - `Remove last selected attachments`
- Attachment-mode wizard launch context (direct file-step start, fixed `email_attachment` share naming with deterministic suffixes, `yyyyMMdd` date prefix).
- Separate password mail flow:
  - new default setting `Send password separately`
  - wizard toggle `Send password in separate email`
  - password-only follow-up HTML and post-send dispatch queue.
- Compose share cleanup lifecycle for unsent drafts (armed/cleared/delayed/delete runtime states).
- Built-in host large-attachment conflict guard with live setting checks in Settings UI and runtime flow gates.
- System-addressbook hardening for Talk defaults and Talk wizard:
  - centralized runtime availability contract (`available`, `error`, `count`, `forceRefresh`)
  - live checks on Talk click, settings open/save, and wizard open
  - deterministic lock state with context-specific tooltips and red warning blocks (Settings + Wizard).

### Changed
- Sharing HTML output in attachment mode:
  - ZIP download URL now uses `/s/<token>/download`
  - permissions row is hidden in attachment mode
  - inline password is hidden when separate password dispatch is enabled.
- Attachment-mode permissions are enforced as read-only regardless of sharing defaults.
- Logging depth for compose sharing flows now includes attachment-evaluation decisions, prompt actions, cleanup lifecycle transitions, and separate-password dispatch outcomes.
- Settings persistence is profile-aware and XML-based:
  - `%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml`
  - encrypted `AppPasswordProtected` (Windows DPAPI, CurrentUser scope)
  - automatic migration from legacy `settings.ini` paths.
- Runtime data path has been consolidated to `%LOCALAPPDATA%\NC4OL`:
  - debug log (`addin-runtime.log`)
  - IFB address book cache.
- Installer defaults have been renamed for consistent product naming:
  - install directory: `C:\Program Files\NC4OL`
  - IFB marker key: `HKLM\Software\NC4OL\HttpUrl`
  - legacy installer key `HKLM\Software\NextcloudTalkOutlookAddIn\HttpUrl` is removed on install.
- Delegation write-flow in appointment `OnWrite` was tightened:
  - pre-step execution order is deterministic (`room name` -> `lobby` -> `description` -> `participants`)
  - per-step pre-sync status is logged explicitly before moderator handover
  - pending delegation is no longer silently blocked by pre-step failures (handover still executes, failures stay visible in runtime logs).

### Fixed
- Debounced attachment evaluation to prevent duplicate triggers during rapid multi-selection.
- Runtime guard enforcement now blocks compose attachment automation not only in UI, but also before evaluation, before threshold action handling, and before attachment-mode finalize.
- Password policy capability parsing is now compatible with current Nextcloud payload variants (`minLength` etc.) and normalized generator URL formats.
- Password policy HTTP handling now enables gzip/deflate decompression and sanitizes JSON wrappers before parsing.
- Cross-client Talk edit interoperability was hardened for Thunderbird-created appointments:
  - missing local room traits are now bootstrapped from Talk server endpoints (`/object`, `/webinar/lobby`)
  - resolved room traits are persisted back into Outlook appointment properties
  - prevents false room-description update failures like `...: event` when editing TB-created event conversations.
- MSI maintenance behavior was aligned for support scenarios:
  - reinstall/update now supports same/older/newer package installation over an existing installation (`AllowDowngrades="yes"`).
- Room creation path no longer falls back from event conversation to standard conversation on create errors:
  - requested room type is now kept deterministic
  - if event conversation prerequisites are missing, creation fails fast with a clear service error.

### Documentation
- Updated `README.md` / `README.de.md` with 2.3.0 operational behavior:
  - profile-based settings XML + legacy migration cleanup
  - consolidated runtime path `%LOCALAPPDATA%\NC4OL`
  - compose cleanup and separate-password follow-up flow semantics.
- Clarified MSI reinstall/update semantics in README/admin docs (same/older/newer package install over existing installation).
- Expanded `docs/ADMIN.md` / `docs/ADMIN.de.md` with compose-sharing lifecycle details (cleanup contract and password follow-up dispatch behavior).
- Expanded `docs/DEVELOPMENT.md` / `docs/DEVELOPMENT.de.md` with explicit 2.3.0 implementation deltas and runtime contracts used for parity work.

## [2.2.7] - 2026-02-13

### Added
- More UI translations (see `Translations.md`).
- Tooltips across Settings and the wizards.
- Optional auto-add of event invitees to the Talk room (Nextcloud users via system address book, others via e-mail).
- Live Nextcloud password policy support for Talk and Sharing (minimum length + generator API, with secure fallback).

### Changed
- Unified, modernized UI across Talk wizard, Sharing wizard, and Settings.
- Legal/branding update: renamed to **NC Connector for Outlook**, with new app icon and header assets.

### Removed
- Buggy muzzle feature.

### Fixed
- Dark mode/theme handling to better follow the Outlook/Office theme.

### Documentation
- Expanded admin and developer documentation.

