# Changelog

All notable changes to **NC Connector for Outlook** will be documented in this file.

This project follows the principles of **Keep a Changelog** and **Semantic Versioning**.

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

