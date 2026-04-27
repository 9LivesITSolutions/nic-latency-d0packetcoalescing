# Changelog

All notable changes to this project will be documented in this file.

---

## [4.2.0] - 2025-01-27

### Fixed
- Syntax error (missing closing brace) caused by heredoc `@' '@` mishandling in `$ApplyBlock` — affected PS parser beyond line 176
- `$isDryRun` typed as `[bool]` via `[bool]$WhatIfPreference` instead of `$WhatIfPreference.IsPresent` which returned an empty string when `-WhatIf` was absent, causing `Invoke-FixOnComputer` parameter binding to fail

### Changed
- Scheduled task command now encoded as **Base64** (`-EncodedCommand`) to avoid quoting and escaping issues in nested PowerShell strings
- Adapter list for scheduled task written to `C:\Windows\Temp\9Lives-NICList.txt` (file-based handoff) instead of being embedded in task command string

---

## [4.1.0] - 2025-01-27

### Added
- Step-by-step verbose logging to `C:\Windows\Temp\9Lives-NICFix.log` on remote targets
- Log is read and printed to console during Pass 2 for full execution trace
- Pass 1 / Pass 2 progress messages in console (`[PC] Pass 1 - Applying fix...`)
- `=== RESULTS ===` section header in final output

### Changed
- Log file and NIC list temp files cleaned up at end of Pass 2

---

## [4.0.0] - 2025-01-27

### Added
- **Scheduled task mechanism** (SYSTEM, one-shot) for NIC restart — survives WinRM session drop caused by WiFi adapter restart
- Adapter list passed via temp file `9Lives-NICList.txt` to scheduled task

### Changed
- Remote mode now uses two distinct WinRM sessions (Pass 1: apply, Pass 2: read) with 15s delay between them
- `$ReadBlock` reads live status via `Get-NetAdapterPowerManagement` after driver reload

### Fixed
- NIC restart was silently skipped because `Restart-NetAdapter` inside a WinRM session was interrupted when the WiFi adapter disconnected, dropping the session before the cmdlet completed

---

## [3.2.0] - 2025-01-27

### Added
- `Restart-NetAdapter` after registry write so driver reloads the new value immediately
- Local mode confirms final status post-restart: `FIXED - Confirmed Disabled` or `FIXED - Restart pending`

### Fixed
- Registry write was succeeding but `Get-NetAdapterPowerManagement` still reported `Enabled` because the driver had not reloaded the registry value

---

## [3.1.0] - 2025-01-27

### Fixed
- Coalescing key names hardcoded inside each scriptblock — `ArgumentList` array serialization over WinRM PS5.1 was silently passing empty values, causing no registry keys to be matched or written

---

## [3.0.0] - 2025-01-27

### Changed
- Replaced `Set-NetAdapterPowerManagement` with **direct registry writes** — the cmdlet does not map to the correct registry key on all Intel driver variants
- Key detection is now driver-aware: `*PacketCoalescing` (AX201/AX211), `*D0PacketCoalescing` (AX200/AC9560), `DMACoalescing` (I225/I219)

### Fixed
- `Set-NetAdapterPowerManagement -D0PacketCoalescing Disabled` had no effect on Intel AX201 (key name mismatch: actual key is `*PacketCoalescing`, not `*D0PacketCoalescing`)

---

## [2.3.0] - 2025-01-27

### Added
- Two-pass WinRM execution: Pass 1 applies fix (session may drop), Pass 2 reconnects and reads status

### Fixed
- Script output (rows) was lost when WinRM session dropped during `Set-NetAdapterPowerManagement` on WiFi-only targets — return value never reached the caller

---

## [2.2.0] - 2025-01-27

### Fixed
- Removed `Where-Object { $_.Status -ne "Not Present" }` filter — adapter status enumeration differs between local and WinRM sessions; filter was silently excluding adapters on remote targets

---

## [2.1.0] - 2025-01-27

### Fixed
- `$using:` syntax invalid in PS5.1 — replaced with `Invoke-FixOnComputer` helper function for PS5.1 sequential path
- `switch` statement now explicitly handles `"Unsupported"` string value returned by `Get-NetAdapterPowerManagement`

---

## [2.0.0] - 2025-01-27

### Added
- Remote mode via WinRM: `-ComputerName` and `-OUPath` parameters
- Auto-detection of execution mode (local vs remote) from parameters
- PS7+ parallel execution via `ForEach-Object -Parallel` with `-ThrottleLimit`
- PS5.1 sequential fallback
- `-ExportCsv` parameter for result export
- `-WhatIf` dry-run support

### Changed
- Single unified script replacing separate local/remote scripts

---

## [1.0.0] - 2025-01-27

### Added
- Initial release
- Local GPO startup script disabling `D0PacketCoalescing` via `Set-NetAdapterPowerManagement`
- Event Viewer logging (source: `9Lives-NetworkPerf`, EventID 1001)
- `Check-NetworkPowerStatus.ps1` audit script with remote `-ComputerName` support
- `README.md`, `LICENSE` (MIT), `.gitignore`
