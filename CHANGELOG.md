# MSP Technician Toolkit – Professional Edition  
*Changelog*

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased / Future]

Planned for **v3.1.4**:
- Cancellable background tasks (using `CancellationPending`).
- "Copy PowerShell command" feature for each tool.
- System Health Dashboard (one-click triage).
- Standardized confirmations for all destructive operations.
- Dynamic resize handling using stored final Y position.

---

## [3.1.3] – 2026-01-25

### Fixed
- Resolved section numbering conflict: "LOG MANAGEMENT" renamed from Section 5 → **Section 10** to avoid overlap with "DISK AND STORAGE HEALTH".
- Eliminated hardcoded button container height (`5500`) — now dynamically sized based on content.
- Moved Teams cache clear, Outlook profile reset, and OneDrive sync fix to **background tasks** to prevent UI freezing.
- Improved event handler parameter clarity in `FormClosing` (`$sender` → `$frmSender`).

### Changed
- Final button container height computed after all controls are added for optimal scrolling.
- All profile-related operations now consistently use `Start-BackgroundTask`.

---

## [3.1.2] – 2026-01-23

### Added
- Full **thread-safe logging system** with color-coded status (PASS/WARN/FAIL) and console fallback.
- **Background task engine** using `System.ComponentModel.BackgroundWorker` with lifecycle tracking.
- **Safe external process execution** via `Invoke-ExternalCommand` (captures exit code, stdout/stderr, auto-cleanup).
- **DPI awareness** for high-resolution displays.
- **Auto-elevation** to Administrator with preserved script path.
- **Form closing handler** that cancels background tasks and kills orphaned processes.
- **Confirmation prompts** for destructive actions (e.g., Outlook reset).
- **Status bar** showing PC name, user, domain, and admin status.

### Improved
- All long-running operations (SFC, DISM, CHKDSK, Battery Report) run in background.
- Robust error handling in every diagnostic function.
- Modern dark-themed UI with hover effects and consistent styling.
- Tooltips with 8-second auto-pop delay for better usability.

### Internal
- Modular architecture with reusable functions (`New-ToolkitButton`, `Show-Confirmation`, etc.).
- Fallback event registration using `Register-ObjectEvent` for compatibility.
- Thread-safe resource tracking with monitor-based locking.
- Centralized external command invocation with process lifecycle management.

---

## [3.1.1] – 2026-01-19

### Fixed
- Fixed crash on startup with corrupted preferences.
- Improved logging and telemetry anonymization.

### Added
- Telemetry opt-out option added to settings.
- Installer now validates checksum before unpack.

### Changed
- Minor UI polish and translations updated.

---

## [3.1.0] – 2026-01-15

### Added
- **Network Traffic Monitor** (per-process bandwidth tracking).
- **WiFi Password Recovery** tool.
- **Battery Health Report** generator.
- **Last Logon Tracker** for security auditing.
- **Complete System Dashboard** for at-a-glance diagnostics.
- **Auto-elevation** on launch for administrative tasks.
- **Dark-themed GUI** with modern aesthetics.

### Improved
- Enhanced error handling across all modules.
- Added tooltips for better user guidance.
- Progress indicators for long-running operations.
- Color-coded logs for improved readability.

---

## [3.0.2] – 2025-12-10

### Fixed
- Print spooler service detection fix for Windows 11 24H2.
- OneDrive for Business sync reset compatibility.
- Prevent connectivity loss during network stack reset.
- Eliminated SMART false positives on NVMe drives.

---

## [3.0.1] – 2025-11-22

### Added
- Custom size threshold for large file finder.
- **Disk health dashboard** with SMART visualization.
- **Blue screen detection** and event-log scanning.

### Fixed
- Fixed temp file cleanup routine.
- Fixed update history date parsing.
- Increased domain test timeout for slower networks.
- Enhanced credential handling and UAC improvements.

---

## [3.0.0] – 2025-10-05

### Added
- **60+ diagnostic and repair functions** covering all major system areas.
- **GUI-based interface** replacing CLI for better usability.
- **Progress bar tracking** for long-running operations.
- **Log export functionality** for documentation and auditing.
- **M365 integration tools** (Teams, Outlook, OneDrive).
- **Printer troubleshooting suite** with comprehensive diagnostics.
- **Advanced networking diagnostics** including traffic monitoring.
- **Windows Update repair toolkit** with component reset.

### Changed
- Complete architectural overhaul from CLI to GUI.
- Unified error handling and logging framework.

---

## [2.5.3] – 2025-08-18

### Added
- **Ghost printer removal** functionality.
- **Pending reboot detection** across multiple registry keys.
- **Public IP address display** tool.

### Fixed
- Fixed SFC progress visualization.
- Fixed memory usage percentage calculation.
- Fixed group policy immediate refresh.

---

## [2.5.0] – 2025-07-01

### Added
- **Outlook profile repair** tool with registry cleanup.
- **OneDrive sync reset** functionality.
- **Mapped drive connectivity diagnostics**.
- **Windows Update component reset** tool.

### Improved
- Enhanced error messages with actionable suggestions.
- Added fallback methods for older Windows versions.

---

## [2.0.1] – 2025-05-10

### Added
- **Print spooler management** tools.
- **Printer inventory listing** with status.
- **Temp file cleanup** automation.
- **Update history viewer**.

### Fixed
- Fixed CHKDSK execution on system drive.
- Fixed memory report crashes for systems with >64GB RAM.
- Fixed network stack reset residual configuration issues.

---

## [2.0.0] – 2025-03-15

### Added
- **System File Checker** automation with progress tracking.
- **CHKDSK scanning** capability with read-only mode.
- **Teams cache cleaner** supporting classic and new Teams.
- **Network stack reset** with comprehensive cleanup.
- **Print spooler clearing** with service restart.
- **Memory-hogging process identification** (top 5).
- **SMART status checking** for disk health.
- **Domain connection testing** with secure channel validation.
- **Group policy updater** with force refresh.
- **Basic logging functionality** with color-coded output.

### Changed
- Initial public release with GUI migration.
- Established baseline feature set for MSP workflows.

---

## [1.0.0] – 2025-09-01

### Added
- Initial release of MSP Technician Toolkit core features.
- Foundation for diagnostic and repair automation.

---

## Version History Summary

| Version | Date       | Type        | Highlights |
|---------|------------|-------------|------------|
| 3.1.3   | 2026-01-26 | Patch       | UI fixes, dynamic sizing, background task improvements |
| 3.1.2   | 2026-01-23 | Minor       | Thread safety, background tasks, modern UI |
| 3.1.1   | 2026-01-19 | Patch       | Stability fixes, telemetry controls |
| 3.1.0   | 2026-01-15 | Minor       | Network monitor, WiFi recovery, dashboard |
| 3.0.2   | 2025-12-10 | Patch       | Windows 11 compatibility |
| 3.0.1   | 2025-11-22 | Patch       | Disk health, BSOD detection |
| 3.0.0   | 2025-10-05 | Major       | 60+ tools, GUI interface, M365 integration |
| 2.5.3   | 2025-08-18 | Patch       | Printer tools, reboot detection |
| 2.5.0   | 2025-07-01 | Minor       | Profile repair, drive diagnostics |
| 2.0.1   | 2025-05-10 | Patch       | Print management, temp cleanup |
| 2.0.0   | 2025-03-15 | Major       | Initial public release with GUI |
| 1.0.0   | 2025-09-01 | Major       | Initial toolkit release |

---

## Semantic Versioning

This project follows [Semantic Versioning](https://semver.org/):
- **MAJOR** version for incompatible API changes
- **MINOR** version for backwards-compatible functionality additions
- **PATCH** version for backwards-compatible bug fixes

---

## Contributing

For feature requests, bug reports, or contributions, please contact the MSP Solutions Team.

---

*Last updated: 2026-01-26*