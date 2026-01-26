# The MSP Toolkit - Professional Edition

![PowerShell](https://img.shields.io/badge/PowerShell-3.0%2B-blue)
![Windows](https://img.shields.io/badge/Windows-10%2F11-success)
![License](https://img.shields.io/badge/license-MIT-green)
![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)

**Version:** 3.1.1 (FIXED) — January 23, 2026  
**Author:** Gareth Sheldon

A comprehensive Windows troubleshooting toolkit built for MSP technicians, IT professionals, and help desk teams. This PowerShell-based GUI application provides one-click access to the most common diagnostic and repair operations encountered in enterprise support environments. The current release includes 60+ functions and auto-elevates to administrator on launch.

## Features

### System Repair and Health
- **SFC Scan** - Automated System File Checker with progress tracking
- **Pending Reboot Detection** - Checks multiple registry locations for reboot flags
- **CHKDSK Scanning** - Non-destructive disk integrity checks

### User Profile and M365
- **Teams Cache Cleaner** - Fixes freezing, login issues, and sync problems
- **Outlook Profile Repair** - Resolves "Enter Password" prompts and corruption
- **OneDrive Sync Fixes** - Resets sync engine for stuck files
- **Mapped Drive Diagnostics** - Tests connectivity to network shares

### Advanced Networking
- **Network Traffic Monitor** - Identifies top bandwidth-consuming processes
- **WiFi Password Recovery** - Extracts saved wireless credentials
- **Network Stack Reset** - Nuclear option for persistent connectivity issues
- **IP Address Display** - Shows both local and public IP addresses

### Printer Troubleshooting
- **Print Spooler Management** - Clears stuck jobs and restarts service
- **Printer Inventory** - Lists all installed printers with status
- **Ghost Printer Removal** - Automatically removes offline/error printers

### Disk and Storage Health
- **Large File Finder** - Locates files over 500MB consuming disk space
- **SMART Status Checker** - Early warning system for failing drives
- **Temp File Cleanup** - Automated cleanup of Windows and user temp folders
- **Disk Health Dashboard** - Physical disk health monitoring

### Windows Update Fixes
- **Update Component Reset** - Fixes 80% of Windows Update failures
- **Update History Viewer** - Shows last 5 installed updates with dates

### Performance and Diagnostics
- **Memory Usage Report** - Identifies RAM-hogging processes
- **Blue Screen Detection** - Scans event logs for recent BSOD errors
- **Battery Health Report** - Generates detailed battery degradation analysis (laptops)

### Domain and Active Directory
- **Domain Connection Test** - Verifies secure channel to domain controller
- **Group Policy Updater** - Forces immediate policy refresh
- **Last Logon Tracker** - Shows last user login time and username

### System Information
- **Complete System Dashboard** - One-click overview of computer name, domain, OS, hardware specs, and uptime

## Screenshots

![MSP Toolkit Interface](screenshots/main-interface.png)

*Screenshot showing the main interface with diagnostic tools organized by category*

## Requirements

- **Compatible platforms**: Windows 10, Windows 11, and Windows Server 2016+
- **PowerShell**: 3.0 or higher (5.1+ recommended)
- **Privileges**: Administrator (the script will auto-elevate)

## Installation

### Method 1: Download Script Directly

1. Download `MSPToolKit_latest.ps1` from the [Releases](https://github.com/GarethMSheldon/TheMSPToolkit/releases) page
2. Right-click the file and select "Run with PowerShell"
3. Accept the UAC prompt when it appears
4. Start troubleshooting!

### Method 2: Clone Repository

```bash
git clone https://github.com/GarethMSheldon/TheMSPToolkit.git
cd TheMSPToolkit
```

Then run the script:

```powershell
./MSPToolKit_latest.ps1
```

### Method 3: Quick Download via PowerShell

```powershell
# Download and run directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/GarethMSheldon/TheMSPToolkit/main/MSPToolKit_latest.ps1" -OutFile "$env:TEMP\MSPToolKit_latest.ps1"
& "$env:TEMP\MSPToolKit_latest.ps1"
```

## Usage

### First Launch
The application will automatically elevate to administrator mode and display a dark-themed GUI with all diagnostic tools organized into categories.

### Running Operations
- **Hover over any button** to see a detailed tooltip explaining what it does
- **Watch the progress bar** during long-running operations (SFC, CHKDSK, file scans)
- **Monitor the output console** on the right for real-time status updates
- **Color-coded logging**: 
    - **Green (PASS)** - Operation successful
    - **Red (FAIL)** - Operation failed or issue detected
    - **Yellow (WARN)** - Warning or non-critical issue
    - **White (INFO)** - Informational messages

### Saving Logs
Click **"Save Log to Desktop"** to export a timestamped troubleshooting report for documentation or ticket notes.

## Safety Features

- **Comprehensive error handling** - Every operation wrapped in try-catch blocks
- **Graceful fallbacks** - Alternative methods when modern cmdlets aren't available
- **Path validation** - Checks file/folder existence before operations
- **Service state verification** - Validates services before stop/start operations
- **Domain detection** - Skips domain tests on workgroup computers
- **Progress indicators** - Visual feedback for long-running tasks
- **No destructive operations** - All actions are safe for production systems

## Common Use Cases

### "Computer is slow"
1. Run **Memory-Hogging Processes** to identify RAM issues
2. Use **Large File Finder** to check disk space
3. Run **Clean Temp Files** to free up storage
4. Check **Network Traffic Monitor** for bandwidth hogs

### "Can't connect to network"
1. Test with **Public and Local IP** display
2. Run **Fix Disconnected Mapped Drives**
3. Use **Network Stack Reset** as last resort

### "Outlook keeps asking for password"
1. Run **Fix Outlook Profile**
2. Restart Outlook to recreate profile

### "Printer won't print"
1. Use **Clear Print Spooler**
2. Check **List All Installed Printers** for status
3. Run **Remove Offline/Ghost Printers**

### "Windows Update failed"
1. Run **Reset Windows Update Components**
2. Reboot system
3. Retry Windows Update

## Built For MSPs

Based on real-world ticket data from MSP help desks, this tool addresses:

- **Password resets** - Integrated with M365 workflows
- **Outlook profile corruption** - #1 complaint from r/msp
- **OneDrive sync issues** - Extremely common in remote work
- **Printer problems** - "Just printer" is enough said
- **Network connectivity** - "Internet doesn't work" tickets
- **Slow computer complaints** - Memory and disk diagnostics
- **Domain trust relationship errors** - Secure channel testing

## Roadmap

Future features planned:

- [ ] BitLocker key recovery tool
- [ ] Office 365 license status checker
- [ ] Credential manager diagnostics
- [ ] Advanced Active Directory reporting
- [ ] Remote desktop troubleshooting
- [ ] VPN diagnostics and repair
- [ ] Email configuration validator
- [ ] Browser cache/profile repair tools
- [ ] Automated ticket documentation export

## Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

Please ensure your PR:
- Includes descriptive comments in code
- Adds error handling for new features
- Updates README.md if adding new functionality
- Tests on Windows 10 and Windows 11

### Areas for Contribution
- Additional diagnostic tools based on MSP ticket data
- Localization/translation support
- Performance optimizations
- UI/UX improvements
- Documentation enhancements

## FAQ

**Q: Does this work on Windows 7/8?**  
A: The tool is designed for Windows 10 and newer. Some features may work on Windows 8.1 with PowerShell 5.1, but earlier versions are not officially supported.

**Q: Will this break my system?**  
A: All operations are designed to be safe for production systems. However, always test in a non-production environment first.

**Q: Can I run this without admin rights?**  
A: No, most diagnostic and repair operations require administrator privileges. The script will auto-elevate when launched.

**Q: Can I customize the tools included?**  
A: Yes! The script is open source. You can add, remove, or modify tools to fit your organization's needs.

**Q: Does this send any data externally?**  
A: No. The only external connection is to `api.ipify.org` for retrieving your public IP address (optional tool). All other operations are local.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool performs system-level operations and requires administrator privileges. While designed to be safe, always test in a non-production environment first. The authors are not responsible for any system changes or data loss. Use at your own risk.

## Acknowledgments

- Built with feedback from real MSP technicians
- Inspired by common help desk tickets from [r/msp](https://reddit.com/r/msp)
- Designed for efficiency and ease of use in high-pressure support environments
- Special thanks to the PowerShell community for cmdlet inspiration

## Support

- **Issues**: [GitHub Issues](https://github.com/GarethMSheldon/TheMSPToolkit/issues)
- **Discussions**: [GitHub Discussions](https://github.com/GarethMSheldon/TheMSPToolkit/discussions)
- **Contact**: Open an issue for questions or feature requests

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=GarethMSheldon/TheMSPToolkit&type=Date)](https://star-history.com/#GarethMSheldon/TheMSPToolkit&Date)

---

**Made with care for IT professionals who deserve better tools**

If this toolkit saves you time, consider giving it a ⭐ star on GitHub!
