MSP Technician IDE - Professional Edition
Show Image
Show Image
Show Image

A comprehensive Windows troubleshooting toolkit built for MSP technicians, IT professionals, and help desk teams. This PowerShell-based GUI application provides one-click access to the most common diagnostic and repair operations encountered in enterprise support environments.

🚀 Features
System Repair & Health
SFC Scan - Automated System File Checker with progress tracking
Pending Reboot Detection - Checks multiple registry locations for reboot flags
CHKDSK Scanning - Non-destructive disk integrity checks
User Profile & M365
Teams Cache Cleaner - Fixes freezing, login issues, and sync problems
Outlook Profile Repair - Resolves "Enter Password" prompts and corruption
OneDrive Sync Fixes - Resets sync engine for stuck files
Mapped Drive Diagnostics - Tests connectivity to network shares
Advanced Networking
Network Traffic Monitor - Identifies top bandwidth-consuming processes
WiFi Password Recovery - Extracts saved wireless credentials
Network Stack Reset - Nuclear option for persistent connectivity issues
IP Address Display - Shows both local and public IP addresses
Printer Troubleshooting
Print Spooler Management - Clears stuck jobs and restarts service
Printer Inventory - Lists all installed printers with status
Ghost Printer Removal - Automatically removes offline/error printers
Disk & Storage Health
Large File Finder - Locates files over 500MB consuming disk space
SMART Status Checker - Early warning system for failing drives
Temp File Cleanup - Automated cleanup of Windows and user temp folders
Disk Health Dashboard - Physical disk health monitoring
Windows Update Fixes
Update Component Reset - Fixes 80% of Windows Update failures
Update History Viewer - Shows last 5 installed updates with dates
Performance & Diagnostics
Memory Usage Report - Identifies RAM-hogging processes
Blue Screen Detection - Scans event logs for recent BSOD errors
Battery Health Report - Generates detailed battery degradation analysis (laptops)
Domain & Active Directory
Domain Connection Test - Verifies secure channel to domain controller
Group Policy Updater - Forces immediate policy refresh
Last Logon Tracker - Shows last user login time and username
System Information
Complete System Dashboard - One-click overview of computer name, domain, OS, hardware specs, and uptime
📋 Requirements
Windows 10/11 or Windows Server 2016+
PowerShell 5.1 or higher
Administrator privileges (script auto-elevates)
🔧 Installation
Download the script:
powershell
   # Clone the repository
   git clone https://github.com/yourusername/msp-technician-ide.git
   cd msp-technician-ide
Run the script:
powershell
   # Right-click ITAssist.ps1 and select "Run with PowerShell"
   # OR run from PowerShell:
   .\ITAssist.ps1
Accept UAC prompt - The script will automatically request administrator privileges
💡 Usage
First Launch
The application will automatically elevate to administrator mode and display a dark-themed GUI with all diagnostic tools organized into categories.

Running Operations
Hover over any button to see a detailed tooltip explaining what it does
Watch the progress bar during long-running operations (SFC, CHKDSK, file scans)
Monitor the output console on the right for real-time status updates
Color-coded logging:
🟢 Green (PASS) - Operation successful
🔴 Red (FAIL) - Operation failed or issue detected
🟡 Yellow (WARN) - Warning or non-critical issue
⚪ White (INFO) - Informational messages
Saving Logs
Click "Save Log to Desktop" to export a timestamped troubleshooting report for documentation or ticket notes.

🛡️ Safety Features
Comprehensive error handling - Every operation wrapped in try-catch blocks
Graceful fallbacks - Alternative methods when modern cmdlets aren't available
Path validation - Checks file/folder existence before operations
Service state verification - Validates services before stop/start operations
Domain detection - Skips domain tests on workgroup computers
Progress indicators - Visual feedback for long-running tasks
📊 Common Use Cases
"Computer is slow"
Run Memory-Hogging Processes to identify RAM issues
Use Large File Finder to check disk space
Run Clean Temp Files to free up storage
Check Network Traffic Monitor for bandwidth hogs
"Can't connect to network"
Test with Public & Local IP display
Run Fix Disconnected Mapped Drives
Use Network Stack Reset as last resort
"Outlook keeps asking for password"
Run Fix Outlook Profile
Restart Outlook to recreate profile
"Printer won't print"
Use Clear Print Spooler
Check List All Installed Printers for status
Run Remove Offline/Ghost Printers
"Windows Update failed"
Run Reset Windows Update Components
Reboot system
Retry Windows Update
🎯 Built For MSPs
Based on real-world ticket data from MSP help desks, this tool addresses:

Password resets (integrated with M365 workflows)
Outlook profile corruption (#1 complaint)
OneDrive sync issues (extremely common)
Printer problems ("just printer" is enough said)
Network connectivity ("internet doesn't work")
Slow computer complaints
Domain trust relationship errors
🤝 Contributing
Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

Areas for Expansion
BitLocker key recovery
Office 365 license checking
Credential manager diagnostics
Advanced Active Directory tools
Remote desktop troubleshooting
VPN diagnostics
📝 License
This project is licensed under the MIT License - see the LICENSE file for details.

⚠️ Disclaimer
This tool performs system-level operations and requires administrator privileges. Always test in a non-production environment first. The authors are not responsible for any system changes or data loss. Use at your own risk.

🙏 Acknowledgments
Built with feedback from real MSP technicians
Inspired by common help desk tickets from r/msp
Designed for efficiency and ease of use in high-pressure support environments
📞 Support
For issues, questions, or feature requests, please open an issue on GitHub.

Made with ❤️ for IT professionals who deserve better tools

