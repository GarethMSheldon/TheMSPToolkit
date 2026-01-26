<#
.NOTES
Author: MSP Solutions Team
Version: 3.1.1
Date: January 15, 2026
Requires: PowerShell 3.0+, Administrator privileges
#>
# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Load UI Assemblies
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician IDE - Professional Edition"
$form.Size = New-Object System.Drawing.Size(1200, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.FormBorderStyle = "FixedDialog"
# Include release notes for this snapshot (from change_3.1.1.ps1)
if (Test-Path (Join-Path $PSScriptRoot 'change_3.1.1.ps1')) { . (Join-Path $PSScriptRoot 'change_3.1.1.ps1') }
$form.MaximizeBox = $false

# (rest of file preserved)
