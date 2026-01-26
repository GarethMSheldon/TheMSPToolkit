#Requires -RunAsAdministrator
<#
.SYNOPSIS
MSP Technician Toolkit — v3.0.1 snapshot
.NOTES
Version: 3.0.1
Date: 2025-11-22
#>

$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (Test-Path (Join-Path $PSScriptRoot 'change_3.0.1.ps1')) { . (Join-Path $PSScriptRoot 'change_3.0.1.ps1') }
