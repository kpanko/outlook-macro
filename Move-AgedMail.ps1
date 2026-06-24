#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(1, 3650)]
    [int]$AgeDays = 40,

    [string]$IncubateFolderPath = "GTD-incubate",
    [string]$CleanupInboxFolderPath = "cleanup/inbox",
    [string]$CleanupSentFolderPath = "cleanup/sent",
    [string]$SentFolderId = "sentitems",
    [switch]$SkipConnect,
    [switch]$InstallMissingGraphModule
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$modulePath = Join-Path $PSScriptRoot "Move-AgedMail.psm1"
Import-Module -Name $modulePath -Force

Invoke-MoveAgedMail `
    -AgeDays $AgeDays `
    -IncubateFolderPath $IncubateFolderPath `
    -CleanupInboxFolderPath $CleanupInboxFolderPath `
    -CleanupSentFolderPath $CleanupSentFolderPath `
    -SentFolderId $SentFolderId `
    -SkipConnect:$SkipConnect `
    -InstallMissingGraphModule:$InstallMissingGraphModule
