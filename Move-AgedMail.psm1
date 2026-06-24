#Requires -Version 7.0

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Initialize-GraphAuthenticationModule {
    [CmdletBinding()]
    param(
        [switch]$InstallMissingModule
    )

    $module = Get-Module `
        -ListAvailable `
        -Name Microsoft.Graph.Authentication |
        Select-Object -First 1

    if (-not $module) {
        if (-not $InstallMissingModule) {
            $message = "Microsoft.Graph.Authentication is not installed. " +
                "Run again with -InstallMissingGraphModule, or run: " +
                "Install-Module Microsoft.Graph.Authentication " +
                "-Scope CurrentUser"
            throw $message
        }

        $nuGetProvider = Get-PackageProvider `
            -Name NuGet `
            -ListAvailable `
            -ErrorAction SilentlyContinue

        if (-not $nuGetProvider) {
            Install-PackageProvider `
                -Name NuGet `
                -Scope CurrentUser `
                -Force | Out-Null
        }

        Install-Module `
            -Name Microsoft.Graph.Authentication `
            -Repository PSGallery `
            -Scope CurrentUser `
            -Force `
            -AllowClobber
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }
    catch {
        $modulePath = $env:PSModulePath -split [IO.Path]::PathSeparator
        $modulePath = $modulePath -join "; "
        $message = "Microsoft.Graph.Authentication could not be loaded. " +
            "Try running this script with -InstallMissingGraphModule so " +
            "it can install the module. Original error: " +
            $_.Exception.Message +
            " PSModulePath: $modulePath"
        throw $message
    }
}

function Connect-MailGraph {
    [CmdletBinding()]
    param(
        [switch]$SkipConnect
    )

    if ($SkipConnect) {
        return
    }

    $connectArgs = @{
        Scopes    = @("Mail.ReadWrite")
        NoWelcome = $true
    }

    Connect-MgGraph @connectArgs
}

function ConvertTo-GraphPathSegment {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    return [uri]::EscapeDataString($Value)
}

function ConvertTo-GraphRequestUri {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    if ($Uri -match "^[a-z][a-z0-9+.-]*://") {
        return $Uri
    }

    if ($Uri -match "^/(v1\.0|beta)(/|$)") {
        return $Uri
    }

    if ($Uri -match "^(v1\.0|beta)(/|$)") {
        return "/$Uri"
    }

    if ($Uri.StartsWith("/")) {
        return "/v1.0$Uri"
    }

    return "/v1.0/$Uri"
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory)]
        [psobject]$InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if (-not $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-NullableDateTime {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    try {
        return [datetime]$Value
    }
    catch {
        return $null
    }
}

function Get-MessageAgeDate {
    param(
        [Parameter(Mandatory)]
        [psobject]$Message,

        [datetime]$FallbackDate = (Get-Date)
    )

    $sentDateTime = Get-ObjectPropertyValue `
        -InputObject $Message `
        -Name sentDateTime
    $sentDateTime = ConvertTo-NullableDateTime -Value $sentDateTime

    if ($sentDateTime) {
        return $sentDateTime
    }

    $createdDateTime = Get-ObjectPropertyValue `
        -InputObject $Message `
        -Name createdDateTime
    $createdDateTime = ConvertTo-NullableDateTime -Value $createdDateTime

    if ($createdDateTime) {
        return $createdDateTime
    }

    return $FallbackDate
}

function Test-MessageIsAged {
    param(
        [Parameter(Mandatory)]
        [psobject]$Message,

        [Parameter(Mandatory)]
        [datetime]$Cutoff
    )

    $messageDate = Get-MessageAgeDate `
        -Message $Message `
        -FallbackDate (Get-Date)

    return $messageDate -lt $Cutoff
}

function Invoke-GraphPagedGet {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $nextUri = $Uri
    while ($nextUri) {
        $requestUri = ConvertTo-GraphRequestUri -Uri $nextUri
        $response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $requestUri `
            -OutputType PSObject

        foreach ($item in @($response.value)) {
            $item
        }

        $nextUri = $null
        if ($response.PSObject.Properties.Name -contains "@odata.nextLink") {
            $nextUri = $response.'@odata.nextLink'
        }
    }
}

function Get-MailFolderMessage {
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolderId,

        [Parameter(Mandatory)]
        [string]$FilterText
    )

    $escapedSourceId = ConvertTo-GraphPathSegment -Value $SourceFolderId
    $filter = [uri]::EscapeDataString($FilterText)
    $query = (
        "`$top=100&" +
        "`$select=id,subject,sentDateTime,createdDateTime&" +
        "`$filter=$filter"
    )

    $messagesUri = "/me/mailFolders/$escapedSourceId/messages?$query"

    return @(Invoke-GraphPagedGet -Uri $messagesUri)
}

function Get-MailFolderChild {
    param(
        [string]$ParentFolderId
    )

    $query = "`$top=100&`$select=id,displayName,parentFolderId"
    if ([string]::IsNullOrWhiteSpace($ParentFolderId)) {
        return @(Invoke-GraphPagedGet -Uri "/me/mailFolders?$query")
    }

    $escapedParentId = ConvertTo-GraphPathSegment -Value $ParentFolderId
    $childFoldersUri = "/me/mailFolders/$escapedParentId/childFolders?$query"

    return @(Invoke-GraphPagedGet -Uri $childFoldersUri)
}

function Resolve-MailFolderPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $currentFolderId = $null
    $parts = @(
        $Path -split "[\\/]" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($parts.Count -eq 0) {
        throw "Folder path cannot be empty."
    }

    foreach ($part in $parts) {
        $children = Get-MailFolderChild -ParentFolderId $currentFolderId
        $matchingFolders = @(
            $children | Where-Object { $_.displayName -ieq $part }
        )

        if ($matchingFolders.Count -eq 0) {
            throw "Could not find mail folder '$part' while resolving '$Path'."
        }

        if ($matchingFolders.Count -gt 1) {
            $message = "Found multiple mail folders named '$part' while " +
                "resolving '$Path'. Rename one or use a unique path."
            throw $message
        }

        $currentFolderId = $matchingFolders[0].id
    }

    return $currentFolderId
}

function Get-AgedMessage {
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolderId,

        [Parameter(Mandatory)]
        [datetime]$Cutoff
    )

    $cutoffText = $Cutoff.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $filters = @(
        "sentDateTime lt $cutoffText"
        "createdDateTime lt $cutoffText"
    )

    $messagesById = [ordered]@{}
    foreach ($filter in $filters) {
        $messages = Get-MailFolderMessage `
            -SourceFolderId $SourceFolderId `
            -FilterText $filter

        foreach ($message in $messages) {
            if (-not $messagesById.Contains($message.id)) {
                $messagesById[$message.id] = $message
            }
        }
    }

    return @(
        $messagesById.Values |
            Where-Object { Test-MessageIsAged -Message $_ -Cutoff $Cutoff }
    )
}

function Move-AgedMessage {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Label,

        [Parameter(Mandatory)]
        [string]$SourceFolderId,

        [Parameter(Mandatory)]
        [string]$DestinationFolderId,

        [Parameter(Mandatory)]
        [datetime]$Cutoff
    )

    $messages = @(
        Get-AgedMessage -SourceFolderId $SourceFolderId -Cutoff $Cutoff
    )
    $moved = 0
    $failed = 0

    foreach ($message in $messages) {
        $subject = $message.subject
        if ([string]::IsNullOrWhiteSpace($subject)) {
            $subject = "(no subject)"
        }

        if ($PSCmdlet.ShouldProcess("$Label :: $subject", "Move message")) {
            $messageId = ConvertTo-GraphPathSegment -Value $message.id
            $moveUri = ConvertTo-GraphRequestUri `
                -Uri "/me/messages/$messageId/move"
            $body = @{
                destinationId = $DestinationFolderId
            } | ConvertTo-Json -Compress

            try {
                Invoke-MgGraphRequest `
                    -Method POST `
                    -Uri $moveUri `
                    -Body $body `
                    -ContentType "application/json" `
                    -OutputType PSObject | Out-Null

                $moved++
            }
            catch {
                $failed++
                Write-Warning (
                    "Failed to move '$subject' from '$Label': " +
                    $_.Exception.Message
                )
            }
        }
    }

    [pscustomobject]@{
        Source     = $Label
        Candidates = $messages.Count
        Moved      = $moved
        Failed     = $failed
    }
}

<#
.SYNOPSIS
Moves aged Outlook messages into cleanup folders.

.DESCRIPTION
Connects to Microsoft Graph, finds messages older than the configured
threshold, and moves them from the incubate and Sent Items folders into
cleanup folders.

.PARAMETER AgeDays
Minimum message age, in days, before a message is eligible to move.

.PARAMETER SkipConnect
Skips Microsoft Graph connection setup. Useful for tests or existing sessions.

.PARAMETER InstallMissingGraphModule
Installs or repairs the Microsoft.Graph.Authentication module if needed.
#>
function Invoke-MoveAgedMail {
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

    Initialize-GraphAuthenticationModule `
        -InstallMissingModule:$InstallMissingGraphModule
    Connect-MailGraph -SkipConnect:$SkipConnect

    $cutoff = (Get-Date).Date.AddDays(-1 * $AgeDays).ToUniversalTime()

    $incubateFolderId = Resolve-MailFolderPath -Path $IncubateFolderPath
    $cleanupInboxFolderId = Resolve-MailFolderPath `
        -Path $CleanupInboxFolderPath
    $cleanupSentFolderId = Resolve-MailFolderPath -Path $CleanupSentFolderPath

    $results = @(
        Move-AgedMessage `
            -Label $IncubateFolderPath `
            -SourceFolderId $incubateFolderId `
            -DestinationFolderId $cleanupInboxFolderId `
            -Cutoff $cutoff
        Move-AgedMessage `
            -Label "Sent Items" `
            -SourceFolderId $SentFolderId `
            -DestinationFolderId $cleanupSentFolderId `
            -Cutoff $cutoff
    )

    $results | Format-Table -AutoSize

    $totalCandidates = ($results |
        Measure-Object -Property Candidates -Sum).Sum
    $totalMoved = ($results | Measure-Object -Property Moved -Sum).Sum
    $totalFailed = ($results | Measure-Object -Property Failed -Sum).Sum

    if ($WhatIfPreference) {
        Write-Output "Would move $totalCandidates message(s)."
        return
    }

    Write-Output "Moved $totalMoved message(s). Failed $totalFailed."
}

Export-ModuleMember -Function Invoke-MoveAgedMail
