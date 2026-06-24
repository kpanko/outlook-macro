#Requires -Version 7.0

$script:ModulePath = Join-Path $PSScriptRoot "Move-AgedMail.psm1"
Import-Module -Name $script:ModulePath -Force

Describe "Move-AgedMail wrapper parameters" {
    It "requires a positive AgeDays value" {
        $scriptPath = Join-Path $PSScriptRoot "Move-AgedMail.ps1"
        $command = Get-Command $scriptPath
        $attributes = $command.Parameters["AgeDays"].Attributes
        $range = $attributes | Where-Object {
            $_.TypeId.Name -eq "ValidateRangeAttribute"
        }

        $range.MinRange | Should -Be 1
        $range.MaxRange | Should -Be 3650
    }
}

InModuleScope Move-AgedMail {
    Describe "Test-MessageIsAged" {
        It "uses sentDateTime before createdDateTime" {
            $cutoff = [datetime]"2026-01-10T00:00:00Z"
            $message = [pscustomobject]@{
                id              = "recent-sent"
                sentDateTime    = "2026-01-11T00:00:00Z"
                createdDateTime = "2020-01-01T00:00:00Z"
            }

            Test-MessageIsAged -Message $message -Cutoff $cutoff |
                Should -BeFalse
        }

        It "falls back to createdDateTime when sentDateTime is missing" {
            $cutoff = [datetime]"2026-01-10T00:00:00Z"
            $message = [pscustomobject]@{
                id              = "created-old"
                sentDateTime    = $null
                createdDateTime = "2020-01-01T00:00:00Z"
            }

            Test-MessageIsAged -Message $message -Cutoff $cutoff |
                Should -BeTrue
        }
    }

    Describe "ConvertTo-GraphRequestUri" {
        It "adds the Graph API version to relative me paths" {
            ConvertTo-GraphRequestUri -Uri "/me/mailFolders" |
                Should -Be "/v1.0/me/mailFolders"
        }

        It "keeps versioned relative paths unchanged" {
            ConvertTo-GraphRequestUri -Uri "/v1.0/me/mailFolders" |
                Should -Be "/v1.0/me/mailFolders"
        }

        It "keeps absolute next links unchanged" {
            $nextLink = "https://graph.microsoft.com/v1.0/me/mailFolders"

            ConvertTo-GraphRequestUri -Uri $nextLink |
                Should -Be $nextLink
        }
    }

    Describe "Get-AgedMessage" {
        BeforeEach {
            $script:requestedUris = @()

            Mock Invoke-GraphPagedGet {
                param(
                    [string]$Uri
                )

                $decodedUri = [uri]::UnescapeDataString($Uri)
                $script:requestedUris += $decodedUri

                if ($decodedUri -match "\`$filter=sentDateTime lt") {
                    return @(
                        [pscustomobject]@{
                            id              = "sent-old"
                            subject         = "sent old"
                            sentDateTime    = "2020-01-01T00:00:00Z"
                            createdDateTime = "2020-01-01T00:00:00Z"
                        }
                    )
                }

                return @(
                    [pscustomobject]@{
                        id              = "sent-old"
                        subject         = "duplicate"
                        sentDateTime    = "2020-01-01T00:00:00Z"
                        createdDateTime = "2020-01-01T00:00:00Z"
                    },
                    [pscustomobject]@{
                        id              = "created-old"
                        subject         = "created old"
                        sentDateTime    = $null
                        createdDateTime = "2020-01-01T00:00:00Z"
                    },
                    [pscustomobject]@{
                        id              = "created-old-sent-new"
                        subject         = "sent new"
                        sentDateTime    = "2026-01-11T00:00:00Z"
                        createdDateTime = "2020-01-01T00:00:00Z"
                    }
                )
            }
        }

        It "queries sent and created dates, dedupes, and applies fallback" {
            $cutoff = [datetime]"2026-01-10T00:00:00Z"

            $result = Get-AgedMessage `
                -SourceFolderId "source" `
                -Cutoff $cutoff

            @($result.id).Count | Should -Be 2
            @($result.id) | Should -Contain "sent-old"
            @($result.id) | Should -Contain "created-old"
            @($result.id) | Should -Not -Contain "created-old-sent-new"

            $uriText = $script:requestedUris -join "`n"
            $uriText | Should -Match "sentDateTime lt"
            $uriText | Should -Match "createdDateTime lt"
        }
    }

    Describe "Move-AgedMessage" {
        BeforeEach {
            Mock Get-AgedMessage {
                @(
                    [pscustomobject]@{
                        id              = "first"
                        subject         = "first"
                        sentDateTime    = "2020-01-01T00:00:00Z"
                        createdDateTime = "2020-01-01T00:00:00Z"
                    }
                )
            }
        }

        It "reports candidates during WhatIf without moving messages" {
            Mock Invoke-MgGraphRequest {
                throw "The move API should not be called."
            }

            $result = Move-AgedMessage `
                -Label "source" `
                -SourceFolderId "source-id" `
                -DestinationFolderId "dest-id" `
                -Cutoff ([datetime]"2026-01-10T00:00:00Z") `
                -WhatIf

            $result.Candidates | Should -Be 1
            $result.Moved | Should -Be 0
            $result.Failed | Should -Be 0

            Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
        }

        It "continues after a per-message move failure" {
            Mock Get-AgedMessage {
                @(
                    [pscustomobject]@{
                        id              = "first"
                        subject         = "first"
                        sentDateTime    = "2020-01-01T00:00:00Z"
                        createdDateTime = "2020-01-01T00:00:00Z"
                    },
                    [pscustomobject]@{
                        id              = "second"
                        subject         = "second"
                        sentDateTime    = "2020-01-01T00:00:00Z"
                        createdDateTime = "2020-01-01T00:00:00Z"
                    }
                )
            }

            $script:moveCalls = 0
            Mock Invoke-MgGraphRequest {
                $script:moveCalls++
                if ($script:moveCalls -eq 1) {
                    throw "temporary Graph failure"
                }

                [pscustomobject]@{}
            }

            $result = Move-AgedMessage `
                -Label "source" `
                -SourceFolderId "source-id" `
                -DestinationFolderId "dest-id" `
                -Cutoff ([datetime]"2026-01-10T00:00:00Z")

            $result.Candidates | Should -Be 2
            $result.Moved | Should -Be 1
            $result.Failed | Should -Be 1

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly `
                -ParameterFilter {
                    $Uri -like "/v1.0/me/messages/*/move"
                }
        }
    }
}
