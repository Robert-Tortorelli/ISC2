<#
.SYNOPSIS
    Exports a GitHub Copilot chat transcript to a Markdown file.

.DESCRIPTION
    Lists all available chat sessions for the current VS Code workspace and
    lets you select one to export.  All user messages and assistant messages
    (including reasoning/thinking text) are written to a Markdown file.

    The output file is named:
        <current folder>.<session title>.<YYYY-MM-DD>.<HH-MM>.md

    Run from any folder inside the open VS Code workspace.

.PARAMETER TranscriptPath
    Full path to a specific .jsonl transcript file.
    If omitted, the script discovers available sessions and prompts for one.

.PARAMETER OutputPath
    Destination Markdown file. Defaults to:
        "<current folder>.<session title>.<YYYY-MM-DD>.<HH-MM>.md"
    in the current directory.

.PARAMETER Title
    H1 heading written at the top of the output file.
    Defaults to the resolved session title.

.PARAMETER SessionTitle
    Title (or substring) of the chat session to export.
    If omitted, the script lists all available sessions and prompts for one.

.EXAMPLE
    .\Export-ChatTranscript.ps1

.EXAMPLE
    .\Export-ChatTranscript.ps1 -SessionTitle "Export-ChatTranscript"

.EXAMPLE
    .\Export-ChatTranscript.ps1 -OutputPath ".\Transcripts\2026-05-09.md"
#>
param(
    [string]$TranscriptPath = "",
    [string]$OutputPath     = "",
    [string]$Title          = "",
    [string]$SessionTitle   = ""
)

# ── Working directory (used for output naming and workspace detection) ─────────
$currentFolder     = (Get-Location).Path.TrimEnd('\')
$currentFolderName = Split-Path $currentFolder -Leaf
$workspaceFolder   = $currentFolder   # used for workspace storage detection

# ── Locate workspace storage ID ──────────────────────────────────────────────
# The storage ID is stable for a given workspace folder; the transcript session
# ID (filename) changes each conversation.  Discover both automatically so the
# script works for every future session without edits.

$copilotRoot = "$env:APPDATA\Code - Insiders\User\workspaceStorage"
if (-not (Test-Path $copilotRoot)) {
    # Stable-release VS Code
    $copilotRoot = "$env:APPDATA\Code\User\workspaceStorage"
}

if (-not $TranscriptPath) {
    # Find the workspaceStorage folder that owns this workspace by looking for
    # a workspace.json whose "folder" points at our working directory.

    $storageId = Get-ChildItem $copilotRoot -Directory | ForEach-Object {
        $wsjson = Join-Path $_.FullName "workspace.json"
        if (-not (Test-Path $wsjson)) { return }
        try {
            $ws = Get-Content $wsjson -Raw | ConvertFrom-Json

            $folderMatches = $false

            # Case 1: "folder" URI — matches if equal to or a parent of current directory
            if ($ws.folder) {
                $raw     = $ws.folder -replace '^file:///', ''
                $decoded = [System.Uri]::UnescapeDataString($raw).TrimEnd('/\').Replace('/', '\')
                if ($workspaceFolder -ieq $decoded -or $workspaceFolder.StartsWith($decoded + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $folderMatches = $true
                }
            }

            # Helper: check if a workspace file's folders[] contains CWD
            $checkWorkspaceFile = {
                param($wsFilePath)
                if (-not (Test-Path $wsFilePath)) { return $false }
                try {
                    $wsFile    = Get-Content $wsFilePath -Raw | ConvertFrom-Json
                    $wsFileDir = Split-Path $wsFilePath -Parent
                    foreach ($f in $wsFile.folders) {
                        $fp = $f.path
                        if (-not [System.IO.Path]::IsPathRooted($fp)) {
                            $fp = Join-Path $wsFileDir $fp
                        }
                        $resolved = [System.IO.Path]::GetFullPath($fp).TrimEnd('\')
                        if ($workspaceFolder -ieq $resolved -or $workspaceFolder.StartsWith($resolved + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                            return $true
                        }
                    }
                } catch {}
                return $false
            }

            # Case 2: "configuration" URI — named .code-workspace file
            if (-not $folderMatches -and $ws.configuration) {
                $raw        = $ws.configuration -replace '^file:///', ''
                $wsFilePath = [System.Uri]::UnescapeDataString($raw).Replace('/', '\')
                $folderMatches = & $checkWorkspaceFile $wsFilePath
            }

            # Case 3: "workspace" URI — VS Code untitled multi-root workspace
            # stored in %APPDATA%\Code*\Workspaces\<id>\workspace.json
            if (-not $folderMatches -and $ws.workspace) {
                $raw        = $ws.workspace -replace '^file:///', ''
                $wsFilePath = [System.Uri]::UnescapeDataString($raw).Replace('/', '\')
                $folderMatches = & $checkWorkspaceFile $wsFilePath
            }

            if (-not $folderMatches) { return }

            # Rank by most recently written transcript file
            $tDir   = Join-Path $_.FullName "GitHub.copilot-chat\transcripts"
            $latest = Get-ChildItem $tDir -Filter "*.jsonl" -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            [PSCustomObject]@{
                StoragePath = $_.FullName
                LatestTime  = if ($latest) { $latest.LastWriteTime } else { [datetime]::MinValue }
            }
        } catch {}
    } | Sort-Object LatestTime -Descending |
        Select-Object -First 1 -ExpandProperty StoragePath

    if (-not $storageId) {
        Write-Error "Could not locate workspace storage for: $workspaceFolder"
        exit 1
    }

    $transcriptsDir  = Join-Path $storageId "GitHub.copilot-chat\transcripts"
    $chatSessionsDir = Join-Path $storageId "chatSessions"

    if (-not (Test-Path $transcriptsDir)) {
        Write-Error "Transcripts folder not found: $transcriptsDir"
        exit 1
    }

    # ── Discover exportable sessions ──────────────────────────────────────────
    # Only sessions that have a matching file in GitHub.copilot-chat\transcripts
    # can be exported.  Titles come from chatSessions JSONL: kind=1 records
    # where k contains "customTitle" and the title value is in v.
    $availableSessions = @(
        Get-ChildItem $transcriptsDir -Filter "*.jsonl" |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            $sid    = $_.BaseName
            $sTitle = $null
            $csFile = Join-Path $chatSessionsDir "$sid.jsonl"
            if (Test-Path $csFile) {
                foreach ($csLine in [System.IO.File]::ReadAllLines($csFile)) {
                    if (-not $csLine.Trim()) { continue }
                    try {
                        $o = $csLine | ConvertFrom-Json -ErrorAction Stop
                        if ($o.kind -eq 1 -and $o.k -and ($o.k -join "") -like "*customTitle*" -and $o.v) {
                            $sTitle = [string]$o.v
                            break
                        }
                    } catch {}
                }
            }
            if (-not $sTitle) { $sTitle = $sid }
            [PSCustomObject]@{
                SessionId      = $sid
                Title          = $sTitle
                TranscriptFile = $_.FullName
                LastWriteTime  = $_.LastWriteTime
            }
        }
    )

    if ($availableSessions.Count -eq 0) {
        Write-Error "No exportable chat sessions found in: $transcriptsDir"
        exit 1
    }

    # ── Select session ─────────────────────────────────────────────────────────
    $chosen = $null
    if (-not $SessionTitle) {
        Write-Host ""
        Write-Host "Tip: Re-run with -SessionTitle <title> to skip this prompt."
        Write-Host ""
        Write-Host "Available chat sessions:"
        Write-Host ""
        for ($i = 0; $i -lt $availableSessions.Count; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $availableSessions[$i].Title)
        }
        Write-Host ""
        $sel = Read-Host "Select a session (1-$($availableSessions.Count))"
        $idx = [int]$sel - 1
        if ($idx -lt 0 -or $idx -ge $availableSessions.Count) {
            Write-Error "Invalid selection '$sel'. Enter a number between 1 and $($availableSessions.Count)."
            exit 1
        }
        $chosen = $availableSessions[$idx]
    } else {
        $chosen = $availableSessions | Where-Object { $_.Title -ilike "*$SessionTitle*" } | Select-Object -First 1
        if (-not $chosen) {
            Write-Error "No session found matching: '$SessionTitle'"
            Write-Host ""
            Write-Host "Available sessions:"
            $availableSessions | ForEach-Object { Write-Host "  - $($_.Title)" }
            exit 1
        }
    }

    $TranscriptPath       = $chosen.TranscriptFile
    $resolvedSessionTitle = $chosen.Title
    Write-Host "Session    : $resolvedSessionTitle"
    Write-Host "Transcript : $TranscriptPath"
} else {
    # TranscriptPath supplied directly; derive session title from -Title or filename
    $resolvedSessionTitle = if ($Title) { $Title } else { [System.IO.Path]::GetFileNameWithoutExtension($TranscriptPath) }
}

# ── Build output path ─────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $dateStr   = Get-Date -Format "yyyy-MM-dd"
    $timeStr   = Get-Date -Format "HH-mm"
    $safeTitle = $resolvedSessionTitle -replace '[\\/:*?"<>|]', '-'
    $OutputPath = Join-Path $currentFolder "$currentFolderName.$safeTitle.$dateStr.$timeStr.md"
}
if (-not $Title) {
    $Title = $resolvedSessionTitle
}

# ── Parse and format ──────────────────────────────────────────────────────────
$lines = [System.IO.File]::ReadAllLines($TranscriptPath)
$out   = [System.Text.StringBuilder]::new()

[void]$out.AppendLine("# $Title")
[void]$out.AppendLine("")

# Extract session metadata from the first session.start entry
foreach ($line in $lines) {
    try {
        $obj = $line | ConvertFrom-Json -ErrorAction Stop
        if ($obj.type -eq "session.start") {
            $d = $obj.data
            [void]$out.AppendLine("| Field | Value |")
            [void]$out.AppendLine("|---|---|")
            if ($d.startTime)      { [void]$out.AppendLine("| Start Time | $($d.startTime) |") }
            if ($d.producer)       { [void]$out.AppendLine("| Producer | $($d.producer) |") }
            if ($d.copilotVersion) { [void]$out.AppendLine("| Copilot Version | $($d.copilotVersion) |") }
            if ($d.vscodeVersion)  { [void]$out.AppendLine("| VS Code Version | $($d.vscodeVersion) |") }
            [void]$out.AppendLine("")
            break
        }
    } catch {}
}

foreach ($line in $lines) {
    if (-not $line.Trim()) { continue }
    try {
        $obj = $line | ConvertFrom-Json -ErrorAction Stop

        # ── User message ──────────────────────────────────────────────────────
        if ($obj.type -eq "user.message" -and $obj.data.content) {
            [void]$out.AppendLine("---")
            [void]$out.AppendLine("")
            [void]$out.AppendLine("## User")
            [void]$out.AppendLine("")
            [void]$out.AppendLine($obj.data.content.Trim())
            [void]$out.AppendLine("")
        }

        # ── Assistant message ─────────────────────────────────────────────────
        elseif ($obj.type -eq "assistant.message") {
            $reasoning = $obj.data.reasoningText
            $content   = $obj.data.content

            $hasReasoning = $reasoning -and $reasoning.Trim() -ne ""
            $hasContent   = $content   -and $content.Trim()   -ne ""

            # Collect run_in_terminal commands from toolRequests
            $terminalCommands = @()
            if ($obj.data.toolRequests) {
                foreach ($tr in $obj.data.toolRequests) {
                    if ($tr.name -eq "run_in_terminal") {
                        try {
                            $args = $tr.arguments | ConvertFrom-Json -ErrorAction Stop
                            $cmd  = $args.command
                            if ($cmd -and $cmd.Trim() -ne "") {
                                $terminalCommands += $cmd.Trim()
                            }
                        } catch {}
                    }
                }
            }

            if ($hasReasoning -or $hasContent -or $terminalCommands.Count -gt 0) {
                [void]$out.AppendLine("## GitHub Copilot")
                [void]$out.AppendLine("")

                if ($hasReasoning) {
                    [void]$out.AppendLine("*Reasoning:*")
                    [void]$out.AppendLine("")
                    [void]$out.AppendLine($reasoning.Trim())
                    [void]$out.AppendLine("")
                }

                if ($hasContent) {
                    [void]$out.AppendLine($content.Trim())
                    [void]$out.AppendLine("")
                }

                foreach ($cmd in $terminalCommands) {
                    [void]$out.AppendLine("*Ran:*")
                    [void]$out.AppendLine("")
                    [void]$out.AppendLine('```')
                    [void]$out.AppendLine($cmd)
                    [void]$out.AppendLine('```')
                    [void]$out.AppendLine("")
                }
            }
        }
    } catch {}
}

# ── Write output ──────────────────────────────────────────────────────────────
$out.ToString() | Set-Content -Path $OutputPath -Encoding UTF8

$info  = Get-Item $OutputPath
$lines = (Get-Content $OutputPath | Measure-Object -Line).Lines
Write-Host "Output     : $OutputPath"
Write-Host "Size       : $([math]::Round($info.Length / 1KB)) KB, $lines lines"
