<#
.SYNOPSIS
    Exports the most recent GitHub Copilot chat transcript to a Markdown file.

.DESCRIPTION
    Reads the JSONL transcript stored by VS Code for this workspace and writes
    all user messages and assistant messages (including the reasoning/thinking
    text) to a Markdown file.

.PARAMETER TranscriptPath
    Full path to a specific .jsonl transcript file.
    If omitted, the script picks the most-recently-written .jsonl file in the
    workspace's Copilot transcripts folder.

.PARAMETER OutputPath
    Destination Markdown file. Defaults to "Chat Transcript.md" in the same
    directory as this script.

.PARAMETER Title
    H1 heading written at the top of the output file.

.EXAMPLE
    .\Export-ChatTranscript.ps1

.EXAMPLE
    .\Export-ChatTranscript.ps1 -OutputPath ".\Transcripts\2026-04-14.md"
#>
param(
    [string]$TranscriptPath = "",
    [string]$OutputPath     = "",
    [string]$Title          = ""
)

# ── Derive defaults from workspace folder name ───────────────────────────────
$workspaceFolder = (Get-Location).Path.TrimEnd('\')
$workspaceName   = Split-Path $workspaceFolder -Leaf
$timestamp       = Get-Date -Format "yyyy-MM-dd-HH-mm"

if (-not $OutputPath) {
    $OutputPath = Join-Path $workspaceFolder "$workspaceName.Chat Transcript.$timestamp.md"
}
if (-not $Title) {
    $Title = "$workspaceName — Chat Transcript"
}

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

    $transcriptsDir = Join-Path $storageId "GitHub.copilot-chat\transcripts"
    if (-not (Test-Path $transcriptsDir)) {
        Write-Error "Transcripts folder not found: $transcriptsDir"
        exit 1
    }

    $TranscriptPath = Get-ChildItem $transcriptsDir -Filter "*.jsonl" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $TranscriptPath) {
        Write-Error "No .jsonl transcript files found in: $transcriptsDir"
        exit 1
    }

    Write-Host "Transcript : $TranscriptPath"
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
