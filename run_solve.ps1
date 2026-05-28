param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$ImagePath,

    [string[]]$Number,
    [string]$OutputDir,
    [string[]]$Formats,
    [string]$Font,
    [switch]$NoInteractive,
    [switch]$SkipAi,
    [string[]]$Markdown,
    [switch]$MergeOnly,
    [string]$MergedName,
    [switch]$NoMerge
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Get-Command python -ErrorAction SilentlyContinue

if (-not $Python) {
    throw 'Python not found. Please install Python 3 and add it to PATH.'
}

$Args = @("$ScriptDir\solve_homework.py")
if ($ImagePath) { $Args += $ImagePath }
if ($Number) { $Args += @('--number') + $Number }
if ($OutputDir) { $Args += @('-o', $OutputDir) }
if ($Formats) { $Args += @('--formats') + $Formats }
if ($Font) { $Args += @('--font', $Font) }
if ($NoInteractive) { $Args += '--no-interactive' }
if ($SkipAi) { $Args += '--skip-ai' }
if ($Markdown) { $Args += @('--markdown') + $Markdown }
if ($MergeOnly) { $Args += '--merge-only' }
if ($MergedName) { $Args += @('--merged-name', $MergedName) }
if ($NoMerge) { $Args += '--no-merge' }

& python @Args
