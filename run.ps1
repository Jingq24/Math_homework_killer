param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [string]$OutputDir,
    [string[]]$Formats,
    [string]$Font,
    [switch]$NoInteractive
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Get-Command python -ErrorAction SilentlyContinue

if (-not $Python) {
    throw 'Python not found. Please install Python 3 and add it to PATH.'
}

$Args = @("$ScriptDir\convert.py", $InputPath)
if ($OutputDir) { $Args += @('-o', $OutputDir) }
if ($Formats) { $Args += @('--formats') + $Formats }
if ($Font) { $Args += @('--font', $Font) }
if ($NoInteractive) { $Args += '--no-interactive' }

& python @Args
