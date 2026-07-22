# Create a desktop shortcut for BUFS HWP Editor.
# Keep this script ASCII-only so it parses on older Windows PowerShell hosts.

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$exe = Join-Path $root "BUFS-HWP-Editor.exe"

if (-not (Test-Path $exe)) {
  throw "Cannot find executable: $exe"
}

$desktop = [Environment]::GetFolderPath("Desktop")
$hwpKo = [string]([char]0xD55C) + [string]([char]0xAE00)
$shortcutTitle = "BUFS $hwpKo Editor"
$shortcutPath = Join-Path $desktop "$shortcutTitle.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = $root
$shortcut.Description = $shortcutTitle
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"
