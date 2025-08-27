# Paths and variables - customize these
$oneNoteProcessName = "ONENOTE"
$solutionPath = "C:\Path\To\YourSolution.sln"
$devenvPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.com"
$setupOutputPath = "C:\Path\To\Setup\Output\YourSetup.msi"
$msiInstallerPath = $setupOutputPath  # assuming MSI installer
$oneNoteExePath = "$Env:ProgramFiles\Microsoft Office\root\Office16\ONENOTE.EXE"  # Adjust Office version accordingly

# 1. Close OneNote
Write-Host "Closing OneNote if running..."
Get-Process -Name $oneNoteProcessName -ErrorAction SilentlyContinue | Stop-Process -Force

# 2. Save all open project files (build solution)
Write-Host "Building solution to rebuild setup project..."
& "$devenvPath" "$solutionPath" /Rebuild Release

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed. Aborting."
    exit 1
}

# 3. Install the setup file
Write-Host "Installing setup for all users..."

# For MSI installing for all users, use ALLUSERS=1 property (common)
# You can run it silently with /qn, or UI with /passive
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiInstallerPath`" ALLUSERS=1 /qn" -Wait

if ($LASTEXITCODE -ne 0) {
    Write-Error "Installation failed."
    exit 1
}

# 4. Start OneNote again
Write-Host "Starting OneNote..."
Start-Process -FilePath $oneNoteExePath

Write-Host "All tasks completed successfully."
