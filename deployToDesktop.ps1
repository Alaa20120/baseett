$DesktopPath = [Environment]::GetFolderPath("Desktop")
$TargetFolder = Join-Path -Path $DesktopPath -ChildPath "InvoiceSystem"
if (!(Test-Path -Path $TargetFolder)) {
    New-Item -ItemType Directory -Force -Path $TargetFolder
}
Copy-Item -Path "c:\Users\ahmed\.gemini\antigravity\scratch\*" -Destination $TargetFolder -Recurse -Force

Start-Process -FilePath (Join-Path -Path $TargetFolder -ChildPath "index.html")
