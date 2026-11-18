# RDCRAMAN
## RDC and RemoteApp Manager

[Download the latest Windows x64 build](build/DesktopAndRemoteApp.exe) · [Changelog](CHANGELOG.md)

The manager hosts active RDP desktops in one window, with a connection tree, live display sizing, and full-screen controls. It opens Windows-published RemoteApps through their native shortcuts. Saved connections can be exported and restored from the Backup menu.

### Build from source

Run the following from the repository root in PowerShell with the Windows .NET Framework 4.x compiler installed.

```powershell
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\vbc.exe'
$root = (Get-Location).Path
$source = Join-Path $root 'src\DesktopAndRemoteApp.vb'
$icon = Join-Path $root 'assets\DesktopAndRemoteApp.ico'
$background = Join-Path $root 'assets\ZombieGreen.png'
New-Item -ItemType Directory -Force -Path (Join-Path $root 'build') | Out-Null
$exe = Join-Path $root 'build\DesktopAndRemoteApp.exe'
& $compiler /nologo /target:winexe /platform:x64 "/out:$exe" "/win32icon:$icon" "/resource:$background,ZombieGreen.png" `
    /reference:System.Windows.Forms.dll,System.Drawing.dll,System.Security.dll,System.Xml.Linq.dll $source
```
