$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$bin = Join-Path $root 'bin'
New-Item -ItemType Directory -Force $bin | Out-Null

# Find all C# source files in the root folder
$sources = Get-ChildItem -Path $root -Filter '*.cs' |
    Sort-Object FullName |
    ForEach-Object { $_.FullName }

# Locate standard .NET Framework 4.0 C# compiler (csc.exe)
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (!(Test-Path $csc)) { $csc = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if (!(Test-Path $csc)) { throw "csc.exe for .NET Framework 4.0 was not found." }

# Resolve SolidWorks Interop files (prefer local lib, fallback to default install path)
$swSldWorks = Join-Path $root 'lib\SolidWorks.Interop.sldworks.dll'
$swConst = Join-Path $root 'lib\SolidWorks.Interop.swconst.dll'

if (!(Test-Path $swSldWorks) -or !(Test-Path $swConst)) {
    $swRoot = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS'
    $swSldWorks = Join-Path $swRoot 'SolidWorks.Interop.sldworks.dll'
    $swConst = Join-Path $swRoot 'SolidWorks.Interop.swconst.dll'
}

if (!(Test-Path $swSldWorks)) { throw "SOLIDWORKS interop DLL not found." }
if (!(Test-Path $swConst)) { throw "SOLIDWORKS interop DLL not found." }

$out = Join-Path $bin 'SWTableConvertor.exe'
Write-Host "Compiling SWTableConvertor..."
& $csc /nologo /target:winexe "/out:$out" `
    /r:System.dll `
    /r:System.Core.dll `
    /r:System.Data.dll `
    /r:System.Drawing.dll `
    /r:System.Windows.Forms.dll `
    /r:System.IO.Compression.dll `
    /r:System.IO.Compression.FileSystem.dll `
    /r:System.Xml.dll `
    /r:Microsoft.CSharp.dll `
    "/r:$swSldWorks" `
    "/r:$swConst" `
    $sources

if ($LASTEXITCODE -ne 0) { throw "csc.exe failed with exit code $LASTEXITCODE." }

Write-Host "Copying SOLIDWORKS Interop assemblies..."
Copy-Item -LiteralPath $swSldWorks -Destination (Join-Path $bin 'SolidWorks.Interop.sldworks.dll') -Force
Copy-Item -LiteralPath $swConst -Destination (Join-Path $bin 'SolidWorks.Interop.swconst.dll') -Force

Write-Host "Build Succeeded! Output at: $out"
