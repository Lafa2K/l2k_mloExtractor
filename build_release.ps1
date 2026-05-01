Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-CscPath {
    $dotnetRoslyn = Get-ChildItem -Path "$env:ProgramFiles\dotnet\sdk\*\Roslyn\bincore\csc.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if ($dotnetRoslyn) {
        return $dotnetRoslyn.FullName
    }

    $frameworkCsc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (Test-Path -LiteralPath $frameworkCsc) {
        return $frameworkCsc
    }

    throw "Could not locate csc.exe. Install the .NET SDK or .NET Framework build tools."
}

function Assert-PathExists([string]$PathValue) {
    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "Required file or folder not found: $PathValue"
    }
}

$root = $PSScriptRoot
$src = Join-Path $root "src"
$lib = Join-Path $root "lib"
$img = Join-Path $root "img"
$build = Join-Path $root "build"
$release = Join-Path $build "Release"
$packageVersion = "1.1.2"
$packageName = "l2k_mloExtractor-$packageVersion-win64"
$packageZip = Join-Path $build ($packageName + ".zip")

$referenceRoot = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
$facadeRoot = Join-Path $referenceRoot "Facades"

$csc = Get-CscPath

$required = @(
    (Join-Path $src "AssemblyInfo.cs"),
    (Join-Path $src "Program.cs"),
    (Join-Path $src "MloExporterForm.cs"),
    (Join-Path $src "YtypPropExporter.cs"),
    (Join-Path $img "banner_idle.png"),
    (Join-Path $img "banner_loading.gif"),
    (Join-Path $img "icon.ico"),
    (Join-Path $lib "CodeWalker.exe"),
    (Join-Path $lib "CodeWalker.Core.dll"),
    (Join-Path $lib "CodeWalker.WinForms.dll"),
    (Join-Path $referenceRoot "mscorlib.dll"),
    (Join-Path $referenceRoot "System.dll"),
    (Join-Path $referenceRoot "System.Core.dll"),
    (Join-Path $referenceRoot "System.Drawing.dll"),
    (Join-Path $referenceRoot "System.Windows.Forms.dll"),
    (Join-Path $referenceRoot "System.Xml.dll"),
    (Join-Path $facadeRoot "netstandard.dll")
)

foreach ($pathValue in $required) {
    Assert-PathExists $pathValue
}

if (Test-Path -LiteralPath $build) {
    Remove-Item -LiteralPath $build -Recurse -Force
}
New-Item -ItemType Directory -Path $release | Out-Null

$outputExe = Join-Path $release "Blender MLO Extractor.exe"

$arguments = @(
    "/nologo",
    "/target:winexe",
    "/platform:x64",
    "/langversion:7.3",
    "/optimize+",
    "/nowarn:0436",
    "/out:$outputExe",
    "/win32icon:$(Join-Path $img 'icon.ico')",
    "/r:$(Join-Path $referenceRoot 'mscorlib.dll')",
    "/r:$(Join-Path $referenceRoot 'System.dll')",
    "/r:$(Join-Path $referenceRoot 'System.Core.dll')",
    "/r:$(Join-Path $referenceRoot 'System.Drawing.dll')",
    "/r:$(Join-Path $referenceRoot 'System.Windows.Forms.dll')",
    "/r:$(Join-Path $referenceRoot 'System.Xml.dll')",
    "/r:$(Join-Path $facadeRoot 'netstandard.dll')",
    "/r:$(Join-Path $lib 'CodeWalker.exe')",
    "/r:$(Join-Path $lib 'CodeWalker.Core.dll')",
    "/r:$(Join-Path $lib 'CodeWalker.WinForms.dll')",
    (Join-Path $src "AssemblyInfo.cs"),
    (Join-Path $src "Program.cs"),
    (Join-Path $src "MloExporterForm.cs"),
    (Join-Path $src "YtypPropExporter.cs")
)

Write-Host "Compiling Blender MLO Extractor..."
& $csc @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Compilation failed with exit code $LASTEXITCODE."
}

Get-ChildItem -LiteralPath $lib -File | Copy-Item -Destination $release -Force

$releaseImg = Join-Path $release "img"
New-Item -ItemType Directory -Path $releaseImg | Out-Null
Get-ChildItem -LiteralPath $img -File | Copy-Item -Destination $releaseImg -Force

Copy-Item -LiteralPath (Join-Path $root "README.md") -Destination (Join-Path $release "README.md") -Force

@'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8" />
  </startup>
</configuration>
'@ | Set-Content -LiteralPath (Join-Path $release "Blender MLO Extractor.exe.config") -Encoding UTF8

Compress-Archive -Path (Join-Path $release "*") -DestinationPath $packageZip -Force

Write-Host ""
Write-Host "Build complete:"
Write-Host $outputExe
Write-Host $packageZip
