# Building and testing

Run commands from this repository's root on Windows with Visual Studio 2022 or Build Tools and the .NET desktop build workload. NuGet access is required for the first restore.

The library and font tool target **.NET Framework 4.5.2**; both test projects target **4.7.2**, matching VisioAutomation. Reference-assembly packages supply the targeting packs. You do not need Visio to build or run the pure unit tests. This alignment does not imply a new Microsoft support guarantee for these legacy targets.

## Build

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$vs = & $vswhere -version '[17.0,18.0)' -products '*' -requires Microsoft.Component.MSBuild -latest -property installationPath
if (-not $vs) { throw 'Visual Studio 2022 was not found.' }
$msbuild = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
& $msbuild VisioAutomationVDX.sln -restore -p:Configuration=Release '-p:Platform=Any CPU' -m
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
```

Use `Debug` for development. Outputs remain in `bin\Debug` and `bin\Release`. Dependencies are centralized in [Directory.Packages.props](Directory.Packages.props).

## Pure tests

```powershell
$vstest = Join-Path $vs 'Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe'
& $vstest TestVisioAutomationVDX.Unit\bin\Release\TestVisioAutomationVDX.Unit.dll /Platform:x64 '/Logger:trx;LogFileName=unit.trx' /ResultsDirectory:TestResults
if ($LASTEXITCODE -ne 0) { throw 'Unit tests failed.' }
```

[CI](.github/workflows/build.yml) builds all four projects in Debug and Release and runs the pure tests. Hosted runners do not run Visio integration tests.

## Visio integration tests

Use installed, activated Visio in an interactive Windows session. Save personal Visio work first. Tests create their own application and temporary files; cleanup runs even on assertion failures. Do not use parallel test execution.

```powershell
& $vstest TestVisioAutomationVDX\bin\Release\TestVisioAutomationVDX.dll /Platform:x64 '/Logger:trx;LogFileName=integration.trx' /ResultsDirectory:TestResults
if ($LASTEXITCODE -ne 0) { throw 'Integration tests failed.' }
```

The suite uses the published VisioAutomation2010 3.0.0 package, not a sibling checkout. Its test-only log adapter normalizes space-padded dates in a private copy because the published parser rejects them. Remove that adapter when consuming a release containing the upstream parser fix. Original Visio log files are never rewritten.

## Font specimen tool

```powershell
.\VisioFontCompare\bin\Release\VisioFontCompare.exe TestResults\specimen.vdx "Calibri"
```

The output directory must exist. With no arguments, output is `specimen.vdx` in the current directory and the font is Calibri. Generation needs no Visio; viewing the result requires a VDX-capable viewer. This console sample is separate from the historical Visio-Font-Compare repository.

## Package and release

Use a current NuGet CLI on PATH, not the historical executable checked into `NuGet`. Build Release first:

```powershell
nuget pack NuGet\VisioAutomation.VDX.nuspec -OutputDirectory TestResults
if ($LASTEXITCODE -ne 0) { throw 'Packaging failed.' }
```

The package includes only the Release VDX DLL under `lib/net452`, plus README and MIT license. It must not include Visio interop, VisioAutomation, or test DLLs.

Packaging is not publishing. The checked-in version remains 1.1.3 for local verification; choose a new release version and update release notes before publishing. Dropping net40 requires an explicit compatibility/version decision. No automatic publishing workflow or package secret is configured by this pass.
