param(
    [string] $RepositoryRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($RepositoryRoot)
$prefix = $root.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar

function Assert-LocalFile([string] $BaseDirectory, [string] $RelativePath) {
    if ([IO.Path]::IsPathRooted($RelativePath) -or $RelativePath -match '\$\(|%[^%]+%') {
        throw "Build input must use a repository-relative path: $RelativePath"
    }
    $path = [IO.Path]::GetFullPath((Join-Path $BaseDirectory $RelativePath))
    if (-not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or
        -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing or external build input: $RelativePath"
    }
    return $path
}

[xml] $solution = Get-Content -Raw -LiteralPath (Join-Path $root 'VisioAutomationVDX.slnx')
if (Test-Path -LiteralPath (Join-Path $root 'VisioAutomationVDX.sln')) {
    throw 'The legacy SLN must not coexist with the canonical SLNX.'
}
if ($solution.SelectNodes('//Configurations | //Properties | //Build').Count) {
    throw 'Unexpected legacy solution settings or project-skipping configuration.'
}
foreach ($file in $solution.SelectNodes('//File')) {
    Assert-LocalFile $root $file.Path | Out-Null
}
$projects = $solution.SelectNodes('//Project')
if ($projects.Count -ne 4) { throw 'The solution must build all four projects.' }
foreach ($project in $projects) {
    $path = Assert-LocalFile $root $project.Path
    [xml] $xml = Get-Content -Raw -LiteralPath $path
    if ($xml.Project.Sdk -ne 'Microsoft.NET.Sdk') { throw "Expected SDK-style project: $path" }
    if ($xml.SelectNodes('//HintPath | //ProjectGuid | //TargetFrameworkVersion | //Import').Count) {
        throw "Legacy or manually resolved build metadata in $path"
    }
    foreach ($input in $xml.SelectNodes('//ProjectReference | //EmbeddedResource')) {
        Assert-LocalFile (Split-Path $path -Parent) $input.Include | Out-Null
    }
}

# XML fixtures can otherwise reopen stencils from the original author's Office installation.
$fixtures = @(
    'VisioAutomation.VDX/Resources/DefaultVDXTemplate.xml',
    'TestVisioAutomationVDX/datafiles/template_router.vdx',
    'TestVisioAutomationVDX/datafiles/vdx_with_warnings_1.vdx'
)
foreach ($fixture in $fixtures) {
    [xml] $xml = Get-Content -Raw -LiteralPath (Assert-LocalFile $root $fixture)
    foreach ($node in $xml.SelectNodes('//@* | //text()')) {
        if ($node.Value -match '^(?:[A-Za-z]:[\\/]|\\\\|file:)') {
            throw "Machine-specific XML reference in $fixture"
        }
    }
}
Write-Host 'Repository checks passed: four portable projects and no machine-specific fixture paths.'
