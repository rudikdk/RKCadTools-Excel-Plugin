[CmdletBinding()]
param(
    [string]$Version,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$projectPath = Join-Path $repoRoot "RKExcelReportCompare.csproj"
$publishDir = Join-Path $repoRoot "bin\x64\Release\net48\publish"
$distDir = Join-Path $repoRoot "dist"

if (-not $Version) {
    [xml]$projectXml = Get-Content $projectPath
    $Version = $projectXml.Project.PropertyGroup.Version
}

if ($Version.StartsWith("v")) {
    $tagVersion = $Version
    $packageVersion = $Version.Substring(1)
} else {
    $tagVersion = "v$Version"
    $packageVersion = $Version
}

if (-not $SkipBuild) {
    dotnet restore $projectPath
    dotnet build $projectPath -c Release /p:Platform=x64
}

$xll32 = Join-Path $publishDir "RKExcelReportCompare-packed.xll"
$xll64 = Join-Path $publishDir "RKExcelReportCompare64-packed.xll"

if (-not (Test-Path $xll32)) {
    throw "Missing release artifact: $xll32"
}

if (-not (Test-Path $xll64)) {
    throw "Missing release artifact: $xll64"
}

New-Item -ItemType Directory -Force $distDir | Out-Null

$releaseName = "RKExcelReportCompare-$tagVersion"
$releaseDir = Join-Path $distDir $releaseName
$zipPath = Join-Path $distDir "$releaseName.zip"

if (Test-Path $releaseDir) {
    Remove-Item -Recurse -Force $releaseDir
}

if (Test-Path $zipPath) {
    Remove-Item -Force $zipPath
}

New-Item -ItemType Directory -Force $releaseDir | Out-Null

Copy-Item $xll32 $releaseDir
Copy-Item $xll64 $releaseDir
Copy-Item (Join-Path $repoRoot "docs\USER_GUIDE.html") $releaseDir
Copy-Item (Join-Path $repoRoot "README.md") $releaseDir
Copy-Item (Join-Path $repoRoot "LICENSE") (Join-Path $releaseDir "LICENSE.txt")
Copy-Item (Join-Path $repoRoot "CREDITS.md") $releaseDir
Copy-Item (Join-Path $repoRoot "CHANGELOG.md") $releaseDir
Copy-Item (Join-Path $repoRoot "THIRD_PARTY_NOTICES.md") $releaseDir

@"
RK Excel Report Compare $tagVersion

Included files:
- RKExcelReportCompare64-packed.xll for 64-bit Excel
- RKExcelReportCompare-packed.xll for 32-bit Excel
- USER_GUIDE.html installation and usage guide
- README.md project overview
- LICENSE.txt MIT license
- CREDITS.md developer and contributor credits
- CHANGELOG.md release history
- THIRD_PARTY_NOTICES.md third-party license notices
- CHECKSUMS.txt SHA256 checksums

Install:
1. Extract this zip file.
2. Right-click the selected .xll file, open Properties, and click Unblock if Windows shows the option.
3. In Excel, go to File -> Options -> Add-ins.
4. Set Manage to Excel Add-ins, click Go, then Browse to the .xll file.
"@ | Set-Content -Path (Join-Path $releaseDir "PACKAGE_README.txt") -Encoding UTF8

@"
RK Excel Report Compare $tagVersion

See CHANGELOG.md for project changes and USER_GUIDE.html for installation and usage instructions.
"@ | Set-Content -Path (Join-Path $releaseDir "RELEASE_NOTES.txt") -Encoding UTF8

$checksumLines = Get-ChildItem $releaseDir -File |
    Sort-Object Name |
    ForEach-Object {
        $hash = Get-FileHash $_.FullName -Algorithm SHA256
        "$($hash.Hash)  $($_.Name)"
    }

$checksumLines | Set-Content -Path (Join-Path $releaseDir "CHECKSUMS.txt") -Encoding ASCII

Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $zipPath -Force

Write-Host "Created release package:"
Write-Host $zipPath
