$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$dist = Join-Path $root 'dist'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
$wordInterop = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word') `
    -Recurse -Filter Microsoft.Office.Interop.Word.dll -ErrorAction Stop |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
$officeInterop = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\office') `
    -Recurse -Filter OFFICE.DLL -ErrorAction Stop |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
$extensibility = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC\Extensibility') `
    -Recurse -Filter extensibility.dll -ErrorAction Stop |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName

if (-not (Test-Path -LiteralPath $compiler)) {
    throw '.NET Framework C# compiler was not found.'
}

New-Item -ItemType Directory -Path $dist -Force | Out-Null
$addin = Join-Path $dist 'WordMoveToHeading.dll'
$setup = Join-Path $dist 'WordMoveToHeading-Setup.exe'

& $compiler /nologo /target:library /out:$addin `
    /reference:$wordInterop /reference:$officeInterop /reference:$extensibility `
    /reference:System.Windows.Forms.dll `
    (Join-Path $root 'src\Connect.cs') `
    (Join-Path $root 'src\MoveToHeadingMenu.cs')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& $compiler /nologo /target:winexe /out:$setup `
    /reference:System.Windows.Forms.dll `
    "/resource:$addin,WordMoveToHeading.Payload.dll" `
    (Join-Path $root 'installer\Installer.cs')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$hash = Get-FileHash -LiteralPath $setup -Algorithm SHA256
Write-Host "Built: $setup"
Write-Host "SHA256: $($hash.Hash)"


