# PowerShell script to compare DOC to TXT converters on all .doc files in a folder

param(
    [string]$InputFolder = ".\examples",
    [string]$OutputFolder = ".\examples\out",
    [string]$DtExePath = ".\examples\dt.exe",
    [string]$AntiwordExePath = "c:\antiword\antiword.exe",
    [string]$AbiwordPath = "C:\Program Files (x86)\AbiWord\bin\abiword.exe",
    [string]$LibreOfficePath = "C:\Program Files\LibreOffice\program\soffice.exe",
    [string]$WvWareNetPath = ".\WvWareNetConsole\bin\Debug\net9.0\WvWareNetConsole.exe"
)

$env:HOME=$env:USERPROFILE

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$AbsoluteOutputFolder = (Resolve-Path $OutputFolder).Path

Get-ChildItem -Path $InputFolder -Filter *.doc | ForEach-Object {
    $docFile = $_.FullName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)

    # doctotext.exe
    $dtOut = Join-Path $OutputFolder "$baseName.dt.txt"
    & $DtExePath $docFile > $dtOut

    # antiword.exe
    $antiwordOut = Join-Path $OutputFolder "$baseName.antiword.txt"
    & $AntiwordExePath $docFile > $antiwordOut

    # Abiword
    $abiOut = Join-Path $OutputFolder "$baseName.abiword.txt"
    & $AbiwordPath --to=txt --to-name=$abiOut $docFile

    # WvWareNet
    $wvOut = Join-Path $OutputFolder "$baseName.wvware.txt"
    & $WvWareNetPath $docFile $wvOut

    # LibreOffice
    Start-Process -FilePath $LibreOfficePath -ArgumentList "--headless", "--norestore", "--quickstart=no", "--convert-to", "txt:Text", "--outdir", "$AbsoluteOutputFolder", "$docFile" -Wait
}

Get-ChildItem -Path $InputFolder -Filter *.doc | ForEach-Object {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)

    # LibreOffice
    $loOut = Join-Path $OutputFolder "$baseName.libreoffice.txt"

    # LibreOffice output will be $OutputFolder\$baseName.txt, rename to match suffix
    $loOrig = Join-Path $OutputFolder "$baseName.txt"
    if (Test-Path $loOrig) {
        Rename-Item -Force $loOrig $loOut #([System.IO.Path]::GetFileName($loOut))
    }

}
