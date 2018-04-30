#Requires -Version 5
Set-StrictMode -Version 5

$Publish = [IO.Path]::Combine($PSScriptRoot, 'publish')
if (-not(Test-Path $Publish)) {
    New-Item $Publish -ItemType Directory
}

$Shortcut = [IO.Path]::Combine($PSScriptRoot, 'Shortcut.vbs')
$Binary = [IO.Path]::Combine($PSScriptRoot, 'OpenCsvByExcel\bin')
$Builds = 'x86','x64'

$Builds|%{
    $WorkingDir = [IO.Path]::Combine($Binary, $_, 'Release')
    Copy-Item $Shortcut -Destination $WorkingDir

    $Files = Get-ChildItem $WorkingDir -Exclude '*.xml'

    $ExePath = $Files|?{$_.Name -eq 'OpenCsvByExcel.exe'}
    $ExeVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath).ProductVersion

    Set-Location $WorkingDir

    $ZipDestination = [IO.Path]::Combine($Publish, "OpenCsvByExcel_$($ExeVersion)_$($_).zip")
    Compress-Archive $Files -DestinationPath $ZipDestination -Force
}