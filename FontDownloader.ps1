﻿<#This script will download a repo of fonts and create a temp file on the desktop.
It will then install all the fonts in that temp folder, once completed it will then delete itself#>

clear


<#Downloading fonts, creating the temp file and unzipping them#>

Invoke-WebRequest "https://github.com/Corbin-Hardy-CMI/CMI-Fonts/archive/refs/heads/main.zip" -OutFile "C:\Users\$env:username\Test.zip"

Expand-Archive -Path "C:\Users\$env:username\Test.zip" -DestinationPath "C:\Users\$env:username\Test"
Expand-Archive -Path "C:\Users\$env:username\Test\CMI-Fonts-main\CMI Fonts.zip" -DestinationPath "C:\Users\$env:username\Test"

Expand-Archive -Path "C:\Users\$env:username\Test\Gotham.zip" -DestinationPath "C:\Users\$env:username\Test"
Expand-Archive -Path "C:\Users\$env:username\Test\Lato.zip" -DestinationPath "C:\Users\$env:username\Test"
Expand-Archive -Path "C:\Users\$env:username\Test\Roboto.zip" -DestinationPath "C:\Users\$env:username\Test"
Expand-Archive -Path "C:\Users\$env:username\Test\OneDrive_2023-05-04.zip" -DestinationPath "C:\Users\$env:username\Test"
Expand-Archive -Path "C:\Users\$env:username\Test\OneDrive_2023-05-04 (1).zip" -DestinationPath "C:\Users\$env:username\Test"

<#Installing all the fonts in the "Test" folder#>

$fontDir = "C:\Users\$env:username\Test\Gotham"
$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)
Get-ChildItem -Path $fontDir -Filter *.otf | ForEach-Object {
    $fontFile = $_.FullName
    $fontName = $_.Name
    $fontsFolder.CopyHere($fontFile)
    Write-Host "Installed font: $fontName"
}

$fontDir = "C:\Users\$env:username\Test\Lato"
$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)
Get-ChildItem -Path $fontDir -Filter *.ttf | ForEach-Object {
    $fontFile = $_.FullName
    $fontName = $_.Name
    $fontsFolder.CopyHere($fontFile)
    Write-Host "Installed font: $fontName"
}

$fontDir = "C:\Users\$env:username\Test\Roboto"
$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)
Get-ChildItem -Path $fontDir -Filter *.ttf | ForEach-Object {
    $fontFile = $_.FullName
    $fontName = $_.Name
    $fontsFolder.CopyHere($fontFile)
    Write-Host "Installed font: $fontName"
}

$fontDir = "C:\Users\$env:username\Test\Proxima Nova"
$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)
Get-ChildItem -Path $fontDir -Filter *.otf | ForEach-Object {
    $fontFile = $_.FullName
    $fontName = $_.Name
    $fontsFolder.CopyHere($fontFile)
    Write-Host "Installed font: $fontName"
}

$fontDir = "C:\Users\$env:username\Test\Proxima Nova Condensed"
$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)
Get-ChildItem -Path $fontDir -Filter *.otf | ForEach-Object {
    $fontFile = $_.FullName
    $fontName = $_.Name
    $fontsFolder.CopyHere($fontFile)
    Write-Host "Installed font: $fontName"
}

<#Force delete the temp folders#>

Remove-Item -Path "C:\Users\$env:username\Test" -Recurse -Force
Remove-Item -Path "C:\Users\$env:username\Test.zip" -Recurse -Force