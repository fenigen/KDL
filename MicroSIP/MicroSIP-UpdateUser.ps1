<# README
Copyright:        fenigen
License:          Apache License 2.0
License URL:      http://www.apache.org/licenses/LICENSE-2.0
Author:           NFetisov

Version:          1.0.0.0
File Description: Обновление программы MicroSIP
#>

<# История изменений:
1.0.0.0 (11.05.2021) - Создание основного функционала (Получение списка где есть MicroSIP, обновление справочника и заготовка к обновлению ПО
1.0.0.3 (12.05.2021) - Правка ошибок
#>

$TimeStart = Get-Date
CLS
# $PSVersionTable.PSVersion
#[System.Windows.Forms.MessageBox]::Show("НАЧАЛО")
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
Write-Host " "

##### START #####
Write-Host ":-) Скрипт запущен из папки: " $PSScriptRoot  -ForegroundColor black -BackgroundColor green

#   Install 7zip module
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
#Set-PSRepository -Name 'PSGallery' -SourceLocation "https://www.powershellgallery.com/api/v2" -InstallationPolicy Trusted
Set-PSRepository -Name 'PSGallery' -SourceLocation "$PSScriptRoot\Install" -InstallationPolicy Trusted
Install-Module -Name 7Zip4PowerShell -Force

#   Extract 7zip file
$sourcefile = "C:\KDL_Prog\Temp\MicroSIP.7z"
Expand-7Zip -ArchiveFileName $sourcefile -TargetPath $PSScriptRoot

Start "$PSScriptRoot\microsip.exe"


##### END #####
Write-Host " "
$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host ":-D ВЫПОЛНЕНО за "$Time -ForegroundColor white -BackgroundColor blue
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
#[System.Windows.Forms.MessageBox]::Show("ГОТОВО")

# Закрытие с кодом возврата (при необходимости)
# [System.Environment]::Exit(0)
