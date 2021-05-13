﻿<# README
Copyright:        fenigen
License:          Apache License 2.0
License URL:      http://www.apache.org/licenses/LICENSE-2.0
Author:           NFetisov

Version:          1.0.0.3
File Description: Обновление телефонного справочника и проверка необходимости обновления программы
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
### Проверка и по необходиомсти подключение сетевого диска
$Networkpath = "N:\" 
$pathExists = Test-Path -Path $Networkpath

If ($pathExists)  {
    Write-host ":-) Путь существует" -ForegroundColor black -BackgroundColor green
}

else {
    (new-object -com WScript.Network).MapNetworkDrive("N:"," serv")
    Write-Host ":-) Путь создан" -ForegroundColor black -BackgroundColor green
}

### Обработка на АРМе
$UserPach = Get-Childitem -Path C:\Users -directory # Получением подкаталоги C:\users
Write-Host ":-) Получили список каталогов" -ForegroundColor black -BackgroundColor green

ForEach ($item in $UserPach.FullName){
    $Pach = $item+'\KDL_Prog\MicroSIP\'
    if (Test-Path $Pach){
        Write-Host ":-| MicroSIP есть в папке: $Pach" -ForegroundColor black -BackgroundColor green

        # Обновляем программу
        if ((get-content \MicroSIP-Update.ini) -ne (get-content $Pach\MicroSIP-Update.ini)) {
            Write-Host ":-( Требуется обновление" -ForegroundColor black -BackgroundColor red
            start "$Pach\MicroSIP-UpdateUser.exe"
            }
        else {Write-Host ":-) Актуальная версия" -ForegroundColor black -BackgroundColor green}

        # Обновляем справочник
        Copy-Item -Path " \Contacts.xml" -Destination "$Pach\Contacts.xml" -Force
        Write-Host ":-) Справочник обновлен" -ForegroundColor black -BackgroundColor green
        }
    }

##### END #####
Write-Host " "
$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host ":-D ВЫПОЛНЕНО за "$Time -ForegroundColor white -BackgroundColor blue
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
#[System.Windows.Forms.MessageBox]::Show("ГОТОВО")

# Закрытие с кодом возврата (при необходимости)
# [System.Environment]::Exit(0)