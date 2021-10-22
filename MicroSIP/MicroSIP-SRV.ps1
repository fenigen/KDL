<#
.Synopsis
    Формирование телефонного справочника для программы MicroSIP.
.Description
    Использование MicroSIP
.Parameter Type
    Update указывает на запуск обновления.
    В прочих случаях (во избежания запуск не из тела основного скрипта) выводится уведомление об ошибке.
.Example

.Inputs
    System.String
.Role
    Domain Admin
.Licence
    Copyright NFetisov@kdl.ru
    Licensed under the Apache License, Version 2.0
.Component	
    RSAT
.Link
    https://github.com/fenigen/KDL/tree/main/MicroSIP
    http://www.apache.org/licenses/LICENSE-2.0
.Notes
    MicroSIP-SRV.exe
#>

[CmdletBinding()]

param (
    [string]$Type
)

$Global:Project = @{
    Version = "1.0.1.1"
    Name = "MicroSIP Contact Update"
    TimeStart = Get-Date
    TimeEnd = ""
    TimeWork = ""
}

Import-module activedirectory

$Global:Argument = @{
    Domain = ""
    OU = ""
    Share = ""
}

&{
    $Global:Log = ""
    $Global:Console_Info = @{ForegroundColor='Magenta'; BackgroundColor='DarkBlue'}
    $Global:Console_OK = @{ForegroundColor="Black"; BackgroundColor = "Green"}
    $Global:Console_Error = @{ForegroundColor="White"; BackgroundColor = "Red"}
    $Global:Console_Wait = @{ForegroundColor="Black"; BackgroundColor = "Yellow"}
    $Global:Console_Mod = @{ForegroundColor="Yellow"; BackgroundColor = "DarkYellow"}
    $Global:Console_Print = @{ForegroundColor="DarkBlue"; BackgroundColor = "Magenta"}
}

Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] ######################### OPEN  #########################") @Console_Info
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  INFO  ] Версия PowerShell" $PSVersionTable.PSVersion) @Global:Console_Info
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  INFO  ] Запущен скрипт: " $PSScriptRoot) @Global:Console_Info

###################################################################################################
function Close {
<# Закрытие приложения #>
    $Global:Project.TimeEnd = Get-Date
    $Global:Project.TimeWork = $Global:Project.TimeEnd - $Global:Project.TimeStart
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] Время работы скрипта: " $Global:Project.TimeWork) @Console_Info

    Write-Host (Write-Output -OutVariable +global:Log "                                                                                ") @Console_Info
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] ########################## END  ##########################") @Console_Info
    
    #Pause
    
    Clear-Host
    Remove-Variable -Name * -Force -ErrorAction SilentlyContinue
    [System.Environment]::Exit(0)
}
###################################################################################################

# Проверка задачи в schedule
if ((Get-ScheduledTask -TaskName "SRV_MicroSIP") -eq $null){
    $Trigger = New-JobTrigger -Weekly -DaysOfWeek 1 -At 11:00am
    $Action = New-ScheduledTaskAction -Execute ("C:\!SRV\MicroSIP-SRV.exe")
    Register-ScheduledTask -TaskName "SRV\SRV_MicroSIP" -Trigger $Trigger -Action $Action –Force

    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Задача создана ") @Global:Console_OK
}
else{Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Задача существует в планировщике ") @Global:Console_OK}

# Проверка связи с доменом
if ((Test-Connection -computer $Global:Argument.Domain -quiet) -ne $True) {
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  ERROR ] Нет связи с доменом: " $Global:Argument.Domain) @Global:Console_Error
    Close
}

# Запрашиваем информацию из домена
$user = Get-ADUser -SearchBase $Global:Argument.OU -Filter {(telephoneNumber -ne "null") -and (Enabled -eq "true")} -Property CN, givenName, sn, telephoneNumber, MobilePhone, ipPhone, mail, streetAddress, City, st, postalCode, Department, Title, Enabled
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Получены данные из " $Global:Argument.OU) @Global:Console_OK

# Формируем XML
$XML = @()

$XML = '<?xml version="1.0"?>
<contacts>
'

# Заполняем XML данными
ForEach ($item in $user){
    $NUMBER = $item.telephoneNumber -replace '[^0-9]+', ""
    $MOBILE = $item.MobilePhone -replace '[^0-9]+', ""
    $temp = '<contact name="'+$item.CN+'" number="'+$NUMBER+'" firstname="'+$item.givenName+'" lastname="'+$item.sn+'" phone="'+$NUMBER+'" mobile="'+$MOBILE+'" email="'+$item.mail+'" address="'+$item.streetAddress+'" city="'+$item.City+'" state="'+$item.st+'" zip="'+$item.postalCode+'" comment="'+$item.Deportament+$item.title+'" id="" info="" presence="0" directory="0"/>
    '
    $temp = -join $temp
    $exp =$exp+$temp
}

$XML = $XML + '</contacts>'
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] XML сформирован " $Global:Argument.OU) @Global:Console_OK

# Обновляем файлы
Remove-Item $Global:Argument.Local
Remove-Item $Global:Argument.Share
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Старые версии удалены " $Global:Argument.OU) @Global:Console_OK

$XML | Out-File -FilePath $Global:Argument.Local -Encoding UTF8
$XML | Out-File -FilePath $Global:Argument.Share -Encoding UTF8
Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Созданы новые " $Global:Argument.OU) @Global:Console_OK

###################################################################################################
Close
