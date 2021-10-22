<#
.Synopsis
    Обеспечение обновления программы MicroSIP.
.Description
    Использование MicroSIP
.Parameter Type
    Update указывает на запуск обновления.
    В прочих случаях (во избежания запуск не из тела основного скрипта) выводится уведомление об ошибке.
.Example

.Inputs
    System.String
.Role
    Domain User
.Licence
    Copyright NFetisov@kdl.ru
    Licensed under the Apache License, Version 2.0
.Component	
    MicroSIP
.Link
    https://github.com/fenigen/KDL/tree/main/MicroSIP
    http://www.apache.org/licenses/LICENSE-2.0
.Notes
    Для работы требуется установленный 7-ZIP
    MicroSIP-Update.exe
#>

[CmdletBinding()]

param (
    [string]$Type
)

$Global:Project = @{
    Version = "1.0.1.15"
    Name = "MicroSIP"
    TimeStart = Get-Date
    TimeEnd = ""
    TimeWork = ""
}

$Global:Argument = @{
    Source_Contacte = ""
    PWD = $env:USERPROFILE + "\MicroSIP"
    Install_EXE = $env:TEMP + "\MicroSIP.exe"
    Update_EXE = ($PWD + "\MicroSIP-Update.exe")
    Status = ""
}

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

# Получаем информацию об установленном 7-ZIP
if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) {
    throw "$env:ProgramFiles\7-Zip\7z.exe needed"
}

set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"



# Обрабатываем
# Имя пользователя из реестра
$Regedit = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI -Name LastLoggedOnDisplayName

if ($Type -like "Update") {

    # Убиваем запущенные процессы
    if ((Get-Process -processname microsip) -ne $null){
        <#
        # Подтверждение
        $wshell = New-Object -ComObject Wscript.Shell
        $Message_Output = $wshell.Popup($Regedit.LastLoggedOnDisplayName + "
Производится обновление программы MicroSIP. 
После нажатия кнопки OK программа будет MicroSIP, будет закрыта
",0,$Global:Project.Name + " v. " + $Global:Project.Version + "",0+32)
#>
        Stop-Process -processname microsip
        $Global:Argument.Status = "true"
        sleep 1
    }

    if ((Get-Process -processname microsip-1) -ne $null){
        Stop-Process -processname microsip-1
        sleep 1
     }

    # Распаковываем
    & {sz x ($Global:Argument.Install_EXE) -aoa -o($Global:Argument.PWD) -y}
    sleep 3
    Copy-Item -Path $Global:Argument.Source_Contacte -Destination ($PWD + "\Contacts.xml") -Force
    & $Global:Argument.Update_EXE 'Clean'

    if ($Global:Argument.Status -eq "true") {
        Start ($PWD + "\microsip.exe")
    }

    # Вывод сообщения
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balmsg.BalloonTipText = 'MicroSIP обновлен. Справочник контактов обновлен.'
    $balmsg.BalloonTipTitle = $Regedit.LastLoggedOnDisplayName
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(1)
}


if ($Type -notlike "Update") {
    # Вывод сообщения
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::error
    $balmsg.BalloonTipText = 'Запуск программы не возможен. Отсутствуют параметры запуска.'
    $balmsg.BalloonTipTitle = $Regedit.LastLoggedOnDisplayName
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(100)
}

Close
