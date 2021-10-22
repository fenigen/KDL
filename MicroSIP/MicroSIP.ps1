<#
.Synopsis
    Обеспечение автоматической конфигурации, обновления программы MicroSIP.
.Description
    Использование MicroSIP
.Parameter Type
    Отсутствует - Вывод информации об ошибке
    Install - Распаковка программы добавления правил в систему.
    Update - указывает на запуск обновления.
    Clean - Очистка временных файлов (после обновления)
.Example

.Inputs
    System.String
.Role
    Domain User, Local User
.Licence
    Copyright NFetisov@kdl.ru
    Licensed under the Apache License, Version 2.0
.Component	
    MicroSIP
.Link
    https://github.com/fenigen/KDL/tree/main/MicroSIP
    http://www.apache.org/licenses/LICENSE-2.0
.Notes
    MicroSIP.exe
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
    PWD = $env:USERPROFILE + "\MicroSIP"
    Source_EXE = ""
    Source_Contacte = ""
    Source_EXE_Update = ""
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
##### Подфункции
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

function Testing-System {
<# Проверка условий работы и получение дополнительной информации #>

    $Domain = ""
    $Testing_Status = @{
        Domain = ""
        UserName = ""
        UserCN = ""
    }

    # Проверка политики PowerShell
    if ((Get-Executionpolicy) -ne "RemoteSigned"){
        Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  MOD   ] Политика PS изменена на RemoteSigned") @Global:Console_Mod
    }
    else {Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Используется политика RemoteSigned") @Global:Console_OK}

    #Проверка типа пользователя
    $LocalList = (Get-LocalUser).Name
    foreach ($item in $LocalList) {
        if ($item -like $env:USERNAME){
            $UserName = ""
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  ERROR ] Пользователь локальный") @Global:Console_Error
            break
        }
    }
    if ($UserName -ne ""){
        $UserName = $env:USERNAME
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   OK   ] Пользователь домена: " $UserName) @Global:Console_OK
    }
    $Testing_Status.UserName = $UserName

    # Проверяем доступность домена
    $Testing_Status.Domain = Test-Connection -computer $Domain -quiet
    if ($Testing_Status.Domain -ne $True) {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  ERROR ] Нет связи с доменом: " $Global:Argument.Domain) @Global:Console_Error

        $Regedit = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI -Name LastLoggedOnDisplayName
        $Testing_Status.UserCN = $Regedit.LastLoggedOnDisplayName
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Из реестра получено: " + $Testing_Status.UserCN)) @Console_Error
    }
    else {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Есть связь с доменом: " $Domain) @Global:Console_OK
        
        $Testing_Status.UserCN = ([adsi]"WinNT://$env:UserDomain/$env:UserName,user").fullname
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Пользователь: " + $Testing_Status.UserCN)) @Console_Wait
    }

    # Возвращаем
    return $Testing_Status
}

function GUI-TelefonNumber {
<# Поле ввода номера телефона если его не удалось получить в AD #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Global:Project.Name + " v. " + $Global:Project.Version + " - Ввод номера телефона"
    $form.Size = New-Object System.Drawing.Size(300,180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'Fixed3D'
    $form.MaximizeBox = $false


    $label_CN = New-Object System.Windows.Forms.Label
    $label_CN.Size = New-Object System.Drawing.Size(($form.Width - 10),20)
    $label_CN.Location = New-Object System.Drawing.Point(10,20)
    $label_CN.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
    $label_CN.Text = "Фетисов Никита Геннадьевич"
    $form.Controls.Add($label_CN)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,50)
    $label.Size = New-Object System.Drawing.Size(($form.Width - 10),20)
    $label.Text = 'Введите Ваш внутренний номер телефона:'
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,70)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Location = New-Object System.Drawing.Point(100,110)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)


    $form.Topmost = $true
    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()


    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        $TelefonNumber = $textBox.Text
    }

    return $TelefonNumber
}

function New-MicroSIP-Config {
<# Получение номера и формирование файла конфигурации #>
    $Argument = @{
        Label = ""
        Server = ""
    }
    
    # шаблон
    $Template = '[Settings]
accountId=0
singleMode=1
ringingSound=
volumeRing=100
audioRingDevice=""
audioOutputDevice=""
audioInputDevice=""
micAmplification=0
swLevelAdjustment=0
audioCodecs=PCMA/8000/1 PCMU/8000/1
VAD=0
EC=1
forceCodec=0
opusStereo=0
disableMessaging=1
rport=1
sourcePort=0
rtpPortMin=0
rtpPortMax=0
dnsSrvNs=
dnsSrv=0
STUN=
enableSTUN=0
recordingPath=Recordings
recordingFormat=mp3
autoRecording=0
recordingButton=1
DTMFMethod=0
autoAnswer=button
autoAnswerDelay=0
autoAnswerNumber=
denyIncoming=button
usersDirectory=
defaultAction=
enableMediaButtons=1
headsetSupport=1
localDTMF=1
enableLog=0
bringToFrontOnIncoming=1
enableLocalAccount=0
randomAnswerBox=0
crashReport=0
callWaiting=1
updatesInterval=never
checkUpdatesTime=1620849581
noResize=0
userAgent=
autoHangUpTime=0
maxConcurrentCalls=0
noIgnoreCall=0
cmdOutgoingCall=
cmdIncomingCall=
cmdCallRing=
cmdCallAnswer=
cmdCallBusy=
cmdCallStart=
cmdCallEnd=
silent=0
portKnockerHost=
portKnockerPorts=
mainX=1021
mainY=215
mainW=259
mainH=465
messagesX=471
messagesY=215
messagesW=0
messagesH=0
ringinX=0
ringinY=0
callsWidth0=0
callsWidth1=0
callsWidth2=0
callsWidth3=0
callsWidth4=0
callsWidth5=0
contactsWidth0=0
contactsWidth1=0
contactsWidth2=0
volumeOutput=100
volumeInput=100
activeTab=0
AA=0
AC=0
DND=0
alwaysOnTop=0
enableShortcuts=0
shortcutsBottom=0
lastCallNumber=
lastCallHasVideo=0

[Account1]
label=#NAMELABEL#
server=#SIPSERVER#
proxy=
domain=#SIPSERVER#
username=#NUMBERUSER#
password=#NUMBERUSER#
authID=#NUMBERUSER#
displayName=#NUMBERUSER#
dialingPrefix=
dialPlan=
hideCID=0
voicemailNumber=
transport=udp
publicAddr=
SRTP=
registerRefresh=300
keepAlive=15
publish=1
ICE=0
allowRewrite=1
disableSessionTimer=1
'

    # Получаем номер телефона
    if (($Global:Testing_Status.Domain -eq $True) -and ($Global:Testing_Status.UserName -ne "")){
        $User = $Global:Testing_Status.UserName
        $Filter = "(&(objectCategory=User)(samAccountName=$User))"
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.Filter = $Filter
        $UserADpatch = $Searcher.FindOne()
        $UserAD = $UserADpatch.GetDirectoryEntry()

        if ($UserAD.telephoneNumber -ne "$null") {
            $TelefonNumber = $UserAD.telephoneNumber -replace '[^0-9]+', ""
            $TelefonNumber = [string]$TelefonNumber
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   OK   ] Получен номер телефона: " $TelefonNumber) @Global:Console_Wait
        }
        else {$TelefonNumber = ""}
    }
    else {$TelefonNumber = ""}

    if ($TelefonNumber -eq "") {
        $TelefonNumber = GUI-TelefonNumber
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Введен номер телефона: " + $TelefonNumber)) @Console_Wait
    }

    # Производим замены
    $Template = $Template -replace "#NUMBERUSER#", $TelefonNumber
    $Template = $Template -replace "#NAMELABEL#", $Argument.Label
    $Template = $Template -replace "#SIPSERVER#", $Argument.Server
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  OK   ] Файл настроек сгенерирован") @Global:Console_OK

    return $Template
}

function Install-MicroSIP {
<# Функция установки MicroSIP #>

    param (
        $ini
    )
    
    # Повышение привелегий
    $user = '\Admin'
    $key = ""
    $pass = ""
    $password = $pass | ConvertTo-SecureString -Key $key
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $password

    # Сохраняем файлы
    $ini | Out-File -FilePath ($Global:Argument.PWD + "\microsip.ini") -Encoding UTF8
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  OK   ] Файл настроек сохранен") @Global:Console_OK
    Copy-Item -Path $Global:Argument.Source_Contacte -Destination ($Global:Argument.PWD + "\Contacts.xml") -Force


    ### Создаем задачи в планировщике на запуск обновлений
    # Задача на обновление
    $Trigger_Updata = New-JobTrigger -Weekly -DaysOfWeek ((Get-Date).DayOfWeek.value__) -At (Get-Date -format "HH:mm")
    $Action_Updata = New-ScheduledTaskAction -Execute ($Global:Argument.PWD + "\MicroSIP.exe") -Argument "Update"
    Register-ScheduledTask -TaskName "MicroSIP_Update" -Trigger $Trigger_Updata -Action $Action_Updata –Force
    
    # Добавляем программу в исключения брендмауэра
    Start ($Global:Argument.PWD + "\microsip.exe")
    sleep 1
    Stop-Process -processname microsip
    sleep 1

    $Rules = Get-NetFirewallRule -DisplayName "MicroSIP"
    if ($Rules -eq $null){
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  OK   ] Создаем правило" ) @Global:Console_OK

        Start-Process -FilePath "powershell" -Credential $credentials -ArgumentList '-noprofile -command &{
            Start-Process -FilePath "powershell" -Verb RunAs {
            $source = $env:USERPROFILE + "\MicroSIP\microsip.exe"
            New-NetFirewallRule -DisplayName “MicroSIP” -Direction Inbound -Program $source -Action Allow
        }'
    }
    else {Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  OK   ] Правило существует" ) @Global:Console_OK}

    # Выводим ярлык
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($env:USERPROFILE + "\Desktop\MicroSIP.lnk")
    $Shortcut.TargetPath = $Global:Argument.PWD + "\microsip.exe"
    $Shortcut.Save()
    
    # Вывод сообщение о завершении установки
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::None
    $balmsg.BalloonTipText = 'на вашем АРМе установлена программа MicroSIP'
    $balmsg.BalloonTipTitle = $Global:Testing_Status.UserCN
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(5)
}

function Update-MicroSIP {
<# Проверка наличия обновлений. Обновление телефонного справочника. Запуск обновления программы #>

    $Version = @{
        Exe_Source = ""
        Contact_Source = ""
        Contact_Use = ""
    }

    # Проверяем необходимость и обновление exe
    if (Test-Path $Global:Argument.Source_EXE){
        $Version.Exe_Source = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Global:Script.Source_EXE).FileVersion
        if ($Version.Exe_Source -notlike $Global:Project.Version){
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  ERROR ] обновляем программу") @Global:Console_Wait

            # Копируем файлы во временную папку
            Copy-Item -Path $Global:Argument.Source_EXE -Destination ($env:TEMP + "\MicroSIP.exe") -Force
            Copy-Item -Path $Global:Argument.Source_EXE -Destination ($env:TEMP + "\MicroSIP-Update.exe") -Force

            # Запуск обновления
            $Apps = ($env:TEMP + "\MicroSIP-Update.exe")
            & $Apps 'Update'

        } # Проверка версии
    } # Проверка пути
    
    ### Обновляем файл контактов
    if (Test-Path $Global:Argument.Source_Contacte){
        $Version.Contact_Source = (Get-Item $Global:Argument.Source_Contacte).LastWriteTime
        $Version.Contact_Use = (Get-Item $Global:Argument.Source_Contacte).LastWriteTime

        if ($Version.Contact_Source -ne $Version.Contact_Use){
            Copy-Item -Path $Global:Argument.Source_Contacte -Destination ($Global:Argument.PWD + "\Contacts.xml") -Force
        } # Сравнение версий
    } # Проверка пути

} # Функция Обновления

###################################################################################################
##### Обработка
Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] ######################### START  #########################") @Console_Info
Write-Host (Write-Output -OutVariable +global:Log "                                                                                ") @Console_Info

$Global:Testing_Status = Testing-System

### Обработчик
if ($Type -like ""){
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::error
    $balmsg.BalloonTipText = 'Запуск программы без ключей не возможен'
    $balmsg.BalloonTipTitle = $Global:Testing_Status.UserCN
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(100)
}

if ($Type -like "Install"){
    # Вывод сообщение о старте
    Add-Type -AssemblyName System.Windows.Forms
    $global:balmsg = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::None
    $balmsg.BalloonTipText = 'на вашем АРМе производится установка программы для телефонии - MicroSIP'
    $balmsg.BalloonTipTitle = $Global:Testing_Status.UserCN
    $balmsg.Visible = $true
    $balmsg.ShowBalloonTip(5)
    
    $INI = New-MicroSIP-Config
    Install-MicroSIP $INI
    Start-ScheduledTask -TaskName "MicroSIP_Update"
}

if ($Type -like "Update") {
    Update-MicroSIP
}

if ($Type -like "Clean") {
<# Удаляем временные файлы после обновления #>
    Remove-Item ($env:TEMP + "\MicroSIP.exe") -Force
    Remove-Item ($env:TEMP + "\MicroSIP-Update.exe") -Force
}

###################################################################################################
Close
