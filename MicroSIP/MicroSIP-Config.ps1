<# README
Copyright:        fenigen
License:          Apache License 2.0 
License URL:      http://www.apache.org/licenses/LICENSE-2.0
Author:           NFetisov

Version:          1.0.0.5
File Description: Автоматическая конфигурация MicroSIP
#>

<# История изменений:
1.0.0.0 (05.05.2021) - Создан скрипт для автоматического создания файла конфигурации MicroSIP (во время установки ПО). Без тестирования
1.0.0.1 (11.05.2021) - Тест, исправление ошибок
1.0.0.2 (12.05.2021) - Обработка параметров.
1.0.0.5 (13.05.2021) - Правка конфигурации. Добавление задания в Планировщик задач. Добавление в автозагрузку.
#>

$TimeStart = Get-Date
CLS
# $PSVersionTable.PSVersion
#[System.Windows.Forms.MessageBox]::Show("НАЧАЛО")
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
Write-Host " "

##### START #####
##### Переменные среды #####
$label = " "
$domain = " "

$User = $env:UserName
$path = "C:\Users\"+$User+"\Prog\MicroSIP"

### ТЕСТЫ #####
#Тест связи с доменом
if ((Test-Connection -computer $domain -quiet) -ne $True){
    Write-Host ":-( Нет связи с доменом $domain" -ForegroundColor white -BackgroundColor red
    [System.Windows.Forms.MessageBox]::Show("Нет связи с доменом! Попробуйте позже.")
    exit
    }
Else {Write-Host ":-) Есть связь с доменом $domain" -ForegroundColor black -BackgroundColor green}

#Проверка что пользователь локальный или доменный
if ((Get-LocalUser $User).Name -eq $User){
    Write-Host "Пользователь локальный :-(" -ForegroundColor white -BackgroundColor red
    $User = Read-Host 'Введите логин: '
    }
Write-Host ":-| Ваш логин: " $User -ForegroundColor black -BackgroundColor yellow
Write-Host ":-) Пользователь доменный" -ForegroundColor black -BackgroundColor green

##### ОБРАБОТКА #####
##### Получаем данные LDAP фильтрами из домена #####
# Формируем запрос пользователя
$Filter = "(&(objectCategory=User)(samAccountName=$User))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$UserADpatch = $Searcher.FindOne()
$UserAD = $UserADpatch.GetDirectoryEntry()

##### Получаем и обрабатываем телефонный номер #####
if ($UserAD.telephoneNumber -ne "$null"){
    $number = $UserAD.telephoneNumber -replace '[^0-9]+', ""
    $number = [string]$number
    }
else {$number = Read-Host 'Введите Ваш внутренний номер телефона в формате 70XXXXX: '}
Write-Host ":-| Внутренний номер телефона: " $number -ForegroundColor black -BackgroundColor yellow
Write-Host ":-| Установка будет произведена по пути: " $path -ForegroundColor black -BackgroundColor yellow

##### Формируем файл конфигурации #####
# Шаблон
$ini = '[Settings]
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
server=SERVER
proxy=
domain=SERVER
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

# Производим замены
$ini = $ini -replace "#NUMBERUSER#", $NUMBER
$ini = $ini -replace "#NAMELABEL#", $label
# Сохраняем
$ini | Out-File -FilePath "$path\microsip.ini" -Encoding UTF8
Write-Host ":-) Конфигурация сохранена в файл INI" -ForegroundColor black -BackgroundColor green

##### СТАРТУЕМ #####
Copy-Item -Path "$path\Install\MicroSIP-Update.exe" -Destination "C:\Prog\MicroSIP-Update.exe" -Force
sleep 2
Start "C:\KDL_Prog\MicroSIP-Update.exe"

# Создаем ярлык
$source = "C:\Users\"+$User+"\Prog\MicroSIP\microsip.exe"
$target = 'C:\Users\'+$User+'\Desktop\MicroSIP.lnk'
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($target)
$Shortcut.TargetPath = $source
$Shortcut.Save()

# Используем конфигурацию
$EXE = $path+'\MicroSIP.exe'
Start-Process -FilePath "powershell" -Verb RunAs {
    New-NetFirewallRule -DisplayName “MicroSIP” -Direction Inbound -Program $EXE -Action Allow
    exit
    }

Start $EXE
Start-Process -FilePath "powershell" -Verb RunAs {
    sleep 1 
    Stop-Process -processname microsip
    exit
}

#Правим планировщик
#if ((schtasks /query /tn KDL_MicroSIP) -eq "*Ошибка*" -or "*Error*"){Write-Host ":-D"}
Start-Process -FilePath "powershell" -Verb RunAs {
    schtasks /DELETE /tn KDL_MicroSIP /F
    exit
    }

#$Task = $path+'\Install\KDL_MicroSIP.xml'
Start-Process -FilePath "powershell" -Verb RunAs {
    #schtasks /Create /XML $Task /tn KDL_MicroSIP
    schtasks /Create /tn KDL_MicroSIP /TR "C:\Prog\MicroSIP-Update.exe" /SC WEEKLY
    exit
    }

# Запускаем
sleep 1
Start $EXE

##### END #####
Write-Host " "
$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host ":-D ВЫПОЛНЕНО за "$Time -ForegroundColor white -BackgroundColor blue
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
[System.Windows.Forms.MessageBox]::Show("MicroSIP установлен в Вашей учетной записи Windows")

# Закрытие с кодом возврата (при необходимости)
# [System.Environment]::Exit(0)