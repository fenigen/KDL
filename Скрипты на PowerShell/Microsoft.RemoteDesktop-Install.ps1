<# README
Имя файла:        1.0.0.4 — Microsoft.RemoteDesktop-Install.ps1

Copyright:        fenigen
License:          Apache License 2.0 
License URL:      http://www.apache.org/licenses/LICENSE-2.0
Author:           

Version:          1.0.0.4
File Description: Автоматическая установка и настройка Microsoft Remote Desktop
Product Name:     RDS
#>

<#
Описание:

#>


<# История изменений:
1.0.0.0 (26.05.2021) - Написан шаблон
1.0.0.1 (28.05.2021) - Тестирование и исправление ошибок
1.0.0.2 (31.05.2021) - Добавлен вывод ярлыка на рабочий стол
1.0.0.3 (31.05.2021) - Корректное добавление RDS
1.0.0.4 (31.05.2021) - Изменения внешнего вида сообщений.
#>

### START ###
CLS
$TimeStart = Get-Date
$PSVersionTable.PSVersion

#[System.Windows.Forms.MessageBox]::Show("НАЧАЛО")
Write-Host "##############################################################" -ForegroundColor Magenta -BackgroundColor DarkBlue
Write-Host "Время запуска программы: " $TimeStart -ForegroundColor Magenta -BackgroundColor DarkBlue
Write-Host ""


# Формируем переменные среды:
$User = $env:UserName                               # Получаем имя УЗ Windows
$TargetFile =  "C:\Windows\explorer.exe"
$ShortcutFile = "$env:USERPROFILE\Desktop\Рабочие приложения.lnk"

#$DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)


#Проверка что пользователь локальный или доменный
if ((Get-LocalUser $User).Name -eq $User){
    Write-Host Get-Date ":-( Пользователь локальный" -ForegroundColor white -BackgroundColor red
    $User = Read-Host 'Введите логин: '
    }
Write-Host Get-Date ":-| Ваш логин: " $User -ForegroundColor black -BackgroundColor yellow
Write-Host Get-Date ":-) Пользователь доменный" -ForegroundColor black -BackgroundColor green

# Проверка версии OS
$sys = wmic os get Version /value #Get-ComputerInfo .OSVersion
if ($sys -like "*10*"){
    Write-Host Get-Date "Версия операционной системы" $sys -ForegroundColor DarkMagenta -BackgroundColor Green
# Проверка наличие приложения:
    if ((get-appxpackage -name Microsoft.RemoteDesktop) -eq $null){
        Write-Host Get-Date ":-| Сейчас будет произведена установка приложения" -ForegroundColor DarkMagenta -BackgroundColor Yellow
# Устанавлиевем приложение
        Add-AppxPackage -Path "C:\Temp\Microsoft.RemoteDesktop.AppxBundle"
        # msiexec.exe /I "C:\Temp\RemoteDesktop_1.2.1954.0_x64.msi" /qn ALLUSERS=1 # для версии MSI
        Write-Host Get-Date "Microsoft Remote Desktop установлена" -ForegroundColor Magenta -BackgroundColor Green
        sleep 5

# Создаем ярлык
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.Arguments="shell:AppsFolder\Microsoft.RemoteDesktop_8wekyb3d8bbwe!App"
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
        Write-Host Get-Date ":-) Ярлык выведен на рабочий стол пользователя" $Version -ForegroundColor DarkMagenta -BackgroundColor Green



        $FrmMain = New-Object 'System.Windows.Forms.Form'
        $FrmMain.TopMost = $True
[System.Windows.Forms.MessageBox]::Show($FrmMain,"Приложение установлено. Для дальнейшей работы необходимо войти на сервера!

Для этого введите пароль и закройте приложение. 
После чего оно будет вызвано повторно. Введите парооль второй раз.

Для продолжения нажимте кнопку ОК

", "Microsoft.RemoteDesktop", 0, 48)


        $proc = Get-Process -Name RdClient.Windows | Sort-Object -Property ProcessName -Unique
        while ( $proc.Responding -eq $True ) {
            sleep 1
            $proc = Get-Process -Name RdClient.Windows | Sort-Object -Property ProcessName -Unique
        }

        Start-Process "ms-rd:subscribe?url=https://АДРЕС СЕРВЕРА/RDWeb/feed/WebFeed.aspx&username=KDL\$User"
        Write-Host Get-Date ":-) в подписки добавлен rds.kdl.ru" -ForegroundColor Magenta -BackgroundColor Green

        $proc = Get-Process -Name RdClient.Windows | Sort-Object -Property ProcessName -Unique
        while ( $proc.Responding -eq $True ) {
            sleep 1
            $proc = Get-Process -Name RdClient.Windows | Sort-Object -Property ProcessName -Unique
        }
        start "ms-rd:subscribe?url=https://АДРЕС ВТОРОГО СЕРВЕРА/RDWeb/feed/WebFeed.aspx&username=KDL\$User"
        Write-Host Get-Date ":-) в подписки добавлен msk-rds-02" -ForegroundColor Magenta -BackgroundColor Green
    }

    Write-Host ":-) Приложение готово к работе" $Version -ForegroundColor DarkMagenta -BackgroundColor Green
}

# Вывод сообщения.
Add-Type -AssemblyName System.Windows.Forms
$global:balmsg = New-Object System.Windows.Forms.NotifyIcon
$path = (Get-Process -id $pid).Path
$balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
$balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
$balmsg.BalloonTipText = 'Клиент для работы с приложениями удаленного рабочего стола установлено. На рабочем столе Вы можете найти ярлык'
$balmsg.BalloonTipTitle = "Уважаемый $User"
$balmsg.Visible = $true
$balmsg.ShowBalloonTip(1000)


# Удаляем лишнее:
Remove-item C:\Temp\Microsoft.RemoteDesktop.AppxBundle -force
Remove-item C:\Temp\Microsoft.RemoteDesktop-Install.exe -force

##### END #####
Write-Host " "
$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host ":-D ВЫПОЛНЕНО за "$Time -ForegroundColor Magenta -BackgroundColor DarkBlue
Write-Host "##############################################################" -ForegroundColor Magenta -BackgroundColor DarkBlue
#[System.Windows.Forms.MessageBox]::Show("Приложение для работы с RDS готово")

# Закрытие с кодом возврата (при необходимости)
[System.Environment]::Exit(0)