<#
.Synopsis
    ManagerCSP - управление сертификатами ЭЦП.
.Description
    Установка сертификатов ЭЦП
.Parameter Type

.Example

.Inputs
    
.Role
    User
.Licence
    Copyright https://github.com/fenigen/
    Licensed under the Apache License, Version 2.0
.Component	
    Крипто ПРО CSP
.Link
    https://github.com/fenigen/
    http://www.apache.org/licenses/LICENSE-2.0
.Notes
    FileName = ManagerCSP

    Использованные материалы:
    https://cmd.readthedocs.io/csptest.html
    https://www.codegrepper.com/code-examples/shell/browse+for+folder+powershell
#>

[CmdletBinding()]

param (
   [string]$Args
)

Add-Type -AssemblyName PresentationFramework
Add-Type -Assembly System.Windows.Forms
Add-Type -Assembly System.Drawing

###################################################################################################
##### Определяем переменные
$Global:Log = ""

$Global:Basic = @{
    ProjectName = "ManagerCSP"           # Имя проекта
    CurrentVersion = "1.0.1.3"               # Текущая версия

    Var = @{
        Temp = $env:TEMP + "\ManagerCSP" # Каталог для временных файлов
        UserDisplayName = $null              # Отображаемое имя пользователя
        PathInt = $null
        Basic = $False                       # Статус разовых базовых операций
        Log = ""
    }

    Update = @{
        Files = ""
        Version = ""                         # Версия на севере
        Status = ""                          # Статус обновления
        Type = "Auto"                        # Тип обновления (Ручной или автоматический)
    }

    BIN = @{
        Icon = $null
        Background = $null
        FileInstruction = $null
    }
}

##################################################################################################################################

### Определяем функции
function Start-BasicFunction {
<# Выполнение базовых операций #>
    [CmdletBinding()]
        param ($Args)
    ##################################################################################################################################
    ##### Подфункции
    function Close () {
        if (Test-Path -Path $Global:Basic.Var.Temp){
            Remove-Item -path $Global:Basic.Var.Temp -Recurse -Force
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Очищен каталог: " + $Global:Basic.Var.Temp)) @Console_WAIT
        }

        Write-Host (Write-Output -OutVariable +Global:Log "") @Global:Console_Info
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] ######################### EXIT  #########################") @Console_Info
#pause        
        # Закрываем открытые формы
        $About_Forms_xam.Close()
        $Log_Forms_xam.Close()
        $xamGUI_Forms.Close()
        
        Clear-Host
        Remove-Variable -Name * -Force -ErrorAction SilentlyContinue
[System.Environment]::Exit(0)
    } <# Закрываем программу #>
    
    function Open-About () {
        [xml]$Global:About_Forms_xml = (
'<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="' + $Global:Basic.ProjectName + '  - About" Name="About_Form" Height="333" Width="367" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid Height="299" Width="349">
        <Image Name="About_Form_Background"  Stretch="UniformToFill" Margin="-12,-42,0,0" Grid.ColumnSpan="3" Grid.RowSpan="2" />
        <Button Name="About_OK_Button" Content="ОК" Height="24" HorizontalAlignment="Left" Margin="124,263,0,0" VerticalAlignment="Top" Width="99" FontSize="12" />
        <TextBlock Text="О программе: ' + $Global:Basic.ProjectName + '" Height="23" HorizontalAlignment="Left" Margin="0,6,0,0" VerticalAlignment="Top" Width="348" TextAlignment="Center" FontWeight="Bold" FontSize="14" />
        <Label Content="Текущая версия:" Height="24" HorizontalAlignment="Left" Margin="0,34,0,0" VerticalAlignment="Top" Width="116" FontSize="11" FontWeight="Bold" />
        <Label Name="About_Current_label" Content="' + $Global:Basic.CurrentVersion + '" FontSize="11" Height="24" HorizontalAlignment="Left" Margin="113,34,0,0" VerticalAlignment="Top" Width="54" />
        <Label Content="Доступная версия:" Height="24" HorizontalAlignment="Left" Margin="0,60,0,0" VerticalAlignment="Top" Width="116" FontSize="11" FontWeight="Bold" />
        <Label Name="About_Update_label" Content="' + $Global:Basic.Update.Version + '" FontSize="11" Height="24" HorizontalAlignment="Right" Margin="0,60,182,0" VerticalAlignment="Top" Width="54" />
        <Button Name="About_Update_Button" Content="Обновить" Height="23" HorizontalAlignment="Left" Margin="173,48,0,0" VerticalAlignment="Top" Width="75" FontSize="10" Visibility="Hidden" />
        <Label Content="Описание:" FontSize="11" Height="24" HorizontalAlignment="Left" Margin="0,97,0,0" Name="label5" VerticalAlignment="Top" Width="128" FontWeight="Bold" />
        <TextBlock Height="110" HorizontalAlignment="Left" Margin="12,119,0,0" Text="Программа предназначена для автоматической установки сертификатов ЭЦП из контейнеров в папке вида XXXXXXX.000. Подробная информацияв документации." VerticalAlignment="Top" Width="325" TextWrapping="Wrap" />
    </Grid>
</Window>
')
        $Global:About_Forms_xam = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $About_Forms_xml))
        $About_Forms_xml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
            Set-Variable -Name ($_.Name) -Value $About_Forms_xam.FindName($_.Name) -Scope Global
        }

        if (($Global:Basic.BIN.Icon -ne $null) -and ($Global:Basic.BIN.Background -ne $null)){
            $About_Form.Icon = ([System.Convert]::FromBase64String($Global:Basic.BIN.Icon))
            $About_Form_Background.source = ([System.Convert]::FromBase64String($Global:Basic.BIN.Background)) # Задний фон
        } # Добавляем графику

        switch ($Global:Basic.Update.Status) {
            'ConnectError' {$About_Update_label.Content = 'Нет связи с сервером' }
            'Update' {
                $About_Current_label.Background = "#fde910"
                $About_Update_label.Background = "#fde910" #Лимонный
                $About_Update_Button.Visibility = 'Visible'} 
            'OK' {$About_Current_label.Background = "#99ff99"
                $About_Update_label.Background = "#99ff99" } # Салатовый
            'Test' {$About_Current_label.Background = "#99ff99"
                $About_Update_label.Background = "#7fc7ff" }
        }
                
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем About") @Global:Console_OK
        $About_Forms_xam.ShowDialog() | Out-Null
    } <# Вывод диалогового окна About#>

    function Open-Help () {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем Help") @Global:Console_OK
        $FileInstruction = $Global:Basic.Var.Temp + "\" + $Global:Basic.ProjectName + ".pdf"

        if ($Global:Basic.BIN.FileInstruction -ne $null){
            [System.Convert]::FromBase64String($Global:Basic.BIN.FileInstruction) | Set-Content $FileInstruction -Encoding Byte
        }
        else {
            $msgBox =  [System.Windows.MessageBox]::Show("Отсутствует файл содержащий cправку", $Global:Var.ProjectName + "- Справка",'OK','Error')
        }

        if (Test-Path $FileInstruction) {
            sleep 1
            & $FileInstruction
            sleep 1
        }
    } <# Открывем справку #>

    function Open-Log () {
        [xml]$Global:Log_Forms_xml = ('<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="' + $Global:Basic.ProjectName + '  - Лог программы" Name="Log_Form" Height="480" Width="612" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid Height="448">
        <Image Name="Log_Form_Background"  Stretch="UniformToFill" />
        <Menu Height="20" HorizontalAlignment="Left" VerticalAlignment="Top" Width="590">
            <MenuItem Name="LogSave_MenuItem" Header="Сохранить в файл"></MenuItem>
        </Menu>
        <Button Name="Log_okButton" Content="ОК" Height="36" HorizontalAlignment="Left" Margin="184,400,0,0" VerticalAlignment="Top" Width="223" />
        <TextBox Name="Log_TextBox" Height="351" HorizontalAlignment="Left" Margin="12,26,0,0" VerticalAlignment="Top" Width="566" />
    </Grid>
</Window>')

        $Global:Log_Forms_xam = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $Log_Forms_xml))
        $Log_Forms_xml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
            Set-Variable -Name ($_.Name) -Value $Log_Forms_xam.FindName($_.Name) -Scope Global
        }
        if (($Global:Basic.BIN.Icon -ne $null) -and ($Global:Basic.BIN.Background -ne $null)){
            $Log_Form.Icon = ([System.Convert]::FromBase64String($Global:Basic.BIN.Icon))
            $Log_Form_Background.source = ([System.Convert]::FromBase64String($Global:Basic.BIN.Background)) # Задний фон
        } # Добавляем графику

        # Заполняем лог
        $Log = [string]$Global:Log -replace '\d\d/\d\d/\d\d\d\d', "`n"
        $Log_TextBox.Text = $Log

        # Обрабатываем кнопки
        $LogSave_MenuItem.Add_Click({
            $saveDlg = New-Object -Typename System.Windows.Forms.SaveFileDialog
            $saveDlg.FileName = $Global:Basic.ProjectName
            $saveDlg.Filter = 'Лог работы программы (*.log)|*.log'
            $saveDlg.ShowDialog()
            $Log | Out-File $saveDlg.FileName
        }) <# Сохраняем лог #>

        $Log_okButton.Add_Click({
            $Log_Forms_xam.Close()
        })

        # Отображаем форму
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем log-программы") @Global:Console_OK
        $Log_Forms_xam.ShowDialog() | Out-Null
    } <# Вывод диалогового окна Log#>

    function Start-Update () {
        $File = $Global:Basic.Var.PathInt
        if ($File -ne "") {
            Rename-Item -path $File -NewName ($File + ".bak") -Force
            sleep 1
            Copy-Item -Path $Global:Basic.Update.Files -Destination $File -Force
            sleep 3

            Add-Type -AssemblyName System.Windows.Forms
            $Notification.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($File)
            $Notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::"info"
            $Notification.BalloonTipText = "Произведено обновление программы " + $Global:Basic.ProjectName + ".`n Актуальная версия: " +  $Global:Basic.Update.Version
            $Notification.BalloonTipTitle = $Global:Basic.Var.UserDisplayName
            $Notification.Visible = $true
            $Notification.ShowBalloonTip(10)
            
            Start-Process $File -ArgumentList "CancelUpdate"
            Close
        }
    } <# Обновляем программу #>

    function Get-Update () {
        if (Test-Path $Global:Basic.Update.Files){
            $Global:Basic.Update.Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Global:Basic.Update.Files).FileVersion
            $CurrentVersion = [System.Version]::Parse($Global:Basic.CurrentVersion)

            if($Global:Basic.Update.Version -gt $CurrentVersion) {
                Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  WAIT  ] Требуется обновление. Актуальная версия: " + $Global:Basic.Update.Version + " Текущая версия: " + $Global:Basic.CurrentVersion)) @Console_Mod
                $Global:Basic.Update.Status = "Update"
                
                if ($Global:Basic.Update.Type -like "Auto"){
                    Start-Update
                <#
                # Вывод уведомления в ModernUI
                Add-Type -AssemblyName System.Windows.Forms
                $Global:Notification = New-Object System.Windows.Forms.NotifyIcon
                $Notification.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Global:Var.PathScript)
                $Notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::"info"
                $Notification.BalloonTipText = "Производится обновление программы " + $Global:Var.ProjectName + ".`n Актуальная версия: " +  $Global:UpdateVersion
				$Notification.BalloonTipTitle = $Global:Var.UserDisplayName
                $Notification.Visible = $true
                $Notification.ShowBalloonTip(10)
                #>
                }
            }
            elseif($Global:Basic.Update.Version -eq $CurrentVersion) {
                Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Используется актуальная версия" $Global:Basic.CurrentVersion) @Global:Console_OK
                $Global:Basic.Update.Status = "OK"
            }
            else {
                Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] Используется тестовая версия: " + $Global:Basic.CurrentVersion + " Актуальная версия: " + $Global:Basic.Update.Version)) @Console_ERROR
                $Global:Basic.Update.Status = "Test"
            }

        } # Проверяем доступность пути
        else {
            $Global:Basic.Update.Status = "ConnectError"
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] Сетевой путь не доступен. Обновление не возможно")) @Console_ERROR
        }

    } <# Проверка возможности обновления #>

    function Send-MailSD () {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Генирируем письмо в SD") @Global:Console_OK
        $Log = [string]$Global:Log -replace '\d\d/\d\d/\d\d\d\d', "%0D"
        $Mail = "mailto:"

        $msgBox =  [System.Windows.MessageBox]::Show($Global:Var.UserDisplayName + "
Вы хотите открыть почтовую программу по умолчанию для формирования обращения в тех.поддержку.

Внимание! Дальнейшая комуникация по возникшей проблемы будет происходить в рамка обращения с использованием почтового клиента!
(Введите описание ошибки и/или приложите скриншоты возникающей ошибки и отправьте письмо)

", $Global:Var.ProjectName + "- Выбор сертификата",'YesNo','Info')
        switch  ($msgBox) {
            'Yes' {Start-Process $Mail}
        }        
    } <# Обращение на SD #>

    ##################################################################################################################################
    ### Получаем переменные
    if ($Global:Basic.Var.Basic -eq $False){
        $Global:Basic.Var.Basic = "True"

        ##### Дополнительная для Logа
        $Global:Console_Info = @{ForegroundColor='Magenta'; BackgroundColor='DarkBlue'}
        $Global:Console_OK = @{ForegroundColor="Black"; BackgroundColor = "Green"}
        $Global:Console_Error = @{ForegroundColor="White"; BackgroundColor = "Red"}

        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] Запущена программа:" $Global:Basic.ProjectName "v." $Global:Basic.CurrentVersion) @Console_Info
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] Пользователь: "$env:USERDOMAIN"\"$env:USERNAME " (Домен:" $env:USERDNSDOMAIN")") @Console_Info
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] Устройство: "$env:LOGONSERVER"\"$env:COMPUTERNAME " ") @Console_Info
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[  INFO  ] Среда исполнения PowerShell версии" $PSVersionTable.PSVersion) @Global:Console_Info

        ##### Дополнительная информация
        # Информация о пользователе
        $Global:Basic.Var.UserDisplayName = ([adsi]"WinNT://$env:UserDomain/$env:UserName,user").fullname
        if ($Global:Basic.Var.UserDisplayName -ne $null){
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Из домена получено отображаемое имя: " + $Global:Basic.Var.UserDisplayName)) @Console_OK
        }
        else {
            $Global:Basic.Var.UserDisplayName = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI -Name LastLoggedOnDisplayName).LastLoggedOnDisplayName
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Из реестра Windows получено отображаемое имя: " + $Global:Basic.Var.UserDisplayName)) @Console_Error
        }

        Get-Update

        # Создаем временную папку
        if (-not(Test-Path $Global:Basic.Var.Temp)){
            New-Item $Global:Basic.Var.Temp -ItemType "directory" -Force | Out-Null
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Создан временный каталог " + $Global:Basic.Var.Temp)) @Console_OK
        }


    } <# Первый запуск. Выполняем самые базовые процедуры #>

    ##################################################################################################################################
    ##### Разбор вызовов
    if ($Args -like "ServiceDesk") {Send-MailSD}
    if ($Args -like "Log")         {Open-Log}
    if ($Args -like "Help")        {Open-Help}
    if ($Args -like "About")       {Open-About}
    if ($Args -like "Close")       {Close}
}

function Test-KeyPath ($Pach) {
<# Проверяем содержимое контейнера #>
    $File = @{
        header = $Pach + "\header.key"
        masks = $Pach + "\masks.key"
        masks2 = $Pach + "\masks2.key"
        name = $Pach + "\name.key"
        primary = $Pach + "\primary.key"
        primary2 = $Pach + "\primary2.key"
    }

    if ((Test-Path $File.header) -and (Test-Path $File.masks) -and (Test-Path $File.masks2) -and (Test-Path $File.name) -and (Test-Path $File.primary) -and (Test-Path $File.primary2)) {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Все файлы на месте") @Global:Console_OK
        return "$True"
    }
    else {
        return "$False"
    }
}

function Install-Sert ($Path) {
<# Устанавливаем контейнер в реестр #>
    $LatterKeyСarrier = "B:"
    $PathKeys = $Global:Basic.Var.Temp + "\Key"
    $PathKeysTemp = $Global:Basic.Var.Temp + "\KeysTMP"

    ### Копируем на отторгаемый носитель
    if (Test-Path -Path $PathKeys){
        Remove-Item -path $PathKeys -Recurse -Force
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Очищен каталог: " + $PathKeys)) @Console_WAIT
    }
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Создаем каталог " $PathKeys) @Global:Console_OK
    New-Item -Path $PathKeys -ItemType Directory -Force

    if (Test-Path -Path $LatterKeyСarrier){
        Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] Имеется устройство " + $KeyСarrier)) @Console_ERROR
        
                $msgBox =  [System.Windows.MessageBox]::Show("Ошибка в работе программы (имеется устройство B:) `n`n`n"+ 
$Global:Var.UserDisplayName + "
Обратитесь в техническую поддержку.
(при обращении сделайте скриншот данного окна)

[ Программа будет закрыта ].
", $Global:Basic.ProjectName + "- Letter",'OK','Error')
        Start-BasicFunction "Close"
    }
    else {
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Формируем отторгаемый носитель информации " $LatterKeyСarrier) @Global:Console_OK
        subst $LatterKeyСarrier $PathKeys
    }

    # Копируем папку на съемный носитель
    Copy-Item -Path $Path -Destination ($LatterKeyСarrier + "\") -Recurse
    
    # Получаем имя контейнера
    $ContList = csptest "-keyset", "-enum_cont", "-verifycontext", "-fqcn", "-machinekeys" #| ConvertTo-Encoding cp866 windows-1251
    $i = 0
    while ((($ContList -split "`r?`n")[$i]) -ne $null) {
        $Cont = (($ContList -split "`r?`n")[$i])
        if ($Cont -match ("\\\\.\\FAT12_B\\*")) {
            $ContName = $Cont -replace '\\', "`n"
            $ContSrc = ($ContName -split "`r?`n")[4]
            Break
        }
        $i = $i + 1
    }
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Получено имя контейнера: " + $ContSrc)) @Console_OK

    # Копируем контейнер в реестр
    $ContDest = "\\.\REGISTRY\" + $ContSrc + "-copy"
    csptest '-keycopy', '-contsrc', $ContSrc, '-contdest', $ContDest, '-pindest=""' #| ConvertTo-Encoding cp866 windows-1251
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Контейнер скопирован в " + $ContDest)) @Console_OK

    ### Чистим
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Очищаем значения ")) @Console_OK
    subst $LatterKeyСarrier '/d'
    sleep 1
    Remove-Item -path $PathKeys -Recurse -Force
    if (Test-Path -Path $PathKeysTemp){
        Remove-Item -path $PathKeysTemp -Recurse -Force
    }

    ### Устанавливаем сертификат
    $ContReg = $ContSrc + "-copy"
    csptest '-property', '-cinstall', '-container', $ContReg #| ConvertTo-Encoding cp866 windows-1251
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[   ОК   ] Установлен сертификат из контейнера " + $ContReg)) @Console_OK

    # Уведомление об окончании установки
    Add-Type -AssemblyName System.Windows.Forms
    $Global:msg_SertInstall = New-Object System.Windows.Forms.NotifyIcon
    $msg_SertInstall.Icon = ([System.Convert]::FromBase64String($Global:Basic.BIN.Icon))
    $msg_SertInstall.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::"info"
    $msg_SertInstall.BalloonTipText = $Global:Argument.UserDisplayName + "`n Вам добавлен сертификат ЭЦП: `n" + $ContReg
    $msg_SertInstall.BalloonTipTitle = $Global:Basic.Var.UserDisplayName # $Global:Var.ProjectName
    $msg_SertInstall.Visible = $true
    $msg_SertInstall.ShowBalloonTip(10)
}

function ConvertTo-Encoding ([string]$From, [string]$To){
# Исходник: https://xaegr.wordpress.com/2007/01/24/decoder/
    Begin{
        $encFrom = [System.Text.Encoding]::GetEncoding($from)
        $encTo = [System.Text.Encoding]::GetEncoding($to)
    }
    Process{
        $bytes = $encTo.GetBytes($_)
        $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)
        $encTo.GetString($bytes)

    }
} # Преобразование кодировок

function Open-Forms () {
<# Формируем и открываем главное окно программы #>
    ### Формируем GUI
    [xml]$Global:xmlWPF_Forms = (
'<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="' + $Global:Basic.ProjectName + '" Height="394" Width="525" Name="Forms" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <Image Name="Forms_Background"  Stretch="UniformToFill" />
        <Menu Height="23" HorizontalAlignment="Left" VerticalAlignment="Top" Width="503" Background="#0D000000">
            <MenuItem Header="Открыть">
                <MenuItem Name="cpanel_Open_MenuItem" Header="Открыть КриптоПРО"></MenuItem>
                <MenuItem Name="certmgr_Open_MenuItem" Header="Открыть раздел Cертификаты Windows"></MenuItem>
            </MenuItem>
            <MenuItem Name="Help_MenuItem" Header="Справка"></MenuItem>
            <MenuItem Header="О программе">
                <MenuItem Name="MailSD_MenuItem" Header="Отправить письмо в SD"></MenuItem>
                <MenuItem Name="Log_MenuItem" Header="Просмотр лога"></MenuItem>
                <MenuItem Name="About_MenuItem" Header="About"></MenuItem>
            </MenuItem>
        </Menu>
        <TextBlock Text="Здравствуйте  ' + $Global:Basic.Var.UserDisplayName + '! 
Вас приветствует программа" FontSize="12" FontWeight="Normal" Height="34" HorizontalAlignment="Center" Margin="10,29,8,0" TextAlignment="Center" TextWrapping="Wrap" VerticalAlignment="Top" Width="485" />

        <TextBlock Height="27" HorizontalAlignment="Center" Margin="12,63,6,0" VerticalAlignment="Top" Width="485" TextWrapping="Wrap" FontWeight="Bold" TextAlignment="Center" FontSize="15"
                   Text="Менеджер управления сертификатами ЭЦП" />
        <Grid Height="248" HorizontalAlignment="Left" Margin="10,95,0,0" VerticalAlignment="Top" Width="487">
            <TextBlock Height="65" HorizontalAlignment="Left" Margin="6,13,0,0"  VerticalAlignment="Top" Width="475" FontSize="14" TextWrapping="Wrap" TextAlignment="Center" Text="Для установки выбирите соответствующей кнопкой путь к архиву или папке содержащей контейнер сертификата.
Для установки нажмите кнопку Установить"/>
            <TextBox Name="PathSrc_textBox" Text="Укажите путь к архиву или папке содержащей контейнер" Height="62" HorizontalAlignment="Left" Margin="6,100,0,0" VerticalAlignment="Top" Width="367" Background="#FFF8F5F5" />
            <Button Name="FolderSrc_button" Content="Выбрать папку" Height="28" HorizontalAlignment="Right" Margin="0,100,6,0" VerticalAlignment="Top" Width="102" />
            <Button Name="FileSrc_button" Content="Выбрать архив" Height="28" HorizontalAlignment="Right" Margin="0,134,6,0" VerticalAlignment="Top" Width="102" />
            <Button Name="ContInstall_button" Content="Установка нового сертифика из контейнера" Height="32" HorizontalAlignment="Left" Margin="53,198,0,0" VerticalAlignment="Top" Width="386" FontWeight="Bold" />
        </Grid>
    </Grid>
</Window>'
)

    $Global:xamGUI_Forms = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF_Forms))
    $xmlWPF_Forms.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        Set-Variable -Name ($_.Name) -Value $xamGUI_Forms.FindName($_.Name) -Scope Global
    }

    if (($Global:Basic.BIN.Icon -ne $null) -and ($Global:Basic.BIN.Background -ne $null)){
        $Forms.Icon = ([System.Convert]::FromBase64String($Global:Basic.BIN.Icon))
        $Forms_Background.source = ([System.Convert]::FromBase64String($Global:Basic.BIN.Background)) # Задний фон
    }
    #######################################
    ### Формируем меню
    $Help_MenuItem.Add_Click({ Start-BasicFunction "Help" })
    $MailSD_MenuItem.Add_Click({ Start-BasicFunction "ServiceDesk" })
    $Log_MenuItem.Add_Click({ Start-BasicFunction "Log" })
    $About_MenuItem.Add_Click({ Start-BasicFunction "About" })

    $cpanel_Open_MenuItem.Add_Click({
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем Крипто ПРО CSP") @Global:Console_OK
        cpanel
    })
    $certmgr_Open_MenuItem.Add_Click({
        Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем Сертификаты в Windows") @Global:Console_OK
        Start-Process certmgr.msc
    })

    #######################################
    ### Обрабатываем кнопки
    $FolderSrc_button.Add_Click({
    # https://www.codegrepper.com/code-examples/shell/browse+for+folder+powershell
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
        $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowserDialog.Description = "Select a folder"
        $folderBrowserDialog.rootfolder = "MyComputer"
        if ($folderBrowserDialog.ShowDialog() -eq "OK") {
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Получили путь к папке: " $folderBrowserDialog.SelectedPath) @Global:Console_OK
            $PathSrc_textBox.Text = $folderBrowserDialog.SelectedPath
        }
    }) # Выбор папки

    $FileSrc_button.Add_Click({
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'Архив (*.zip)|*.zip'
        }
        if ('OK' -eq $FileBrowser.ShowDialog()) {
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Получили путь к файлу: " $FileBrowser.FileName) @Global:Console_OK
            $PathSrc_textBox.Text = $FileBrowser.FileName
        }
    }) # Выбор архива

    $ContInstall_button.Add_Click({
        

        if ($PathSrc_textBox -like '*.zip') {
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Обрабатываем архив"  $PathSrc_textBox.Text) @Global:Console_OK
            $SourceFile = $PathSrc_textBox.Text
            $PathKeysTemp = $Global:Basic.Var.Temp + "\KeysTMP"

            & {sz x "$SourceFile" -aoa -o"$PathKeysTemp" -y}

            $cat = Get-ChildItem $PathKeysTemp -Directory -Recurse
            $i = ($cat.Length) - 1
            Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Имя контейнера"  $cat[$i].Name) @Global:Console_OK

            if (Test-KeyPath $cat[$i].FullName) {
                Install-Sert $cat[$i].FullName
            }

        } <# Обрабатываем если архив #>
        elseif(Test-Path $PathSrc_textBox.Text) {
            if (Test-KeyPath $PathSrc_textBox.Text){
                Install-Sert $PathSrc_textBox.Text
            }
        } <# Обрабатываем если папка #>
        else {
            Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] Контейнер не выбран")) @Console_ERROR
            
            $msgBox =  [System.Windows.MessageBox]::Show($Global:Basic.Var.UserDisplayName + "

Вы не выбрали путь содержащий контейнер сертифика. Дальнейшая работа не возможна.
", $Global:Basic.ProjectName + "- Выбор сертификата",'OK','Error')
        }
    }) # Установка

    #######################################
    ### Выводим форму
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] Открываем GUI") @Global:Console_OK
    $xamGUI_Forms.ShowDialog() | Out-Null
} # GUI

##################################################################################################################################
##################################################################################################################################
sleep 1
$Global:Basic.Var.PathInt = (Get-Process -Id (Get-CimInstance win32_process -Filter "ProcessId = $PID").ParentProcessId).Path
Write-Host (Write-Output -OutVariable +global:Log (Get-Date) "[  INFO  ] Путь инициирующего процесса: " $Global:Basic.Var.PathInt) @Console_Info
Start-BasicFunction

### Проверяем наличие ПО и формируем вызовы
# Крипто ПРО CSP
 $csptest = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\csptest.exe")."(default)"
if (Test-Path $csptest) {
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] КриптоПРО CSP установлен") @Global:Console_OK
   
    Set-Alias csptest $csptest
    Set-Alias cpanel ($csptest -replace 'csptest.exe', 'cpanel.cpl')
}
else {
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] Крипто ПРО не установлен" + $Global:Argument.LatterKeyСarrier)) @Console_ERROR
    
    $msgBox =  [System.Windows.MessageBox]::Show("Ошибка в работе программы (КриптоПРО CSP не установлен) `n`n`n"+ 
$Global:Var.UserDisplayName + "
Обратитесь в техническую поддержку.
(при обращении сделайте скриншот данного окна)

[ Программа будет закрыта ].
", $Global:Basic.ProjectName + "- Выбор сертификата",'OK','Error')    
    Start-BasicFunction "Close"
}

# 7-ZIP
if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
    Write-Host (Write-Output -OutVariable +Global:Log (Get-Date) "[   ОК   ] 7-ZIP установлен") @Global:Console_OK
    Set-Alias sz "C:\Program Files\7-Zip\7z.exe"
}
elseif(Test-Path "C:\Program Files (x86)\7-Zip\7z.exe"){
    Set-Alias sz "C:\Program Files (x86)\7-Zip\7z.exe"
}
else {
    Write-Host (Write-Output -OutVariable +global:Log (Get-Date) ("[  ERROR ] 7-ZIP не установлен" + $Global:Argument.LatterKeyСarrier)) @Console_ERROR

    $msgBox =  [System.Windows.MessageBox]::Show("Ошибка в работе программы (7-ZIP не установлен) `n`n`n"+ 
$Global:Var.UserDisplayName + "
Обратитесь в техническую поддержку.
(при обращении сделайте скриншот данного окна)

[ Программа будет закрыта ].
", $Global:Basic.ProjectName + "- Выбор сертификата",'OK','Error')
    Start-BasicFunction "Close"
}


### Оброботчик
if ($Args -like "") {
    Open-Forms
}

if ($Args -like "CancelUpdate") {
    Remove-Item ($Global:Basic.Var.PathInt + ".bak") -Force
    Open-Forms
}

Start-BasicFunction "Close"


# SIG # Begin signature block
# MIII0wYJKoZIhvcNAQcCoIIIxDCCCMACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU8ZrHVc987Ivq5pECbNVRhUNM
# bmigggYvMIIGKzCCBROgAwIBAgITUgAAAqIGyIV3fcz/6AAAAAACojANBgkqhkiG
# 9w0BAQsFADBWMRIwEAYKCZImiZPyLGQBGRYCcnUxGDAWBgoJkiaJk/IsZAEZFghr
# ZGwtdGVzdDESMBAGCgmSJomT8ixkARkWAmFkMRIwEAYDVQQDEwlNU0stQ0EtMDEw
# HhcNMjExMjIwMTQzODE5WhcNMjYxMjE5MTQzODE5WjCBvjESMBAGCgmSJomT8ixk
# ARkWAnJ1MRgwFgYKCZImiZPyLGQBGRYIa2RsLXRlc3QxEjAQBgoJkiaJk/IsZAEZ
# FgJhZDESMBAGA1UECwwJS0RMX1VzZXJzMRUwEwYDVQQLDAxtb3Njb3dfdXNlcnMx
# EjAQBgNVBAsTCUVtcGxveWVlczE7MDkGA1UEAwwy0KTQtdGC0LjRgdC+0LIg0J3Q
# uNC60LjRgtCwINCT0LXQvdC90LDQtNGM0LXQstC40YcwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQDv89gz9GwwHUs46Cd5XKK77lPIhxUF39qv9y5TlBUB
# NeAEk0g8r4L55jSu3/6aCe3GzqFuPKm2QE0e7+n+X6ag65jjZ62KMfqod99ByW1J
# LJRHtjQgan1Ol4sjK00jVQeWUzFFANdBatvFXYYN0BGW9YGWtIuU5ZhV9PjEu8EV
# oi+3NR0gZFBY2qibLvJNx36t/3Wu5PWJR3/F8s83NN9bN2FdsYAmopkP2MA1liE+
# coIDOxo3bcHyacs2CLutAa+AnMaSlKXx/HNk46Hg7f8NUobJSLvUurTNYdrERkE9
# nzTSjS99+02r72Ff5z0u7r6COctBcQUnBoI3sV/Z6ULZAgMBAAGjggKHMIICgzA9
# BgkrBgEEAYI3FQcEMDAuBiYrBgEEAYI3FQiHqeZwhMPrYIbRjzSF+tUugtmLGmSG
# 6/wXhKG9QAIBZAIBEjATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNVHQ8BAf8EBAMC
# B4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUk2Dw3TvJ
# VfepYpjHRthg3DYZ4cUwHwYDVR0jBBgwFoAUqEvvdzWqpUDDe9W6dLEWf9lCWSow
# gc8GA1UdHwSBxzCBxDCBwaCBvqCBu4aBuGxkYXA6Ly8vQ049TVNLLUNBLTAxLENO
# PW1zay1jYS0wMSxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049
# U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1hZCxEQz1rZGwtdGVzdCxEQz1y
# dT9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JM
# RGlzdHJpYnV0aW9uUG9pbnQwgcEGCCsGAQUFBwEBBIG0MIGxMIGuBggrBgEFBQcw
# AoaBoWxkYXA6Ly8vQ049TVNLLUNBLTAxLENOPUFJQSxDTj1QdWJsaWMlMjBLZXkl
# MjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWFkLERD
# PWtkbC10ZXN0LERDPXJ1P2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1j
# ZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MCoGA1UdEQQjMCGgHwYKKwYBBAGCNxQCA6AR
# DA9ORmV0aXNvdkBrZGwucnUwDQYJKoZIhvcNAQELBQADggEBADfxaN5pdXjTMPKm
# e/F/6lbA/4hXeFD/LkBTKzDnnJDqpMcoiqBPDJtAePCJ0KVuAGQ3BwmrHmAzQQoY
# eu9zesso2yDQAdLlYPJ5xnrdlw2JxkeBWUreksohZaUdY2CfFSywx0OEvP+YEjb/
# 1YiSsM6qTdobQs9eN0MAeyyp2tqOxb/7ILQk5E4mugeSC5UF05BmagMz0dShTHYh
# JLdC3RymHMlpujfbqdiy+pDWCKZQ0PeNmkuKIVdCYpdAgk0W+0gqJPEquUcSm9OI
# 8xZ5K0/MiyvPXWCxL8D/MwrqMW0gCj6s4Nu9yi26Pk5/C8EcRUJ86eKmGkOjYOiT
# En3gQNYxggIOMIICCgIBATBtMFYxEjAQBgoJkiaJk/IsZAEZFgJydTEYMBYGCgmS
# JomT8ixkARkWCGtkbC10ZXN0MRIwEAYKCZImiZPyLGQBGRYCYWQxEjAQBgNVBAMT
# CU1TSy1DQS0wMQITUgAAAqIGyIV3fcz/6AAAAAACojAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# /Ac7rYH2r2+fU+4jwAFKDpJ4V3MwDQYJKoZIhvcNAQEBBQAEggEAvw96JMCypPQ3
# ivrKajNBb54T6R5AOtAbgtcd6WfPC+hsor8/tG4YJJ2f8QbdJs9t8yk+CseOeJX0
# pvXbzzuWFKEC0PH/QQg4u7aQvsV6wTOPkFfVgTlQc1z6s146JLW66lvmmqhmgz9V
# v74Krnc0yUKUGCTjzOUBZtSjiRn16Tv/cY41YYBm7Fs0OnMeIlaUZREbX/3JW+oZ
# skB6hwni33Pijzfn8DaX6oOQRdSpV9F7CZBrLqtScBYV+hoRc5MUTLF6UUk+/o3y
# jcCSmUXUUWkd2eB2MJVsu7z9UAqWiPwKn0lcEeqSPqAVYXX2air1MXaWRLwVT/bp
# L0dpvoe4bA==
# SIG # End signature block
