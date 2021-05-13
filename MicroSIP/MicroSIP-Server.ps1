<# README
Copyright:        fenigen
License:          Apache License 2.0
License URL:      http://www.apache.org/licenses/LICENSE-2.0
Author:           NFetisov

Version:          1.0.0.2
File Description: Выгрузка телефонного справочника для MicroSIP
#>

<# История изменений:
1.0.0.0 (05.05.2021) - Создан скрипт для автоматического создания файла конфигурации MicroSIP (во время установки ПО). Без тестирования
1.0.0.1 (11.05.2021) - Проверка работы. Исправлены пути.
1.0.0.2 (11.05.2021) - Небольшие правки. Исправление ошибки перезаписи файла.
1.0.0.3 (12.05.2021) - Небольшие правки.
#>

$TimeStart = Get-Date
CLS
# $PSVersionTable.PSVersion
#[System.Windows.Forms.MessageBox]::Show("НАЧАЛО")
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
Write-Host " "

##### START #####
# Подключение подулей
Import-module activedirectory

# Объявляем переменные
$ou = " "
$file = " "
$exp = @()

##### ФОРМИРОВАНИЕ #####
# Получаем данные
$user = Get-ADUser -SearchBase $ou -Filter {(telephoneNumber -ne "null") -and (Enabled -eq "true")} -Property CN, givenName, sn, telephoneNumber, MobilePhone, ipPhone, mail, streetAddress, City, st, postalCode, Department, Title, Enabled 
Write-Host ":-) Получены данные из $ou" -ForegroundColor Black -BackgroundColor green

# Формируем XML
$exp = '<?xml version="1.0"?>
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

# Формируем XML
$exp = $exp + '</contacts>'

Write-Host ":-) Подготовлен XML" -ForegroundColor Black -BackgroundColor green
# Сохраняем файл
Remove-Item $file
Write-Host ":-) Старый файл удален в  $file" -ForegroundColor Black -BackgroundColor green
$exp | Out-File -FilePath $file -Encoding UTF8
Write-Host ":-) Новый файл сохранен в $file" -ForegroundColor Black -BackgroundColor green

##### END #####
Write-Host " "
$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host ":-D ВЫПОЛНЕНО за "$Time -ForegroundColor white -BackgroundColor blue
Write-Host "##############################################################" -ForegroundColor white -BackgroundColor blue
#[System.Windows.Forms.MessageBox]::Show("ГОТОВО")

# Закрытие с кодом возврата (при необходимости)
# [System.Environment]::Exit(0)
