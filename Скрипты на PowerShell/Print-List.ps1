<# Сборка
Имя файла: 1.0.0.1 — Print-List

Автор:            fenigen
Лицензия:         Apache License 2.0
Version:          1.0.0.1
File Description: Выгрузка с Print Server списка принтеров
Product Name:     Print-List
Copyright:        fenigen

За основу взят матерал с ресурса http://did5.ru/it/windows/powershell-sobiraem-informaciyu-o-printerax.html
#>

cls
$TimeStart = Get-Date
# Write-Host $TimeStart "######### Запуск #########" -ForegroundColor Black -BackgroundColor green

# Формируем переменные
$search = "Критерий" # Критерий отбора
$margin = "Comment" # Поле для отбора варианты: Name, PortName, Comment, Location
$Path = "C:\Temp\printers.csv"

# Объявление переменных
$PrintServer = "print-server" # Адрес или IP принт-сервера
$List = @()
$ServerList = Get-Printer -ComputerName $PrintServer #| Format-Table Name, PortName, Comment, Location

# Формирование выборки
ForEach ($item in $ServerList){
    if ($item.$margin -like "*$search*") {
        $fl = Select-Object -Property Name, PortName, Comment, Location -InputObject $item
        $List = $List + $fl
        }
}

#Сохраняем csv понятный Excel
 #$List | Export-Csv -NoClobber -Delimiter ';' -Encoding utf8 -Path $Path

# Выводим результаты в CMD
$List

$TimeEnd = Get-Date
$Time = $TimeStart - $TimeEnd
Write-Host "ВЫПОЛНЕНО за "$Time -ForegroundColor Black -BackgroundColor green

# Закрытие с кодом возврата (при необходимости)
# [System.Environment]::Exit(0)