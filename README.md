# octoprint_upload
small vbs script for uploading  .gcode files to OctoPrint  (for Windows only)

Загрузка файлов gcode в Octoprint c поддержкой русских имен файлов.
В начале файла нужно исправить настройки
' Настройки 
OctoPrint_ApiKey = "01D43472B88F4D1C879AF6EFE8073C1B"    ' OctoPrint API Key
OctoPrint_URL    = "http://192.168.1.104:5000"           ' Адрес OctoPrint
OctoPrint_Select = "true"                                ' true или false.  Выбор файла сразу после загрузки
OctoPrint_Print  = "false"                               ' true или false.  Печать файла сразу после загрузки
ShowStatistics   = "true"                                ' true или false.  Показывать статистику работы скрипта

Для автоматической загрузки из Symplify3D на закладке Scripts нужно добавить следующую строку
wscript c:\utils\octoprint_upload.vbs "[output_filepath]"
где c:\utils\ папка где расположен скрипт
