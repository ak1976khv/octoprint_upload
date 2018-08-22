' Настройки 
OctoPrint_ApiKey = "01D43472B88F4D1C879AF6EFE8073C1B"    ' OctoPrint API Key
OctoPrint_URL    = "http://192.168.1.104:5000"           ' Адрес OctoPrint
OctoPrint_Select = "true"                                ' true или false.  Выбор файла сразу после загрузки
OctoPrint_Print  = "false"                               ' true или false.  Печать файла сразу после загрузки
ShowStatistics   = "true"                                ' true или false.  Показывать статистику работы

' Основная часть скрипта

StartTime = Timer 
LogStr = ""

Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
   MsgBox "Не указано имя файла." & vbCrLf & vbCrLf & "octoprint_upload.vbs filename"
   WScript.Quit
End If


strPath = objArgs(0)

Dim strFile, strExt, strContentType, strBoundary, bytData, bytPayLoad

On Error Resume Next
With CreateObject("Scripting.FileSystemObject")
	If .FileExists(strPath) Then
		strFile = .GetFileName(strPath)
	Else
		MsgBox "Файл '" & strPath & "' не найден"
		WScript.Quit
	End IF
End With
With CreateObject("ADODB.Stream")
	.Type = 1
	.Mode = 3
	.Open
	.LoadFromFile strPath
	If Err.Number <> 0 Then
		MsgBox "Ошибка чтения файла: " & Err.Description & " (" & Err.Number & ")"
		WScript.Quit
	End If
	bytData = .Read
	.Close
End With
LogStr = (Timer - StartTime) & " Чтение файла."

strBoundary = String(4, "-") & Replace(Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36), "-", "")

With CreateObject("ADODB.Stream")
	.Type = 2
	.Mode = 3
	.Charset = "UTF-8"
	.Open
	.WriteText strFile
	.Position = 0
	.Charset = "Windows-1251"
	strRFC5987 = URLEncode(.ReadText)
	.Close
End With
LogStr = LogStr & vbCrLf & (Timer - StartTime) & " Кодировка имени."  

With CreateObject("ADODB.Stream")
	.Mode = 3
	.Charset = "utf-8"
	.Open
	.Type = 2
	.WriteText "--" & strBoundary & vbCrLf
	.WriteText "Content-Disposition: form-data; name=""file""; filename*=UTF-8''" & strRFC5987 & vbCrLf
	.WriteText "Content-Type: ""application/octet-stream""" & vbCrLf & vbCrLf
	.Position = 0
	.Type = 1
	.Position = .Size
	.Write bytData
	.Position = 0
	.Type = 2
	.Position = .Size
	.WriteText vbCrLf & "--" & strBoundary & vbCrLf
	.WriteText "Content-Disposition: form-data; name=""select""" & vbCrLf & vbCrLf
	.WriteText OctoPrint_Select
	.WriteText vbCrLf & "--" & strBoundary & vbCrLf
	.WriteText "Content-Disposition: form-data; name=""print""" & vbCrLf & vbCrLf
	.WriteText OctoPrint_Print
	.WriteText vbCrLf & "--" & strBoundary & "--" & vbCrLf
	.Position = 0
	.Type = 1
	bytPayLoad = .Read
	.Close
End With
LogStr = LogStr & vbCrLf & (Timer - StartTime) & " Формирование запроса."  

REM With CreateObject("ADODB.Stream")
	REM .Type = 1
	REM .Mode = 3
	REM '.Charset = "Windows-1251"
	REM .Open
	REM .Write bytPayLoad
	REM .Position = 0
	REM .SaveToFile "C:\utils\curl\1.log", 2
	REM .Close
REM End With
REM WScript.Quit

With CreateObject("MSXML2.ServerXMLHTTP") 
	.SetTimeouts 0, 60000, 300000, 300000
	.Open "POST", OctoPrint_URL & "/api/files/local", False 
	.SetRequestHeader "X-Api-Key", OctoPrint_ApiKey
	.SetRequestHeader "Content-type", "multipart/form-data; boundary=" & strBoundary
	.Send bytPayLoad
	If Err.Number <> 0 Then
		MsgBox Err.Description & " (" & Err.Number & ")"
	End If
	If .Status <> "200" AND .Status <> "201" Then MsgBox .Status & " " & StatusText & " " & .ResponseText
End With
LogStr = LogStr & vbCrLf & (Timer - StartTime) & " Отправка файла."  
If ShowStatistics = "true" Then MsgBox LogStr


Function URLEncode( StringVal )
	Dim i, CharCode, Char, Space
	Dim StringLen

	StringLen = Len(StringVal)
	ReDim result(StringLen)

	Space = "%20"

	For i = 1 To StringLen
		Char = Mid(StringVal, i, 1)
		CharCode = Asc(Char)
		If 97 <= CharCode And CharCode <= 122 _
				Or 64 <= CharCode And CharCode <= 90 _
				Or 48 <= CharCode And CharCode <= 57 _
				Or 45 = CharCode _
				Or 46 = CharCode _
				Or 95 = CharCode _
				Or 126 = CharCode Then
			result(i) = Char
		ElseIf 32 = CharCode Then
			result(i) = Space
		Else
			result(i) = "%" & Hex(CharCode)
		End If
	Next
	URLEncode = Join(result, "")
End Function
