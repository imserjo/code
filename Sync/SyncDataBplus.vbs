VersionScript=47
' Дата  модификации: 30.09.2016
' Начало разработки: 26.01.2017
' Скрипт - синхронизатор данных, файлов, структуры БД
'Option Explicit
Dim wsh, cnn, fso, msg, comp, MSSQL, WebAd
Set wsh = CreateObject("WScript.Shell")
Set cnn = CreateObject("adodb.connection")
Set fso = CreateObject("Scripting.FileSystemObject")

' --------------------------------
MSSQL = "Provider=SQLOLEDB;Server=tcp:127.0.0.1,3300;Database=master;Integrated Security=SSPI;"
WebAd = "http://mysite.ru/sync/" '
comp = wsh.Environment("Process")("ComputerName")

Dim FileLocker
Set FileLocker = new cScriptFileLocker

'Проверяем состояние блокировки
if Not FileLocker.Locked Then 
    'Блокируем файл скрипта
    FileLocker.Lock
	wsh.Popup "Скрипт версии " & VersionScript & " стартовал!", 5, "Удача (Popup 5 сек)", vbInformation
Else
    'В ином случае сообщаем о том, что скрипт уже выполняется
	'wsh.Popup "Это вторая копия, запуск отменен!", 5, "СТОП (Popup 5 сек)", vbCritical
	WScript.Quit
End if

'GetNewVersionScript

Do Until 1 = 0
	WebAd = GetServerWeb()
	If WebAd<>"" Then 
	    msg = ""
		Err.Clear: GetNewVersionScript
	    Err.Clear: SyncStrukData
	'    Err.Clear: SyncStruk
	'    Err.Clear: SyncData
	    Err.Clear: SyncFile
	    Err.Clear: LogSave
	End If
    WScript.sleep 300000
Loop

'---------------------------------------------------------
Class cScriptFileLocker
    Private oTs
    'Блокировка файла
    Public Sub Lock()
        if Not isObject(oTs) Then Set oTs = fso.OpenTextFile(WScript.ScriptFullName,1,False)
    End Sub
    'Разблокировка файла
    Sub Unlock()
        Set oTs = Nothing
    End Sub
    'Получение состояния блокировки
    Public Property Get Locked()
        On Error Resume Next
        fso.MoveFile WScript.ScriptFullName, WScript.ScriptFullName
        'Если при попытке переместить файл произошла ошибка 70, значит файл уже открыт на чтение
        If Err.number = 70 Then 
            Locked = True
        Else
            Locked = False
        End if
    End Property
End Class

Function tUrl(asp, table, ts)
    tUrl = WebAd & asp & "?table=" & table & "&ts=" & ts & "&comp=" & comp & "&xlam=" & Timer
End Function

Function mmm(txt)
    msg = msg & txt & ";" & vbNewLine
    Err.Clear
End Function

Function LogSave()
On Error Resume Next
    Set z = CreateObject("MSXML2.XMLHTTP")
    url = Left(tUrl("ssLast.asp", "", "") & "&ll=Версия синхронизатора: " & VersionScript & ";" & vbnewline & msg, 1024)
    z.Open "POST", url, False
    z.send
End Function

Function SyncStrukData()
On Error Resume Next
    Set z = CreateObject("MSXML2.XMLHTTP")
    z.Open "POST", tUrl("ssFirst.asp", "", ""), False
    z.send

    TableCount = 0
    
    If cnn.state <> 1 Then cnn.Open MSSQL
    Set r = CreateObject("ADODB.RecordSet")
    r.Open z.ResponseStream
If Err <> 0 Then
    mmm "Таблиц в синхронизации: " & TableCount
    r.Close
    cnn.Close
    Err.Clear
    Exit Function
End If
    Do While Not r.EOF
        If Err <> 0 Then
            mmm "SyncStrukData [1] " & Err.Description
            Exit Do
        End If

		if r("scriptvalue").Value<>"" then
'            uExe r("bdlocal"), r("scriptvalue"), r("scriptname")
            uExe r("bdlocal").Value, r("scriptvalue").Value, r("tablename").Value
		ElseIf LCase(Left(r("tablename").Value, 3)) = "log" Then
            uLog r("bdserver").Value, r("bdlocal").Value, r("tablename").Value
        Else
            uTable r("bdserver").Value, r("bdlocal").Value, r("tablename").Value, r("alls").Value, r("max_ts").Value
        End If
        r.MoveNext
    Loop
    TableCount = r.RecordCount
    r.Close
    cnn.Close

    ' логирование конец
    mmm "Таблиц в синхронизации: " & TableCount
End Function

Function SyncFile()
On Error Resume Next
    If cnn.state <> 1 Then cnn.Open MSSQL
    
    Set r = CreateObject("ADODB.Recordset")
    r.Open "Select top 20 f.* From [File].dbo.F where f.sys_s<>0 and f.f is not null order by lid", cnn, 1, 3
If Err <> 0 Then
    mmm "SyncFile [1] " & Err.Description
    r.Close
    cnn.Close
    Err.Clear
    Exit Function
End If
    TableCount = 0
    Do While Not r.EOF
        If zgSendFile(tUrl("SetFile.asp", "", "") & "&sys_id=" & r("SYS_ID").Value, r("f").Value) Then
            cnn.Execute "update [File].dbo.F set SYS_S=NULL where SYS_ID='" & r("SYS_ID").Value & "'"
            If Err <> 0 Then mmm "SyncFile [3] " & Err.Description
        End If
        r.MoveNext
If Err <> 0 Then
    mmm "SyncFile [4] " & Err.Description
    Exit Function
End If
    Loop
    TableCount = r.RecordCount
    r.Close
    cnn.Close

    r.Close
    cnn.Close

    ' логирование конец
    mmm "Файлов выгружено: " & TableCount
End Function

Function zgSendFile(url, Data)
On Error Resume Next
    Dim z
    Set z = CreateObject("MSXML2.XMLHTTP")
    z.Open "POST", url, False
    z.setRequestHeader "Content-type", "multipart/form-data"
    z.send Data
    zgSendFile = False
    zgSendFile = (z.Status = "200") And (z.responseText = "OK")
End Function

Function uExe(bdlocal, scriptvalue, scriptname)
On Error Resume Next
'    ExeS = False
    cnn.Execute "use [" & bdlocal & "]"
If Err <> 0 Then
    mmm "ErrorStruk [2] " & Err.Description
    Exit Function
End If
    cnn.Execute Replace(scriptvalue, "go", "")
'    ms = Split(scriptvalue, vbNewLine & "go" & vbNewLine)
'    For i = LBound(ms) To UBound(ms)
'        If Trim(ms(i)) <> "" Then cnn.Execute ms(i)
'		If Err <> 0 Then
'			mmm "ErrorStruk [2] " & i & " " & Err.Description
'			err.clear
'		End If
'    Next
	
If Err <> 0 Then
    mmm "ErrorStruk [3] " & Err.Description
    Exit Function
End If
    mmm "Выполнен скрипт [" & bdlocal & "] " & scriptname
'    ExeS = True
End Function

Function uLog(bdserver, bdlocal, t)
On Error Resume Next
    t0 = bdserver & ".dbo." & t
    t = bdlocal & ".dbo." & t
    ID = cnn.Execute("Select max(ID) from " & t)(0)
    If Len(ID & "") = 0 Then Exit Function
    where = " where ID<=" & ID

' наверное нет таблицы или поля ID
If Err <> 0 Then
    mmm t & " [2] " & Err.Description
    Exit Function
End If

    rcOut = cnn.Execute("Select count(*) from " & t & where)(0)
    If Len(rcOut & "") = 0 Then Exit Function
If Err <> 0 Then
    mmm t & " [3] " & Err.Description
    Exit Function
End If

    Dim u, x
    Set u = CreateObject("MSXML2.DOMDocument")
    cnn.Execute("Select * from " & t & where & " order by ID").Save u, 1 'adPersistXML

If Err <> 0 Then
    mmm t & " [4] " & Err.Description
    Exit Function
End If
    
    Set x = CreateObject("MSXML2.XMLHTTP")
    x.Open "POST", tUrl("SetLog.asp", t0, 0), False
    x.setRequestHeader "Content-Type", "text/xml"
    x.send u.XML

If Err <> 0 Then
    mmm t & " [5] " & Err.Description
    Exit Function
End If

    If x.responseText = "OK" Then cnn.Execute "delete from " & t & where
    
    mmm t & " " & rcOut
    
If Err <> 0 Then
    mmm t & " [9] " & Err.Description
End If
End Function



Function uTable(bdserver, bdlocal, t, alls, max_ts)
On Error Resume Next
    t0 = bdserver & ".dbo." & t
    t = bdlocal & ".dbo." & t
    ts = 0
    rcOut = 0
    rcIn = 0
    ts = cnn.Execute("Select max(sys_t) from " & t)(0)

	if not isnumeric(ts) then ts=0
	if not isnumeric(max_ts) then max_ts=0

	if ts>max_ts then ts=max_ts

	
	
' наверное нет таблицы или поля SYS_T
If Err <> 0 Then
    mmm t & " [2] " & Err.Description
    Exit Function
End If

    fsysd = cnn.Execute("Select SYS_D FROM " & bdlocal & ".dbo.SYS_D")(0)
    If Len(fsysd) = 0 Then fsysd = "ALL"
    Err.Clear
    
'   ' а все ли SYS_D заполнены? - если по нему фильтруем - то и запоняем
    If Not alls Then cnn.Execute "update " & t & " set SYS_D='" & fsysd & "' where SYS_D Is Null"
If Err <> 0 Then
    mmm t & " [3] " & Err.Description
End If

    If alls Then
        lfsysd = "ALL"
        'where = " where (SYS_S<>0 OR Код=-1)"
        where = " where (SYS_S<>0 OR Код=-1 OR SYS_T IS NULL OR SYS_T>" & ts & ")"
    Else
        lfsysd = fsysd
        'where = " where (SYS_S<>0 OR Код=-1) and (SYS_D='" & fsysd & "')"
        where =" where (SYS_S<>0 OR Код=-1 OR SYS_T IS NULL OR SYS_T>" & ts & ") and (SYS_D='" & fsysd & "')"
    End If

    Dim u, x, k, w
    rcOut  = cnn.Execute("Select count(*) from " & t & where )(0)
        
If Err <> 0 Then
    mmm t & " [4] " & Err.Description
    Exit Function
End If

    If rcOut > 0 Then
        Set u = CreateObject("MSXML2.DOMDocument")
'        cnn.Execute("Select * from " & t & where).Save u, 1 'adPersistXML
        cnn.Execute("Select top 10000 * from " & t & where & " order by ID").Save u, 1 'adPersistXML
    End If

If Err <> 0 Then
    mmm t & " [4.5] " & Err.Description & " rcOut=" & rcOut & " t0=" & t0 & " ts=" & ts & " max_ts=" & max_ts & " lfsysd=" & lfsysd
    Exit Function '
End If

	' есть что отправить, или на сервере свежие данные
	If rcOut>0 or max_ts>ts Then
		Set x = CreateObject("MSXML2.XMLHTTP")
		x.Open "POST", tUrl("sync.asp", t0, ts) & "&rcout=" & rcOut & "&fsysd=" & lfsysd, False
		x.setRequestHeader "Content-Type", "text/xml"
		If rcOut > 0 Then x.send u.XML Else x.send
	else
		exit function
	end if
	
If Err <> 0 Then
'    mmm t & " [5] " & Err.Description
    mmm t & " [5] " & Err.Description & " rcOut=" & rcOut & " t0=" & t0 & " ts=" & ts & " max_ts=" & max_ts & " lfsysd=" & lfsysd
    Exit Function '
End If

    If x.responseText = "rc0" Then
    ' данных нет
    Else
        Set k = CreateObject("ADODB.Recordset")
        k.Open x.ResponseStream
If Err <> 0 Then
    mmm t & " [6] " & Err.Description
    Exit Function ' наверное мусор прилетел
End If
        Set w = CreateObject("ADODB.Recordset")
        Do While Not k.EOF
            If w.state = 1 Then w.Close
            w.Open "Select * from " & t & " where SYS_ID='" & k("SYS_ID").Value & "'", cnn, 2, 3
            If w.EOF Then w.AddNew
            
If Err <> 0 Then
	mmm t & " [7] " & Err.Description
	Exit Do
End If
            
            For i = 0 To w.Fields.Count - 1
                mSave w, k, UCASE(w(i).Name)
            Next
If Err <> 0 Then Err.Clear

            w("Код" ).Value  = k("Код" ).Value
            w("SYS_T").Value = k("syst").Value
            w("SYS_S").Value = Null
            w.Update
            k.MoveNext

If Err <> 0 Then
	mmm t & " [9] " & Err.Description
	Exit Do
End If
            
        Loop
        w.Close
        rcIn = k.RecordCount
        
        ' а все ли поля есть на клиенте?
        w.Open "Select * from " & t & " where 1=2", cnn, 2, 3
        sErr = ""
        For i = 0 To k.Fields.Count - 1
            p = UCase(k(i).Name)
            If Not (p = "ID" Or p = "КОД" Or p = "SYS_T" Or p = "SYS_S" Or p = "SYST") Then
                p = w(p).Name
                If Err <> 0 Then
                    sErr = sErr & p & ", "
                    Err.Clear
                End If
            End If
        Next
        If sErr <> "" Then mmm t & " [10] нет полей в таблице: " & sErr

    End If
    If rcOut <> 0 Or rcIn <> 0 Then mmm t & " " & rcOut & "/" & rcIn
        
If Err <> 0 Then
    mmm t & " [11] " & Err.Description
End If
End Function

Function mSave(w,k,e)
on error resume Next
    If e<>"ID" And e<>"КОД" And e<>"SYS_T" And e<>"SYS_S" Then
        tmp=k(e).Name ' а есть ли поле в рекордсете с сервера?
		If Err<>0 Then Exit Function
        w(e).Value = k(e).Value
    End If
End Function

Function GetNewVersionScript()
On Error Resume Next
	dim dd, s, z
	dd = WScript.ScriptFullName '& ".txt"

    Set z = CreateObject("MSXML2.XMLHTTP")
    z.Open "GET", tUrl(WScript.ScriptName,"",""), False
    z.send

	' версия - первая строка, после знака =
	vs = clng(trim(split(split(z.ResponseText,VbNewLine)(0),"=")(1)))

	if not isnumeric(vs) then exit function
	
	if vs > VersionScript then
		' разблокировать файл'
		FileLocker.Unlock
		' переписать скрипт
		Set s = CreateObject("ADODB.Stream")
		s.Mode = 3
		s.Type = 1
		s.Open
		s.Write z.responseBody
		s.SaveToFile dd, 2
		' запустить переписанный скрипт 
		wsh.Run dd
		' завершить текущий скрипт
		wscript.Quit
	end if

	if err<>0 then
		wsh.popup err.description, 5, "GetNewVersionScript (Popup 5 сек)", vbCritical
	end if
	
End Function








Function GetServerWeb() 
On Error Resume Next
    If s = "" Then s = GetServerWebTest("http://192.168.0.13/sync/")            ' внутренний IP
    If s = "" Then s = GetServerWebTest("http://240.240.240.240/sync/")         ' внешний IP 1
    If s = "" Then s = GetServerWebTest("http://195.195.195.195/sync/")         ' внешний IP 2
    GetServerWeb = s
End Function

Function GetServerWebTest(url) 
On Error Resume Next
	GetServerWebTest = ""
    adr = Split(Split(Split(url, "://")(1), ":")(0), "/")(0)
    If Not WMIPing(adr) Then Exit Function
    
    Set z = CreateObject("MSXML2.XMLHTTP")
    z.Open "GET", Url & "TestInet.asp?hlam=" & Timer(), False
    z.send
    If z.Status = 200 And z.statusText = "OK" and z.ResponseText = "OK" Then GetServerWebTest = url
End Function

Function WMIPing(CompIP)
On Error Resume Next
	WMIPing = False
    If CompIP = "" Then Exit Function
    
    Dim objPing, objStatus
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & CompIP & "' and Timeout=4000")
    If Err <> 0 Then Exit Function
    For Each objStatus In objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            WMIPing = False 'Computer is not accessible
        Else
            WMIPing = True 'Computer status OK
        End If
    Next
End Function