Attribute VB_Name = "modMAD"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''' Manage Audit Data (MAD) module ''''''''''''''''''''''''
        ''''''''''''''''''' МОДУЛЬ РАБОТЫ С АУДИТ-ДАННЫМИ ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода
Public isSQLChecked As Boolean
Public tmpSQLAvailable As Boolean

Public Function RegGetAuditData(ByVal AuditProp As String) As String
Dim AuditValue As String
'получаем переменные из функции и передаем их в функцию модуля работы с реестром
'если данных нет - так и пишем
Call fReadValue("HKLM", "Software\LARS", AuditProp, "S", "Нет данных", AuditValue)
RegGetAuditData = AuditValue
End Function

Public Function RegPutAuditData(ByVal AuditProp As String, ByVal auditdata As String)
'данная функция вызывается методом Call, использует два параметра - соответственно
'куда класть данные и какие именно данные класть
Call fWriteValue("HKLM", "Software\LARS", AuditProp, "S", auditdata)
End Function

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

Public Function SQLExecute(ByVal SQLRequestString As String, SQLMode As laSQLMode, Optional ByVal ParameterToRead As String) As Variant
On Error GoTo SQL_error
Dim SQLResponse As Variant
Dim SQL As New ADODB.Connection
    Dim SQLData As New ADODB.Recordset
    Dim SQLRequest As String, SQLAPRequest As String
    SQL.Open SQLConnString
    WriteToLog "Исполняю функцию SQLAuditData SQLExecute. Строка исполнения:" & vbCrLf & SQLRequestString
    SQLData.Open SQLRequestString, SQL, adOpenKeyset
        If SQLMode = laRX Then
            SQLExecute = SQLData.Fields(ParameterToRead).Value
        End If
    SQL.Close
    Exit Function
'frmWriteAuditData.tDeb.Text = frmWriteAuditData.tDeb.Text & vbCrLf &
SQL_error:
Dim SQLErrNumber As Long, SQLErrDescription As String
SQLErrNumber = Err.Number
SQLErrDescription = Err.description
    If SQLErrNumber <> 0 Then
        WriteToLog " "
        WriteToLog "МОДУЛЬ SQL СООБЩИЛ ОБ ОШИБКЕ:"
        WriteToLog "Ошибка SQL " & SQLErrNumber & ":" & vbCrLf & SQLErrDescription
    End If
SQLExecute = SQLErrNumber
End Function

Public Function isSQLAvailable() As Boolean
Dim sqlCheckIfAvailable As Long
If isSQLChecked = False Then
    sqlCheckIfAvailable = SQLExecute("SELECT * FROM dbo.larspc", laRX)
        If sqlCheckIfAvailable = -2147467259 Then
            tmpSQLAvailable = False
            WriteToLog " "
            WriteToLog "Модуль проверки SQL сообщил: СЕРВЕР SQL НЕ ДОСТУПЕН"
            WriteToLog " "
        Else
            tmpSQLAvailable = True
            WriteToLog " "
            WriteToLog "Модуль проверки SQL сообщил: УСПЕШНОЕ СОЕДИНЕНИЕ"
            WriteToLog " "
        End If
    isSQLChecked = True
End If
isSQLAvailable = tmpSQLAvailable
End Function
