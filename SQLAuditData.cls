VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLAuditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Data fetched from SQL Server"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCompany As String 'local copy
Private mvarWSName As String 'local copy
Private mvarWSSerial As String 'local copy
Private mvarWindowsVersion As String 'local copy
Private mvarWindowsLicenseModel As String 'local copy
Private mvarWindowsOLPSerial As String 'local copy
Private mvarOfficeVersion As String 'local copy
Private mvarOfficeLicenseModel As String 'local copy
Public Enum laSQLMode
    laTX
    laRX
End Enum


Public Property Let OfficeLicenseModel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeLicenseModel = 5
    mvarOfficeLicenseModel = vData
End Property

Public Property Get OfficeLicenseModel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OfficeLicenseModel
    OfficeLicenseModel = mvarOfficeLicenseModel
End Property

Public Sub SQLSave(ByVal CurrentHostname As String)
'Метод для сохранения данных в SQL из класс-объекта
Dim ibName As Variant
Dim ibValue As String

    'Переменных, хранящих аудит-данные столько же, сколько у нас
    'информационных полей на форме. Отталкиваемся от кол-ва переменных
    'в массиве InfoBoxes.
    'Для каждой переменной из массива ДИНАМИЧЕСКИ запрашиваем с класса AuditData
    '(здесь он обращается сам к себе) соответствующее значение.
    For Each ibName In InfoBoxes
        ibValue = CallByName(Me, ibName, VbGet)
        Call SQLWrite(ibName, HostName, ibValue)
    Next
End Sub

Public Sub SQLLoad(ByVal CurrentHostname As String)
Dim ibName As Variant
Dim ibValue As String

    For Each ibName In InfoBoxes
        ibValue = GetSQLData(ibName, CurrentHostname)
        CallByName Me, ibName, VbLet, ibValue
    Next
End Sub

Public Property Let OfficeVersion(ByVal vData As String)
Attribute OfficeVersion.VB_Description = "Версия и редакция Microsoft Office"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeVersion = 5
    mvarOfficeVersion = vData
End Property


Public Property Get OfficeVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OfficeVersion
    OfficeVersion = mvarOfficeVersion
End Property



Public Property Let WindowsOLPSerial(ByVal vData As String)
Attribute WindowsOLPSerial.VB_Description = "Серийный номер OLP ОС"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsOLPSerial = 5
    mvarWindowsOLPSerial = vData
End Property


Public Property Get WindowsOLPSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowsOLPSerial
    WindowsOLPSerial = mvarWindowsOLPSerial
End Property



Public Property Let WindowsLicenseModel(ByVal vData As String)
Attribute WindowsLicenseModel.VB_Description = "Модель лицензирования ОС"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsLicenseModel = 5
    mvarWindowsLicenseModel = vData
End Property


Public Property Get WindowsLicenseModel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowsLicenseModel
    WindowsLicenseModel = mvarWindowsLicenseModel
End Property



Public Property Let WindowsVersion(ByVal vData As String)
Attribute WindowsVersion.VB_Description = "Версия и редакция Windows"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsVersion = 5
    mvarWindowsVersion = vData
End Property


Public Property Get WindowsVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowsVersion
    WindowsVersion = mvarWindowsVersion
End Property



Public Property Let WSSerial(ByVal vData As String)
Attribute WSSerial.VB_Description = "Номер ПК с наклейки"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSSerial = 5
    mvarWSSerial = vData
End Property


Public Property Get WSSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WSSerial
    WSSerial = mvarWSSerial
End Property



Public Property Let WSName(ByVal vData As String)
Attribute WSName.VB_Description = "Сетевое имя ПК"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSName = 5
    mvarWSName = vData
End Property


Public Property Get WSName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WSName
    WSName = mvarWSName
End Property



Public Property Let Company(ByVal vData As String)
Attribute Company.VB_Description = "Название компании"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Company = 5
    mvarCompany = vData
End Property


Public Property Get Company() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Company
    Company = mvarCompany
End Property


Private Function GetSQLData(ByVal AuditParameter As String, ByVal CurrentHostname As String) As String
On Error GoTo GetSqlData_error
Dim SQLResponse As Variant
Dim SQL As New ADODB.Connection 'ADODB Connection Object
    Dim SQLData As New ADODB.Recordset 'Recordset Object
    Dim SQLRequest As String, SQLAPRequest As String ' String variable to store sql command
    
    SQL.Open _
        "Provider = SQLNCLI11.1;" & _
        "Data Source=WS0006\SQLEXPRESS;" & _
        "Initial Catalog=AIDA;" & _
        "User ID=sa;" & _
        "Password=happyness;"
        
'' Вот эту ебаную индусятину
'    SQLRequest = "select " & _
'                "[Юрлицо] as Company," & _
'                "[Office License Model] as OfficeLicenseModel," & _
'                "[Версия Офис] as OfficeVersion," & _
'                "[Windows License Model] as WindowsLicenseModel," & _
'                "[Номер лицензии Windows OLP] as WindowsOLPSerial," & _
'                "[Версия Windows] as WindowsVersion," & _
'                "[Сетевое имя] as WSName," & _
'                "[Серийный номер] as WSSerial " & _
'                "from aida.dbo.larspc where [Сетевое имя] = '" & CurrentHostname & "';"
'' тоже реализуем с небольшой динамикой

SQLAPRequest = CallByName(Ru, AuditParameter, VbGet)
SQLRequest = "SELECT " & SQLAPRequest & " AS " & AuditParameter & " FROM aida.dbo.larspc WHERE [Сетевое имя] = '" & CurrentHostname & "';"
Debug.Print SQLRequest
    SQLData.Open SQLRequest, SQL, adOpenKeyset
    GetSQLData = SQLData.Fields(AuditParameter).Value
    SQL.Close
    Exit Function
    
GetSqlData_error:
enumSQLFields = enumSQLFields - 1
End Function

Private Function SQLWrite(ByVal AuditParameter As String, ByVal CurrentHostname As String, ByVal AuditValue As String) As Boolean
Dim SQLAPRequest As String, SQLRequest As String
SQLAPRequest = CallByName(Ru, AuditParameter, VbGet)
SQLRequest = "SELECT [Сетевое имя] AS WSName FROM aida.dbo.larspc WHERE [Сетевое имя] = '" & CurrentHostname & "';"
    If (SQLExecute(SQLRequest, laRX, "WSName") = 3265) Or (SQLExecute(SQLRequest, laRX, "WSName") = 3021) Then
        SQLRequest = "INSERT INTO dbo.larspc ([Сетевое имя]) VALUES ('" & CurrentHostname & "');"
        SQLExecute SQLRequest, laTX
    Else
        SQLRequest = "UPDATE aida.dbo.larspc SET " & SQLAPRequest & " = '" & AuditValue & "' WHERE [Сетевое имя] = '" & CurrentHostname & "';"
        SQLExecute SQLRequest, laTX
    End If
End Function

'Public Function SQLExecute(ByVal SQLRequestString As String, SQLMode As laSQLMode, Optional ByVal ParameterToRead As String) As Variant
'On Error GoTo SQL_error
'Dim SQLResponse As Variant
'Dim SQL As New ADODB.Connection
'    Dim SQLData As New ADODB.Recordset
'    Dim SQLRequest As String, SQLAPRequest As String
'
'    SQL.Open _
'        "Provider = SQLNCLI11.1;" & _
'        "Data Source=WS0006\SQLEXPRESS;" & _
'        "Initial Catalog=AIDA;" & _
'        "User ID=sa;" & _
'        "Password=happyness;"
'    Debug.Print "Исполняю функцию SQLAuditData SQLExecute. Строка исполнения:" & vbCrLf & SQLRequestString
'    SQLData.Open SQLRequestString, SQL, adOpenKeyset
'        If SQLMode = laRX Then
'            SQLExecute = SQLData.Fields(ParameterToRead).Value
'        End If
'    SQL.Close
'    Exit Function
'
'SQL_error:
'Debug.Print "Ошибка SQL " & Err.Number & ":" & vbCrLf & Err.Description
'SQLExecute = Err.Number
'End Function

Private Sub Class_Initialize()
Call CheckSQLAvailability
End Sub

Public Function CheckSQLAvailability() As Boolean
Screen.MousePointer = 11
frmwait.Show
frmwait.Reason.Caption = "Проверка SQL"
Dim sqlCheckIfAvailable As Long
sqlCheckIfAvailable = SQLExecute("SELECT * FROM dbo.larspc", laRX)
    If sqlCheckIfAvailable = -2147467259 Then
        frmWriteAuditData.chkSQL.Value = 0
        frmwait.Hide
        MsgBox "Сервер SQL не доступен." & vbCrLf _
            & " Работа проводится только с реестром!", vbExclamation & vbOKOnly, LARSver
        frmWriteAuditData.chkSQL.Enabled = False
    End If
frmwait.Hide
Screen.MousePointer = 0
End Function

