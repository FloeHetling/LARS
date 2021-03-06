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
Private mvarWsName As String 'local copy
Private mvarWSSerial As String 'local copy
Private mvarWindowsVersion As String 'local copy
Private mvarWindowsLicenseModel As String 'local copy
Private mvarWindowsOLPSerial As String 'local copy
Private mvarOfficeVersion As String 'local copy
Private mvarOfficeLicenseModel As String 'local copy
Private mvarOfficeOLPSerial As String
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
'Syntax: WriteToLogX.OfficeLicenseModel
    OfficeLicenseModel = mvarOfficeLicenseModel
End Property

Public Sub SQLSave(ByVal CurrentHostname As String)
'����� ��� ���������� ������ � SQL �� �����-�������
Dim ibName As Variant
Dim ibValue As String

    '����������, �������� �����-������ ������� ��, ������� � ���
    '�������������� ����� �� �����. ������������� �� ���-�� ����������
    '� ������� InfoBoxes.
    '��� ������ ���������� �� ������� ����������� ����������� � ������ AuditData
    '(����� �� ���������� ��� � ����) ��������������� ��������.
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
Attribute OfficeVersion.VB_Description = "������ � �������� Microsoft Office"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeVersion = 5
    mvarOfficeVersion = vData
End Property


Public Property Get OfficeVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.OfficeVersion
    OfficeVersion = GetOfficeVersion
End Property



Public Property Let WindowsOLPSerial(ByVal vData As String)
Attribute WindowsOLPSerial.VB_Description = "�������� ����� OLP ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsOLPSerial = 5
    mvarWindowsOLPSerial = vData
End Property


Public Property Get WindowsOLPSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WindowsOLPSerial
    WindowsOLPSerial = mvarWindowsOLPSerial
End Property

Public Property Let OfficeOLPSerial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeOLPSerial = 5
    mvarOfficeOLPSerial = vData
End Property


Public Property Get OfficeOLPSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.OfficeOLPSerial
    OfficeOLPSerial = mvarOfficeOLPSerial
End Property

Public Property Let WindowsLicenseModel(ByVal vData As String)
Attribute WindowsLicenseModel.VB_Description = "������ �������������� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsLicenseModel = 5
    mvarWindowsLicenseModel = vData
End Property


Public Property Get WindowsLicenseModel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WindowsLicenseModel
    WindowsLicenseModel = mvarWindowsLicenseModel
End Property



Public Property Let WindowsVersion(ByVal vData As String)
Attribute WindowsVersion.VB_Description = "������ � �������� Windows"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsVersion = 5
    mvarWindowsVersion = vData
End Property


Public Property Get WindowsVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WindowsVersion
    WindowsVersion = GetWindowsVersion
End Property



Public Property Let WSSerial(ByVal vData As String)
Attribute WSSerial.VB_Description = "����� �� � ��������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSSerial = 5
    mvarWSSerial = vData
End Property


Public Property Get WSSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WSSerial
    WSSerial = mvarWSSerial
End Property



Public Property Let WsName(ByVal vData As String)
Attribute WsName.VB_Description = "������� ��� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSName = 5
    mvarWsName = vData
End Property


Public Property Get WsName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WSName
    WsName = HostName
End Property



Public Property Let Company(ByVal vData As String)
Attribute Company.VB_Description = "�������� ��������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Company = 5
    mvarCompany = vData
End Property


Public Property Get Company() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.Company
    Company = mvarCompany
End Property


Private Function GetSQLData(ByVal AuditParameter As String, ByVal CurrentHostname As String) As String
On Error GoTo GetSqlData_error
Dim SQLResponse As Variant
Dim SQL As New ADODB.Connection 'ADODB Connection Object
    Dim SQLData As New ADODB.Recordset 'Recordset Object
    Dim SQLRequest As String, SQLAPRequest As String ' String variable to store sql command
    
    SQL.Open SQLConnString
        
'' ��� ��� ������ ����������
'    SQLRequest = "select " & _
'                "[������] as Company," & _
'                "[Office License Model] as OfficeLicenseModel," & _
'                "[������ ����] as OfficeVersion," & _
'                "[Windows License Model] as WindowsLicenseModel," & _
'                "[����� �������� Windows OLP] as WindowsOLPSerial," & _
'                "[������ Windows] as WindowsVersion," & _
'                "[������� ���] as WSName," & _
'                "[�������� �����] as WSSerial " & _
'                "from aida.dbo.larspc where [������� ���] = '" & CurrentHostname & "';"
'' ���� ��������� � ��������� ���������

SQLAPRequest = CallByName(Ru, AuditParameter, VbGet)
SQLRequest = "SELECT " & SQLAPRequest & " AS " & AuditParameter & " FROM lars.dbo.larspc WHERE [������� ���] = '" & CurrentHostname & "';"
WriteToLog SQLRequest
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
SQLRequest = "SELECT [������� ���] AS WSName FROM lars.dbo.larspc WHERE [������� ���] = '" & CurrentHostname & "';"
    If (SQLExecute(SQLRequest, laRX, "WSName") = 3265) Or (SQLExecute(SQLRequest, laRX, "WSName") = 3021) Then
        SQLRequest = "INSERT INTO dbo.larspc ([������� ���]) VALUES ('" & CurrentHostname & "');"
        SQLExecute SQLRequest, laTX
    Else
        SQLRequest = "UPDATE lars.dbo.larspc SET " & SQLAPRequest & " = '" & AuditValue & "' WHERE [������� ���] = '" & CurrentHostname & "';"
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
'    WriteToLog"�������� ������� SQLAuditData SQLExecute. ������ ����������:" & vbCrLf & SQLRequestString
'    SQLData.Open SQLRequestString, SQL, adOpenKeyset
'        If SQLMode = laRX Then
'            SQLExecute = SQLData.Fields(ParameterToRead).Value
'        End If
'    SQL.Close
'    Exit Function
'
'SQL_error:
'WriteToLog"������ SQL " & Err.Number & ":" & vbCrLf & Err.Description
'SQLExecute = Err.Number
'End Function


'Public Function CheckSQLAvailability() As Boolean
'Screen.MousePointer = 11
'frmwait.Show
'frmwait.Reason.Caption = "�������� SQL"
'Dim sqlCheckIfAvailable As Long
'sqlCheckIfAvailable = SQLExecute("SELECT * FROM dbo.larspc", laRX)
'    If sqlCheckIfAvailable = -2147467259 Then
'        frmWriteAuditData.chkSQL.value = 0
'        frmwait.Hide
'        MsgBox "������ SQL �� ��������." & vbCrLf _
'            & " ������ ���������� ������ � ��������!", vbExclamation & vbOKOnly, LARSver
'        frmWriteAuditData.chkSQL.Enabled = False
'    End If
'frmwait.Hide
'Screen.MousePointer = 0
'End Function

