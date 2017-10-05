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
'local variable(s) to hold property value(s)
Private mvarCompany As String 'local copy
Private mvarWSName As String 'local copy
Private mvarWSSerial As String 'local copy
Private mvarWindowsVersion As String 'local copy
Private mvarWindowsLicenseModel As String 'local copy
Private mvarWindowsOLPSerial As String 'local copy
Private mvarOfficeVersion As String 'local copy
Private mvarOfficeLicenseModel As Variant 'local copy
Public Sub SQLSave()
MsgBox "TODO: ����������� �������� ������ � SQL. �����?"
MsgBox "������ ������! ������ ���!"
End Sub

Public Sub SQLLoad(ByVal CurrentHostname As String)
Dim ibName As Variant
Dim ibValue As String

    For Each ibName In InfoBoxes
        ibValue = GetSQLData(ibName, CurrentHostname)
        CallByName Me, ibName, VbLet, ibValue
    Next
End Sub

Public Property Let OfficeLicenseModel(ByVal vData As Variant)
Attribute OfficeLicenseModel.VB_Description = "������ �������������� Microsoft Office"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeLicenseModel = 5
    mvarOfficeLicenseModel = vData
End Property


Public Property Get OfficeLicenseModel() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OfficeLicenseModel
    If IsObject(mvarOfficeLicenseModel) Then
        Set OfficeLicenseModel = mvarOfficeLicenseModel
    Else
        OfficeLicenseModel = mvarOfficeLicenseModel
    End If
End Property



Public Property Let OfficeVersion(ByVal vData As String)
Attribute OfficeVersion.VB_Description = "������ � �������� Microsoft Office"
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
Attribute WindowsOLPSerial.VB_Description = "�������� ����� OLP ��"
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
Attribute WindowsLicenseModel.VB_Description = "������ �������������� ��"
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
Attribute WindowsVersion.VB_Description = "������ � �������� Windows"
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
Attribute WSSerial.VB_Description = "����� �� � ��������"
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
Attribute WSName.VB_Description = "������� ��� ��"
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
Attribute Company.VB_Description = "�������� ��������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Company = 5
    mvarCompany = vData
End Property


Public Property Get Company() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Company
    Company = mvarCompany
End Property


Private Function GetSQLData(ByVal AuditParameter As String, ByVal CurrentHostname As String)
On Error GoTo GetSqlData_error

Dim SQL As New ADODB.Connection 'ADODB Connection Object
    Dim SQLData As New ADODB.Recordset 'Recordset Object
    Dim SQLRequest As String ' String variable to store sql command
    
    SQL.Open _
        "Provider = SQLNCLI11.1;" & _
        "Data Source=WS0006\SQLEXPRESS;" & _
        "Initial Catalog=AIDA;" & _
        "User ID=sa;" & _
        "Password=happyness;"
          
    SQLRequest = "select " & _
                "[������] as Company," & _
                "[Office License Model] as OfficeLicenseModel," & _
                "[������ ����] as OfficeVersion," & _
                "[Windows License Model] as WindowsLicenseModel," & _
                "[����� �������� Windows OLP] as WindowsOLPSerial," & _
                "[������ Windows] as WindowsVersion," & _
                "[������� ���] as WSName," & _
                "[�������� �����] as WSSerial " & _
                "from aida.dbo.larspc where [������� ���] = '" & CurrentHostname & "';"
    
    SQLData.Open SQLRequest, SQL, adOpenKeyset
    GetSQLData = SQLData.Fields(AuditParameter).Value
    SQL.Close
    Exit Function
    
GetSqlData_error:
GetSQLData = "sql_err_nodata"
Dim i
For i = 0 To SQLData.Fields.Count - 1
Debug.Print SQLData.Fields(i).Name & " " & SQLData.Fields(i).Type
Next
End Function
