VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AuditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "�����, ���������� ������ ������ �� �������� ��"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' ����� �������� �����-������ '''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
Option Explicit '�������� "�������������" ����������� - ����������� ���������� ����
'��������� ���������� ��� ���������� �������� ������ ������
Private mvarCompany As String
Private mvarWsName As String
Private mvarWSSerial As String
Private mvarWindowsVersion As String
Private mvarWindowsLicenseModel As String
Private mvarWindowsOLPSerial As String
Private mvarOfficeVersion As String
Private mvarOfficeLicenseModel As String
Private mvarOfficeOLPSerial As String
'local variable(s) to hold property value(s)

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




Public Sub RegSave()
'����� ��� ���������� ������ � ������ �� �����-�������
Dim ibName As Variant
Dim ibValue As String

    '����������, �������� �����-������ ������� ��, ������� � ���
    '�������������� ����� �� �����. ������������� �� ���-�� ����������
    '� ������� InfoBoxes.
    '��� ������ ���������� �� ������� ����������� ����������� � ������ AuditData
    '(����� �� ���������� ��� � ����) ��������������� ��������.
    For Each ibName In InfoBoxes
        ibValue = CallByName(Me, ibName, VbGet)
        '� �������� ���� ���������� - �������� � ��������� RegPutAuditData
        '������ modMAD
        Call RegPutAuditData(ibName, ibValue)
    Next
End Sub

Public Sub RegLoad()
'����� ��� ������ ������ � �����-������ �� ������� Windows
Dim ibName As Variant
Dim ibValue As String

    For Each ibName In InfoBoxes
        ibValue = RegGetAuditData(ibName)
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
    mvarWsName = HostName
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
