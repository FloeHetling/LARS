VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AliasLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public Enum laWriteMode
    laWriteToRegistry
    laWriteToSQL
    laWriteEverywhere
End Enum

Public Property Let OfficeLicenseModel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeLicenseModel = 5
    mvarOfficeLicenseModel = vData
End Property

Public Property Get OfficeLicenseModel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.OfficeLicenseModel
    OfficeLicenseModel = mvarOfficeLicenseModel
End Property

Public Property Let OfficeVersion(ByVal vData As String)
Attribute OfficeVersion.VB_Description = "Версия и редакция Microsoft Office"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeVersion = 5
    mvarOfficeVersion = vData
End Property


Public Property Get OfficeVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.OfficeVersion
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.WindowsOLPSerial
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.WindowsLicenseModel
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.WindowsVersion
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.WSSerial
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.WSName
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
'Syntax: frmwriteauditdata.tdeb.Text = frmwriteauditdata.tdeb.Text & vbcrlf & X.Company
    Company = mvarCompany
End Property

Private Sub Class_Initialize()
Company = "[Юрлицо]"
OfficeLicenseModel = "[Office License Model]"
OfficeVersion = "[Версия Офис]"
WindowsLicenseModel = "[Windows License Model]"
WindowsOLPSerial = "[Номер лицензии Windows OLP]"
WindowsVersion = "[Версия Windows]"
WSName = "[Сетевое имя]"
WSSerial = "[Серийный номер]"
End Sub
