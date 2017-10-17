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
Attribute VB_Description = "Класс, содержащий данные аудита по текущему ПК"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' КЛАСС ХРАНЕНИЯ АУДИТ-ДАННЫХ '''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода
'Локальные переменные для временного хранения данных модуля
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
'Syntax: debug.printX.OfficeOLPSerial
    OfficeOLPSerial = mvarOfficeOLPSerial
End Property

Public Property Let OfficeLicenseModel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OfficeLicenseModel = 5
    mvarOfficeLicenseModel = vData
End Property


Public Property Get OfficeLicenseModel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.OfficeLicenseModel
    OfficeLicenseModel = mvarOfficeLicenseModel
End Property




Public Sub RegSave()
'Метод для сохранения данных в реестр из класс-объекта
Dim ibName As Variant
Dim ibValue As String

    'Переменных, хранящих аудит-данные столько же, сколько у нас
    'информационных полей на форме. Отталкиваемся от кол-ва переменных
    'в массиве InfoBoxes.
    'Для каждой переменной из массива ДИНАМИЧЕСКИ запрашиваем с класса AuditData
    '(здесь он обращается сам к себе) соответствующее значение.
    For Each ibName In InfoBoxes
        ibValue = CallByName(Me, ibName, VbGet)
        'и передаем пары Переменная - Значение в процедуру RegPutAuditData
        'модуля modMAD
        Call RegPutAuditData(ibName, ibValue)
    Next
End Sub

Public Sub RegLoad()
'Метод для чтения данных в класс-объект из реестра Windows
Dim ibName As Variant
Dim ibValue As String

    For Each ibName In InfoBoxes
        ibValue = RegGetAuditData(ibName)
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
'Syntax: debug.printX.OfficeVersion
    OfficeVersion = GetOfficeVersion
End Property



Public Property Let WindowsOLPSerial(ByVal vData As String)
Attribute WindowsOLPSerial.VB_Description = "Серийный номер OLP ОС"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowsOLPSerial = 5
    mvarWindowsOLPSerial = vData
End Property


Public Property Get WindowsOLPSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.WindowsOLPSerial
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
'Syntax: debug.printX.WindowsLicenseModel
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
'Syntax: debug.printX.WindowsVersion
    WindowsVersion = GetWindowsVersion
End Property



Public Property Let WSSerial(ByVal vData As String)
Attribute WSSerial.VB_Description = "Номер ПК с наклейки"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSSerial = 5
    mvarWSSerial = vData
End Property


Public Property Get WSSerial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.WSSerial
    WSSerial = mvarWSSerial
End Property



Public Property Let WsName(ByVal vData As String)
Attribute WsName.VB_Description = "Сетевое имя ПК"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WSName = 5
    mvarWsName = HostName
End Property


Public Property Get WsName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.WSName
    WsName = HostName
End Property



Public Property Let Company(ByVal vData As String)
Attribute Company.VB_Description = "Название компании"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Company = 5
    mvarCompany = vData
End Property


Public Property Get Company() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.Company
    Company = mvarCompany
End Property
