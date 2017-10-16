VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HardAndSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс для хранения данных по железу и некоторых данных по софту"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCPUName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCPUSocket As String 'local copy
Private mvarRAMType As String 'local copy
Private mvarRAMTotalSlots As String 'local copy
'local variable(s) to hold property value(s)
Private mvarRAMUsedSlots As String 'local copy
Private mvarRAMSlotStat As String 'local copy
Private mvarRAMValue As String 'local copy
Private mvarMBName As String 'local copy
Private mvarMBChipset As String 'local copy
Private mvarMBSouth As String 'local copy
Private mvarGPUName As String 'local copy
Private mvarMonitors As String 'local copy
Private mvarMonitorDims As String 'local copy
Private mvarHDD As String 'local copy
'local variable(s) to hold property value(s)
Private mvarWsName As String 'local copy
Private mvarHDDCount As String 'local copy
Private mvarHDDOverallSize As String 'local copy


Public Property Let HDDOverallSize(ByVal vData As String)
Attribute HDDOverallSize.VB_Description = "Всего места на ПК"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDDOverallSize = 5
    mvarHDDOverallSize = vData
End Property


Public Property Get HDDOverallSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.HDDOverallSize
    HDDOverallSize = mvarHDDOverallSize
End Property



Public Property Let HDDCount(ByVal vData As String)
Attribute HDDCount.VB_Description = "Кол-во ЖД"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDDCount = 5
    mvarHDDCount = vData
End Property


Public Property Get HDDCount() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.HDDCount
    HDDCount = mvarHDDCount
End Property



Public Property Let WsName(ByVal vData As String)
Attribute WsName.VB_Description = "Сетевое имя ПК"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WsName = 5
    mvarWsName = vData
End Property


Public Property Get WsName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.WsName
    WsName = HostName
End Property



Public Function Reset()
With Me
    .CPUName = ""
    .CPUSocket = ""
    .GPUName = ""
    .HDD = ""
    .HDDCount = ""
    .HDDOverallSize = ""
    .MBChipset = ""
    .MBName = ""
    .Monitors = ""
    .RAMSlotStat = ""
    .RAMTotalSlots = ""
    .RAMType = ""
    .RAMUsedSlots = ""
    .RAMValue = ""
    .Monitors = ""
End With
End Function
Public Function GetSummary() As String
Dim SummaryMessage As String
With Me
    SummaryMessage = "Имя ПК: " & .WsName
    SummaryMessage = SummaryMessage & vbCrLf & "Процессор: " & .CPUName
    SummaryMessage = SummaryMessage & vbCrLf & "Сокет: " & .CPUSocket
    SummaryMessage = SummaryMessage & vbCrLf & "Видеокарта: " & .GPUName
    SummaryMessage = SummaryMessage & vbCrLf & "Жесткие диски: " & .HDD
    SummaryMessage = SummaryMessage & vbCrLf & "Всего ЖД: " & .HDDCount
    SummaryMessage = SummaryMessage & vbCrLf & "Общий размер ЖД: " & .HDDOverallSize
    SummaryMessage = SummaryMessage & vbCrLf & "Чипсет: " & .MBChipset
    SummaryMessage = SummaryMessage & vbCrLf & "Материнская плата: " & .MBName
    SummaryMessage = SummaryMessage & vbCrLf & "Мониторы: " & .Monitors
    SummaryMessage = SummaryMessage & vbCrLf & "Оперативная память: " & .RAMSlotStat
    SummaryMessage = SummaryMessage & vbCrLf & "Всего слотов ОЗУ: " & .RAMTotalSlots
    SummaryMessage = SummaryMessage & vbCrLf & "Тип ОЗУ: " & .RAMType
    SummaryMessage = SummaryMessage & vbCrLf & "Занято слотов ОЗУ: " & .RAMUsedSlots
    SummaryMessage = SummaryMessage & vbCrLf & "Общий размер ОЗУ: " & .RAMValue
End With
GetSummary = SummaryMessage
End Function
Public Function WriteToSQL() As Boolean
End Function

Public Property Let HDD(ByVal vData As String)
Attribute HDD.VB_Description = "Информация о ЖД"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDD = 5
    mvarHDD = vData
End Property


Public Property Get HDD() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.HDD
    HDD = mvarHDD
End Property

Public Property Let Monitors(ByVal vData As String)
Attribute Monitors.VB_Description = "Информация о мониторе"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Monitors = 5
    mvarMonitors = vData
End Property


Public Property Get Monitors() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.Monitors
    Monitors = mvarMonitors
End Property

Public Property Let GPUName(ByVal vData As String)
Attribute GPUName.VB_Description = "Информация о графическом процессе"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GPUName = 5
    mvarGPUName = vData
End Property


Public Property Get GPUName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.GPUName
    GPUName = mvarGPUName
End Property

Public Property Let MBChipset(ByVal vData As String)
Attribute MBChipset.VB_Description = "Информация о чипсете"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBChipset = 5
    mvarMBChipset = vData
End Property


Public Property Get MBChipset() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.MBChipset
    MBChipset = mvarMBChipset
End Property



Public Property Let MBName(ByVal vData As String)
Attribute MBName.VB_Description = "Инф-ция о материнской плате"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBName = 5
    mvarMBName = vData
End Property


Public Property Get MBName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.MBName
    MBName = mvarMBName
End Property



Public Property Let RAMValue(ByVal vData As String)
Attribute RAMValue.VB_Description = "Кол-во ОЗУ"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMValue = 5
    mvarRAMValue = vData
End Property


Public Property Get RAMValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.RAMValue
    RAMValue = mvarRAMValue
End Property



Public Property Let RAMSlotStat(ByVal vData As String)
Attribute RAMSlotStat.VB_Description = "Статистика по используемым слотам ОЗУ"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMSlotStat = 5
    mvarRAMSlotStat = vData
End Property


Public Property Get RAMSlotStat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.RAMSlotStat
    RAMSlotStat = mvarRAMSlotStat
End Property



Public Property Let RAMUsedSlots(ByVal vData As String)
Attribute RAMUsedSlots.VB_Description = "Кол-во использованных слотов ОЗУ"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMUsedSlots = 5
    mvarRAMUsedSlots = vData
End Property


Public Property Get RAMUsedSlots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.RAMUsedSlots
    RAMUsedSlots = mvarRAMUsedSlots
End Property



Public Property Let RAMTotalSlots(ByVal vData As String)
Attribute RAMTotalSlots.VB_Description = "Кол-во слотов под ОЗУ (всего)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMTotalSlots = 5
    mvarRAMTotalSlots = vData
End Property


Public Property Get RAMTotalSlots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.RAMTotalSlots
    RAMTotalSlots = mvarRAMTotalSlots
End Property



Public Property Let RAMType(ByVal vData As String)
Attribute RAMType.VB_Description = "Тип ОЗУ (DDR)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMType = 5
    mvarRAMType = vData
End Property


Public Property Get RAMType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.RAMType
    RAMType = mvarRAMType
End Property



Public Property Let CPUSocket(ByVal vData As String)
Attribute CPUSocket.VB_Description = "Сокет ЦП"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPUSocket = 5
    mvarCPUSocket = vData
End Property


Public Property Get CPUSocket() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.CPUSocket
    CPUSocket = mvarCPUSocket
End Property



Public Property Let CPUName(ByVal vData As String)
Attribute CPUName.VB_Description = "Процессор"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPUName = 5
    mvarCPUName = vData
End Property


Public Property Get CPUName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: debug.printX.CPUName
    CPUName = mvarCPUName
End Property



