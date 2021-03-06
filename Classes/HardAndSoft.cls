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
Attribute VB_Description = "����� ��� �������� ������ �� ������ � ��������� ������ �� �����"
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
Attribute HDDOverallSize.VB_Description = "����� ����� �� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDDOverallSize = 5
    mvarHDDOverallSize = vData
End Property


Public Property Get HDDOverallSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.HDDOverallSize
    HDDOverallSize = mvarHDDOverallSize
End Property



Public Property Let HDDCount(ByVal vData As String)
Attribute HDDCount.VB_Description = "���-�� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDDCount = 5
    mvarHDDCount = vData
End Property


Public Property Get HDDCount() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.HDDCount
    HDDCount = mvarHDDCount
End Property



Public Property Let WsName(ByVal vData As String)
Attribute WsName.VB_Description = "������� ��� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WsName = 5
    mvarWsName = vData
End Property


Public Property Get WsName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.WsName
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
    SummaryMessage = "��� ��: " & .WsName
    SummaryMessage = SummaryMessage & vbCrLf & "���������: " & .CPUName
    SummaryMessage = SummaryMessage & vbCrLf & "�����: " & .CPUSocket
    SummaryMessage = SummaryMessage & vbCrLf & "����������: " & .GPUName
    SummaryMessage = SummaryMessage & vbCrLf & "������� �����: " & .HDD
    SummaryMessage = SummaryMessage & vbCrLf & "����� ��: " & .HDDCount
    SummaryMessage = SummaryMessage & vbCrLf & "����� ������ ��: " & .HDDOverallSize
    SummaryMessage = SummaryMessage & vbCrLf & "������: " & .MBChipset
    SummaryMessage = SummaryMessage & vbCrLf & "����������� �����: " & .MBName
    SummaryMessage = SummaryMessage & vbCrLf & "��������: " & .Monitors
    SummaryMessage = SummaryMessage & vbCrLf & "����������� ������: " & .RAMSlotStat
    SummaryMessage = SummaryMessage & vbCrLf & "����� ������ ���: " & .RAMTotalSlots
    SummaryMessage = SummaryMessage & vbCrLf & "��� ���: " & .RAMType
    SummaryMessage = SummaryMessage & vbCrLf & "������ ������ ���: " & .RAMUsedSlots
    SummaryMessage = SummaryMessage & vbCrLf & "����� ������ ���: " & .RAMValue
End With
GetSummary = SummaryMessage
End Function
Public Function WriteToSQL() As Boolean
End Function

Public Property Let HDD(ByVal vData As String)
Attribute HDD.VB_Description = "���������� � ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HDD = 5
    mvarHDD = vData
End Property


Public Property Get HDD() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.HDD
    HDD = mvarHDD
End Property

Public Property Let Monitors(ByVal vData As String)
Attribute Monitors.VB_Description = "���������� � ��������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Monitors = 5
    mvarMonitors = vData
End Property


Public Property Get Monitors() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.Monitors
    Monitors = mvarMonitors
End Property

Public Property Let GPUName(ByVal vData As String)
Attribute GPUName.VB_Description = "���������� � ����������� ��������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GPUName = 5
    mvarGPUName = vData
End Property


Public Property Get GPUName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.GPUName
    GPUName = mvarGPUName
End Property

Public Property Let MBChipset(ByVal vData As String)
Attribute MBChipset.VB_Description = "���������� � �������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBChipset = 5
    mvarMBChipset = vData
End Property


Public Property Get MBChipset() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.MBChipset
    MBChipset = mvarMBChipset
End Property



Public Property Let MBName(ByVal vData As String)
Attribute MBName.VB_Description = "���-��� � ����������� �����"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MBName = 5
    mvarMBName = vData
End Property


Public Property Get MBName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.MBName
    MBName = mvarMBName
End Property



Public Property Let RAMValue(ByVal vData As String)
Attribute RAMValue.VB_Description = "���-�� ���"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMValue = 5
    mvarRAMValue = vData
End Property


Public Property Get RAMValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.RAMValue
    RAMValue = mvarRAMValue
End Property



Public Property Let RAMSlotStat(ByVal vData As String)
Attribute RAMSlotStat.VB_Description = "���������� �� ������������ ������ ���"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMSlotStat = 5
    mvarRAMSlotStat = vData
End Property


Public Property Get RAMSlotStat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.RAMSlotStat
    RAMSlotStat = mvarRAMSlotStat
End Property



Public Property Let RAMUsedSlots(ByVal vData As String)
Attribute RAMUsedSlots.VB_Description = "���-�� �������������� ������ ���"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMUsedSlots = 5
    mvarRAMUsedSlots = vData
End Property


Public Property Get RAMUsedSlots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.RAMUsedSlots
    RAMUsedSlots = mvarRAMUsedSlots
End Property



Public Property Let RAMTotalSlots(ByVal vData As String)
Attribute RAMTotalSlots.VB_Description = "���-�� ������ ��� ��� (�����)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMTotalSlots = 5
    mvarRAMTotalSlots = vData
End Property


Public Property Get RAMTotalSlots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.RAMTotalSlots
    RAMTotalSlots = mvarRAMTotalSlots
End Property



Public Property Let RAMType(ByVal vData As String)
Attribute RAMType.VB_Description = "��� ��� (DDR)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAMType = 5
    mvarRAMType = vData
End Property


Public Property Get RAMType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.RAMType
    RAMType = mvarRAMType
End Property



Public Property Let CPUSocket(ByVal vData As String)
Attribute CPUSocket.VB_Description = "����� ��"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPUSocket = 5
    mvarCPUSocket = vData
End Property


Public Property Get CPUSocket() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.CPUSocket
    CPUSocket = mvarCPUSocket
End Property



Public Property Let CPUName(ByVal vData As String)
Attribute CPUName.VB_Description = "���������"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPUName = 5
    mvarCPUName = vData
End Property


Public Property Get CPUName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: WriteToLogX.CPUName
    CPUName = mvarCPUName
End Property



