Attribute VB_Name = "modStartup"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' ��������� ������ "�����" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'���������� ����������
    Public LARSver As String
    Public InfoBoxes() As String
    Public SQLBoxes() As String
    Public HnSArgs() As Variant
    Public HostName As String
    Public enumSQLFields As Integer '���� ����� � ������ SQLAuditData
      
'���������� ���������
''����������
    Public Enum laColorConstants
        laLightGreen = 12648384
        laSand = 12648447
        laLightRed = 12632319
        laDarkGreen = 32768
        laDarkRed = 192
        laDarkBlue = 12936533
        laBlack = 0
    End Enum
    
'���� ������������
    Public Enum laHardware
        laCPU
        laRAM
        laMotherboard
        laChipset
        laSouthBridge
        laUSBHost
        laGPU
        laMonitor
        laHDD
    End Enum
    
'���������� �������
''���������� ����� AuditData, �������� ��� thisPC
    Public thisPC As New auditdata
    Public thisPCSQL As New SQLAuditData '�� �� �����, ������ ��� ��������� � SQL
    Public HnS As New HardAndSoft
    Public Ru As New AliasLibrary '���������� ������� ��� SQL ��������
    
'������������ ���������� ���������� ������
''���������� � ������� ��������� ����� ��
    Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer
    
'    Sub Form_Load()
'
'   If IsUserAnAdmin() = 0 Then
'     MsgBox "Not admin"
'   Else
'     MsgBox "Admin"
'   End If
'
'End Sub
    
'���������� ������ ����������� ��� SQL
    Public SQLConnString As String

'��������� ��������� ������
Public CLIArg As String
        
Sub Main()
CLIArg = Command$
HnSArgs = Array("WSNAME", "CPUNAME", "RAMTYPE", "RAMTOTALSLOTS", "RAMUSEDSLOTS", "RAMSLOTSTAT", "RAMVALUE", "MBNAME", "MBCHIPSET", "GPUNAME", "MONITORS", "HDD", "HDDCOUNT", "HDDOVERALLSIZE", "CPUSOCKET")
SQLConnString = "Provider = SQLOLEDB.1;" & _
        "Data Source=tcp:192.168.78.39,1433[oledb];" & _
        "Persist Security Info=False;" & _
        "Initial Catalog=AIDA;" & _
        "User ID=sa;" & _
        "Connect Timeout=2;" & _
        "Password=happyness;"
isSQLChecked = False
'���������� � ���������� ���������� �������� � ������ ��
LARSver = App.ProductName & ", ������ " & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.CompanyName
'���������, ������� �� ������ ���������
'���� �� - ��������� ����� ������
    If App.PrevInstance = True Then
        MsgBox "���� ��� ��������! ����������, ������� ���������." & vbCrLf & "��� ������� ������ �� ����� ��������� �����", vbExclamation, LARSver
        Exit Sub
        End
    End If

'�������� � ���������� ���������� ������� ��� ��
Dim dwLen As Long
    '������� �����
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    HostName = String(dwLen, "X")
    '�������� ��� ��
    GetComputerName HostName, dwLen
    '������� ������ (�������) �������
    HostName = Left(HostName, dwLen)


'������� ������ ��������� �� ����� �������� � ���������� �� � ��������� ������
''
'' ����� ������ ��� �����:
'' ��� ������ � ��������� 8 ����� �� 8 ���������� ������.
'' � ������, � �������, ����� ����� 20. ��� 3.
'' ������� �� �������������� �� ������� ��������� � ��������.
'' ������ ������� �� ��������� �����. ��������������, ����� ��������� ��� ������ �������
'' ���������� �������� ������� ����� � �������� ��������������� �������� ������� AuditData � SQLAuditData
''
Dim Ctrl As Control
Dim ibIndex As Integer
Dim ibName As String
ibIndex = 0

    For Each Ctrl In frmWriteAuditData.Controls         '�������, ��� ������� �������� �����
        If InStr(1, Ctrl.Tag, "infobox") <> 0 Then      '������� ����� ����� infobox � �������� Tag
            ReDim Preserve InfoBoxes(ibIndex)           '�� ��������� ������ Infoboxes ����� ���������
            Dim InfoboxTag() As String
            InfoboxTag = Split(Ctrl.Tag, ",")
            ibName = InfoboxTag(1)                      '������� ������� � �������� Tag ����� ����� "infobox,"
            InfoBoxes(ibIndex) = ibName                 '���������� ���� ��� � ���� ������ �������� ������� � �������� �� �������
            ibIndex = ibIndex + 1                       '��������, ����� ��������� ������� ����� � ���������/��������� ���
        End If
    Next                                                '�� ������ � ��� ���� ������ ����� ��� ���� ����� - ������ InfoBoxes,
                                                        '�� �������� �� � ����� ���������� � ���������� � ��������
Dim SQLibName As String
ibIndex = 0
    For Each Ctrl In frmWriteAuditData.Controls
        If InStr(1, Ctrl.Tag, "SQLbox") <> 0 Then
            ReDim Preserve SQLBoxes(ibIndex)
            Dim SQLBoxTag() As String
            SQLBoxTag = Split(Ctrl.Tag, ",")
            SQLibName = SQLBoxTag(1)
            SQLBoxes(ibIndex) = SQLibName
            ibIndex = ibIndex + 1
        End If
    Next
    
'���������� ��������� ���������� ������ � ���������� � ������ ��
CLIArg = Command$
    Select Case CLIArg
        
        Case "/edit"
            If IsUserAnAdmin() = 1 Then
                frmWriteAuditData.Show
            Else
                MsgBox "��������� �� � ������� ��������������!", vbExclamation, LARSver
                End
            End If
        Case "/wmi"
        frmWMIQL.Show
        Case Else
            If IsUserAnAdmin() = 1 Then
                Call PopulateAuditData
                Exit Sub
            Else
                MsgBox "��������� �� � ������� ��������������!", vbExclamation, LARSver
                End
            End If
    End Select
End Sub
