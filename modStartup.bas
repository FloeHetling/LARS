Attribute VB_Name = "modStartup"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' ��������� ������ "�����" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'���������� ����������
    Public LARSver As String
    Public InfoBoxes() As String
    Public HostName As String
    Public enumSQLFields As Integer '���� ����� � ������ SQLAuditData
      
'���������� ���������
''����������
    Public Const Lime = 12648384
    Public Const Sand = 12648447
    Public Const Red = 12632319

'���������� �������
''���������� ����� AuditData, �������� ��� thisPC
    Public thisPC As New auditdata
    Public thisPCSQL As New SQLAuditData

'������������ ���������� ���������� ������
''���������� � ������� ��������� ����� ��
    Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
Dim CLIArg As String

Sub Main()
'���������� � ���������� ���������� �������� � ������ ��
LARSver = App.ProductName & ", ������ " & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.CompanyName
'���������, ������� �� ������ ���������
'���� �� - ��������� ����� ������
    If App.PrevInstance = True Then
        Exit Sub
        End
    End If

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

'���������� ��������� ���������� ������ � ���������� � ������ ��
'CLIArg = Command$
CLIArg = "/edit"
    Select Case CLIArg
        
        Case "/edit"
        frmWriteAuditData.Show
        
        Case Else
        Call PopulateAuditData
                
    End Select

'�������� � ���������� ���������� ������� ��� ��
Dim dwLen As Long
    '������� �����
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    HostName = String(dwLen, "X")
    '�������� ��� ��
    GetComputerName HostName, dwLen
    '������� ������ (�������) �������
    HostName = Left(HostName, dwLen)
End Sub
