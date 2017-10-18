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
    Public HostName As String, LARSINIPath As String
    Public enumSQLFields As Integer '���� ����� � ������ SQLAuditData
    Public SilentRun As Boolean, isAllSettingsProvided As Boolean
    
'������������ ����������, �������� �� INI
    Public INIParameters As New Collection

'�����
''��������� Winsock
    Public Enum WinsockControlState
        MAIL_CONNECT
        MAIL_HELO
        MAIL_FROM
        MAIL_RCPTTO
        MAIL_DATA
        MAIL_HEADER
        MAIL_DOT
        MAIL_QUIT
    End Enum

    Public WinsockState As WinsockControlState, SMTPServer As String, SMTPPort As String
    
''���������� ���������� �����
    Public FromEmail As String, _
            ToEmail As String, _
            EmailSubject As String, _
            MailMessage As String, _
            EmailServer As String, _
            EmailServerPort As String
      
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

Public Function isSettingsIntegrityOK() As Boolean
Dim SettingsArray() As String, ParamsArray() As String, SIndex As Integer, SErr As Integer, Setting As Variant

'�����������, ����� ��������� �� INI �� ���������
    With INIParameters
                .Add "DataSource"
                .Add "SMTPServer"
                .Add "SMTPPort"
                .Add "FromEmail"
                .Add "ToEmail"
                .Add "EmailServer"
                .Add "EmailServerPort"
    End With
    
'���������, ���������� �� ���� � �����
'���� �� ���������� - ����� ������ False � ������� �� ���������
    If CheckPath(LARSINIPath) <> True Then
            '''''������� ��������� �����
            Dim iFileNo As Integer
            iFileNo = FreeFile
        
            Open LARSINIPath For Output As #iFileNo
            Print #iFileNo, ";Only Windows-1251 Codepage is allowed!"
            Print #iFileNo, ";���� �� ������ �������� ��� ������, ���� ��������� ����������� ���������"
            Print #iFileNo, ""
            Close #iFileNo
            '''''' � ������� ������.
        isSettingsIntegrityOK = False
        Exit Function
    End If
     
'���� �� ��������� �� ����� - ��������� ������ �������� �� ���������
'��� ���� ������� �������
    SErr = 0
    For Each Setting In INIParameters
        If INIQuery("MAIN", Setting) = "" Then SErr = SErr + 1
    Next Setting

'���� �� �������� �������� ���� ���� ���-������ - ����������� �������� ���� ��������.
    If SErr <> 0 Then isSettingsIntegrityOK = False Else isSettingsIntegrityOK = True

End Function

Public Function INIQuery(ByVal Div As String, ByVal Param As String) As String
Dim INIReadResult As String
Call fReadValue(LARSINIPath, Div, Param, "S", "", INIReadResult)
INIQuery = INIReadResult
End Function

Sub Main()
CLIArg = Command$

''���������� � ���������� ���������� �������� � ������ ��
    LARSver = App.ProductName & " " & _
                App.Major & "." & App.Minor & _
                "." & App.Revision & " - " & _
                App.CompanyName

''�������� � ���������� ���������� ���� �� ����� ��������
LARSINIPath = Replace(App.Path & "\lars.ini", "\\", "\")

''�������� � ���������� ���������� ������� ��� ��
    Dim dwLen As Long
        '������� �����
        dwLen = MAX_COMPUTERNAME_LENGTH + 1
        HostName = String(dwLen, "X")
        '�������� ��� ��
        GetComputerName HostName, dwLen
        '������� ������ (�������) �������
        HostName = Left(HostName, dwLen)
        
isAllSettingsProvided = False

'''''''''''''''''''''''������ � ������ INI'''''''''''''''''''''''
    '��������� ���� Attended ����� - ���� ��������� �������� ���� - ���������� ��������, ����� - ������ ������ �������� � �� � ����������
    If CLIArg <> "" Then
        If isSettingsIntegrityOK = False Then
            MsgBox "�� ������ ���� � ����������� ��" & vbCrLf & _
            "���� �� ��� ��������� ������� ���������." & vbCrLf & vbCrLf & _
            "����������, ��������� ������������� ���������!", vbExclamation, LARSver
            frmSettings.Show vbModal
            If isAllSettingsProvided = False Then End
        Else
            isAllSettingsProvided = True
        End If
    Else
        If isSettingsIntegrityOK = False Then End
    End If
    
    '��������� ���� UnAttended - ���� ��������� �������� ������� - ������ ������ �������� � ��
    If CLIArg = "" Then
        If isSettingsIntegrityOK = True Then isAllSettingsProvided = True
    End If
If isAllSettingsProvided = False Then End '���� � ����� ����� �� ��� ��������� - ����� �� ����������!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
''������ ��������� �����:
        SMTPServer = INIQuery("MAIN", "SMTPServer")
        SMTPPort = INIQuery("MAIN", "SMTPPort")
        FromEmail = INIQuery("MAIN", "FromEmail")
        ToEmail = INIQuery("MAIN", "ToEmail")
        EmailServer = INIQuery("MAIN", "EmailServer")
        EmailServerPort = INIQuery("MAIN", "EmailServerPort")
        EmailSubject = "����: ����� �� ������ ������� ������� """ & HostName & """"
        SendFormCallOnly = False

''������ ��������� ��������� ������
    CLIArg = Command$
    
''��������� ������ ���������� �� ������
    HnSArgs = Array("WSNAME", "CPUNAME", "RAMTYPE", _
                "RAMTOTALSLOTS", "RAMUSEDSLOTS", "RAMSLOTSTAT", _
                "RAMVALUE", "MBNAME", "MBCHIPSET", "GPUNAME", _
                "MONITORS", "HDD", "HDDCOUNT", _
                "HDDOVERALLSIZE", "CPUSOCKET")
     
''������ ���������� ��������� ����������� � SQL
    SQLConnString = "Provider = SQLOLEDB.1;" & _
                "Data Source=" & INIQuery("MAIN", "DataSource") & "" & _
                "Persist Security Info=False;" & _
                "Initial Catalog=AIDA;" & _
                "User ID=sa;" & _
                "Connect Timeout=2;" & _
                "Password=happyness;"
    isSQLChecked = False

''���������, ������� �� ������ ���������
'���� �� - ��������� ����� ������
    If App.PrevInstance = True Then
        MsgBox "���� ��� ��������! ����������, ������� ���������." & vbCrLf & "��� ������� ������ �� ����� ��������� �����", vbExclamation, LARSver
        Exit Sub
        End
    End If

''�� ��������� ������ ������ ���
    SilentRun = False

'������� ������ ��������� �� ����� �������� � ���������� �� � ��������� ������
''
'' ����� ������ ��� �����:
'' ��� ������ � ��������� 8 ����� �� 8 ���������� ������. EDIT: � ������ �� 21. �� � ��� ��� ������������ ������� � ��� ���� ��$%#�?
'' � ������, � �������, ����� ����� 20. ��� 3.              EDIT: � ��� ��� � ����, �����, ������
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
        For Each Ctrl In frmWriteAuditData.Controls         '� �� �� ����� - ��� SQL �����
            If InStr(1, Ctrl.Tag, "SQLbox") <> 0 Then
                ReDim Preserve SQLBoxes(ibIndex)
                Dim SQLBoxTag() As String
                SQLBoxTag = Split(Ctrl.Tag, ",")
                SQLibName = SQLBoxTag(1)
                SQLBoxes(ibIndex) = SQLibName
                ibIndex = ibIndex + 1
            End If
        Next
    
''''''''''''''''''''''''''''''''��� ��������� ������ ���� ������ �� ���� ������''''''''''''''''''''''''''''''''
''���������� ��������� ���������� ������ � ���������� � ������ ��
    Select Case CLIArg
        
        Case "/edit"
            If IsUserAnAdmin() = 1 Then '���������� � WinAPI ��� ���� ����� ������, ���������� �� ���� ������������
                frmWriteAuditData.Show
            Else
                MsgBox "��������� �� � ������� ��������������!", vbExclamation, LARSver '���� ������������ ������������ ����
                End                                                                     '�� �� ���������� � ���������� � "��, ��".
            End If
        Case "/wmi"
        frmWMIQL.Show
        Case Else
            If IsUserAnAdmin() = 1 Then
                AuditorOnly = True
                SilentRun = True
                Call PopulateAuditData
                Exit Sub
            Else
                MsgBox "��������� �� � ������� ��������������!", vbExclamation, LARSver
                End
            End If
    End Select

End Sub
