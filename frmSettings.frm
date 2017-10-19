VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������� ����"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6315
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "������"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "��������� � ����������"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame container 
      Caption         =   "��������� �����"
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6135
      Begin VB.CommandButton cmdCheckEmail 
         Caption         =   "���������"
         Height          =   645
         Left            =   4800
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox tTo 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox tFrom 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox tMailPort 
         Height          =   285
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox tMailServer 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label description 
         Caption         =   "����� ����������:"
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label description 
         Caption         =   "����� ����:"
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label description 
         Caption         =   "���� �������:"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label description 
         Caption         =   "FQDN ��� IP ������� SMTP:"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame container 
      Caption         =   "������ SQL (TCP-IP)"
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdTestSQL 
         Caption         =   "���� �����������"
         Height          =   285
         Left            =   4080
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox tIPPort 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox tIPAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label description 
         Caption         =   "����:"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label description 
         Caption         =   "IP �����:"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''����� ������ ������� isAllSettingsProvided As Boolean

Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCheckEmail_Click()
'Set SMTP = New CSocketMaster
'On Error GoTo MAIL_CONNECT_ERROR
'
'SMTP.Connect Trim(tMailServer.Text), Val(tMailPort.Text)
'WinsockState = MAIL_CONNECT
End Sub

Private Sub cmdSubmit_Click()
Dim DataIsCorrect As Boolean
Dim Excpt As String

DataIsCorrect = True
'��������� ���������� ����� IP ������
    Dim asIP() As String, i As Integer
    
    asIP = Split(tIPAddress.Text, ".")
        If UBound(asIP) = 3 Then
            For i = 0 To 3
                 If (CInt(asIP(i)) > 255) Or (CInt(asIP(i)) < 0) Then
                        Excpt = Excpt & vbCrLf & "IP-����� SQL-������� - ����� IP-������ �� �����"
                        DataIsCorrect = False
                 End If
            Next
        Else
            Excpt = Excpt & vbCrLf & "IP-����� SQL-������� - ������ �� �� �����"
            DataIsCorrect = False
        End If
'��������� ���������� ������
    If Val(tIPPort.Text) > 65535 Then
        DataIsCorrect = False
        Excpt = Excpt & vbCrLf & "����������� ������ ���� SQL"
    End If
    If Val(tMailPort.Text) > 65535 Then
        DataIsCorrect = False
        Excpt = Excpt & vbCrLf & "����������� ������ ���� ��������� �������"
    End If

'��������� ���������� �������� �������
    Dim iAt, iDot As Integer
    
    '�������� ���� ���
    iAt = InStr(1, tFrom.Text, "@")
    iDot = InStr(1, tFrom.Text, ".")
    
    If (iAt = 0) Or (iDot = 0) Or (iAt > iDot) Then
        DataIsCorrect = False
        Excpt = Excpt & vbCrLf & "������ ������ �������� - ���-�� ����������"
    End If
    
    '�������� ���� ���
    iAt = InStr(1, tTo.Text, "@")
    iDot = InStr(1, tTo.Text, ".")
    
    If (iAt = 0) Or (iDot = 0) Or (iAt > iDot) Then
        DataIsCorrect = False
        Excpt = Excpt & vbCrLf & "����� �������� �� �������. ���� �����???"
    End If

''�������� ���������
''����� � INI����

    If DataIsCorrect = True Then
        If SQLPreTest = True Then
            fWriteValue LARSINIPath, "MAIN", "DataSource", "S", "tcp:" & tIPAddress.Text & "," & tIPPort.Text & "[oledb]"
            fWriteValue LARSINIPath, "MAIN", "SMTPServer", "S", tMailServer.Text
            fWriteValue LARSINIPath, "MAIN", "SMTPPort", "S", tMailPort.Text
            fWriteValue LARSINIPath, "MAIN", "EmailServer", "S", tMailServer.Text
            fWriteValue LARSINIPath, "MAIN", "FromEmail", "S", tFrom.Text
            fWriteValue LARSINIPath, "MAIN", "ToEmail", "S", tTo.Text
            fWriteValue LARSINIPath, "MAIN", "EmailServerPort", "S", tMailPort.Text
            MsgBox "��������� ��������� � ��������." & vbCrLf & "���������� ������", vbInformation, LARSver
            isAllSettingsProvided = True
            Unload Me
        Else
            MsgBox "�� ���� ������������ � SQL." & vbCrLf & "��� ������� ��� ��������!", vbCritical, LARSver
        End If
    Else
        MsgBox "��������� ������ �� ����� ������� �����������. ��� ������: " & vbCrLf & Excpt, vbCritical, LARSver
    End If
End Sub

Private Sub cmdTestSQL_Click()
Call SQLPreTest
End Sub

Private Function SQLPreTest() As Boolean
cmdTestSQL.Enabled = False
Dim ConnStrDataSource As String
isSQLChecked = False
ConnStrDataSource = "tcp:" & tIPAddress.Text & "," & tIPPort.Text & "[oledb]"
SQLConnString = "Provider = SQLOLEDB.1;" & _
                "Data Source=" & ConnStrDataSource & "" & _
                "Persist Security Info=False;" & _
                "Initial Catalog=LARS;" & _
                "User ID=sa;" & _
                "Connect Timeout=2;" & _
                "Password=happyness;"
    If isSQLAvailable = True Then
        SQLPreTest = True
        MsgBox "���������� ������ �������", vbInformation, LARSver
    Else
        SQLPreTest = False
        MsgBox "�� ������� ����������� � �������� SQL!", vbCritical, LARSver
    End If
    
isSQLChecked = False
cmdTestSQL.Enabled = True
End Function

Private Sub Form_Load()
Dim mSQLDataSource As String, mSQLSettings() As String

    If CheckPath(LARSINIPath) = True Then
        mSQLDataSource = INIQuery("MAIN", "DataSource")
        mSQLDataSource = Replace(Replace(mSQLDataSource, "tcp:", ""), "[oledb]", "")
        mSQLSettings = Split(mSQLDataSource, ",")
            If mSQLDataSource <> "" Then
                tIPAddress.Text = mSQLSettings(0)
                tIPPort.Text = mSQLSettings(1)
            End If
        tMailServer.Text = INIQuery("MAIN", "SMTPServer")
        tMailPort.Text = INIQuery("MAIN", "SMTPPort")
        tFrom.Text = INIQuery("MAIN", "FromEmail")
        tTo.Text = INIQuery("MAIN", "ToEmail")
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If isAllSettingsProvided <> True Then
        If MsgBox("��������� ������ �� ����!" & vbCrLf _
                & "�� ����� ������ ������� ����?" & vbCrLf _
                & "��� �������� ������ ���������", _
                vbQuestion & vbYesNo, LARSver) = vbYes Then
            End
        Else
            Cancel = 1
        End If
    End If
End Sub
