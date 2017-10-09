VERSION 5.00
Begin VB.Form frmWriteAuditData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������������� ������ ��� ������"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  '�������
   ScaleWidth      =   526
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tDelayedWriteData 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   6720
      Top             =   2760
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "����������������"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox chkSQL 
      Caption         =   "������������ SQL"
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Timer tDelayedReadData 
      Interval        =   20
      Left            =   6240
      Top             =   2760
   End
   Begin VB.Timer tResetColor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   2760
   End
   Begin VB.Frame frInfo 
      Caption         =   "����������"
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5535
      Begin VB.Frame frCommon 
         Caption         =   "�����"
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5295
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   2
            ItemData        =   "frmWriteAuditData.frx":0000
            Left            =   1800
            List            =   "frmWriteAuditData.frx":0002
            TabIndex        =   3
            Tag             =   "infobox,WSSerial"
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   2
            Tag             =   "infobox,WSName"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   0
            ItemData        =   "frmWriteAuditData.frx":0004
            Left            =   1800
            List            =   "frmWriteAuditData.frx":0006
            TabIndex        =   1
            Tag             =   "infobox,Company"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblInfo 
            Caption         =   "����� � ��������:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "��� ��:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "�����������:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frOffice 
         Caption         =   "Office"
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   5295
         Begin VB.ComboBox cbinfo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   8
            Tag             =   "infobox,OfficeLicenseModel"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   7
            Tag             =   "infobox,OfficeVersion"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblInfo 
            Caption         =   "������ ��������:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "�������� Office:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frWindows 
         Caption         =   "Windows"
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   5295
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   6
            Tag             =   "infobox,WindowsOLPSerial"
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   4
            Left            =   1800
            TabIndex        =   5
            Tag             =   "infobox,WindowsLicenseModel"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbinfo 
            Height          =   315
            Index           =   3
            Left            =   1800
            TabIndex        =   4
            Tag             =   "infobox,WindowsVersion"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblInfo 
            Caption         =   "����� OLP:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "������ ��������:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "�������� Windows:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame frMisc 
      Caption         =   "������"
      Height          =   975
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdOptions 
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdLaunchAIDA 
         Height          =   600
         Left            =   1320
         Picture         =   "frmWriteAuditData.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "��������� AIDA64 �� �������� ���������"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdLaunchCLI 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   720
         Picture         =   "frmWriteAuditData.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "��������� ���������� ������ ���������� ��"
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmWriteAuditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '�������� "�������������" ����������� - ����������� ���������� ����

Dim ctlInfobox As Control
Dim isDataChanged As Boolean, isSQLSyncCompleted As Boolean


Private Function LoadAuditData()
Dim ctlIBValue As String, cbAuditValue As String, cbAuditValueSQL As String

tResetColor.Enabled = True
enumSQLFields = UBound(InfoBoxes) - LBound(InfoBoxes) + 1
'��������� ������
thisPC.RegLoad
If chkSQL.Value = 1 Then     '����� ���������� � ����� ��. ��� ����� � ����������
    thisPCSQL.SQLLoad (HostName)    '� ��������� ������ modStartup (Sub Main). ��� �� ���������� ������ ������ SQLAuditData
End If
    For Each ctlInfobox In Me.Controls                              '������ ��������� ��� ������� �������� �� ��������������� � sub_main �������
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then            '���� � ������ ����� ���� ��� ���������
            Dim InfoboxTag() As String
            InfoboxTag = Split(ctlInfobox.Tag, ",")
            ctlIBValue = InfoboxTag(1)                              '��������� �� ���� ��� ���������
            cbAuditValue = CallByName(thisPC, ctlIBValue, VbGet)    '� ������� ������ AuditData �������� �������� � ���� ��������
                
                ' "infobox,"
                '
                ' ������ �� ���������� � ������
                ' ��������� ������ MAD cbExists ���������, ���� �� ���� ������� � ����������
                ' ���� ���� - �� ���������
                '
                If cbExists(cbAuditValue, ctlInfobox) = False Then
                    With ctlInfobox
                     .AddItem (cbAuditValue)
                     .BackColor = Sand
                     .ListIndex = 0
                    End With
                End If
                '
                '��������� ������ ��������� �� ������������ �������� �� SQL ����
                '���� �� ����������� � cbAuditValue � �� ���������� � ��������� - ��������� � �������� � ������ ��� �������
                '
                If chkSQL.Value = 1 Then '������ ��� ������ ���� ����� ������ "�������� � SQL"
                    cbAuditValueSQL = CallByName(thisPCSQL, ctlIBValue, VbGet)
                    If (cbAuditValueSQL <> cbAuditValue) _
                        And (cbExists(cbAuditValueSQL, ctlInfobox) = False) _
                        And Not cbAuditValueSQL = "sql_err_nodata" Then
                            With ctlInfobox
                                .AddItem (cbAuditValueSQL)
                                .BackColor = Red
                                .Tag = .Tag + ",noreset"
                                .ListIndex = 0
                            End With
                    End If
                End If
        End If
    Next
    
    '' ����� ������ � ������ ������� �������
    '' � ����������� ������ � ���� ������ � ������
    If enumSQLFields = 0 Then
            If MsgBox("� �� �� ������� ������� �������� � " & HostName & "!" & vbCrLf & _
                        "������� �������� �������� � ������� ����� ��� ����� ������ � ��?", _
                        vbQuestion & vbYesNo, LARSver) = vbYes Then
                        
                        '�������� ������ �����
                            Dim cbiCount As Integer, NullFieldWarning As Boolean
                            For cbiCount = 0 To cbinfo().UBound
                                If cbinfo(cbiCount).Text = "��� ������" Or _
                                cbinfo(cbiCount).Text = "" Then _
                                NullFieldWarning = True Else _
                                NullFieldWarning = False
                            Next
                        If NullFieldWarning = True Then
                            If MsgBox("���� ��� ��������� ����� �� ����� �� ���������. ����������?", _
                            vbQuestion & vbYesNo, LARSver) = vbYes Then _
                            Call SaveAuditData(laWriteToSQL)
                        End If
            End If
    End If
End Function

Private Function SaveAuditData(ByVal WriteMode As laWriteMode)
Dim ctlIBVariable As String
Dim ctlIBValue As String
tResetColor.Enabled = True
    For Each ctlInfobox In Me.Controls
        '
        '��� ���� ��������� ����� � ����� infobox
        '�� ��������� ��� � �������� ������ AuditData
        '�����, �������� ����� �� ����� ����������
        '� �������� � ��� ���������� ��������������� ���� ��
        '���������, ������� � ������ ������ ��������� � �����
        '
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            Dim InfoboxTag() As String
            InfoboxTag = Split(ctlInfobox.Tag, ",")
            ctlIBVariable = InfoboxTag(1)
            ctlInfobox.BackColor = Lime
            ctlIBValue = ctlInfobox.Text
          '
          ' ctlIBValue = ctlInfobox.List(ctlInfobox.ListIndex) ����� ����� ������ �� ����
          '
            If WriteMode = laWriteToRegistry Then CallByName thisPC, ctlIBVariable, VbLet, ctlIBValue
            If WriteMode = laWriteToSQL Then CallByName thisPCSQL, ctlIBVariable, VbLet, ctlIBValue
        End If
    Next
    
    '��������� ��� �������� � ����� infobox � �������� ��� ���������� ������ AuditData
    '��������� ���������� ��������� ������, ������������ ������ � ������ Windo
        Select Case WriteMode
            Case laWriteEverywhere
                thisPC.RegSave
                thisPCSQL.SQLSave (HostName)
            Case laWriteToRegistry
                thisPC.RegSave
            Case laWriteToSQL
                thisPCSQL.SQLSave (HostName)
        End Select

isDataChanged = False
End Function

Private Sub cbinfo_Change(Index As Integer)
isDataChanged = True
End Sub

Private Sub cbinfo_Click(Index As Integer)
If cbinfo(Index).BackColor = Red Then
    With cbinfo(Index)
    .BackColor = vbWhite
    .Tag = Replace(.Tag, ",noreset", "")
    End With
End If
End Sub

Private Sub cbinfo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(cbinfo(Index), KeyAscii)
End Sub

Private Sub cmdLoad_Click()

End Sub

Private Sub cmdSubmit_Click()
    
End Sub

Private Sub chkSQL_Click()
' Debug.Print SQLExecute("SELECT * FROM dbo.larspc", laRX) ������ ���� �� ����� -2147467259
End Sub

Private Sub cmdLaunchAIDA_Click()
Shell "\\zdc5\work\Administrator\AIDA\aida64.exe", vbNormalFocus
End Sub

Private Sub cmdLaunchCLI_Click()
Shell "cmd.exe", vbNormalFocus
End Sub

Private Sub cmdSync_Click()
Call LoadAuditData
tDelayedWriteData.Enabled = True
cmdSync.Enabled = False
    Dim cbInfoCount As Integer
        For cbInfoCount = 0 To cbinfo().UBound
            cbinfo(cbInfoCount).Enabled = False
        Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If isDataChanged = True Then
    If MsgBox("���� ������������� ��������� �������" & vbCrLf & "�� ����� ������ �����?", vbQuestion & vbYesNo, LARSver) = vbNo Then Cancel = 1
End If
'If isSQLSyncCompleted = False Then
'    if msgbox("������ ��������������� ������ � SQL �� ����� ��?")
'
''' ��� ������ �� ���������...
End Sub

Private Sub tDelayedReadData_Timer()
Call LoadAuditData
tDelayedReadData.Enabled = False
End Sub

Private Sub tDelayedWriteData_Timer()
Select Case chkSQL.Value
        Case 0
            Call SaveAuditData(laWriteToRegistry)
        Case 1
            Call SaveAuditData(laWriteToSQL)
End Select
tDelayedWriteData.Enabled = False
cmdSync.Enabled = True
    Dim cbInfoCount As Integer
        For cbInfoCount = 0 To cbinfo().UBound
            cbinfo(cbInfoCount).Enabled = True
        Next
End Sub

Private Sub tResetColor_Timer()
Dim ibColor As Integer
    For Each ctlInfobox In Me.Controls
    If (InStr(1, ctlInfobox.Tag, "infobox") <> 0) And Not (InStr(1, ctlInfobox.Tag, "noreset") <> 0) Then ctlInfobox.BackColor = vbWhite
    Next
tResetColor.Enabled = False
End Sub


