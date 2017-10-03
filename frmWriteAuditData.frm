VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmWriteAuditData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Редактирование данных для аудита"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc SQL 
      Height          =   330
      Left            =   5760
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer tDelayedReadData 
      Interval        =   20
      Left            =   6240
      Top             =   1080
   End
   Begin VB.Timer tResetColor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   1080
   End
   Begin VB.Frame frInfo 
      Caption         =   "Информация"
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5535
      Begin VB.Frame frCommon 
         Caption         =   "Общая"
         Height          =   1455
         Left            =   120
         TabIndex        =   22
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
            Caption         =   "Номер с наклейки:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Имя ПК:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Организация:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frOffice 
         Caption         =   "Office"
         Height          =   1095
         Left            =   120
         TabIndex        =   19
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
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Office:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frWindows 
         Caption         =   "Windows"
         Height          =   1455
         Left            =   120
         TabIndex        =   15
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
            Caption         =   "Номер OLP:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Модель лицензии:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Редакция Windows:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame frMisc 
      Caption         =   "Прочее"
      Height          =   975
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdOptions 
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdLaunchAIDA 
         Height          =   600
         Left            =   1320
         Picture         =   "frmWriteAuditData.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Запустить AIDA64 из сетевого хранилища"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton cmdLaunchCLI 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   720
         Picture         =   "frmWriteAuditData.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Запустить коммандную строку локального ПК"
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame frRegistry 
      Caption         =   "Реестр"
      Height          =   1815
      Left            =   5760
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
      Begin VB.CheckBox chkSQLCompare 
         Caption         =   "Сравнить с SQL"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "За&писать"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Прочитать"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmWriteAuditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода

Dim ctlInfobox As Control

Private Function LoadAuditData()
Dim ctlIBValue As String, cbAuditValue As String

tResetColor.Enabled = True
thisPC.RegLoad
    For Each ctlInfobox In Me.Controls
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            ctlIBValue = Replace(ctlInfobox.Tag, "infobox,", "")
            cbAuditValue = CallByName(thisPC, ctlIBValue, VbGet)
                '
                ' защита от дубликатов в списке
                ' процедура модуля MAD cbExists проверяет, есть ли этот элемент в комбобоксе
                ' если есть - не добавляет
                '
                If cbExists(cbAuditValue, ctlInfobox) = False Then
                    With ctlInfobox
                     .AddItem (cbAuditValue)
                     .BackColor = Sand
                     .ListIndex = 0
                    End With
                End If
        End If
    Next
End Function

Private Function SaveAuditData()
Dim ctlIBVariable As String
Dim ctlIBValue As String
tResetColor.Enabled = True
    For Each ctlInfobox In Me.Controls
        '
        'для всех элементов формы с тегом infobox
        'мы конвертим тег в свойство класса AuditData
        'затем, вызываем класс по имени экземпляра
        'и помещаем в его переменные соответствующую инфу из
        'инфобокса, который в данный момент участвует в цикле
        '
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            ctlIBVariable = Replace(ctlInfobox.Tag, "infobox,", "")
            ctlInfobox.BackColor = Lime
            ctlIBValue = ctlInfobox.Text
          '
          ' ctlIBValue = ctlInfobox.List(ctlInfobox.ListIndex) Этого здесь нахрен не надо
          '
            CallByName thisPC, ctlIBVariable, VbLet, ctlIBValue
        End If
    Next
    
    'обработав все элементы с тегом infobox и заполнив все переменные класса AuditData
    'запускаем внутреннюю процедуру класса, записывающую данные в реестр Windows
    
thisPC.RegSave
End Function

Private Sub cbinfo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(cbinfo(Index), KeyAscii)
End Sub

Private Sub cmdLoad_Click()
Call LoadAuditData
End Sub

Private Sub cmdSubmit_Click()
Call SaveAuditData
End Sub

Private Sub cmdLaunchAIDA_Click()
Shell "\\zdc5\work\Administrator\AIDA\aida64.exe", vbNormalFocus
End Sub

Private Sub cmdLaunchCLI_Click()
Shell "cmd.exe", vbNormalFocus
End Sub

Private Sub tDelayedReadData_Timer()
Call LoadAuditData
tDelayedReadData.Enabled = False
End Sub

Private Sub tResetColor_Timer()
    For Each ctlInfobox In Me.Controls
    If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then ctlInfobox.BackColor = vbWhite
    Next
tResetColor.Enabled = False
End Sub
