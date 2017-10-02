VERSION 5.00
Begin VB.Form frmWriteAuditData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Редактирование данных для аудита"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
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
         Begin VB.ComboBox cbWSSerial 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Tag             =   "infobox,WSSerial"
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox cbWSName 
            Height          =   315
            Left            =   1800
            TabIndex        =   2
            Tag             =   "infobox,WSName"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbCompany 
            Height          =   315
            ItemData        =   "frmWriteAuditData.frx":0000
            Left            =   1800
            List            =   "frmWriteAuditData.frx":0002
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
         Begin VB.ComboBox cbOfficeLicenseModel 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Tag             =   "infobox,OfficeLicenseModel"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbOfficeVersion 
            Height          =   315
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
         Begin VB.ComboBox cbWindowsOLPSerial 
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Tag             =   "infobox,WindowsOLPSerial"
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox cbWindowsLicenseModel 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Tag             =   "infobox,WindowsLicenseModel"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cbWindowsVersion 
            Height          =   315
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
   Begin VB.Frame frLaunch 
      Caption         =   "Запустить"
      Height          =   975
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdLaunchAIDA 
         Height          =   600
         Left            =   840
         Picture         =   "frmWriteAuditData.frx":0004
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
         Left            =   120
         Picture         =   "frmWriteAuditData.frx":1046
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
      Height          =   1455
      Left            =   5760
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "За&писать"
         Default         =   -1  'True
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Прочитать"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
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
Dim ctlIBValue As String
tResetColor.Enabled = True
thisPC.RegLoad
    For Each ctlInfobox In Me.Controls
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            ctlIBValue = Replace(ctlInfobox.Tag, "infobox,", "")
            ctlInfobox.BackColor = Sand
            ctlInfobox.Text = CallByName(thisPC, ctlIBValue, VbGet)
        End If
    Next
End Function

Private Function SaveAuditData()
Dim ctlIBVariable As String
Dim ctlIBValue As String
tResetColor.Enabled = True
    For Each ctlInfobox In Me.Controls
    
        'для всех элементов формы с тегом infobox
        'мы конвертим тег в свойство класса AuditData
        'затем, вызываем класс по имени экземпляра
        'и помещаем в его переменные соответствующую инфу из
        'инфобокса, который в данный момент участвует в цикле
        
        If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then
            ctlIBVariable = Replace(ctlInfobox.Tag, "infobox,", "")
            ctlInfobox.BackColor = Lime
            ctlIBValue = ctlInfobox.Text
            CallByName thisPC, ctlIBVariable, VbLet, ctlIBValue
        End If
    Next
    
    'обработав все элементы с тегом infobox и заполнив все переменные класса AuditData
    'запускаем внутреннюю процедуру класса, записывающую данные в реестр Windows
    
thisPC.RegSave
End Function

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

''''''''''''''''''''''''''''''''''''
' заготовка для автокомплита

'Option Explicit
'Private Sub cb1_KeyPress(KeyAscii As Integer)
'   KeyAscii = AutoMatchCBBox(cb1, KeyAscii)
'End Sub
'
'Private Sub Form_Initialize()
'    Dim count As Integer, index  As Integer, aDate As Date
'    Randomize
'    count = Int((25 - 5 + 1) * Rnd) + 5
'    aDate = Date
'    Do While count > 0
'        Randomize
'        cb1.AddItem Format(aDate + Int(365 * Rnd), "mmm dd, yyyy")
'        count = count - 1
'    Loop
'    cb1.ListIndex = 0
'End Sub
''''''''''''''''''''''''''''''''''''

Private Sub tResetColor_Timer()
    For Each ctlInfobox In Me.Controls
    If InStr(1, ctlInfobox.Tag, "infobox") <> 0 Then ctlInfobox.BackColor = vbWhite
    Next
tResetColor.Enabled = False
End Sub
