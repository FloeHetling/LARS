VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   0  'Нет
   Caption         =   "ЛАРС: Выполняется длительная процедура"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKillApp 
      Caption         =   "Чтобы завершить работу ПО - нажмите здесь"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5415
   End
   Begin VB.Timer Animate 
      Interval        =   40
      Left            =   5160
      Top             =   120
   End
   Begin VB.PictureBox container 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Shape ticker 
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Заливка
         Height          =   255
         Index           =   4
         Left            =   1560
         Shape           =   5  'Скругленный квадрат
         Top             =   150
         Width           =   255
      End
      Begin VB.Shape ticker 
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Заливка
         Height          =   255
         Index           =   3
         Left            =   1200
         Shape           =   5  'Скругленный квадрат
         Top             =   150
         Width           =   255
      End
      Begin VB.Shape ticker 
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Заливка
         Height          =   255
         Index           =   2
         Left            =   840
         Shape           =   5  'Скругленный квадрат
         Top             =   150
         Width           =   255
      End
      Begin VB.Shape ticker 
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Заливка
         Height          =   255
         Index           =   1
         Left            =   480
         Shape           =   5  'Скругленный квадрат
         Top             =   150
         Width           =   255
      End
      Begin VB.Shape ticker 
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Заливка
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   5  'Скругленный квадрат
         Top             =   150
         Width           =   255
      End
      Begin VB.Label Reason 
         Caption         =   "Поиск SQL-базы"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   105
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmwait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TickerIndex As Integer, KillTime As Integer
Dim OddEven As Boolean

Private Sub Animate_Timer()
    If KillTime = 0 Then
        Me.Height = 645
        container.Height = 615
    End If
DoEvents
    If TickerIndex = 0 Then
        If OddEven = False Then OddEven = True Else OddEven = False
    End If
    
    Select Case OddEven
        Case True
        ticker(TickerIndex).FillStyle = 1
        Case False
        ticker(TickerIndex).FillStyle = 0
    End Select
    
TickerIndex = TickerIndex + 1
    If TickerIndex = 5 Then TickerIndex = 0
KillTime = KillTime + 1
If KillTime > 200 Then
frmwait.Height = 1230
container.Height = 1215
End If
End Sub


Private Sub cmdKillApp_Click()
End
End Sub

Public Sub Form_Load()
DoEvents
Me.Height = 645
container.Height = 615
End Sub

Private Sub Form_LostFocus()
KillTime = 0
Animate.Enabled = False
Me.Height = 645
container.Height = 615
Unload Me
End Sub

Private Sub Form_Terminate()
Animate.Enabled = False
KillTime = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Animate.Enabled = False
KillTime = 0
End Sub

