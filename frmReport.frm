VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отправить отчет о различиях (редактирование)"
   ClientHeight    =   6120
   ClientLeft      =   11805
   ClientTop       =   5460
   ClientWidth     =   7140
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Тема"
      Top             =   120
      Width           =   5535
   End
   Begin VB.TextBox txtBody 
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   1
      Text            =   "frmReport.frx":058A
      Top             =   600
      Width           =   6855
   End
   Begin VB.CommandButton txtSend 
      Caption         =   "Отправить"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
