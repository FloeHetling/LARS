VERSION 5.00
Begin VB.Form frmWMIQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Прямой запрос WMI"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tWQLResult 
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   3
      Top             =   1080
      Width           =   7335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Выполнить"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox tWQLItem 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Введите имя параметра"
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox tWQLRequest 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Введите WQL класс"
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmWMIQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuery_Click()
On Error Resume Next
If tWQLRequest.Text = "Введите WQL класс" Then tWQLRequest.Text = ""
If tWQLItem.Text = "Введите имя параметра" Then tWQLItem.Text = ""
    Dim HW_query As String
    Dim HW_results As Object
    Dim HW_info As Object
'' ОБРАЗЕЦ
If tWQLRequest.Text <> "" And tWQLItem.Text <> "" Then
    HW_query = "SELECT * FROM " & Trim(tWQLRequest.Text)
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        tWQLResult.Text = tWQLResult.Text & vbCrLf & CallByName(HW_info, tWQLItem.Text, VbGet)
    Next HW_info
End If
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tWQLItem_Click()
If tWQLItem.Text = "Введите имя параметра" Then tWQLItem.Text = ""
End Sub

Private Sub tWQLRequest_Click()
If tWQLRequest.Text = "Введите WQL класс" Then tWQLRequest.Text = ""
End Sub

