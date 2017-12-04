VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отправить отчет о различиях (редактирование)"
   ClientHeight    =   3570
   ClientLeft      =   11805
   ClientTop       =   5460
   ClientWidth     =   7050
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7050
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
      Height          =   2895
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
'    Public FromEmail As String, _
'            ToEmail As String, _
'            EmailSubject As String, _ - txtsubject.text
'            MailMessage As String, _ - txtbody.text
'            EmailServer As String, _
'            EmailServerPort As String
Option Explicit

Dim WithEvents SMTP As CSocketMaster
Attribute SMTP.VB_VarHelpID = -1

Private Sub Form_Load()
Set SMTP = New CSocketMaster
txtSubject.Text = EmailSubject
txtBody.Text = MailMessage
'txtBody.Text = Replace(MailMessage, "<br>", vbCrLf) 'без br

    If SendFormCallOnly = True Then
        Me.Visible = False
        SMTP.Connect Trim(SMTPServer), Val(SMTPPort)
        WinsockState = MAIL_CONNECT
    Else
    Me.Visible = True
    End If
End Sub

Private Sub txtSend_Click()
EmailSubject = txtSubject.Text
'MailMessage = Replace(txtBody.Text, vbCrLf, "<br>") без бр обойдемся
MailMessage = txtBody.Text

SMTP.Connect Trim(SMTPServer), Val(SMTPPort)
    'reset the state so our sequence works right
    WinsockState = MAIL_CONNECT
    
cmdSendEmailError:
    ' Show a detailed error message if needed
    If Err.Number <> 0 And AuditorOnly = False Then MsgBox "Ошибка отправки почты: " & vbCrLf & " Error Number: " & Err.Number & _
    vbCrLf & "Error Description: " & Err.description & ".", vbOKOnly + vbCritical, ""
End Sub

Private Sub SMTP_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    
    
    'Retrive data from winsock buffer
    SMTP.GetData strServerResponse
    
    ' Update our text box so we know whats going on.
    WriteToLog strServerResponse
    
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    
    'Only these three codes from the server tell us that the command was accepted
    If strResponseCode = "250" Or strResponseCode = "220" Or strResponseCode = "354" Then
        Select Case WinsockState
            Case MAIL_CONNECT
                WinsockState = MAIL_HELO
                'Remove blank spaces
                strDataToSend = Trim$(FromEmail)
                'Get just the email part of the from line
                strDataToSend = Mid(strDataToSend, 1 + InStr(1, strDataToSend, "<"))
                ' Then get just the account part
                strDataToSend = Left$(strDataToSend, InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                SMTP.SendData "HELO " & strDataToSend & vbCrLf
            Case MAIL_HELO
                WinsockState = MAIL_FROM
                'Send MAIL FROM command to the server so it knows from who the message comes
                SMTP.SendData "MAIL FROM: " & Mid(FromEmail, InStr(1, FromEmail, "<")) & vbCrLf
            Case MAIL_FROM
                WinsockState = MAIL_RCPTTO
                'Send RCPT TO command to the server so it knows where to send the message
                SMTP.SendData "RCPT TO: " & Mid(ToEmail, InStr(1, ToEmail, "<")) & vbCrLf
            Case MAIL_RCPTTO
                WinsockState = MAIL_DATA
                'Send DATA command to the server so it knows that we want to send the message
                SMTP.SendData "DATA" & vbCrLf
            Case MAIL_DATA
                WinsockState = MAIL_DOT
                'Send header and subject
                SMTP.SendData "Return-Path: <" & FromEmail & ">" & vbCrLf & _
                "Content-type: text/html; charset=Windows-1251" & vbCrLf & _
                "Priority: normal" & vbCrLf & _
                "To: " & ToEmail & vbCrLf & _
                "From: " & FromEmail & vbCrLf & _
                "Subject:" & EmailSubject & vbLf & vbCrLf
                
                '''''
                WriteToLog "Return-Path: <" & FromEmail & ">" & vbCrLf & _
                "Content-type: text/html; charset=UTF-8" & vbCrLf & _
                "Priority: normal" & vbCrLf & _
                "To: " & ToEmail & vbCrLf & _
                "From: " & FromEmail & vbCrLf & _
                "Subject:" & EmailSubject & vbLf & vbCrLf
                '''''
                
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                
                SMTP.SendData MailReport & vbCrLf & "." & vbCrLf
                WriteToLog MailReport & vbCrLf & "." & vbCrLf
            Case MAIL_DOT
                WinsockState = MAIL_QUIT
                'Send QUIT command
                SMTP.SendData "QUIT" & vbCrLf
                WriteToLog "QUIT" & vbCrLf
            Case MAIL_QUIT
                'Close the connection to the smtp server
                SMTP.CloseSck
        End Select
    Else
        'Check if an error occured
        SMTP.CloseSck
        If Not WinsockState = MAIL_QUIT Then
            'If yes then print the error
            If Left$(strServerResponse, 3) = 421 Then
                MsgBox "The from email address is invalid for this mail server.  Please check it and try again"
            Else
                MsgBox "Error: " & strServerResponse, vbCritical, "Error"
            End If
        Else
            'if the message sent successfully, print it
            WriteToLog "Отчет успешно отправлен"
            fWriteValue "HKLM", "Software\LARS", "Reported", "S", Date$
            If SilentRun = False Then MsgBox "Отчет отправлен", vbOKOnly + vbInformation, LARSver
            Unload frmReport
            If AuditorOnly = True Then End
        End If
    End If
End Sub

Private Sub SMTP_Error(ByVal Number As Integer, description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Tell the user that an error occured.  There is more then the numbers below but this is a good starting point.
    If SilentRun <> True Then
        If Number = 10049 Then
            MsgBox "Не могу отправить отчет о ПК - неправильный адрес сервера или порт!", vbCritical, LARSver
        ElseIf Number = 10061 Then
            MsgBox "Сервер почты отклонил мое сообщение. Отчет по ПК не отправлен!", vbCritical, LARSver
        ElseIf Number <> 0 Then
            MsgBox "Ошибка соединения с почтовым сервером: " & Number & vbCrLf & description & vbCrLf & vbCrLf & "Отчет по ПК не был отправлен." & vbCrLf & "Пожалуйста, сообщите об этой ошибке в отдел системного администрирования.", vbExclamation, LARSver
        End If
    End If
    
    SMTP.CloseSck
    
End Sub
