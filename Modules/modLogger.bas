Attribute VB_Name = "modLogger"
' Модуль логгирования
' Использует глобальную переменную LARSLogFile

Enum LoggerMode
    ContinueReport
    StartNewReport
End Enum

Option Explicit

Public Function WriteToLog(ByVal TextLine As String, Optional LoggerMode As LoggerMode)
On Error GoTo LOG_ERROR
Dim iLogFile As Integer

    If LoggerMode = StartNewReport Then
        iLogFile = FreeFile
            Open LARSLogFile For Output As #iLogFile
                Print #iLogFile, TextLine
            Close #iLogFile
    Else
        iLogFile = FreeFile
            Open LARSLogFile For Append As #iLogFile
                Print #iLogFile, TextLine
            Close #iLogFile
    End If

Exit Function
LOG_ERROR:
End
End Function
