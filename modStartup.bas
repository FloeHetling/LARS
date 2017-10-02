Attribute VB_Name = "modStartup"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' СТАРТОВЫЙ МОДУЛЬ "ЛАРСА" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Глобальные переменные
    Public LARSver As String
    Public InfoBoxes() As String
      
'Глобальные константы
''Оформление
    Public Const Lime = 12648384
    Public Const Sand = 12648447

'Глобальные объекты
''Используем класс AuditData, обзываем его thisPC
    Public thisPC As New AuditData


Dim CLIArg As String

Sub Main()
'записываем в глобальную переменную зазвание и версию ПО
LARSver = App.ProductName & ", версия " & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.CompanyName
'проверяем, запущен ли другой экземпляр
'если да - прибиваем агент нахрен
    If App.PrevInstance = True Then
        Exit Sub
        End
    End If

'создаем список имеющихся на форме инфоокон и запихиваем их в публичный массив
Dim Ctrl As Control
Dim ibIndex As Integer
Dim ibName As String
ibIndex = 0

    For Each Ctrl In frmWriteAuditData.Controls
        If InStr(1, Ctrl.Tag, "infobox") <> 0 Then
            ReDim Preserve InfoBoxes(ibIndex)
            ibName = Replace(Ctrl.Tag, "infobox,", "")
            InfoBoxes(ibIndex) = ibName
            ibIndex = ibIndex + 1
        End If
    Next

'отправляем параметры коммандной строки в переменную и парсим их
CLIArg = Command$
    Select Case CLIArg
        
        Case "/edit"
        frmWriteAuditData.Show
        
        Case Else
        Call PopulateAuditData
                
    End Select
        
End Sub
