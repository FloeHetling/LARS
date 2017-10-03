Attribute VB_Name = "modMAD"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''' Manage Audit Data (MAD) module ''''''''''''''''''''''''
        ''''''''''''''''''' МОДУЛЬ РАБОТЫ С АУДИТ-ДАННЫМИ ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода

Public Function PopulateAuditData()
MsgBox "ОТЛАДКА: Параметры не переданы. Запустите с параметром /edit для редактирования аудит-информации", vbInformation, LARSver
MsgBox "TODO: модуль автозагрузки информации!!!", vbExclamation, LARSver
End Function

Public Function RegGetAuditData(ByVal AuditProp As String) As String
Dim AuditValue As String
'получаем переменные из функции и передаем их в функцию модуля работы с реестром
'если данных нет - так и пишем
Call fReadValue("HKLM", "Software\LARS", AuditProp, "S", "Нет данных", AuditValue)
RegGetAuditData = AuditValue
End Function

Public Function RegPutAuditData(ByVal AuditProp As String, ByVal auditdata As String)
'данная функция вызывается методом Call, использует два параметра - соответственно
'куда класть данные и какие именно данные класть
Call fWriteValue("HKLM", "Software\LARS", AuditProp, "S", auditdata)
End Function

Public Function cbExists(SearchString As String, ComboBoxForCheck As ComboBox) As Boolean
Dim cItem As Integer
                For cItem = 0 To ComboBoxForCheck.ListCount Step 1
                     If SearchString = ComboBoxForCheck.List(cItem) Then
                     cbExists = True
                     Exit Function
                     End If
                Next cItem
                cbExists = False
End Function

Public Function SQLGetAuditData(ByVal AuditProp As String, ByVal auditdata As String)

End Function
