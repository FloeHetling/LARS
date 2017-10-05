Attribute VB_Name = "modStartup"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' СТАРТОВЫЙ МОДУЛЬ "ЛАРСА" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Глобальные переменные
    Public LARSver As String
    Public InfoBoxes() As String
    Public HostName As String
    Public enumSQLFields As Integer 'учет нулей в классе SQLAuditData
      
'Глобальные константы
''Оформление
    Public Const Lime = 12648384
    Public Const Sand = 12648447
    Public Const Red = 12632319

'Глобальные объекты
''Используем класс AuditData, обзываем его thisPC
    Public thisPC As New auditdata
    Public thisPCSQL As New SQLAuditData

'Подключаемые библиотеки стартового модуля
''Библиотека и функция получения имени ПК
    Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
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
''
'' Зачем вообще это нужно:
'' Вот сейчас в программе 8 полей по 8 параметрам данных.
'' а завтра, к примеру, нужно будет 20. или 3.
'' Поэтому ПО абстрагировано от прямого обращения к объектам.
'' Список ведется по элементам формы. Соответственно, чтобы расширить или сузить выборку
'' необходимо изменить главную форму и добавить соответствующее свойство классам AuditData и SQLAuditData
''
Dim Ctrl As Control
Dim ibIndex As Integer
Dim ibName As String
ibIndex = 0

    For Each Ctrl In frmWriteAuditData.Controls         'Поэтому, для каждого элемента формы
        If InStr(1, Ctrl.Tag, "infobox") <> 0 Then      'который имеет слово infobox в свойстве Tag
            ReDim Preserve InfoBoxes(ibIndex)           'мы обновляем массив Infoboxes типом инфобокса
            Dim InfoboxTag() As String
            InfoboxTag = Split(Ctrl.Tag, ",")
            ibName = InfoboxTag(1)                      'который написан в свойстве Tag после слова "infobox,"
            InfoBoxes(ibIndex) = ibName                 'присваивая этот тип в виде строки элементу массива с индексом по порядку
            ibIndex = ibIndex + 1                       'присвоив, берем следующий элемент формы и проверяем/добавляем его
        End If
    Next                                                'на выходе у нас есть список полей для этой формы - массив InfoBoxes,
                                                        'по которому мы и будем обращаться к процедурам и функциям

'отправляем параметры коммандной строки в переменную и парсим их
'CLIArg = Command$
CLIArg = "/edit"
    Select Case CLIArg
        
        Case "/edit"
        frmWriteAuditData.Show
        
        Case Else
        Call PopulateAuditData
                
    End Select

'получаем в глобальную переменную текущее имя ПК
Dim dwLen As Long
    'Создаем буфер
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    HostName = String(dwLen, "X")
    'Получаем имя ПК
    GetComputerName HostName, dwLen
    'Убираем лишние (нулевые) символы
    HostName = Left(HostName, dwLen)
End Sub
