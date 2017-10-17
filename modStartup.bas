Attribute VB_Name = "modStartup"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''' СТАРТОВЫЙ МОДУЛЬ "ЛАРСА" ''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Глобальные переменные
    Public LARSver As String
    Public InfoBoxes() As String
    Public SQLBoxes() As String
    Public HnSArgs() As Variant
    Public HostName As String
    Public enumSQLFields As Integer 'учет нулей в классе SQLAuditData
    Public SilentRun As Boolean
    
'Почта
''Параметры Winsock
    Public Enum WinsockControlState
        MAIL_CONNECT
        MAIL_HELO
        MAIL_FROM
        MAIL_RCPTTO
        MAIL_DATA
        MAIL_HEADER
        MAIL_DOT
        MAIL_QUIT
    End Enum

    Public WinsockState As WinsockControlState, SMTPServer As String, SMTPPort As String
    
''Глобальные переменные почты
    Public FromEmail As String, _
            ToEmail As String, _
            EmailSubject As String, _
            MailMessage As String, _
            EmailServer As String, _
            EmailServerPort As String
      
'Глобальные константы
''Оформление
    Public Enum laColorConstants
        laLightGreen = 12648384
        laSand = 12648447
        laLightRed = 12632319
        laDarkGreen = 32768
        laDarkRed = 192
        laDarkBlue = 12936533
        laBlack = 0
    End Enum
    
'Типы оборудования
    Public Enum laHardware
        laCPU
        laRAM
        laMotherboard
        laChipset
        laSouthBridge
        laUSBHost
        laGPU
        laMonitor
        laHDD
    End Enum
    
'Глобальные объекты
''Используем класс AuditData, обзываем его thisPC
    Public thisPC As New auditdata
    Public thisPCSQL As New SQLAuditData 'то же самое, только для обращения к SQL
    Public HnS As New HardAndSoft
    Public Ru As New AliasLibrary 'Библиотека алиасов для SQL запросов
    
'Подключаемые библиотеки стартового модуля
''Библиотека и функция получения имени ПК
    Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer
    
'    Sub Form_Load()
'
'   If IsUserAnAdmin() = 0 Then
'     MsgBox "Not admin"
'   Else
'     MsgBox "Admin"
'   End If
'
'End Sub
    
'Глобальная строка подключения для SQL
    Public SQLConnString As String

'Аргументы командной строки
Public CLIArg As String
        
Sub Main()
''получаем в глобальную переменную текущее имя ПК
    Dim dwLen As Long
        'Создаем буфер
        dwLen = MAX_COMPUTERNAME_LENGTH + 1
        HostName = String(dwLen, "X")
        'Получаем имя ПК
        GetComputerName HostName, dwLen
        'Убираем лишние (нулевые) символы
        HostName = Left(HostName, dwLen)

''Задаем параметры почты:
        SMTPServer = "mail.zdravservice.ru"
        SMTPPort = "25"
        FromEmail = "ЛАРС Аудитор <larsmailer@zdravservice.ru>"
        ToEmail = "Даниил Кузнецов <kuznecov_dm@zdravservice.ru>"
        EmailServer = "mail.zdravservice.ru"
        EmailServerPort = "25"
        EmailSubject = "ЛАРС: Отчет по аудиту рабочей станции """ & HostName & """"

''Задаем параметры командной строки
    CLIArg = Command$
    
''Формируем массив параметров по железу
    HnSArgs = Array("WSNAME", "CPUNAME", "RAMTYPE", _
                "RAMTOTALSLOTS", "RAMUSEDSLOTS", "RAMSLOTSTAT", _
                "RAMVALUE", "MBNAME", "MBCHIPSET", "GPUNAME", _
                "MONITORS", "HDD", "HDDCOUNT", _
                "HDDOVERALLSIZE", "CPUSOCKET")
     
''Задаем глобальные параметры подключения к SQL
    SQLConnString = "Provider = SQLOLEDB.1;" & _
                "Data Source=tcp:192.168.78.39,1433[oledb];" & _
                "Persist Security Info=False;" & _
                "Initial Catalog=AIDA;" & _
                "User ID=sa;" & _
                "Connect Timeout=2;" & _
                "Password=happyness;"
    isSQLChecked = False

''записываем в глобальную переменную зазвание и версию ПО
    LARSver = App.ProductName & " " & _
                App.Major & "." & App.Minor & _
                "." & App.Revision & " - " & _
                App.CompanyName

''проверяем, запущен ли другой экземпляр
'если да - прибиваем агент нахрен
    If App.PrevInstance = True Then
        MsgBox "ЛАРС уже работает! Пожалуйста, немного подождите." & vbCrLf & "Или снимите задачу ПО через Диспетчер задач", vbExclamation, LARSver
        Exit Sub
        End
    End If

''По умолчанию тихого режима нет
    SilentRun = False

'создаем список имеющихся на форме инфоокон и запихиваем их в публичный массив
''
'' Зачем вообще это нужно:
'' Вот сейчас в программе 8 полей по 8 параметрам данных. EDIT: А теперь их 21. Ну и кто тут дальновидная сволочь и кто кого на$%#л?
'' а завтра, к примеру, нужно будет 20. или 3.              EDIT: И вот как в воду, сцуко, глядел
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
    Dim SQLibName As String
    ibIndex = 0
        For Each Ctrl In frmWriteAuditData.Controls         'и то же самое - для SQL полей
            If InStr(1, Ctrl.Tag, "SQLbox") <> 0 Then
                ReDim Preserve SQLBoxes(ibIndex)
                Dim SQLBoxTag() As String
                SQLBoxTag = Split(Ctrl.Tag, ",")
                SQLibName = SQLBoxTag(1)
                SQLBoxes(ibIndex) = SQLibName
                ibIndex = ibIndex + 1
            End If
        Next
    
''''''''''''''''''''''''''''''''ВСЕ ПАРАМЕТРЫ ДОЛЖНЫ БЫТЬ ЗАДАНЫ ДО ЭТОЙ СТРОКИ''''''''''''''''''''''''''''''''
''отправляем параметры коммандной строки в переменную и парсим их
CLIArg = Command$
    Select Case CLIArg
        
        Case "/edit"
            If IsUserAnAdmin() = 1 Then 'обращаемся к WinAPI для того чтобы узнать, достаточно ли прав пользователь
                frmWriteAuditData.Show
            Else
                MsgBox "Запустите ПО с правами администратора!", vbExclamation, LARSver 'если пользователь недостаточно прав
                End                                                                     'то не стесняемся в выражениях и "ой, всё".
            End If
        Case "/wmi"
        frmWMIQL.Show
        Case Else
            If IsUserAnAdmin() = 1 Then
                AuditorOnly = True
                SilentRun = True
                Call PopulateAuditData
                Exit Sub
            Else
                MsgBox "Запустите ПО с правами администратора!", vbExclamation, LARSver
                End
            End If
    End Select
End Sub
