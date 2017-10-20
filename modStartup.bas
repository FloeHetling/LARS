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
    Public HostName As String, LARSINIPath As String, LARSLogFile As String
    Public enumSQLFields As Integer 'учет нулей в классе SQLAuditData
    Public SilentRun As Boolean, isAllSettingsProvided As Boolean
    
'Перечисление параметров, читаемых из INI
    Public INIParameters As New Collection

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

Public Function isSettingsIntegrityOK() As Boolean
On Error GoTo INTEGRITYCHECK_ERROR
WriteToLog " "
WriteToLog "К работе приступил модуль проверки настроек программы"
Dim SettingsArray() As String, ParamsArray() As String, SIndex As Integer, SErr As Integer, Setting As Variant

'Перечисляем, какие настройки из INI мы проверяем
    With INIParameters
                .Add "DataSource"
                .Add "SMTPServer"
                .Add "SMTPPort"
                .Add "FromEmail"
                .Add "ToEmail"
                .Add "EmailServer"
                .Add "EmailServerPort"
    End With
    
'Проверяем, существует ли файл в целом
'Если не существует - сразу ставим False и выходим из процедуры
    If CheckPath(LARSINIPath) <> True Then
            WriteToLog " "
            WriteToLog "Не найден файл настроек. Создаем пустой, чтобы модулю сохранения было куда писать настройки"
            WriteToLog " "
            '''''Создаем структуру файла
            Dim iFileNo As Integer
            iFileNo = FreeFile
        
            Open LARSINIPath For Output As #iFileNo
            Print #iFileNo, ";Only Windows-1251 Codepage is allowed!"
            Print #iFileNo, ";Если вы можете прочесть эту строку, ваша кодировка установлена правильно"
            Print #iFileNo, ""
            Close #iFileNo
            '''''' И поехали дальше.
        isSettingsIntegrityOK = False
        Exit Function
    End If
     
'Если из процедуры не вышли - проверяем каждый параметр из коллекции
'При этом считаем провалы
    SErr = 0
    For Each Setting In INIParameters
        If INIQuery("MAIN", Setting) = "" Then SErr = SErr + 1
    Next Setting

'Если на счетчике провалов есть хоть что-нибудь - целостность настроек явно нарушена.
    If SErr <> 0 Then
        isSettingsIntegrityOK = False
        WriteToLog " "
        WriteToLog "Модуль проверки настроек сообщил, что настройки не корректны."
        WriteToLog "Упс."
        WriteToLog "Их надо срочно исправить. Они глобальные! Поправьте файл INI в директории программы"
        WriteToLog "Или запустите ЛАРС с параметром /edit. С любого ПК."
    Else
        isSettingsIntegrityOK = True
        WriteToLog "Модуль проверки настроек проблем не обнаружил."
    End If

WriteToLog "Завершение проверки настроек"
WriteToLog "======================================================="
WriteToLog " "
Exit Function
INTEGRITYCHECK_ERROR:
WriteToLog " "
WriteToLog "Модуль проверки корректности настроек сообщил о критической ошибке " & Err.Number & ":"
WriteToLog Err.description
WriteToLog "Однозначная пасхалка. Чтобы вызвать эту ошибку - надо очень постараться"
WriteToLog "======================================================="
End
End Function

Public Function INIQuery(ByVal Div As String, ByVal Param As String) As String
Dim INIReadResult As String
Call fReadValue(LARSINIPath, Div, Param, "S", "", INIReadResult)
INIQuery = INIReadResult
End Function

Sub Main()
On Error GoTo ERR_STARTUP
CLIArg = Command$
''получаем в глобальную переменную текущее имя ПК
    Dim dwLen As Long
        'Создаем буфер
        dwLen = MAX_COMPUTERNAME_LENGTH + 1
        HostName = String(dwLen, "X")
        'Получаем имя ПК
        GetComputerName HostName, dwLen
        'Убираем лишние (нулевые) символы
        HostName = Left(HostName, dwLen)


    If InStr(1, CLIArg, "/logpath") <> 0 Then
            Dim logStrArray() As String, logStrArrayIdx As Integer
                logStrArray = Split(CLIArg, " ")
                For logStrArrayIdx = 0 To UBound(logStrArray)
                    If logStrArray(logStrArrayIdx) = "/logpath" Then
                        If logStrArrayIdx + 1 <= UBound(logStrArray) Then
                            LARSLogFile = logStrArray(logStrArrayIdx + 1)
                            LARSLogFile = Replace(LARSLogFile, "%20", " ") & "\" & HostName & ".log"
                        End If
                    End If
                Next logStrArrayIdx
    Else
        LARSLogFile = App.Path & "\" & HostName & ".log"
    End If


    If InStr(1, CLIArg, "/inifile") <> 0 Then
        Dim IniStrArray() As String, IniStrArrayIdx As Integer
            IniStrArray = Split(CLIArg, " ")
            For IniStrArrayIdx = 0 To UBound(IniStrArray)
                If IniStrArray(IniStrArrayIdx) = "/inifile" Then
                    If IniStrArrayIdx + 1 <= UBound(IniStrArray) Then
                        LARSINIPath = IniStrArray(IniStrArrayIdx + 1)
                        LARSINIPath = Replace(LARSINIPath, "%20", " ")
                    End If
                End If
            Next IniStrArrayIdx
    Else
        LARSINIPath = App.Path & "\lars.ini"
    End If

''записываем в глобальную переменную зазвание и версию ПО
    LARSver = App.ProductName & " " & _
                App.Major & "." & App.Minor & _
                "." & App.Revision & " - " & _
                App.CompanyName

''Вывести справку если в аргументах какая-нибудь херня и закончить работу
Dim msgHelp As String

        msgHelp = _
        LARSver & vbCrLf & vbCrLf & _
        "Допустимые параметры командной строки:" & vbCrLf & vbCrLf & _
        "Без параметров - Запустить Аудитор в тихом режиме" & vbCrLf & _
        "/edit - Проверить настройки ПО и запустить Редактор" & vbCrLf & _
        "/wmi - Открыть окно прямой работы с WMI" & vbCrLf & _
        "/? - Показать данное окно" & vbCrLf & _
        "--------------------------" & vbCrLf & vbCrLf & _
        "Переопределение параметров:" & vbCrLf & vbCrLf & _
        "/inifile - Путь до файла с настройками ПО *" & vbCrLf & _
        "/logpath - Путь до папки для логов * **" & vbCrLf & vbCrLf & _
        "* - путь без кавычек, пробелы заменены на ""%20""" & vbCrLf & _
        "** - без слеша ( \ ) в конце пути к папке"
        
        If CLIArg = "/?" Then
                MsgBox msgHelp, vbInformation, "Справка"
                End
        End If


''Получаем в глобальную переменную путь до файла настроек
WriteToLog "=============== " & Date & " " & Time & " ===============", StartNewReport
WriteToLog "LARS APP LAUNCHED. Logfile Codepage is Windows-1251", ContinueReport
WriteToLog "               VERSION " & App.Major & "." & App.Minor & ", build " & App.Revision & "               "
WriteToLog "==================================================="
WriteToLog "Читаю файл конфигурации " & LARSINIPath
        
isAllSettingsProvided = False

'''''''''''''''''''''''РАБОТА С ФАЙЛОМ INI'''''''''''''''''''''''
    'Исполнить если Attended режим - если интегрити получает фейл - исправляем ситуацию, иначе - ставим статус настроек в ОК и продолжаем
    If CLIArg <> "" Then
        If isSettingsIntegrityOK = False Then
            MsgBox "Не найден файл с настройками ПО" & vbCrLf & _
            "Либо не все настройки указаны корректно." & vbCrLf & vbCrLf & _
            "Пожалуйста, заполните отсутствующие настройки!", vbExclamation, LARSver
            frmSettings.Show vbModal
            If isAllSettingsProvided = False Then End
        Else
            isAllSettingsProvided = True
        End If
    Else
        If isSettingsIntegrityOK = False Then End
    End If
    
    'Исполнить если UnAttended - если интегрити проходит успешно - ставим статус настроек в ок
    If CLIArg = "" Then
        If isSettingsIntegrityOK = True Then isAllSettingsProvided = True
    End If
If isAllSettingsProvided = False Then End 'Если и после этого не все заполнено - прога не запустится!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
''Задаем параметры почты:
        SMTPServer = INIQuery("MAIN", "SMTPServer")
        SMTPPort = INIQuery("MAIN", "SMTPPort")
        FromEmail = INIQuery("MAIN", "FromEmail")
        ToEmail = INIQuery("MAIN", "ToEmail")
        EmailServer = INIQuery("MAIN", "EmailServer")
        EmailServerPort = INIQuery("MAIN", "EmailServerPort")
        EmailSubject = "ЛАРС: Отчет по аудиту рабочей станции """ & HostName & """"
        SendFormCallOnly = False

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
                "Data Source=" & INIQuery("MAIN", "DataSource") & "" & _
                "Persist Security Info=False;" & _
                "Initial Catalog=LARS;" & _
                "User ID=lars;" & _
                "Connect Timeout=2;" & _
                "Password=XzlOq2JNh8;"
    isSQLChecked = False

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

Dim StartArgs As String
If InStr(1, CLIArg, "/wmi") <> 0 Then StartArgs = "wmi"
If InStr(1, CLIArg, "/edit") <> 0 Then StartArgs = "edit"

    Select Case StartArgs
        Case "edit"
            If IsUserAnAdmin() = 1 Then 'обращаемся к WinAPI для того чтобы узнать, достаточно ли прав пользователь
                frmWriteAuditData.Show
            Else
                MsgBox "Запустите ПО с правами администратора!", vbExclamation, LARSver 'если пользователь недостаточно прав
                End                                                                     'то не стесняемся в выражениях и "ой, всё".
            End If
        Case "wmi"
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
Exit Sub
ERR_STARTUP:
WriteToLog "На самом старте программы возникла ошибка " & Err.Number & ":"
WriteToLog Err.description
WriteToLog "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
End
End Sub
