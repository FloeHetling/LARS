Attribute VB_Name = "modAuditor"
Option Explicit

'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
'Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
'Private Const REG_BINARY As Long = 3
'Private Const ERROR_SUCCESS As Long = 0&
'Private Const KEY_WOW64_64KEY As Long = &H100&
'Private Const KEY_READ As Long = &H20019
''''' SoftwareLicensingProduct class, WMI
Public cWinVer As String
Public cWinSer As String
Public cOffVer As String
Public cProcVer As String
Public MailReport As String, MRBaseString As String
Public Const RegNoData = "Нет данных"
Public Const SqlNoData = ""
Public AuditorOnly As Boolean, SendFormCallOnly As Boolean


Public Function PopulateAuditData()
On Error GoTo ERR_AUDITOR
WriteToLog " "
WriteToLog "================================================="
WriteToLog Date & " " & Time & " - Отчет по исполнению модуля ЛАРС: Аудитор"
WriteToLog "-------------------------------------------------"
WriteToLog "Старт функции PopulateAuditData"
MailReport = "Отчет ЛАРС об отсутствующих параметрах на рабочей станции " & HostName & ":" & vbCrLf
MRBaseString = MailReport
On Error Resume Next
    Dim HW_query As String
    Dim HW_results As Object
    Dim HW_info As Object
    Dim SQLErr As Integer
HnS.Reset
'' ОБРАЗЕЦ
'    HW_query = "SELECT * FROM "
'    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
'    For Each info In HW_results
'        var = info.
'    Next info
WriteToLog "Беру информацию о CPU"
''CPUNAME
    HW_query = "SELECT * FROM Win32_Processor"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.CPUName = HW_info.name
    Next HW_info
    
WriteToLog "Беру информацию о сокете"
''CPUSOCKET
    HW_query = "SELECT * FROM win32_processor"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.CPUSocket = SocketLibrary(HW_info.upgrademethod)
    Next HW_info

WriteToLog "Опрашиваю тип памяти"
''RAMTYPE
    HW_query = "SELECT * FROM Win32_PhysicalMemory"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.RAMType = RAMLibrary(HW_info.MemoryType)
    Next HW_info

WriteToLog "Опрашиваю ОЗУ"
'RAM
''RAMTOTALSLOT
    Dim ramSlotQuery As String, ramModuleQuery As String
    Dim ramSlot As Object, ramSlots As Object
    Dim ramModule As Object, ramModules As Object
    Dim RamModulesCount As Integer
    
    ramSlotQuery = "SELECT * FROM win32_PhysicalMemoryArray"
    Set ramSlots = GetObject("Winmgmts:").ExecQuery(ramSlotQuery)
    For Each ramSlot In ramSlots
        HnS.RAMTotalSlots = ramSlot.MemoryDevices
    Next ramSlot

WriteToLog "Получаю статистику по слотам"
''RAMSLOTSTAT
    ramModuleQuery = "SELECT * FROM win32_PhysicalMemory"
    RamModulesCount = 0
    Set ramModules = GetObject("Winmgmts:").ExecQuery(ramModuleQuery)
    For Each ramModule In ramModules
        If HnS.RAMSlotStat <> "" Then HnS.RAMSlotStat = HnS.RAMSlotStat & ","
        HnS.RAMSlotStat = HnS.RAMSlotStat & ramModule.DeviceLocator & "@" & toGB(ramModule.capacity) & "GB"
        RamModulesCount = RamModulesCount + 1
    Next ramModule

WriteToLog "Получаю количество модулей"
''RAMUSEDSLOTS
    HnS.RAMUsedSlots = RamModulesCount

WriteToLog "Получаю количество ОЗУ"
''RAMVALUE
    HW_query = "SELECT * FROM Win32_ComputerSystem"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.RAMValue = toGB(HW_info.TotalPhysicalMemory) & " GB."
    Next HW_info

WriteToLog "Запрашиваю модель материнской платы"
''MBNAME
    HW_query = "SELECT * FROM Win32_BaseBoard"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.MBName = HW_info.Manufacturer & " " & HW_info.Product & " REV. " & HW_info.Version
    Next HW_info

WriteToLog "Пытаюсь выяснить что за чипсет стоит на ПК"
''MBCHIPSET
    Dim isChipset As String
    HW_query = "SELECT * FROM Win32_PnPEntity"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        If isChipset <> "" Then isChipset = isChipset & ","
        isChipset = isChipset + HW_info.Caption
    Next HW_info
        HnS.MBChipset = DeviceEnum(isChipset, laChipset)

WriteToLog "Опрашиваю видеокарту"
''GPUNAME
    Dim GPUQuery As String, GPUsCount As Integer
    Dim GPU As Object, GPUs As Object
    GPUQuery = "SELECT * FROM Win32_VideoController"
    GPUsCount = 0
    Set GPUs = GetObject("Winmgmts:").ExecQuery(GPUQuery)
    For Each GPU In GPUs
        If HnS.GPUName <> "" Then HnS.GPUName = HnS.GPUName & ","
        HnS.GPUName = HnS.GPUName & GPU.Caption
        GPUsCount = GPUsCount + 1
    Next GPU

WriteToLog "Запрашиваю с мониторов строку производителя и расшифровываю ее"
''Monitors
    HnS.Monitors = GetMonitorInfo

WriteToLog "Опрашиваю винчестеры"
''HDD
    Dim HDDQuery As String, HDDsQuery As String, HDDModel As String
    Dim HDDisk As Object, HDDisks As Object
    Dim HDDrive As Object, HDDrives As Object
    Dim HDDrivesCount As Integer
    Dim HDDOverallSpace As Integer
    
    HDDsQuery = "SELECT * FROM Win32_DiskDrive"
    HDDrivesCount = 0
    Set HDDrives = GetObject("Winmgmts:").ExecQuery(HDDsQuery)
    For Each HDDrive In HDDrives
        If HnS.HDD <> "" Then HnS.HDD = HnS.HDD & ","
        HDDModel = Replace(HDDrive.Caption, " ATA Device", "")
        HnS.HDD = HnS.HDD & HDDModel & "@" & toGB(HDDrive.Size) & "GB"
        HDDrivesCount = HDDrivesCount + 1
        HDDOverallSpace = Int(HDDOverallSpace) + Int(toGB(HDDrive.Size))
    Next HDDrive
    HnS.HDDCount = HDDrivesCount
    HnS.HDDOverallSize = HDDOverallSpace

If isSQLAvailable = True Then
    WriteToLog " "
    WriteToLog "Мне удалось подключиться к SQL"
    Dim SQLAPRequest As String, SQLRequest As String, HnSValue As String, HnSArgsIndex As Integer
    SQLRequest = "SELECT WSName FROM lars.dbo.hwinfo WHERE WSName = '" & HostName & "';"
    
    'Проверяем, есть ли в базе индекс с нашим именем ПК
    'Если нет - добавляем
    SQLErr = SQLExecute(SQLRequest, laRX, "WSName")
    If (SQLErr = 3265) Or (SQLErr = 3021) Then
            SQLRequest = "INSERT INTO dbo.hwinfo (WSName) VALUES ('" & HostName & "');"
            WriteToLog "Исполняю TRANSACT-SQL запрос: " & SQLRequest
            SQLExecute SQLRequest, laTX
    End If
    
    'Исполняем для всех собранных параметров
    For HnSArgsIndex = 0 To UBound(HnSArgs)
        'Получаем имя параметра в SQLAPRequest
            SQLAPRequest = HnSArgs(HnSArgsIndex)
        'Получаем значение параметра в HnSValue
            HnSValue = CallByName(HnS, HnSArgs(HnSArgsIndex), VbGet)
        'Обновляем таблицу HWInfo - ставим значение SQLAPRequest равным HnSValue если выбранная запись содержит имя ПК
            SQLRequest = "UPDATE lars.dbo.hwinfo SET " & SQLAPRequest & " = '" & HnSValue & "' WHERE WSName = '" & HostName & "';"
            WriteToLog "Записал в базу " & HnSValue
            SQLExecute SQLRequest, laTX
    Next HnSArgsIndex
WriteToLog "Закончил запись в SQL"
WriteToLog " "
End If



'Составляем отчет на основе отсутствующих записей в реестре
'Если значения в реестре нет - пишем из SQL в реестр
'Иначе пишем значения из реестра в SQL
'Если и в реестре и в базе значений нет - добавляем строчку в отчет

    If isSQLAvailable = True Then
        'Читаем значения из реестра
        thisPC.RegLoad
        WriteToLog "Загрузил значения из реестра"
        'Читаем значения из SQL
        thisPCSQL.SQLLoad (HostName)
        WriteToLog "Загрузил значения из SQL"
        WriteToLog "Проверяю соответствия"
        WriteToLog " "
        
        If thisPC.Company = RegNoData Then
            If thisPCSQL.Company = SqlNoData Then
                MailReport = MailReport & "<br>Не заполнена информация о компании"
            Else
                thisPC.Company = thisPCSQL.Company
            End If
        Else
        thisPCSQL.Company = thisPC.Company
        End If
        
        If thisPC.WsName = RegNoData Then
            If thisPCSQL.WsName = SqlNoData Then
                MailReport = MailReport & "<br>Не заполнена информация о сетевом имени компа." & "<br>Чего в принципе быть не может. Значит в проге ошибка!"
            Else
                thisPC.WsName = thisPCSQL.WsName
            End If
        Else
        thisPCSQL.WsName = thisPC.WsName
        End If
        
        If thisPC.WSSerial = RegNoData Then
            If thisPCSQL.WSSerial = SqlNoData Then
                MailReport = MailReport & "<br>Не введен номер ПК с этикетки"
            Else
                thisPC.WSSerial = thisPCSQL.WSSerial
            End If
        Else
        thisPCSQL.WSSerial = thisPC.WSSerial
        End If
        
        If thisPC.WindowsVersion = RegNoData Then
            If thisPCSQL.WindowsVersion = SqlNoData Then
                MailReport = MailReport & "<br>Не заполнена информация о версии Windows. ОШИБКА: быть этого не может!!!"
            Else
                thisPC.WindowsVersion = thisPCSQL.WindowsVersion
            End If
        Else
        thisPCSQL.WindowsVersion = thisPC.WindowsVersion
        End If
        
        If thisPC.WindowsLicenseModel = RegNoData Then
            If thisPCSQL.WindowsLicenseModel = SqlNoData Then
                MailReport = MailReport & "<br>Нет информации о модели лицензирования ОС"
            Else
                thisPC.WindowsLicenseModel = thisPCSQL.WindowsLicenseModel
            End If
        Else
        thisPCSQL.WindowsLicenseModel = thisPC.WindowsLicenseModel
        End If
        
        If thisPC.WindowsOLPSerial = RegNoData And thisPC.WindowsLicenseModel = "OLP" Then
            If (thisPCSQL.WindowsOLPSerial = SqlNoData) Or (thisPCSQL.WindowsOLPSerial = RegNoData) Then
                MailReport = MailReport & "<br>Нет информации о номере OLP Windows"
            Else
                thisPC.WindowsOLPSerial = thisPCSQL.WindowsOLPSerial
            End If
        Else
        thisPCSQL.WindowsOLPSerial = thisPC.WindowsOLPSerial
        End If
        
        If thisPC.OfficeVersion = RegNoData Then
            If thisPCSQL.OfficeVersion = SqlNoData Then
                MailReport = MailReport & "<br>Не заполнена информация о версии Офиса. Или он просто не установлен."
            Else
                thisPC.OfficeVersion = thisPCSQL.OfficeVersion
            End If
        Else
        thisPCSQL.OfficeVersion = thisPC.OfficeVersion
        End If
        
        If thisPC.OfficeLicenseModel = RegNoData Then
            If thisPCSQL.OfficeLicenseModel = SqlNoData Then
                MailReport = MailReport & "<br>Не заполнена информация о модели лицензирования Офиса"
            Else
                thisPC.OfficeLicenseModel = thisPCSQL.OfficeLicenseModel
            End If
        Else
        thisPCSQL.OfficeLicenseModel = thisPC.OfficeLicenseModel
        End If
        
        If thisPC.OfficeOLPSerial = RegNoData And thisPC.OfficeLicenseModel = "OLP" Then
            If (thisPCSQL.OfficeOLPSerial = SqlNoData) Or (thisPCSQL.OfficeOLPSerial = RegNoData) Then
                MailReport = MailReport & "<br>Нет номера OLP Офиса"
            Else
                thisPC.OfficeOLPSerial = thisPCSQL.OfficeOLPSerial
            End If
        Else
        thisPCSQL.OfficeOLPSerial = thisPC.OfficeOLPSerial
        End If
        WriteToLog "Закончил проверять соответствия. Сформировал по ним отчет"
        WriteToLog "-----------------------------------"
        WriteToLog MailReport
        WriteToLog "-----------------------------------"
        WriteToLog " "
        WriteToLog "Выполняю синхронизацию SQL и реестра"
        thisPC.RegSave
        thisPCSQL.SQLSave (HostName)
        WriteToLog "Синхронизация завершена"
        WriteToLog " "
    End If
    
'Закончили сбор данных
'Обрабатываем собранное и если есть чего - отправляем по почте

WriteToLog "Сравниваем почтовую строку чтобы выяснить, надо ли отправлять отчет"

    If MailReport <> MRBaseString Then
        WriteToLog "Отчет не пуст, пытаемся отправить его почтой"
        MailReport = MailReport & "<br><br>Отчет сформирован " & Time & _
                                    " " & Date & _
                                    "." & "<br>Советую поставить задачу на заполнение отсутствующих данных в ручном режиме!"
        MailReport = Replace(MailReport, "<br>", vbCrLf)
        If AuditorOnly = True Then
            WriteToLog "Запускаю отправку через сервер " & SMTPServer & ", порт " & SMTPPort
            
            'Проверка существующего отчета
                Dim Reported As String
                Call fReadValue("HKLM", "Software\LARS", "Reported", "S", "", Reported)
                
                If Reported <> "" Then 'если значение пусто, т.е., если отчет еще не отправлялся
                    If AuditorOnly = False Then 'при режиме редактора и прочих - спросить
                        If MsgBox("Уже есть отправленный отчет от " & Reported & _
                                    ". Повторить?", vbQuestion & vbYesNo, LARSver) = vbYes Then 'и если ответ утвердительный
                            frmReport.Show vbModal                                              'то показать форму отправки. Модально.
                        End If
                    Else    'если включен режим аудитора - при отправленном отчете
                        WriteToLog " "
                        WriteToLog Date & " " & Time & " " & "ОК: Функция PopulateAuditData исполнена. Отчет отправлялся ранее."
                        End  'Просто пишем что отчет отправлялся ранее и завершаем работу программы
                    End If
                
                Else 'если строка Reported все же пустая
                    If AuditorOnly = True Then 'в режиме аудитора
                        SendFormCallOnly = True 'включаем тихий режим
                        Load frmReport          'и грузим форму отчета с соответствующими процедурами
                    Else                        'в режиме редактора и прочих
                        frmReport.Show vbModal  'грузим форму отправки почты. модально.
                    End If
                End If
                
          End If
    Else
        WriteToLog "Отчет пуст. Идем дальше"
        WriteToLog " "
    End If

    If AuditorOnly = False Then
        MailMessage = MailReport
            If MsgBox("Аудитор завершил процедуру." & vbCrLf & "Отправить результаты проверки?", vbQuestion & vbYesNo, LARSver) = vbYes Then
                frmReport.Show vbModal
            Else
                MsgBox "Вы можете выполнить отправку позже используя форму обратной связи." & vbCrLf & "Для этого выполните Инструменты - Сообщить о различиях" & vbCrLf & "или воспользуйтесь сочетанием клавиш Ctrl+E", vbInformation, LARSver
                frmWriteAuditData.cmdReport.Enabled = True
            End If
    End If

WriteToLog " "
WriteToLog Date & " " & Time & " " & "ОК: Функция PopulateAuditData исполнена"
If MailReport = MRBaseString And AuditorOnly = True Then End
Exit Function

ERR_AUDITOR:
WriteToLog Date & " " & Time & "Во время работы Аудитора возникла ошибка " & Err.Number & ":"
WriteToLog Err.description
WriteToLog "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
End
End Function

Public Function toGB(Bytes As Double) As Double
   'This function gives an estimate to two decimal
   'places.  For a more precise answer, format to
   'more decimal places or just return dblAns
 
  Dim dblAns As Double
  dblAns = ((Bytes / 1024) / 1024) / 1024
  toGB = Format(dblAns, "###,###,###")
End Function

Public Function GetWindowsVersion() As String
'Запрашиваем инфу о выпуске системы
    Dim WVerString As String
    Dim WProcArch As String
        Call fReadValue("HKLM", "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "S", "", WVerString)
        Call fReadValue("HKLM", "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", "S", "", WProcArch)
        If WProcArch <> "x86" Then WProcArch = "x64"
    GetWindowsVersion = WVerString & " " & WProcArch
End Function

Public Function GetOfficeVersion() As String
'Запрашиваем инфу о версии и выпуске Microsoft Office
    Dim OVerString As String, OVerIndex As Variant
    Dim OVerArray() As String
        Call fReadValue("HKCR", "Word.Application\CurVer", "", "S", "не установлен", OVerString)
        OVerArray = Split(OVerString, ".")
        For Each OVerIndex In OVerArray
            If OVerIndex = Val(OVerIndex) Then OVerString = OVerIndex
        Next
        Select Case OVerString
                Case "7"
                OVerString = "95"
                Case "8"
                OVerString = "97"
                Case "9"
                OVerString = "2000"
                Case "10"
                OVerString = "2002"
                Case "11"
                OVerString = "2003"
                Case "12"
                OVerString = "2007"
                Case "14"
                OVerString = "2010"
                Case "15"
                OVerString = "2013"
                Case "16"
                OVerString = "2016"
                Case "17"
                OVerString = "2019"
                Case Else
                OVerString = "365"
        End Select
        GetOfficeVersion = "Microsoft Office " & OVerString
End Function


Public Function SocketLibrary(ByVal UpgradeMethodIndex As Integer) As String
If UpgradeMethodIndex <> 0 Then UpgradeMethodIndex = UpgradeMethodIndex - 1
'Дружно скажем большое спасибо дядюшке Гейтсу за то, что WMI НЕ ИМЕЕТ НИКАКИХ значений для прямого вывода строки типа сокета
Dim arSocketTypes() As Variant

arSocketTypes = Array("Other", "Unknown", "Daughter Board", "ZIF Socket", _
                        "Replacement/Piggy Back", "None", "LIF Socket", "Slot 1", _
                        "Slot 2", "370 Pin Socket", "Slot A", "Slot M", "Socket 423", _
                        "Socket A (Socket 462)", "Socket 478", "Socket 754", "Socket 940", _
                        "Socket 939", "Socket mPGA604", "Socket LGA771", "Socket LGA775", _
                        "Socket S1", "Socket AM2", "Socket F (1207)", "Socket LGA1366", _
                        "Socket G34", "Socket AM3", "Socket C32", "Socket LGA1156", _
                        "Socket LGA1567", "Socket PGA988A", "Socket BGA1288", "rPGA988B", _
                        "BGA1023", "BGA1224", "LGA1155", "LGA1356", "LGA2011", "Socket FS1", _
                        "Socket FS2", "Socket FM1", "Socket FM2", "Socket LGA2011-3", _
                        "Socket LGA1356-3", "Socket LGA1150", "Socket BGA1168", "Socket BGA1234", _
                        "Socket BGA1364", "Socket AM4", "Socket LGA1151", "Socket BGA1356", _
                        "Socket BGA1440", "Socket BGA1515")

If UpgradeMethodIndex <> 9 Then SocketLibrary = arSocketTypes(UpgradeMethodIndex) Else SocketLibrary = "LGA1156"
End Function

Public Function RAMLibrary(ByVal RAMIndex As Integer) As String
'Дружно СНОВА скажем большое спасибо дядюшке Гейтсу за то, что WMI НЕ ИМЕЕТ НИКАКИХ значений для прямого вывода строки типа оперативки
Dim arRAMTypes() As Variant

arRAMTypes = Array("Unknown", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM", "EDO", _
                    "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM", "FEPROM", _
                    "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM", "DDR", "DDR-2", _
                    "BRAM", "FB-DIMM", "DDR3", "FBD2", "DDR4", "LPDDR", "LPDDR2", "LPDDR3", _
                    "LPDDR4", "DMTF Reserved", "Vendor Reserved")

If RAMIndex <> 0 Then RAMLibrary = arRAMTypes(RAMIndex) Else RAMLibrary = "Нет данных"
End Function

Public Function DeviceEnum(ByVal CSVString As String, Optional ByVal HardwareType As laHardware) As String
Dim PnPArray() As String, PnPAIndex As Integer, PnPAItem As Variant, PnPDeviceType As String
    PnPArray = Split(CSVString, ",")
    For Each PnPAItem In PnPArray
    
            Select Case HardwareType
                Case laChipset
                PnPDeviceType = "chipset"
            End Select
            
            If InStr(1, PnPAItem, PnPDeviceType, vbTextCompare) <> 0 Then
                DeviceEnum = Left(PnPAItem, InStr(1, PnPAItem, PnPDeviceType, vbTextCompare) + Len(PnPDeviceType))
                Exit Function
            End If
            
    Next PnPAItem
End Function
