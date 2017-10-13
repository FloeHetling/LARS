Attribute VB_Name = "modAutoAwesome"
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


Public Function PopulateAuditData()
Debug.Print "Старт функции PopulateAuditData"
On Error Resume Next
    Dim HW_query As String
    Dim HW_results As Object
    Dim HW_info As Object
HnS.Reset
'' ОБРАЗЕЦ
'    HW_query = "SELECT * FROM "
'    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
'    For Each info In HW_results
'        var = info.
'    Next info

''CPUNAME
    HW_query = "SELECT * FROM Win32_Processor"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.CPUName = HW_info.name
    Next HW_info
    
''CPUSOCKET
    HW_query = "SELECT * FROM win32_processor"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.CPUSocket = SocketLibrary(HW_info.upgrademethod)
    Next HW_info

''RAMTYPE
    HW_query = "SELECT * FROM Win32_PhysicalMemory"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.RAMType = RAMLibrary(HW_info.MemoryType)
    Next HW_info

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

''RAMSLOTSTAT
    ramModuleQuery = "SELECT * FROM win32_PhysicalMemory"
    RamModulesCount = 0
    Set ramModules = GetObject("Winmgmts:").ExecQuery(ramModuleQuery)
    For Each ramModule In ramModules
        If HnS.RAMSlotStat <> "" Then HnS.RAMSlotStat = HnS.RAMSlotStat & ","
        HnS.RAMSlotStat = HnS.RAMSlotStat & ramModule.DeviceLocator & "@" & toGB(ramModule.capacity) & "GB"
        RamModulesCount = RamModulesCount + 1
    Next ramModule
''RAMUSEDSLOTS
    HnS.RAMUsedSlots = RamModulesCount
  
''RAMVALUE
    HW_query = "SELECT * FROM Win32_ComputerSystem"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.RAMValue = toGB(HW_info.TotalPhysicalMemory) & " GB."
    Next HW_info
   
''MBNAME
    HW_query = "SELECT * FROM Win32_BaseBoard"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        HnS.MBName = HW_info.Manufacturer & " " & HW_info.Product & " REV. " & HW_info.Version
    Next HW_info
    
''MBCHIPSET
    Dim isChipset As String
    HW_query = "SELECT * FROM Win32_PnPEntity"
    Set HW_results = GetObject("Winmgmts:").ExecQuery(HW_query)
    For Each HW_info In HW_results
        If isChipset <> "" Then isChipset = isChipset & ","
        isChipset = isChipset + HW_info.Caption
    Next HW_info
        HnS.MBChipset = DeviceEnum(isChipset, laChipset)
        
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
    
''Monitors
    HnS.Monitors = GetMonitorInfo

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

Dim SQLAPRequest As String, SQLRequest As String, HnSValue As String, HnSArgsIndex As Integer
SQLRequest = "SELECT WSName FROM aida.dbo.hwinfo WHERE WSName = '" & HostName & "';"

'Проверяем, есть ли в базе индекс с нашим именем ПК
'Если нет - добавляем
If (SQLExecute(SQLRequest, laRX, "WSName") = 3265) Or (SQLExecute(SQLRequest, laRX, "WSName") = 3021) Then
        SQLRequest = "INSERT INTO dbo.hwinfo (WSName) VALUES ('" & HostName & "');"
        SQLExecute SQLRequest, laTX
End If

'Исполняем для всех собранных параметров
For HnSArgsIndex = 0 To UBound(HnSArgs)
    'Получаем имя параметра в SQLAPRequest
        SQLAPRequest = HnSArgs(HnSArgsIndex)
    'Получаем значение параметра в HnSValue
        HnSValue = CallByName(HnS, HnSArgs(HnSArgsIndex), VbGet)
    'Обновляем таблицу HWInfo - ставим значение SQLAPRequest равным HnSValue если выбранная запись содержит имя ПК
        SQLRequest = "UPDATE aida.dbo.hwinfo SET " & SQLAPRequest & " = '" & HnSValue & "' WHERE WSName = '" & HostName & "';"
        SQLExecute SQLRequest, laTX
Next HnSArgsIndex



Debug.Print "ОК: Функция PopulateAuditData исполнена" & vbCrLf
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
        Call fReadValue("HKCR", "Word.Application\CurVer", "", "S", "NA", OVerString)
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
        End Select
        GetOfficeVersion = "Microsoft Office " & OVerString
End Function

Public Function GetWindowsKey() As String

Dim hKey As Long, lDataSize As Long, szBinData As String, nIndx As Integer
Dim strComputer, message
Dim intMonitorCount
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3
Dim sValue
Dim iRC, iRC2, iRC3
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintDProdID
Dim strRawEDID
Dim ByteValue, strSerFind, strMdlFind
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmpser, tmpmdl, tmpctr
Dim batch, bHeader

strComputer = HostName
strComputer = UCase(strComputer)

Dim HexBuf() As Byte

intMonitorCount = 0
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'get a handle to the WMI registry object
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "/root/default:StdRegProv")
 
        If Err <> 0 Then
            If batch Then
                EchoAndLog strComputer & ",,,,," & Err.Description
            Else
                MsgBox "Failed. " & Err.Description, vbCritical + vbOKOnly, strComputer
                Exit Function
            End If
        End If
         

        sBaseKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        oRegistry.GetBinaryValue HKLM, sBaseKey, "DigitalProductId", HexBuf
'            For nIndx = 0 To UBound(HexBuf()) Step 1
'                szBinData = szBinData + Right("00" + Hex(HexBuf(nIndx)), 2) + " "
'            Next nIndx
'        szBinData = Trim(szBinData)

        Dim tmp As String
        tmp = ""

        Dim l As Integer
        For l = LBound(HexBuf) To UBound(HexBuf)
            tmp = tmp & " " & Hex(HexBuf(l))
        Next

        Dim StartOffset As Integer
        StartOffset = 52
        Dim EndOffset As Integer
        EndOffset = 67
        Dim Digits(24) As String

        Digits(0) = "B": Digits(1) = "C": Digits(2) = "D": Digits(3) = "F"
        Digits(4) = "G": Digits(5) = "H": Digits(6) = "J": Digits(7) = "K"
        Digits(8) = "M": Digits(9) = "P": Digits(10) = "Q": Digits(11) = "R"
        Digits(12) = "T": Digits(13) = "V": Digits(14) = "W": Digits(15) = "X"
        Digits(16) = "Y": Digits(17) = "2": Digits(18) = "3": Digits(19) = "4"
        Digits(20) = "6": Digits(21) = "7": Digits(22) = "8": Digits(23) = "9"

        Dim dLen As Integer
        dLen = 29
        Dim sLen As Integer
        sLen = 15
        Dim HexDigitalPID(15) As String
        Dim Des(30) As String

        Dim tmp2 As String
        tmp2 = ""
        Dim i As Integer
        For i = StartOffset To EndOffset
            HexDigitalPID(i - StartOffset) = HexBuf(i)
            tmp2 = tmp2 & " " & Hex(HexDigitalPID(i - StartOffset))
        Next

        Dim KEYSTRING As String
        KEYSTRING = ""
        For i = dLen - 1 To 0 Step -1
            If ((i + 1) Mod 6) = 0 Then
                Des(i) = "-"
                KEYSTRING = KEYSTRING & "-"
            Else
                Dim HN As Integer
                HN = 0
                Dim n As Integer
                For n = (sLen - 1) To 0 Step -1
                    Dim value As Integer
                    value = ((HN * 2 ^ 8) Or HexDigitalPID(n))
                    HexDigitalPID(n) = value \ 24
                    HN = (value Mod 24)

                Next

                Des(i) = Digits(HN)
                KEYSTRING = KEYSTRING & Digits(HN)
            End If
        Next

GetWindowsKey = StrReverse(KEYSTRING)
End Function

Public Function SocketLibrary(ByVal UpgradeMethodIndex As Integer) As String
If UpgradeMethodIndex <> 0 Then UpgradeMethodIndex = UpgradeMethodIndex - 1
'Дружно скажем большое спасибо дядюшке Гейтсу за то, что WMI НЕ ИМЕЕТ НИКАКИХ значений для прямого вывода строки типа сокета
Dim arSocketTypes() As Variant
arSocketTypes = Array("Other", "Unknown", "Daughter Board", "ZIF Socket", "Replacement/Piggy Back", "None", "LIF Socket", "Slot 1", "Slot 2", "370 Pin Socket", "Slot A", "Slot M", "Socket 423", "Socket A (Socket 462)", "Socket 478", "Socket 754", "Socket 940", "Socket 939", "Socket mPGA604", "Socket LGA771", "Socket LGA775", "Socket S1", "Socket AM2", "Socket F (1207)", "Socket LGA1366", "Socket G34", "Socket AM3", "Socket C32", "Socket LGA1156", "Socket LGA1567", "Socket PGA988A", "Socket BGA1288", "rPGA988B", "BGA1023", "BGA1224", "LGA1155", "LGA1356", "LGA2011", "Socket FS1", "Socket FS2", "Socket FM1", "Socket FM2", "Socket LGA2011-3", "Socket LGA1356-3", "Socket LGA1150", "Socket BGA1168", "Socket BGA1234", "Socket BGA1364", "Socket AM4", "Socket LGA1151", "Socket BGA1356", "Socket BGA1440", "Socket BGA1515")
SocketLibrary = arSocketTypes(UpgradeMethodIndex)
End Function

Public Function RAMLibrary(ByVal RAMIndex As Integer) As String
'Дружно СНОВА скажем большое спасибо дядюшке Гейтсу за то, что WMI НЕ ИМЕЕТ НИКАКИХ значений для прямого вывода строки типа оперативки
Dim arRAMTypes() As Variant
arRAMTypes = Array("Unknown", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM", "EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM", "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM", "DDR", "DDR-2", "BRAM", "FB-DIMM", "DDR3", "FBD2", "DDR4", "LPDDR", "LPDDR2", "LPDDR3", "LPDDR4", "DMTF Reserved", "Vendor Reserved")
If RAMIndex <> 0 Then RAMLibrary = arRAMTypes(RAMIndex) Else RAMLibrary = "DDR3"
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
