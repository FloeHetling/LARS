Attribute VB_Name = "modRegistry"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''' МОДУЛЬ ДЛЯ РАБОТЫ С РЕЕСТРОМ '''''''''''''''''''''''
        ''''''''''''''''''''''' И ФАЙЛАМИ НАСТРОЕК (ini) '''''''''''''''''''''''''
        ''''''''''''''''''' ВЗЯТ С http://www.thescarms.com ''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit 'повышаем "придирчивость" компилятора - увеличиваем надежность кода

Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Boolean
End Type

Public Const MAX_SIZE = 2048
Public Const MAX_INISIZE = 8192
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

' Registry security attributes
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8

Declare Function RegEnumValue Lib "advapi32.dll" _
        Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" _
        Alias "RegDeleteValueA" _
        (ByVal hKey As Long, ByVal lpValueName As String) _
        As Long

Declare Function RegDeleteKey Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, _
        ByVal dwOptions As Long, ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
        lpdwDisposition As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpszValueName As String, _
        ByVal lpdwReserved As Long, lpdwType As Long, _
        lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegEnumKey Lib "advapi32.dll" Alias _
        "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpName As String, ByVal cbName As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

Declare Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal _
        lpFileName As String) As Long
        
Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        
Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) _
        As Long

Declare Function GetPrivateProfileInt Lib "kernel32" _
        Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName _
        As String) As Long

Public Function fDeleteKey(ByVal sTopKey As String, _
    ByVal sSubKey As String, ByVal sKeyName As String) As Long

Dim lTopKey As Long
Dim lHandle As Long
Dim lResult As Long

On Error GoTo fDeleteKeyError
lResult = 99
lTopKey = fTopKey(sTopKey)
If lTopKey = 0 Then GoTo fDeleteKeyError

lResult = RegOpenKeyEx(lTopKey, sSubKey, 0, KEY_CREATE_SUB_KEY, lHandle)
If lResult = ERROR_SUCCESS Then
    lResult = RegDeleteKey(lHandle, sKeyName)
End If

If lResult = ERROR_SUCCESS Or lResult = ERROR_FILE_NOT_FOUND Then
    fDeleteKey = ERROR_SUCCESS
Else
    fDeleteKey = lResult
End If
Exit Function

fDeleteKeyError:
    fDeleteKey = lResult
End Function

Public Function fDeleteValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String) As Long

Dim lTopKey As Long
Dim lHandle As Long
Dim lResult As Long

On Error GoTo fDeleteValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fDeleteValueError

If lTopKey = 1 Then
    lResult = WritePrivateProfileString(sSubKeyOrSection, sValueName, "", sTopKeyOrFile)
Else
    lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_SET_VALUE, lHandle)
    If lResult = ERROR_SUCCESS Then
        lResult = RegDeleteValue(lHandle, sValueName)
    End If
    
    If lResult = ERROR_SUCCESS Or lResult = ERROR_FILE_NOT_FOUND Then
        fDeleteValue = ERROR_SUCCESS
    Else
        fDeleteValue = lResult
    End If
End If
Exit Function

fDeleteValueError:
    fDeleteValue = lResult
End Function

Public Function fEnumKey(ByVal sTopKey As String, _
    ByVal sSubKey As String, sValues As String) As Long

Dim bDone    As Boolean
Dim lTopKey  As Long
Dim lHandle  As Long
Dim lResult  As Long
Dim lIndex   As Long
Dim sKeyName As String

On Error GoTo fEnumKeyError
lResult = 99
lTopKey = fTopKey(sTopKey)
If lTopKey = 0 Then GoTo fEnumKeyError

lResult = RegOpenKeyEx(lTopKey, sSubKey, 0, KEY_ENUMERATE_SUB_KEYS, lHandle)
If lResult <> ERROR_SUCCESS Then GoTo fEnumKeyError

Do While Not bDone
    sKeyName = Space$(MAX_SIZE)
    lResult = RegEnumKey(lHandle, lIndex, sKeyName, MAX_SIZE)
    
    If lResult = ERROR_SUCCESS Then
        sValues = sValues & Trim$(sKeyName)
        lIndex = lIndex + 1
    Else
        bDone = True
    End If
Loop
sValues = sValues & vbNullChar
If Len(sValues) = 1 Then sValues = sValues & vbNullChar

fEnumKey = RegCloseKey(lHandle)
Exit Function

fEnumKeyError:
    fEnumKey = lResult
End Function
Public Function fEnumValue(ByVal sTopKeyOrIniFile As String, _
    ByVal sSubKeyOrSection As String, sValues As String) As Long

Dim lTopKey    As Long
Dim lHandle    As Long
Dim lResult    As Long
Dim lValueLen  As Long
Dim lIndex     As Long
Dim lValue     As Long
Dim lValueType As Long
Dim lData      As Long
Dim lDataLen   As Long
Dim bDone      As Boolean
Dim sValueName As String
Dim sValue     As String

On Error GoTo fEnumValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrIniFile)
If lTopKey = 0 Then GoTo fEnumValueError

If lTopKey = 1 Then
    
    sValues = Space$(MAX_INISIZE)
    lResult = GetPrivateProfileSection(sSubKeyOrSection, sValues, Len(sValues), sTopKeyOrIniFile)
Else
    
    lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_QUERY_VALUE, lHandle)
    If lResult <> ERROR_SUCCESS Then GoTo fEnumValueError
    
    Do While Not bDone
        lDataLen = MAX_SIZE
        lValueLen = lDataLen
        sValueName = Space$(lDataLen)
        
        lResult = RegEnumValue(lHandle, lIndex, sValueName, lValueLen, 0, lValueType, ByVal lData, lDataLen)
        If lResult = ERROR_SUCCESS Then
            Select Case lValueType
                Case REG_SZ
                    sValue = Space$(lDataLen)
                    sValueName = Left$(sValueName, lValueLen)
                    lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, ByVal sValue, lDataLen)
                    If lResult = ERROR_SUCCESS Then
                        sValues = sValues & sValueName & "=" & sValue
                    Else
                        GoTo fEnumValueError
                    End If
                Case REG_DWORD
                    lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_NONE, lValue, lDataLen)
                    If lResult = ERROR_SUCCESS Then
                        sValueName = Left$(sValueName, lValueLen)
                        sValues = sValues & sValueName & "=" & lValue & vbNullChar
                    Else
                        GoTo fEnumValueError
                    End If
                Case REG_BINARY
                    If lDataLen <= 2 Then
                        lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_NONE, lValue, lDataLen)
                        If lResult = ERROR_SUCCESS Then
                            sValueName = Left$(sValueName, lValueLen)
                            sValues = sValues & sValueName & "=" & lValue & vbNullChar
                        Else
                            GoTo fEnumValueError
                        End If
                    End If
                Case Else
            End Select
            lIndex = lIndex + 1
        Else
            bDone = True
        End If
    Loop
    sValues = sValues & vbNullChar
    If Len(sValues) = 1 Then sValues = sValues & vbNullChar
   
    lResult = RegCloseKey(lHandle)
    fEnumValue = lResult
End If
Exit Function

fEnumValueError:
    fEnumValue = lResult
End Function






Public Function fReadIniFuzzy(ByVal sIniFile As String, _
    sSection As String, ByVal sIniEntry As String, _
    ByVal sDefault As String, sValue As String) As Long

Dim sNextChar    As String
Dim sLine        As String
Dim sEntry       As String
Dim sSectionName As String
Dim iLen         As Integer
Dim iLocOfEq     As Integer
Dim iFnum        As Integer
Dim bDone        As Boolean
Dim bFound       As Boolean
Dim bNewSection  As Boolean

On Error GoTo fReadIniFuzzyError
fReadIniFuzzy = 99
bDone = False
sValue = sDefault
sEntry = UCase$(sIniEntry)
sSection = UCase$(sSection)
iLen = Len(sSection)

iFnum = FreeFile
Open sIniFile For Input Access Read As iFnum

Line Input #iFnum, sLine
Do While Not EOF(iFnum) And Not bDone
    sLine = UCase$(Trim$(sLine))
    bNewSection = False
    
    If Left$(sLine, 1) = "[" Then
        
        sSectionName = sLine
        Dim iPos As Integer
        iPos = InStr(1, sLine, sSection)
        If iPos > 0 Then
           
            sNextChar = Mid$(sLine, iPos + iLen, 1)
            If sNextChar = " " Or sNextChar = "]" Then
                
                Line Input #iFnum, sLine
                bFound = False
                bNewSection = False
                Do While Not EOF(iFnum) And Not bFound
                    
                    sLine = UCase$(Trim$(sLine))
                    If Left$(sLine, 1) = "[" Then
                        bNewSection = True
                        Exit Do
                    End If
                    
                    If InStr(1, sLine, sEntry) = 1 Then
                        
                        iLocOfEq = InStr(1, sLine, "=")
                        If iLocOfEq <> 0 Then
                            sValue = Mid$(sLine, iLocOfEq + 1)
                            sSection = Mid$(sSectionName, 2, InStr(1, sSectionName, "]") - 2)
                            bFound = True
                            bDone = True
                            fReadIniFuzzy = 0
                        End If
                    End If
                    If Not bFound Then
                        Line Input #iFnum, sLine
                    End If
                Loop
                If EOF(iFnum) Then bDone = True
                sSection = Mid$(sSectionName, 2, InStr(1, sSectionName, "]") - 2)
            End If
        End If
    End If
    If Not bNewSection And Not bDone Then
        Line Input #iFnum, sLine
    End If
Loop
Close iFnum
Exit Function

fReadIniFuzzyError:
    fReadIniFuzzy = 99
End Function
Public Function fReadValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String, _
    ByVal sValueType As String, ByVal vDefault As Variant, _
    vValue As Variant) As Long

Dim lTopKey     As Long
Dim lHandle     As Long
Dim lLenData    As Long
Dim lResult     As Long
Dim lDefault    As Long
Dim lValue      As Long
Dim sValue      As String
Dim sSubKeyPath As String
Dim sDefaultStr As String
Dim bValue      As Boolean

On Error GoTo fReadValueError
lResult = 99
vValue = vDefault
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fReadValueError

If lTopKey = 1 Then
    '
    ' Read the .ini file value.
    '
    If UCase$(sValueType) = "S" Then
        lLenData = 255
        sDefaultStr = vDefault
        sValue = Space$(lLenData)
        lResult = GetPrivateProfileString(sSubKeyOrSection, sValueName, sDefaultStr, sValue, lLenData, sTopKeyOrFile)
        vValue = Left$(sValue, lResult)
    Else
        lDefault = 0
        lResult = GetPrivateProfileInt(sSubKeyOrSection, sValueName, lDefault, sTopKeyOrFile)
    End If
Else
    
    lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_QUERY_VALUE, lHandle)
    If lResult <> ERROR_SUCCESS Then
        fReadValue = lResult
        Exit Function
    End If
    
    Select Case UCase$(sValueType)
        Case "S"
            
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, "", lLenData)
            If lResult = ERROR_MORE_DATA Then
                sValue = Space(lLenData)
                lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, ByVal sValue, lLenData)
            End If
            If lResult = ERROR_SUCCESS Then  'Remove null character.
                vValue = Left$(sValue, lLenData - 1)
            Else
                GoTo fReadValueError
            End If
        Case "B"
            lLenData = Len(bValue)
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_BINARY, bValue, lLenData)
            If lResult = ERROR_SUCCESS Then
                vValue = bValue
            Else
                GoTo fReadValueError
            End If
        Case "D" 'короче, здесь какая то ЕБАНАЯ ХУЙНЯ, полный разброд в данных. надо разобрать
            lLenData = 32
            lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_DWORD, lValue, lLenData)
            If lResult = ERROR_SUCCESS Then
                vValue = lValue
            Else
                GoTo fReadValueError
            End If
    End Select
    
    lResult = RegCloseKey(lHandle)
    fReadValue = lResult
End If
Exit Function

fReadValueError:
    fReadValue = lResult
End Function


Private Function fTopKey(ByVal sTopKeyOrFile As String) As Long
Dim sDir   As String

On Error GoTo fTopKeyError
fTopKey = 0
Select Case UCase$(sTopKeyOrFile)
    Case "HKCU"
        fTopKey = HKEY_CURRENT_USER
    Case "HKLM"
        fTopKey = HKEY_LOCAL_MACHINE
    Case "HKU"
        fTopKey = HKEY_USERS
    Case "HKDD"
        fTopKey = HKEY_DYN_DATA
    Case "HKCC"
        fTopKey = HKEY_CURRENT_CONFIG
    Case "HKCR"
        fTopKey = HKEY_CLASSES_ROOT
    Case Else
        On Error Resume Next
        sDir = Dir$(sTopKeyOrFile)
        If Err.Number = 0 And sDir <> "" Then fTopKey = 1
End Select
Exit Function

fTopKeyError:
End Function

Public Function fWriteValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String, _
    ByVal sValueType As String, ByVal vValue As Variant) As Long

Dim hKey                As Long
Dim lTopKey             As Long
Dim lOptions            As Long
Dim lsamDesired         As Long
Dim lHandle             As Long
Dim lDisposition        As Long
Dim lLenData            As Long
Dim lResult             As Long
Dim lValue              As Long
Dim sClass              As String
Dim sValue              As String
Dim sSubKeyPath         As String
Dim bValue              As Boolean
Dim tSecurityAttributes As SECURITY_ATTRIBUTES

On Error GoTo fWriteValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fWriteValueError

If lTopKey = 1 Then
    
    If UCase$(sValueType) = "S" Then
        sValue = vValue
        lResult = WritePrivateProfileString(sSubKeyOrSection, sValueName, sValue, sTopKeyOrFile)
    Else
        GoTo fWriteValueError
    End If
Else
    sClass = ""
    lOptions = REG_OPTION_NON_VOLATILE
    lsamDesired = KEY_CREATE_SUB_KEY Or KEY_SET_VALUE
    
    lResult = RegCreateKeyEx(lTopKey, sSubKeyOrSection, 0, sClass, lOptions, _
                  lsamDesired, tSecurityAttributes, lHandle, lDisposition)
    If lResult <> ERROR_SUCCESS Then GoTo fWriteValueError
    
    Select Case UCase$(sValueType)
        Case "S"
            sValue = vValue
            lLenData = Len(sValue)
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_SZ, ByVal sValue, lLenData)
        Case "B"
            bValue = vValue
            lLenData = Len(bValue)
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_BINARY, bValue, lLenData)
        Case "D"
            lValue = CInt(vValue)
            lLenData = 4
            lResult = RegSetValueEx(lHandle, sValueName, 0, REG_DWORD, lValue, lLenData)
    End Select
   
    If lResult = ERROR_SUCCESS Then
        lResult = RegCloseKey(lHandle)
        fWriteValue = lResult
        Exit Function
    End If
End If
Exit Function

fWriteValueError:
    fWriteValue = lResult
End Function
