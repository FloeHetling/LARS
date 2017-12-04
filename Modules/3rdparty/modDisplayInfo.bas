Attribute VB_Name = "modDisplayInfo"
'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: Vijay kumar Bhudala
'
' AUTHOR: FTT User ,
' DATE  : 11/17/2009
'
' COMMENT: Синтаксис переработан для исполнения кода в ПО ЛАРС
'
'==========================================================================
 
Option Explicit

Public Function GetMonitorInfo() As String
Dim strComputer, message
 
Dim intMonitorCount
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3
Dim sValue
Dim i, iRC, iRC2, iRC3
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintEDID
Dim strRawEDID
Dim ByteValue, strSerFind, strMdlFind
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmp, tmpser, tmpmdl, tmpctr
Dim batch, bHeader
batch = True
 

strComputer = HostName
strComputer = UCase(strComputer)

 
Dim strarrRawEDID()
intMonitorCount = 0
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'get a handle to the WMI registry object
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "/root/default:StdRegProv")
 
If Err <> 0 Then
    If batch Then
        EchoAndLog strComputer & ",,,,," & Err.description
    Else
        MsgBox "Failed. " & Err.description, vbCritical + vbOKOnly, strComputer
        Exit Function
    End If
End If
 
 
sBaseKey = "SYSTEM\CurrentControlSet\Enum\DISPLAY\"
'enumerate all the keys HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\
iRC = oRegistry.EnumKey(HKLM, sBaseKey, arSubKeys)
For Each skey In arSubKeys
     'we are now in the registry at the level of:
     'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\
     'we need to dive in one more level and check the data of the "HardwareID" value
    sBaseKey2 = sBaseKey & skey & "\"
    iRC2 = oRegistry.EnumKey(HKLM, sBaseKey2, arSubKeys2)
    For Each skey2 In arSubKeys2
          'now we are at the level of:
          'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\<PNP_ID>\
          'so we can check the "HardwareID" value
        oRegistry.GetMultiStringValue HKLM, sBaseKey2 & skey2 & "\", "HardwareID", sValue
        For tmpctr = 0 To UBound(sValue)
            If LCase(Left(sValue(tmpctr), 8)) = "monitor\" Then
                    'If it is a monitor we will check for the existance of a control subkey
                    'that way we know it is an active monitor
                    sBaseKey3 = sBaseKey2 & skey2 & "\"
                    iRC3 = oRegistry.EnumKey(HKLM, sBaseKey3, arSubKeys3)
                    For Each skey3 In arSubKeys3
                    'Kaplan edit
                    strRawEDID = ""
                         If skey3 = "Control" Then
                              'If the Control sub-key exists then we should read the edid info
                              oRegistry.GetBinaryValue HKLM, sBaseKey3 & "Device Parameters\", "EDID", arrintEDID
                           If varType(arrintEDID) <> 8204 Then 'and If we don't find it...
                                   strRawEDID = "EDID Not Available" 'store an "unavailable message
                              Else
                                   For Each ByteValue In arrintEDID 'otherwise conver the byte array from the registry into a string (for easier processing later)
                                        strRawEDID = strRawEDID & Chr(ByteValue)
                                   Next
                              End If
                              'now take the string and store it in an array, that way we can support multiple monitors
                              ReDim Preserve strarrRawEDID(intMonitorCount)
                              strarrRawEDID(intMonitorCount) = strRawEDID
                              intMonitorCount = intMonitorCount + 1
                          End If
                    Next
            End If
        Next
    Next
Next
'*****************************************************************************************
'now the EDID info for each active monitor is stored in an array of strings called strarrRawEDID
'so we can process it to get the good stuff out of it which we will store in a 5 dimensional array
'called arrMonitorInfo, the dimensions are as follows:
'0=VESA Mfg ID, 1=VESA Device ID, 2=MFG Date (M/YYYY),3=Serial Num (If available),4=Model Descriptor
'5=EDID Version
'*****************************************************************************************
On Error Resume Next
Dim arrMonitorInfo()
ReDim arrMonitorInfo(intMonitorCount - 1, 5)
Dim location(3)
For tmpctr = 0 To intMonitorCount - 1
     If strarrRawEDID(tmpctr) <> "EDID Not Available" Then
          '*********************************************************************
          'first get the model and serial numbers from the vesa descriptor
          'blocks in the edid.  the model number is required to be present
          'according to the spec. (v1.2 and beyond)but serial number is not
          'required.  There are 4 descriptor blocks in edid at offset locations
          '&H36 &H48 &H5a and &H6c each block is 18 bytes long
          '*********************************************************************
          location(0) = Mid(strarrRawEDID(tmpctr), &H36 + 1, 18)
          location(1) = Mid(strarrRawEDID(tmpctr), &H48 + 1, 18)
          location(2) = Mid(strarrRawEDID(tmpctr), &H5A + 1, 18)
          location(3) = Mid(strarrRawEDID(tmpctr), &H6C + 1, 18)
     
          'you can tell If the location contains a serial number If it starts with &H00 00 00 ff
          strSerFind = Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&HFF)
          'or a model description If it starts with &H00 00 00 fc
          strMdlFind = Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&HFC)
     
          intSerFoundAt = -1
          intMdlFoundAt = -1
          For findit = 0 To 3
               If InStr(location(findit), strSerFind) > 0 Then
                    intSerFoundAt = findit
               End If
               If InStr(location(findit), strMdlFind) > 0 Then
                    intMdlFoundAt = findit
               End If
          Next
           
          'If a location containing a serial number block was found then store it
          If intSerFoundAt <> -1 Then
               tmp = Right(location(intSerFoundAt), 14)
               If InStr(tmp, Chr(&HA)) > 0 Then
                    tmpser = Trim(Left(tmp, InStr(tmp, Chr(&HA)) - 1))
               Else
                    tmpser = Trim(tmp)
               End If
               'although it is not part of the edid spec it seems as though the
               'serial number will frequently be preceeded by &H00, this
               'compensates for that
               If Left(tmpser, 1) = Chr(0) Then tmpser = Right(tmpser, Len(tmpser) - 1)
          Else
               tmpser = "Not Found"
          End If
      
          'If a location containing a model number block was found then store it
          If intMdlFoundAt <> -1 Then
               tmp = Right(location(intMdlFoundAt), 14)
               If InStr(tmp, Chr(&HA)) > 0 Then
                    tmpmdl = Trim(Left(tmp, InStr(tmp, Chr(&HA)) - 1))
               Else
                    tmpmdl = Trim(tmp)
               End If
               'although it is not part of the edid spec it seems as though the
               'serial number will frequently be preceeded by &H00, this
               'compensates for that
               If Left(tmpmdl, 1) = Chr(0) Then tmpmdl = Right(tmpmdl, Len(tmpmdl) - 1)
          Else
               tmpmdl = "Not Found"
          End If
 
          '**************************************************************
          'Next get the mfg date
          '**************************************************************
          Dim tmpmfgweek, tmpmfgyear, tmpmdt
          'the week of manufacture is stored at EDID offset &H10
          tmpmfgweek = Asc(Mid(strarrRawEDID(tmpctr), &H10 + 1, 1))
           
          'the year of manufacture is stored at EDID offset &H11
          'and is the current year -1990
          tmpmfgyear = (Asc(Mid(strarrRawEDID(tmpctr), &H11 + 1, 1))) + 1990
           
          'store it in month/year format
          tmpmdt = Month(DateAdd("ww", tmpmfgweek, DateValue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear
           
          '**************************************************************
          'Next get the edid version
          '**************************************************************
          'the version is at EDID offset &H12
          Dim tmpEDIDMajorVer, tmpEDIDRev, tmpVer
          tmpEDIDMajorVer = Asc(Mid(strarrRawEDID(tmpctr), &H12 + 1, 1))
           
          'the revision level is at EDID offset &H13
          tmpEDIDRev = Asc(Mid(strarrRawEDID(tmpctr), &H13 + 1, 1))
           
          'store it in month/year format
          tmpVer = Chr(48 + tmpEDIDMajorVer) & "." & Chr(48 + tmpEDIDRev)
           
          '**************************************************************
          'Next get the mfg id
          '**************************************************************
          'the mfg id is 2 bytes starting at EDID offset &H08
          'the id is three characters long.  using 5 bits to represent
          'each character.  the bits are used so that 1=A 2=B etc..
          '
          'get the data
          Dim tmpEDIDMfg, tmpMfg
          Dim Char1, Char2, Char3
          Dim Byte1, Byte2
          tmpEDIDMfg = Mid(strarrRawEDID(tmpctr), &H8 + 1, 2)
          Char1 = 0: Char2 = 0: Char3 = 0
          Byte1 = Asc(Left(tmpEDIDMfg, 1)) 'get the first half of the string
          Byte2 = Asc(Right(tmpEDIDMfg, 1)) 'get the first half of the string
          'now shift the bits
          'shift the 64 bit to the 16 bit
          If (Byte1 And 64) > 0 Then Char1 = Char1 + 16
          'shift the 32 bit to the 8 bit
          If (Byte1 And 32) > 0 Then Char1 = Char1 + 8
          'etc....
          If (Byte1 And 16) > 0 Then Char1 = Char1 + 4
          If (Byte1 And 8) > 0 Then Char1 = Char1 + 2
          If (Byte1 And 4) > 0 Then Char1 = Char1 + 1
 
          'the 2nd character uses the 2 bit and the 1 bit of the 1st byte
          If (Byte1 And 2) > 0 Then Char2 = Char2 + 16
          If (Byte1 And 1) > 0 Then Char2 = Char2 + 8
          'and the 128,64 and 32 bits of the 2nd byte
          If (Byte2 And 128) > 0 Then Char2 = Char2 + 4
          If (Byte2 And 64) > 0 Then Char2 = Char2 + 2
          If (Byte2 And 32) > 0 Then Char2 = Char2 + 1
 
          'the bits for the 3rd character don't need shifting
          'we can use them as they are
          Char3 = Char3 + (Byte2 And 16)
          Char3 = Char3 + (Byte2 And 8)
          Char3 = Char3 + (Byte2 And 4)
          Char3 = Char3 + (Byte2 And 2)
          Char3 = Char3 + (Byte2 And 1)
          tmpMfg = Chr(Char1 + 64) & Chr(Char2 + 64) & Chr(Char3 + 64)
           
          '**************************************************************
          'Next get the device id
          '**************************************************************
          'the device id is 2bytes starting at EDID offset &H0a
          'the bytes are in reverse order.
          'this code is not text.  it is just a 2 byte code assigned
          'by the manufacturer.  they should be unique to a model
          Dim tmpEDIDDev1, tmpEDIDDev2, tmpDev
           
          tmpEDIDDev1 = Hex(Asc(Mid(strarrRawEDID(tmpctr), &HA + 1, 1)))
          tmpEDIDDev2 = Hex(Asc(Mid(strarrRawEDID(tmpctr), &HB + 1, 1)))
          If Len(tmpEDIDDev1) = 1 Then tmpEDIDDev1 = "0" & tmpEDIDDev1
          If Len(tmpEDIDDev2) = 1 Then tmpEDIDDev2 = "0" & tmpEDIDDev2
          tmpDev = tmpEDIDDev2 & tmpEDIDDev1
           
          '**************************************************************
          'finally store all the values into the array
          '**************************************************************
          'Kaplan adds code to avoid duplication...
           
          If Not InArray(tmpser, arrMonitorInfo, 3) Then
              arrMonitorInfo(tmpctr, 0) = tmpMfg
              arrMonitorInfo(tmpctr, 1) = tmpDev
              arrMonitorInfo(tmpctr, 2) = tmpmdt
              arrMonitorInfo(tmpctr, 3) = tmpser
              arrMonitorInfo(tmpctr, 4) = tmpmdl
              arrMonitorInfo(tmpctr, 5) = tmpVer
          End If
     End If
Next
 
'For now just a simple screen print will suffice for output.
'But you could take this output and write it to a database or a file
'and in that way use it for asset management.
i = 0
For tmpctr = 0 To intMonitorCount - 1
     If arrMonitorInfo(tmpctr, 1) <> "" And arrMonitorInfo(tmpctr, 0) <> "PNP" Then
'         If batch Then
'             EchoAndLog strComputer & "," & arrMonitorInfo(tmpctr, 4) & "," & _
'             arrMonitorInfo(tmpctr, 3) & "," & arrMonitorInfo(tmpctr, 0) & "," & _
'             arrMonitorInfo(tmpctr, 2)
'          Else
             message = message & "Монитор " & Chr(i + 65) & ")" & vbCrLf & _
             "Модель: " & arrMonitorInfo(tmpctr, 4) & vbCrLf & _
             "Серийный номер: " & arrMonitorInfo(tmpctr, 3) & vbCrLf & _
             "VESA ID: " & arrMonitorInfo(tmpctr, 0) & vbCrLf & _
             "Дата производства: " & arrMonitorInfo(tmpctr, 2) & vbCrLf & vbCrLf
             'WriteToLog".........." & "Device ID: " & arrMonitorInfo(tmpctr,1)
             'WriteToLog".........." & "EDID Version: " & arrMonitorInfo(tmpctr,5)
                 i = i + 1
'         End If
     End If
Next
 
'If Not batch Then
'    MsgBox message, vbInformation + vbOKOnly, strComputer & " Monitor Info"
'End If
GetMonitorInfo = message
End Function
Function InArray(strValue, List, Col)
    Dim i
    For i = 0 To UBound(List)
        If List(i, Col) = CStr(strValue) Then
            InArray = True
            Exit Function
        End If
    Next
    InArray = False
End Function
 
Sub EchoAndLog(message)
'Echo output and write to log
    WriteToLog message
End Sub

