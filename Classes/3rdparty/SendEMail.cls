VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'********                                      CRYPTO CLS                                                  *******
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptDuplicateHash Lib "advapi32.dll" (ByVal hHash As Long, pdwReserved As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32.dll" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvReserved As Long, ByVal phCryptProv As Long, ByVal pdwKeySpec As Long, ByVal pfCallerFreeProv As Long) As Long



Private Const CAQ As String = "CryptAcquireContext"
Private Const CCH As String = "CryptCreateHash"
Private Const CDrK As String = "CryptDeriveKey"
Private Const CDuH As String = "CryptDuplicateHash"
Private Const CE As String = "CryptEncrypt"
Private Const CD As String = "CryptDecrypt"
Private Const CEK As String = "CryptExportKey"
Private Const CIK As String = "CryptImportKey"
Private Const CGHP As String = "CryptGetHashParam"
Private Const CGK As String = "CryptGenKey"
Private Const CGPP As String = "CryptGetProvParam"
Private Const CGUK As String = "CryptGetUserKey"
Private Const CHD As String = "CryptHashData"
Private Const CSHP As String = "CryptSetHashParam"
Private Const CSKP As String = "CryptSetKeyParam"
Private Const CSH As String = "CryptSignHash"

Private Const abClientFinished = "client finished"
Private Const abServerFinished = "server finished"

'================================
'EVENTS
'================================
Public Event CryptError(ByVal Number As Long, Description As String, ByVal Source As String)

' exported key blob definitions
Private Const SIMPLEBLOB As Byte = &H1
Private Const CUR_BLOB_VERSION As Byte = &H2
Private Const PUBLICKEYBLOB As Byte = &H6
Private Const T_PUBLICKEYBLOBLEN = 276
Private Const T_PRIVATEKEYBLOBLEN = 1172
Private Const T_SIMPLEKEYBLOBLEN = 268

Private Type T_PUBLICKEYBLOB
    bType    As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 256) As Byte
End Type

Private Type T_EXP_PUBLICKEYBLOB
    bPublicKey(1 To T_PUBLICKEYBLOBLEN) As Byte
End Type

Private Type T_PRIVATEKEYBLOB
    bType    As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 256)  As Byte
    prime1(1 To 128)   As Byte
    prime2(1 To 128)   As Byte
    exponent1(1 To 128)       As Byte
    exponent2(1 To 128)       As Byte
    coefficient(1 To 128)     As Byte
    privateExponent(1 To 256) As Byte
End Type

Private Type T_SIMPLEKEYBLOB
    bType    As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
    Algid    As Long
    modulus(1 To 256) As Byte
End Type

Private Type T_EXP_SIMPLEKEYBLOB
    bSimpleKey(1 To T_SIMPLEKEYBLOBLEN) As Byte
End Type

Private Type SCHANNEL_ALG
    dwUse           As Long  '4 bytes
    Algid           As Long  '4 bytes
    cBits           As Long  '4 bytes
    dwFlags         As Long  '4 bytes
    dwReserved      As Long  '4 bytes
End Type                     'Total 20 bytes

Private Type CRYPT_DATA_BLOB
    cbData As Long          'The count of bytes in the buffer pointed to by pbData.
    pbData As Long          'A pointer to a block of data bytes.
End Type

Private Type HMAC_INFO
    ALG_ID As Long          'HashAlgid
    pbIStr(1 To 64) As Byte '*pbInnerString (BYTE*)0x36
    cbIStr As Long          'len(cbInnerString)
    pbOStr(1 To 64) As Byte '*pbOuterString (BYTE*)0x5C
    cbOStr As Long          'len(cbOuterString)
End Type

Private Type SHAMD5_HASH
    bHash(1 To 36) As Byte
End Type

'Class Variables
Private sContainer As String
Private sProvider As String
Private sProvType As Long
Private PublicKey As T_EXP_PUBLICKEYBLOB
Private NewPubKey As T_PUBLICKEYBLOB
Private fClient As Boolean
Private bClientRandom() As Byte
Private bServerRandom() As Byte
Private bPreMasterKey() As Byte
Private Alg_Hash As Long


'API memory functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const MS_DEF_RSA_SCHANNEL_PROV As String = "Microsoft RSA SChannel Cryptographic Provider"
                                                  
Private Const PROV_RSA_SCHANNEL As Long = 12
Private Const CRYPT_NEWKEYSET = &H8
Private Const PP_CONTAINER = 6
Private Const AT_KEYEXCHANGE = 1
Private Const PP_ENUMALGS As Long = 1

'Algorithm Classes
Private Const ALG_CLASS_ANY              As Long = 0
Private Const ALG_CLASS_SIGNATURE        As Long = (1 * 2 ^ 13) '8192
Private Const ALG_CLASS_MSG_ENCRYPT      As Long = (2 * 2 ^ 13) '16384
Private Const ALG_CLASS_DATA_ENCRYPT     As Long = (3 * 2 ^ 13) '24576
Private Const ALG_CLASS_HASH             As Long = (4 * 2 ^ 13) '32768
Private Const ALG_CLASS_KEY_EXCHANGE     As Long = (5 * 2 ^ 13) '40960

Private Const ALG_TYPE_ANY              As Long = 0
Private Const ALG_TYPE_RSA              As Long = (2 * 2 ^ 9) '1024
Private Const ALG_TYPE_BLOCK            As Long = (3 * 2 ^ 9) '1536
Private Const ALG_TYPE_STREAM           As Long = (4 * 2 ^ 9) '2048
Private Const ALG_TYPE_SECURECHANNEL    As Long = (6 * 2 ^ 9) '3072

Private Const ALG_SID_RC2               As Long = 2
Private Const ALG_SID_RC4               As Long = 1
Private Const ALG_SID_MD2               As Long = 1
Private Const ALG_SID_MD4               As Long = 2
Private Const ALG_SID_MD5               As Long = 3
Private Const ALG_SID_SHA1              As Long = 4
Private Const ALG_SID_SCHANNEL_MASTER_HASH As Long = 2
Private Const ALG_SID_SCHANNEL_MAC_KEY  As Long = 3
Private Const ALG_SID_TLS1_MASTER       As Long = 6
Private Const ALG_SID_SCHANNEL_ENC_KEY  As Long = 7
Private Const ALG_SID_SSL3SHAMD5        As Long = 8
Private Const ALG_SID_HMAC              As Long = 9
Private Const ALG_SID_TLS1PRF           As Long = 10
Private Const ALG_SID_RSA_ANY           As Long = 0

Private Const CALG_SCHANNEL_MASTER_HASH = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MASTER_HASH) '19458
Private Const CALG_SCHANNEL_MAC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MAC_KEY) '19459
Private Const CALG_TLS1_MASTER = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_TLS1_MASTER) '19462
Private Const CALG_SCHANNEL_ENC_KEY = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_ENC_KEY) '19463
Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5) '32771
Private Const CALG_SHA1 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1) '32772
Private Const CALG_TLS1PRF = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_TLS1PRF) '32778
Private Const CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2) '26114
Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4) '26625
Private Const CALG_RSA_KEYX = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY) '41984

Private Const SCHANNEL_MAC_KEY As Long = 0
Private Const SCHANNEL_ENC_KEY As Long = 1
Private Const HP_HASHVAL As Long = 2 'Hash value
Private Const HP_TLS1PRF_LABEL As Long = 6
Private Const HP_TLS1PRF_SEED As Long = 7
Private Const CRYPT_SERVER = 1024 '0x00000400

'constants from WinErr.h
Private Const NTE_NO_KEY As Long = &H8009000D              ' Key does not exist

'clsCryptoFilterBox constants
Private Const CFB_BUSY = 0
Private Const CFB_READY = 1

Private Const CRYPT_EXPORTABLE = 1

'KP Constants
Private Const KP_SCHANNEL_ALG = 20
Private Const KP_CLIENT_RANDOM = 21 'for setting the Secure Channel client random data
Private Const KP_SERVER_RANDOM = 22 'for setting the Secure Channel server random data

'private property buffers
Private lCipher As Long
Private sInBuffer As String
Private sOutBuffer As String
Private sRndBuffer As String
Private sPublicKey As String
Private sPassword As String
Private sSignature As String
Private lStatus As Long
Private sMsgType As Byte
Private sMsgVerMaj As Byte
Private sMsgVerMin As Byte

'TLS Key Handles
Private hCryptProv As Long
Private hMasterKey As Long
Private hReadKey As Long
Private hWriteKey As Long
Private hReadMAC As Long
Private hWriteMAC As Long
Private WRITE_MAC_KEY As String
Private READ_MAC_KEY As String
Private WRITE_KEY As String
Private READ_KEY As String

'Counters
Private SEND_SEQ_NUM(2) As Long
Private RECV_SEQ_NUM(2) As Long

Private flgFinal As Boolean

Private hSESSION_MD5 As Long 'Handle to MD5 Session Hash
Private hSESSION_SHA1 As Long 'Handle to SHA1 Session Hash
Private hCOPY_MD5 As Long 'Handle to Copy of MD5 Session Hash
Private hCOPY_SHA1 As Long 'Handle to Copy of SHA1 Session Hash
Private CLIENT_RANDOM As String
Private SERVER_RANDOM As String
Private PRE_MASTER_KEY As String
Private PUBLIC_KEY As String
Private RSA_KEY_LEN As Long
Private RSA_KEY_256V3 As String
Private RSA_KEY_128V3 As String

Private SNI As String

'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'********                                      SENDER CLS                                                  *******
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
Private Const cBoundaryA As String = "Boundary_A_3435FE2_6617A_AA"
Private Const cBoundaryB As String = "Boundary_B_3435FE2_6617B_BB"
Private Const cAST As String = "*"
Private Const cDASH2 As String = "--"
Private Const cDASH As String = "-"
Private Const cDOT As String = "."
Private Const cSEMIC As String = ";"
Private Const LEN8192 As Long = 8192
Private Const cSRC = "SRC="
Private Const R_BRACKET = ">"
Private Const cSPACE = " "
Private Const cEQUAL = "="


Private Const INTERNET_CONNECTION_PROXY As Long = &H4

Private Const INTERNET_RFC1123_FORMAT = 0&
Private Const INTERNET_RFC1123_BUFSIZE = 30&

Private Const INADDR_NONE = &HFFFFFFFF
Private Const WSAEWOULDBLOCK = 10035
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128

Private Const WSABASEERR = 10000

Private Const WSASYSNOTREADY = (WSABASEERR + 91)
Private Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Private Const WSAEFAULT = (WSABASEERR + 14)
Private Const WSAEPROCLIM = (WSABASEERR + 67)
Private Const WSAEINPROGRESS = (WSABASEERR + 36)

Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_CONNECT = &H10&

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Const AF_INET = 2
Private Const SOCK_STREAM = 1

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_QUERY_VALUE As Long = &H1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (pGuid As guid, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSAAsyncSelect Lib "WSOCK32.DLL" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WCreateSocket Lib "WSOCK32.DLL" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function WConnect Lib "WSOCK32.DLL" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function WSend Lib "WSOCK32.DLL" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function WCloseSocket Lib "WSOCK32.DLL" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function ioctlsocket Lib "WSOCK32.DLL" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long

Private Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer
Private Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function recv Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function recvfrom Lib "WSOCK32.DLL" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long
Private Declare Function getpeername Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef PearName As sockaddr, ByRef namelen As Long) As Long

'These API's are added to sustain unicode filenames
Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, _
ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, _
ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
'*************************************************

Private Const FILE_SHARE_READ = &H1
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4


Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetTimeFromSystemTime Lib "wininet.dll" (ByRef pst As SYSTEMTIME, ByVal dwRFC As Long, ByVal lpszTime As String, ByVal cbTime As Long) As Long

Public Event Connected()
Public Event OnSendComplete(Success As Boolean, sErr As String)
Public Event DisConnected()

Private Qt2 As String
Private cNULL As String

Private HSock As Long  '--socket handle. lngsockethandle
Private SockValue As Long '--  will be 100 if winsock started.
Private BooHooked As Boolean
Private BooPlain As Boolean
Private sOutBuf As String
Private sRecip As String
Private sInBuf As String
Private sFromEmail As String
Private sServer As String
Private sFromName As String
Private sUserName As String
Private sUserName64 As String
Private sAttach As String
Private sPassword64 As String
Private sPortNumber As Long
Private sPLAINLogon64 As String
Private sToEmail As String
Private sBodyText As String
Private sSubjectLine As String
Private iErrCode As Long
Private m_lngRecvBufferLen  As Long
Private flgTLS As Boolean
Private txtOut As String
Private Const VERSION_MAJOR As Byte = &H3
Private Const VERSION_MINOR_1 As Byte = &H1
Private Const TLS1_RT_CHANGE_CIPHER_SPEC As Byte = &H14 '20
Private Const TLS1_RT_ALERT As Byte = &H15              '21
Private Const TLS1_RT_HANDSHAKE As Byte = &H16          '22
Private Const TLS1_RT_APPLICATION_DATA As Byte = &H17   '23
Private Const TLS1_MT_HELLO_REQUEST As Byte = &H0       '0
Private Const TLS1_MT_SERVER_HELLO As Byte = &H2        '2
Private Const TLS1_MT_CERTIFICATE As Byte = &HB         '11
Private Const TLS1_MT_SERVER_DONE As Byte = &HE         '14
Private Const TLS1_MT_CLIENT_KEY_EXCHANGE As Byte = &H10 '16
Private OutArray(0 To 5)    As Variant  'save sent data
Public SysBusy As Boolean

Private Enum LCur_State               'The smtp states enum
    Server_CONNECT      '1
    Server_CONNECTED    '2
    Server_EHLO         '3
    Server_AUTH         '4
    Server_LOGIN        '5
    Server_PLAIN        '6
    Server_BEGINSEND    '7
    Server_FROM         '8
    Server_TO           '9
    Server_DATA         '10
    Server_END          '11
    Server_QUIT         '12
End Enum
Private Layer As Long
Private SESSION_ID As String

Private LCurState           As LCur_State   'the smtp states
Private MsgID               As String   'Message-ID from Date/Time
Private UTCDate             As String   'Coordinated Universal Time & date
Private FileAttach()        As String   'Buffer for file names
Private Recipient()         As String   'buffer for recipient names
Private FileType()          As String   'Buffer for Content-Type
Private Bas64buffer()       As String   'buffer for attachmentheader and base 64 encoded attachments

'******************************************************************************************************
'******************************************************************************************************
'***********                     SUBCLASSING                                            ***************
'******************************************************************************************************
'******************************************************************************************************
'------------------------------------------------------------------------------------------------------
'SUBCLASSING:   The WSAAsyncSelect api that controls the callback notification events of the WinSock
'               needs to look at the Mainform.hWnd for its information. But because my server is build
'               as a class module, you can't get to the Mainform.hWnd and its notification events.
'               To do this, you need to subclass to a module of the Mainform.
'------------------------------------------------------------------------------------------------------

Public HWndToMsg As Long            'Holds address of the MainForm.hWnd
Private hWndProcOrg As Long         'Holds original address of the Window Procedure for this window
Private hWndSubClassed As Long      'Holds address of the subclassed Window Procedure



Private m_blnEncr               As Boolean      'True for secure sockets
Private m_strRecvBuffer As String   'local outgoing buffer
Private Const sckUnsupported = 40018
Private Const SOCKET_ERROR     As Long = -1
Private Const INVALID_SOCKET   As Long = SOCKET_ERROR

Private Sub SubClass()

    If hWndProcOrg Then
        Exit Sub
    End If
    'hWndProcOrg = SetWindowLong(HWndToMsg, GWL_WNDPROC, AddressOf SubWndProc)
    hWndSubClassed = HWndToMsg
    'Call SetWindowLong(HWndToMsg, GWL_USERDATA, ObjPtr(Me))
End Sub

Private Sub UnSubClass()
    
    If Not hWndProcOrg Then Exit Sub
    'SetWindowLong hWndSubClassed, GWL_WNDPROC, hWndProcOrg
    hWndProcOrg = 0
End Sub

'************************************************************************************************************************
'************************************************************************************************************************
'*******                                                                                                          *******
'*******                                        SENDER CLS                                                        *******
'*******                                                                                                          *******
'************************************************************************************************************************
'************************************************************************************************************************

'******************************
'******************************
'******    PROPERTIES    ******
'******************************
'******************************

Public Property Let EncrFlg(ByVal blnEncr As Boolean)
    m_blnEncr = blnEncr
End Property

Public Property Let Bodytext(val As String)
    sBodyText = Trim$(val)
End Property

Public Property Let Attach(val As String)
    sAttach = Trim$(val)
End Property

Public Property Let PortNumber(val As Long)
    sPortNumber = val
End Property

Public Property Let Recip(val As String)
    sRecip = Trim$(val)
End Property

Public Property Let Plaintext(IfPlain As Boolean)
    BooPlain = IfPlain
End Property

Public Property Let FromName(sFrom As String)
    sFromName = Trim$(sFrom)
End Property

Public Property Let FromEmail(sFrom As String)
    sFromEmail = Trim$(sFrom)
End Property

Public Property Let ToEmail(sTo As String)
    sToEmail = Trim$(sTo)
End Property

Public Property Let Subject(sSubj As String)
    sSubjectLine = Trim$(sSubj)
End Property

Public Property Let RSA_KEY_256(Stri As String)
    RSA_KEY_256V3 = Trim$(Stri)
End Property

Public Property Let RSA_KEY_128(Stri As String)
    RSA_KEY_128V3 = Trim$(Stri)
End Property

Public Property Let UserName(sUName As String)
    Dim s1 As String
   
    On Error Resume Next
    sUserName = Trim$(sUName)
    sUserName64 = ConvertToBase64(sUserName, False, False)
    If (Len(sPassword) > 0) And (Len(sPLAINLogon64) = 0) Then
        s1 = sUserName & cNULL & sUserName & cNULL & sPassword
        sPLAINLogon64 = ConvertToBase64(s1, False, False)
    End If
End Property

Public Property Let Password(sPWord As String)
    Dim s1 As String
   
    On Error Resume Next
    sPassword = Trim$(sPWord)
    sPassword64 = ConvertToBase64(sPassword, False, False)
    If (Len(sUserName) > 0) And (Len(sPLAINLogon64) = 0) Then
        s1 = sUserName & cNULL & sUserName & cNULL & sPassword
        sPLAINLogon64 = ConvertToBase64(s1, False, False)
    End If
End Property

Public Property Let Server(sServ1 As String)
    sServer = Trim$(sServ1)
End Property

Public Function IsOnline() As Boolean
    Dim lFlags As Long
    
    On Error Resume Next
    IsOnline = InternetGetConnectedState(lFlags, 0&)
End Function

Public Function IsConnectionProxy() As Boolean
    Dim lFlags As Long, lRet As Long
        
    On Error Resume Next
    lRet = InternetGetConnectedState(lFlags, 0&)

    IsConnectionProxy = lFlags And INTERNET_CONNECTION_PROXY
End Function

Public Function StartUp(sErr As String) As Long '--returns 100 if OK.
    Dim sRet As String
    
    iErrCode = 0
    SockValue = StartWinsock()
  
'******************************************************************************************************
'******************************************************************************************************
    Call SubClass 'eventueel hier nog error trappen
'******************************************************************************************************
'******************************************************************************************************
    If (iErrCode <> 0) Then
        sErr = LoadResString(iErrCode)
    End If
    StartUp = SockValue
End Function

Public Sub CloseDown() '-- close socket.
    Dim lRet As Long
   
    On Error Resume Next
   
'******************************************************************************************************
'******************************************************************************************************
    Call UnSubClass 'Here you can trap errors, if needed
'******************************************************************************************************
'******************************************************************************************************
   
    If (HSock > 0) Then
         DisConnect
    End If
    If (SockValue = 100) Then
        lRet = WSACleanup()
        SockValue = 0
    End If
End Sub

Public Sub ResetEmail()
    
    On Error Resume Next
    sInBuf = ""
    sOutBuf = ""
    iErrCode = 0
    DisConnect
End Sub

Public Function Connect(sErr As String) As Long
    
    Dim lRet As Long, LRet2 As Long, LEvents As Long, LPort587 As Long
    Dim iVal As Long
    Dim s1 As String
    Dim Boo1 As Boolean
    Dim sck As sockaddr
     
    Debug.Print "SEND.Connect"
    On Error Resume Next
    sErr = ""
    If (SockValue <> 100) Then
        Debug.Print "SOCK NOT STARTED !!!!! "
        iVal = iErrCode
        GoTo QuitError
    End If
    ResetEmail
    If (IsOnline = False) Then
        iVal = 106
        GoTo QuitError
    End If
    If (IsConnectionProxy = True) Then
        iVal = 107  '-- Internet connection is via proxy server. Unable to send email via proxy.
        GoTo QuitError
    End If
    If (Len(sServer) = 0) Then
        iVal = 108  ' Server address string missing.
        GoTo QuitError
    End If
    If (Len(sPassword) = 0) Then
        iVal = 109  ' Password string missing.   '--skip this. check password when sending if nec.
        GoTo QuitError
    End If
    If (Len(sFromEmail) = 0) Then
        iVal = 110  ' From email address missing.
        GoTo QuitError
    ElseIf (InStr(1, sFromEmail, "@") = 0) Then
        iVal = 111  ' Invalid From email address.
        GoTo QuitError
    End If
    If (Len(sToEmail) = 0) Then
        iVal = 112    ' To email address missing.
        GoTo QuitError
    ElseIf (InStr(1, sToEmail, "@") = 0) Then
        iVal = 113  ' Invalid To email address.
        GoTo QuitError
    End If
    If (Len(sBodyText) = 0) Then
        iVal = 114 ' Text of email is missing.
        GoTo QuitError
    End If
    If (Len(sUserName) = 0) Then
        lRet = InStr(1, sFromEmail, "@")
        sUserName = Left$(sFromEmail, (lRet - 1))
        sUserName64 = ConvertToBase64(sUserName, False, False)
        If (Len(sPassword) > 0) And (Len(sPLAINLogon64) = 0) Then
            s1 = sUserName & cNULL & sUserName & cNULL & sPassword
            sPLAINLogon64 = ConvertToBase64(s1, False, False)
        End If
    End If
    Err.Clear
    HSock = WCreateSocket(AF_INET, SOCK_STREAM, 0&)
        If (HSock = -1) Then
           iVal = 115
           GoTo QuitError   ' Winsock error: Failed to create socket.
        End If
    With sck
      .sin_family = AF_INET
      .sin_addr = ResolveServer(sServer)
      .sin_port = htons(sPortNumber)
      .sin_zero = String$(8, 0)
   End With
   If (sck.sin_addr = INADDR_NONE) Then
      lRet = WCloseSocket(HSock)
      HSock = 0
      iVal = 116  ' Failed to resolve server IP address.
      GoTo QuitError
   End If
   Err.Clear
   lRet = WConnect(HSock, sck, Len(sck))
   If (lRet = -1) Then
       lRet = WCloseSocket(HSock)
       HSock = 0
       iVal = 117  ' Failed to connect.
       GoTo QuitError
   End If
   LCurState = Server_CONNECT
   Err.Clear
   LEvents = FD_READ Or FD_WRITE Or FD_CONNECT ' Or FD_CLOSE Or FD_ACCEPT
'******************************************************************************************************
'******************************************************************************************************
   lRet = WSAAsyncSelect(HSock, HWndToMsg, ByVal 1025, ByVal LEvents)
'******************************************************************************************************
'******************************************************************************************************
   If (lRet = -1) Then
       lRet = WCloseSocket(HSock)
       HSock = 0
       iVal = 118   ' failed with WSAAsyncSelect function.
       GoTo QuitError
   End If
   BooHooked = True
   iErrCode = 0
   Connect = 100
   Debug.Print "Event CONNECTED is raised"
   RaiseEvent Connected
   Exit Function

QuitError:
   sErr = LoadResString(iVal)
   Debug.Print sErr
   If (iVal = 115) Or (iVal = 117) Or (iVal = 118) Then
       sErr = sErr & CStr(Err.LastDllError)
   End If
   Connect = iVal
   Debug.Print "CONNECT = " & iVal & " : " & sErr
   
End Function

Public Sub DisConnect()
    Dim lRet As Long
    
    On Error Resume Next
    If (HSock > 0) Then
        If (BooHooked = True) Then
'******************************************************************************************************
'******************************************************************************************************
            lRet = WSAAsyncSelect(HSock, 0&, 1025, 0&)
'******************************************************************************************************
'******************************************************************************************************
            BooHooked = False
        End If
        lRet = WCloseSocket(HSock)
        HSock = 0
        RaiseEvent DisConnected
    End If
   
End Sub

Private Function ResolveServer(sServ As String) As Long
    
    On Error Resume Next
    Dim lRet As Long, LRet2 As Long
    Dim HEnt As HOSTENT
    Dim LAddrList As Long
   
    lRet = inet_addr(ByVal sServ)  '--try converting numeric first.
        If (lRet = -1) Then
            LRet2 = gethostbyname(ByVal sServ)
            If (LRet2 > 0) Then
                CopyMemory HEnt, ByVal LRet2, Len(HEnt)
                CopyMemory LAddrList, ByVal HEnt.h_addr_list, 4
                CopyMemory lRet, ByVal LAddrList, HEnt.h_length
            Else
                lRet = INADDR_NONE
            End If
        End If
    
    ResolveServer = lRet
End Function

Private Function StartWinsock() As Long  '--return 100 if success.
 
    Dim lRet As Long
    Dim WData  As WSAData
     
    On Error Resume Next
    If (SockValue = 100) Then '--already started.
        StartWinsock = 100
        Exit Function
    End If
    lRet = WSAStartup(&H101, WData)
    If (lRet <> 0) Then
        Select Case lRet
              Case WSASYSNOTREADY
                 iErrCode = 101  ' "System not ready for network communication."
                 StartWinsock = 1  ' system not ready for network communication.
              Case WSAVERNOTSUPPORTED
                 iErrCode = 102  '"Winsock version 1.1 not supported."
                 StartWinsock = 2  ' version 1.1 not supported.
              Case WSAEINPROGRESS
                 iErrCode = 103  ' "A blocking winsock is already in use."
                 StartWinsock = 3  ' a blocking winsock already in use.
              Case WSAEPROCLIM
                 iErrCode = 104   ' "The limit on number of winsock tasks has been reached."
                 StartWinsock = 4  ' limit on number of winsock tasks has been reached.
              Case WSAEFAULT
                 iErrCode = 105 '  "Error with parameter of WSAStartup call."
                 StartWinsock = 5  '  WData is not a valid pointer.
        End Select
    Else
       StartWinsock = 100
    End If
End Function
 
Private Function SendData(sData As String) As Long   'Private Function SendData(sData As String) As Long

    Dim lRet As Long, LenDat As Long
    Dim AData() As Byte
      
    If InStr(1, sServer, "live") Then
        Sleep 100 'we need to slow down the process a little, to avoid fatal alert 20
    Else
        Sleep 60
    End If
    On Error Resume Next
    sOutBuf = sOutBuf & sData
    LenDat = Len(sOutBuf)
    If LenDat = 0 Then
        SendData = -1
        Exit Function
    End If
    If (LenDat > 0) Then
        AData() = StrConv(sOutBuf, vbFromUnicode)
        lRet = WSend(HSock, AData(0), LenDat, 0&)
        If (lRet = -1) Then
            If Not (Err.LastDllError = WSAEWOULDBLOCK) Then
                SendData = -1
                Exit Function
            End If
        Else
            If (LenDat > lRet) Then
                sOutBuf = Mid$(sOutBuf, lRet + 1)
            Else
                sOutBuf = ""
            End If
        End If
    End If
    SendData = 0
End Function

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim WSAEvent, LNum, WSAError, lRet, RemotePort As Long
    Dim sBuf, sTxt, s3 As String
    Dim ABuf() As Byte
    Dim newSa As sockaddr
    Dim Sin As sockaddr
    
    
    On Error Resume Next
    Select Case uMsg
        Case 1025
            WSAEvent = WSAGetSelectEvent(lParam)
            WSAError = WSAGetAsyncError(lParam)
            Select Case WSAEvent
            Case FD_CONNECT
                'Arrival of this message means that the connection initiated by the call
                'of the connect Winsock API function was successfully established.
                Debug.Print "FD_CONNECT " & HSock
                If LCurState <> Server_CONNECT Then
                    Exit Function
                End If
                'Get the connection local end-point parameters
                lRet = getpeername(HSock, newSa, LenB(newSa))
                LCurState = Server_CONNECTED
                ServerConnect
            Case FD_READ 'Incoming data
                'find a coupling between  lparam and ServerMessageProc(iRet As Integer, sRet As String)
                ReDim ABuf(0 To 8191) As Byte
                LNum = recv(HSock, ABuf(0), LEN8192, 0&)
                If (LNum > 0) Then   '-- 0-closed. -1-error.
                    sBuf = StrConv(ABuf(), vbUnicode)
                    sTxt = Left$(sBuf, LNum)    '-- get text of server response.
                    sInBuf = sTxt
                    m_strRecvBuffer = sInBuf
                    If Len(sBuf) > 0 Then
                        If m_blnEncr Then
                            Call EncrDataArrival(Len(m_strRecvBuffer), sTxt)
                        Else
                            Call DataArrival(Len(m_strRecvBuffer), sTxt)
                        End If
                    End If
                Else
                    DisConnect
                End If
            Case FD_WRITE  '--   &H2   - ready for writing.
                    If (Len(sOutBuf) > 0) Then
                       SendData sOutBuf
                    End If
                'This message means that the socket is in a write-able
                'state, that is, buffer for outgoing data of the transport
                'service is empty and ready to receive data to send through
                'the network.
                
            End Select
        Case Else
          '  WindowProc = CallWindowProc(hWndProcOrg, hWnd, uMsg, wParam, ByVal lParam)
    End Select
    
HandleError:
End Function

Private Sub ServerConnect()
    'Send Client Hello (Unencrypted)
    Dim strSend As String
    Screen.MousePointer = vbDefault
    hSESSION_MD5 = 0
    hSESSION_SHA1 = 0
    'Prepare middle 3 bytes of MAC Header
    Let msgType = TLS1_RT_HANDSHAKE
    Let msgVerMaj = VERSION_MAJOR
    Let msgVerMin = VERSION_MINOR_1
    If flgTLS Then
        strSend = GetRandomString
        hSESSION_MD5 = MD5_AddToHash(strSend, hSESSION_MD5)
        hSESSION_SHA1 = SHA1_AddToHash(strSend, hSESSION_SHA1)
        SendData AddRecHeader(strSend)
    End If
End Sub

Private Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function

Private Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Private Sub ExitConnection(ByVal sErr As String)
   Layer = 0
   DisConnect
   flgTLS = False
   Let EncrFlg = flgTLS
   If (LCurState <> Server_QUIT) Then
      LCurState = Server_CONNECT
      RaiseEvent OnSendComplete(False, sErr)
   Else
      LCurState = Server_CONNECT
      RaiseEvent OnSendComplete(True, "")
   End If
End Sub

Public Function ConvertToBase64(sPathOrString As String, IfFile As Boolean, AddReturns As Boolean) As String
  Static Enc() As Byte
  Dim B1() As Byte, B2() As Byte, B76() As Byte
  Dim nReadRetVal As Long, ret As Long, handle As Long, I As Long, i2 As Long, i3 As Long, LFil As Long, NumReturns As Long
  Dim FF2 As Integer
  Dim lOverLapped As OVERLAPPED
     
  On Error Resume Next
  
  If (Not val(Not Enc)) = 0 Then
    Enc = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
  End If
  
  If (IfFile = True) Then  '-- if converting picture file.
        handle = CreateFileW(StrPtr(sPathOrString), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, _
        FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM, 0)
        If handle <> -1 Then
            LFil = GetFileSize(handle, 0)
            ReDim B1(0 To (LFil - 1)) As Byte
            nReadRetVal = ReadFile(handle, B1(0), LFil, ret, lOverLapped)
            If nReadRetVal = 0 Then
                End
            End If
        End If
  Else   '-- converting a string.
     B1() = StrConv(sPathOrString, vbFromUnicode)
     LFil = UBound(B1) + 1  '-- added as correction 12-2-04
  End If
  
  
  ReDim Preserve B1(0 To ((LFil - 1) \ 3) * 3 + 2)
  ReDim Preserve B2(0 To (UBound(B1) \ 3) * 4 + 3)
  For I = 0 To UBound(B1) - 1 Step 3
    B2(i2) = Enc(B1(I) \ 4)
      i2 = i2 + 1
    B2(i2) = Enc((B1(I + 1) \ 16) Or (B1(I) And 3) * 16)
      i2 = i2 + 1
    B2(i2) = Enc((B1(I + 2) \ 64) Or (B1(I + 1) And 15) * 4)
      i2 = i2 + 1
    B2(i2) = Enc(B1(I + 2) And 63)
      i2 = i2 + 1
  Next I
    For I = 1 To I - LFil
       B2(UBound(B2) - I + 1) = 61
    Next I
    
   If (AddReturns = True) And (LFil > 76) Then
      '-- add returns every 76 characters before converting to string:
        NumReturns = ((UBound(B2) + 1) \ 76)
        LFil = (UBound(B2) + (NumReturns * 2)) '--make B76 B2 plus 2 spots for each vbcrlf.
         ReDim B76(0 To LFil) As Byte
          i2 = 0
          i3 = 0
        For I = 0 To UBound(B2)
           B76(i2) = B2(I)
            i2 = i2 + 1
            i3 = i3 + 1
           If (i3 = 76) And (i2 < (LFil - 2)) Then   '--extra check. make sure there are still
              B76(i2) = 13                      '-- 2 spots left for return if at end.
              B76(i2 + 1) = 10
              i2 = i2 + 2
              i3 = 0
           End If
        Next
     ConvertToBase64 = StrConv(B76, vbUnicode)
   Else
     ConvertToBase64 = StrConv(B2, vbUnicode)
   End If
   If handle <> -1 Then
        CloseHandle handle
   End If
End Function

Private Function UCFileExists(sFilPath As String) As Boolean
    On Error GoTo Hell
    If Dir(sFilPath, vbHidden Or vbArchive Or vbReadOnly Or vbSystem) <> "" Then
        UCFileExists = True
    Else
        UCFileExists = False
    End If
    Exit Function
Hell:
    UCFileExists = False
End Function

Private Sub Class_Initialize()
     'SENDER CLS
     On Error Resume Next
     Qt2 = Chr$(34)
     cNULL = Chr$(0)
     BooPlain = True  '-- default to plain text.
    
    'CRYPTO CLS
    sContainer = "JAC SSL Container" & vbNullChar
    sProvider = MS_DEF_RSA_SCHANNEL_PROV
    sProvType = PROV_RSA_SCHANNEL
End Sub

Private Sub Class_Terminate()
    Dim lRet As Long
   
    On Error Resume Next
    If (HSock > 0) Then
        DisConnect
    End If
    If (SockValue = 100) Then
        lRet = WSACleanup()
        SockValue = 0
    End If
    '******************* CRYPT CLS **********************
    CryptReleaseContext hCryptProv, 0
End Sub

Private Sub GetData(ByRef Data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    maxLen = Len(m_strRecvBuffer)
    Dim lngBytesRecibidos  As Long
    lngBytesRecibidos = RecvData(Data, False, varType, maxLen)
End Sub

Private Function RecvData(ByRef Data As Variant, ByVal blnPeek As Boolean, Optional varClass As Variant, Optional maxLen As Variant) As Long
'This function is to retrieve data from the buffer. If we are using
'TCP then the data is retrieved from a local buffer (m_strRecvBuffer).
'If we are using UDP the data is retrieved from winsock buffer.
'It can be called by two public methods of the class - GetData and PeekData.
'Behavior of the function is defined by the blnPeek argument. If a value of
'that argument is TRUE, the function returns number of bytes in the
'buffer, and copy data from that buffer into the data argument.
'If a value of the blnPeek is FALSE, then this function returns number of
'bytes received, and move data from the buffer into the data
'argument. MOVE means that data will be removed from the buffer.
    Dim blnMaxLenMiss   As Boolean
    Dim blnClassMiss As Boolean
    Dim strRecvData     As String
    Dim lngBufferLen    As Long
    Dim arrBuffer()     As Byte
    Dim lngErrorCode    As Long
        lngBufferLen = Len(m_strRecvBuffer)
    blnMaxLenMiss = IsMissing(maxLen)
    blnClassMiss = IsMissing(varClass)
    'Select type of data
    If varType(Data) = vbEmpty Then
        If blnClassMiss Then varClass = vbArray + vbByte
    Else
        varClass = varType(Data)
    End If
    'As stated on Winsock control documentation if the
    'data type passed is string or byte array type then
    'we must take into account maxLen argument.
    'If it is another type maxLen is ignored.
    If varClass = vbString Or varClass = vbArray + vbByte Then
        If blnMaxLenMiss Then 'if maxLen argument is missing
            If lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                Data = arrBuffer
                Exit Function
            Else
                RecvData = lngBufferLen
                arrBuffer = BuildArray(lngBufferLen, blnPeek, lngErrorCode)
            End If
        Else 'if maxLen argument is not missing
            If maxLen = 0 Or lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                Data = arrBuffer
                Exit Function
            ElseIf maxLen > lngBufferLen Then
                RecvData = lngBufferLen
                arrBuffer = BuildArray(lngBufferLen, blnPeek, lngErrorCode)
            Else
                RecvData = CLng(maxLen)
                arrBuffer() = BuildArray(CLng(maxLen), blnPeek, lngErrorCode)
            End If
        End If
    End If
        Select Case varClass
            Case vbString
                Dim strData As String
                strData = StrConv(arrBuffer(), vbUnicode)
                Data = strData
                
            Case Else
                Err.Raise sckUnsupported, "CSocket2.RecvData", "Unsupported variant type."
        End Select
    'if BuildArray returns an error is handled here
    'If lngErrorCode <> 0 Then
        'err.Raise lngErrorCode, "CSocket2.RecvData", GetErrorDescription(lngErrorCode)
        'Debug.Print "Error"
    'End If
End Function

Public Sub CloseSck()
    If HSock = INVALID_SOCKET Then Exit Sub
    DestroySocket
    m_strRecvBuffer = ""
End Sub

Private Sub DestroySocket()
    'Destroys the socket if it exists and unregisters it
    'from control list.
    If Not HSock = INVALID_SOCKET Then
        Dim lngResult As Long
        lngResult = WCloseSocket(HSock)
        If lngResult = SOCKET_ERROR Then
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
        Else
            UnregisterSocket HSock
            HSock = INVALID_SOCKET
        End If
    End If
End Sub

Private Sub UnregisterSocket(ByVal lngSocket As Long)
    'Removes the socket from the m_colSocketsInst collection
    'If it is the last socket in that collection, the window
    'and colection will be destroyed as well.
    On Error Resume Next
        UnSubClass
End Sub

Private Function BuildArray(ByVal size As Long, ByVal blnPeek As Boolean, ByRef lngErrorCode As Long) As Byte()
'Returns a byte array of Size bytes filled with incoming buffer data.
    Dim strData As String
        strData = Left$(m_strRecvBuffer, CLng(size))
        BuildArray = StrConv(strData, vbFromUnicode)
        If Not blnPeek Then
            m_strRecvBuffer = Mid$(m_strRecvBuffer, size + 1)
        End If
End Function

Private Sub DataArrival(ByVal bytesTotal As Long, ByVal sRet As String)
    Dim M%, N%, I%
    Dim Response$
    Dim result As Long
    Static strBuffer As String
    Static iPntr As Integer
    Dim strArrival As String
    
    On Error GoTo DataArrivalErr
    M% = 1
    N% = InStr(sRet, Chr$(13))
    
    Do Until N% = 0 'There shouldn't be more than 1 line, but check anyway
        strBuffer = strBuffer + Mid$(sRet, M%, N% - M%)
        Response$ = ParseInput(strBuffer, " ")
        Debug.Print "RESPONSE: ~" & strBuffer & "|"
        If Len(Response$) > 0 Then
            Select Case val(Response$)
                Case 221 'QUIT
                    ExitConnection "Message was sent"
                    LCurState = Server_CONNECT '0
                Case 220 'Greeting/STARTTLS
                    If LCurState = Server_CONNECT Then '0
                        'HELO is used instead of EHLO because EHLO returns
                        'multiple lines. Additional code would be required here.
                        txtOut = "HELO " & Chr(32) & sFromName '"HELO "
                        Debug.Print "SENDING: " & txtOut
                        'Save SMTP conversation
                        SendData (txtOut) & vbCrLf
                        LCurState = Server_EHLO '1
                    ElseIf LCurState = Server_EHLO And flgTLS = True Then 'Start TLS
                        Let msgType = TLS1_RT_HANDSHAKE
                        Let msgVerMaj = VERSION_MAJOR
                        Let msgVerMin = VERSION_MINOR_1
                        Let EncrFlg = flgTLS
                        txtOut = GetRandomString
                        Debug.Print "SENDING random string, length = " & Len(txtOut)
                        hSESSION_MD5 = 0 'Clear Session Hashes
                        hSESSION_SHA1 = 0
                        hSESSION_MD5 = MD5_AddToHash(txtOut, hSESSION_MD5)
                        hSESSION_SHA1 = SHA1_AddToHash(txtOut, hSESSION_SHA1)
                        'Prepare middle 3 bytes of MAC Header
                        SendData AddRecHeader(txtOut)
                    End If
                Case 250 'Most other responses
                    flgTLS = True
                    txtOut = "STARTTLS"
                    OutArray(4) = OutArray(4) & strBuffer & vbCrLf & txtOut & vbCrLf
                    SendData (txtOut) & vbCrLf
                    Debug.Print "SENDING: " & txtOut
                Case Else
                    'Call CloseSck
                    SysBusy = False
                    LCurState = Server_CONNECT '0
                    OutArray(4) = OutArray(4) & strBuffer & vbCrLf & "Send Error" & vbCrLf
                    OutArray(3) = False 'Sent flag set to false
                    ExitConnection "Message did not Send Properly!"
                    Exit Sub
            End Select
        End If
        strBuffer = ""
        If Mid$(strArrival, N% + 1, 1) = Chr(10) Then
            M% = N% + 2
        Else
            M% = N% + 1
        End If
        N% = InStr(M% + 1, strArrival, Chr$(13))
    Loop
    'Copy what's left to buffer
    strBuffer = Mid$(strArrival, M%)
    Exit Sub
DataArrivalErr:
    'Debug.Print Error$(err) & ": " & strArrival
End Sub

Private Sub EncrDataArrival(ByVal bytesTotal As Long, ByVal sRet As String)
    Dim sData As String
    'This routine is re-entrant, hence the next 3 variable must be static
    Static InBuff As String
    Static Header() As Byte
    Static RecLen As Long
    Dim MsgLen As Long
    Static iPntr As Integer
    Dim DeCryptData As String
    Dim SessionLen As Long
    Dim sCipher As String
    Dim sTemp As String
    Dim lTemp As Long
    Dim strToSend As String
    Dim strSHAMD5 As String
    Dim sCharset As String
    Dim ErrCode As Long
    Dim Response$
    Dim I%
    
    GetData sData
    InBuff = InBuff & sData 'Add data to buffer
GetNextRecord:
    If Len(InBuff) < 6 Then Exit Sub 'If no record length yet then exit & wait
    If RecLen = 0 Then 'New record
        Header() = StrConv(Left$(InBuff, 5), vbFromUnicode)
        Select Case Header(0)
            Case TLS1_RT_CHANGE_CIPHER_SPEC 'Not considered a Crypto record
            Case TLS1_RT_ALERT, TLS1_RT_HANDSHAKE, TLS1_RT_APPLICATION_DATA 'Update Msg Type
                Let msgType = Header(0)
                Let msgVerMaj = Header(1)
                Let msgVerMin = Header(2)
            Case Else 'Ignore record
                InBuff = ""
                RecLen = 0
                GoTo Done
        End Select
        RecLen = Header(3) * 256 + Header(4)
        InBuff = Mid$(InBuff, 6) 'Strip header from buffer
    End If
    If Len(InBuff) >= RecLen Then
        Let InBuffer = Left$(InBuff, RecLen)
        InBuff = Mid$(InBuff, RecLen + 1)
    Else
        Exit Sub 'Wait for complete record
    End If
    'record complete - Process it
    Select Case Layer
        Case 0 'this layer processes the Server Hello
            sTemp = sInBuffer
            'Hash data
            hSESSION_MD5 = MD5_AddToHash(sTemp, hSESSION_MD5)
            hSESSION_SHA1 = SHA1_AddToHash(sTemp, hSESSION_SHA1)
ProcessMore:
            MsgLen = Asc(Mid$(sTemp, 3, 1)) * 256 + Asc(Mid$(sTemp, 4, 1)) + 4
            Select Case Asc(sTemp)
                Case TLS1_MT_SERVER_HELLO
                    Debug.Print "RESPONSE: ENCRYPTED DATA"
                    SERVER_RANDOM = Mid$(sTemp, 7, 32)
                    SessionLen = Asc(Mid$(sTemp, 39, 1))
                    SESSION_ID = Mid$(sTemp, 40, SessionLen)
                    Let CipherToUse = Mid$(sTemp, 40 + SessionLen, 2)
                    If Len(sTemp) > MsgLen Then
                        sTemp = Mid$(sTemp, MsgLen + 1)
                        GoTo ProcessMore
                    End If
                Case TLS1_MT_CERTIFICATE
                    Debug.Print "RESPONSE: ENCRYPTED CERTIFICATE"
                    If Not VerifySignature(sTemp) Then
                        Debug.Print "Unable to Verify Signature!", vbExclamation
                        GoTo Done
                    End If
                    If Len(sTemp) > MsgLen Then
                        sTemp = Mid$(sTemp, MsgLen + 1)
                        GoTo ProcessMore
                    End If
                Case TLS1_MT_SERVER_DONE
                    Debug.Print "SENDING PRE-MASTER KEY "
                    strToSend = GetClientExchange
                    SendData strToSend
                    IncSendSeqNum
                    Layer = Layer + 1
            End Select
        Case 1 'this layer processes the Server Key Exchange record
            Debug.Print "RESPONSE: ENCRYPTED SERVER KEY EXCHANGE RECORD"
            sTemp = sInBuffer
            If Header(0) = &H14 Then
                'Cipher Change Spec; Do nothing
                IncRecvSeqNum (True) 'Reset Recv Sequence Number
            ElseIf Header(0) = &H15 Then 'Problem?
                ErrCode = Asc(Mid$(sTemp, 2, 1))
                Select Case Asc(sTemp)
                    Case 1
                        If ErrCode > 0 Then
                            ExitConnection "ALERT WARNING " & Chr(32) & CStr(ErrCode)
                            Debug.Print "EXIT CONNECTION BECAUSE OF ALERT WARNING"
                        End If
                    Case 2
                        ExitConnection "FATAL ALERT " & Chr(32) & CStr(ErrCode)
                        Debug.Print "EXIT CONNECTION BECAUSE OF FATAL ALERT"
                    Case Else
                        ExitConnection "UNKOWN ERROR"
                        Debug.Print "EXIT CONNECTION BECAUSE OF UNKNOWN ERROR"
                End Select
                GoTo Done
            Else 'Must be Session Handshake
                sTemp = sInBuffer
                Call TLSDecrypt
                IncRecvSeqNum
                'Save Decrypted Session Hash
                strSHAMD5 = sOutBuffer
                Let InBuffer = MD5_GetHash(hSESSION_MD5) _
                    & SHA1_GetHash(hSESSION_SHA1)
                Call Hash_SHAMD5(True)
                Layer = Layer + 1
                Let msgType = TLS1_RT_APPLICATION_DATA
                Let msgVerMaj = VERSION_MAJOR
                Let msgVerMin = VERSION_MINOR_1
                If sPortNumber = 587 Then 'Carry on with secure sockets
                    Debug.Print "SENDING ENCRYPTED: AUTH PLAIN " & sPassword
                    txtOut = "AUTH PLAIN " & Chr(32) & sPassword
                    OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                    sInBuffer = txtOut & vbCrLf
                    TLSEncrypt
                    SendData AddRecHeader(sOutBuffer)
                    IncSendSeqNum
                    LCurState = Server_AUTH
                End If
            End If
        Case Else 'Application data starts here
            Call TLSDecrypt
            Debug.Print "RESPONSE: ENCRYPTED APPLICATION DATA"
            IncRecvSeqNum
            sTemp = sOutBuffer
            If Header(0) = &H15 Then 'Problem?
                ErrCode = Asc(Mid$(sTemp, 2, 1))
                Select Case Asc(sTemp)
                    Case 1
                        If ErrCode > 0 Then
                            ExitConnection "ALERT WARNING " & Chr(32) & CStr(ErrCode)
                            Debug.Print "PROBLEM: ALERT WARNING"
                        End If
                    Case 2
                        If Not (CStr(ErrCode) = "20") Then
                            ExitConnection "FATAL ALERT " & Chr(32) & CStr(ErrCode)
                            Debug.Print "PROBLEM: FATAL ALERT"
                        Else
                            LCurState = Server_QUIT
                            ExitConnection "MAIL WAS SENT"
                            Debug.Print "SUCCESS: MAIL WAS SENT"
                        End If
                    Case Else
                        ExitConnection "UNKOWN ERROR"
                        Debug.Print "PROBLEM: UNKNOWN ERROR"
                End Select
                GoTo Done
            End If
            Response$ = ParseInput(sTemp, " ")
            Debug.Print "RESPONSE: ENCRYPTED " & Response
            If Len(Response$) > 0 Then
                Select Case val(Response$)
                    Case 221 'QUIT
                        ExitConnection "MAIL WAS SENT"
                        SysBusy = False
                        LCurState = Server_CONNECT '0
                        OutArray(3) = True 'Message was sent
                    Case 220 'Greeting
                        If LCurState = Server_CONNECT Then '0
                            txtOut = "EHLO" & Chr(32) & sUserName
                            Debug.Print "SENDING ENCRYPTED: " & txtOut
                            'Save SMTP conversation
                            OutArray(4) = sTemp & vbCrLf & txtOut & vbCrLf
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            LCurState = Server_EHLO '1
                        End If
                    Case 250 'Most other responses
                        If LCurState = Server_EHLO Then '1
                            If Len(sFromEmail) > 0 Then
                                txtOut = "AUTH PLAIN " & Chr(32) & sPassword
                                Debug.Print "SENDING ENCRYPTED: " & txtOut
                                OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                                sInBuffer = txtOut & vbCrLf
                                TLSEncrypt
                                SendData AddRecHeader(sOutBuffer)
                                IncSendSeqNum
                                LCurState = Server_AUTH '2
                            Else
                                txtOut = "MAIL FROM: " & Chr(32) & sUserName
                                Debug.Print "SENDING ENCRYPTED: " & txtOut
                                OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                                sInBuffer = txtOut & vbCrLf
                                TLSEncrypt
                                SendData AddRecHeader(sOutBuffer)
                                IncSendSeqNum
                                LCurState = Server_FROM '3
                            End If
                        ElseIf LCurState = Server_FROM Then '3
                            If iPntr < UBound(Recipient) Then
                                txtOut = "RCPT TO: " & Chr(32) & Recipient(iPntr)
                                iPntr = iPntr + 1
                            Else
                                txtOut = "RCPT TO: " & Chr(32) & Recipient(iPntr)
                                LCurState = Server_TO '4
                                iPntr = 0
                            End If
                            OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                            Debug.Print "SENDING ENCRYPTED: " & txtOut
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                        ElseIf LCurState = Server_TO Then '4
                            txtOut = "DATA"
                            Debug.Print "SENDING ENCRYPTED: " & txtOut
                            OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            LCurState = Server_DATA '5
                        ElseIf LCurState = Server_QUIT Then '6
                            txtOut = "QUIT"
                            Debug.Print "SENDING ENCRYPTED: " & txtOut
                            OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                        End If
                    Case 354 'Data response
                        txtOut = "From: " & Chr(32) & sFromEmail & vbCrLf
                        txtOut = txtOut & "To: " & Chr(32) & Recipient(0) & vbCrLf
                        txtOut = txtOut & "Subject: " & Chr(32) & sSubjectLine & vbCrLf
                        txtOut = txtOut & "X-Mailer: AutoMail 2.0" & vbCrLf
                        txtOut = txtOut & "Date: " & Chr(32) & GetDateString() & vbCrLf
                        txtOut = txtOut & "Message-ID:" & " <" & GetGUID() & "@" & "automail" & ">" & vbCrLf
                        txtOut = txtOut & "MIME-Version: 1.0" + vbCrLf
                        sCharset = "charset=" & Qt2 & "ISO-8859-1" & Qt2 & vbCrLf
                        If Len(sAttach) > 0 Then 'Attachments found
                            Debug.Print "ATTACHMENTS FOUND"
                            txtOut = txtOut & "Content-Type: multipart/mixed;" & vbCrLf
                            txtOut = txtOut & vbTab & "boundary=" & Qt2 & cBoundaryA & Qt2 & vbCrLf
                            txtOut = txtOut & "X-Mailer: AutoMail" & vbCrLf
                            Debug.Print "SENDING: " & txtOut
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                            txtOut = cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/plain;" & vbCrLf
                            txtOut = txtOut & vbTab & sCharset
                            txtOut = txtOut & "Content-Transfer-Encoding: 7bit" & vbCrLf & vbCrLf
                            Debug.Print "SENDING: " & txtOut
                            sInBuffer = txtOut & vbCrLf
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            sInBuffer = sBodyText & vbCrLf & vbCrLf
                            Debug.Print "SENDING: " & sInBuffer
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            For I% = 0 To UBound(FileAttach)
                                Do
                                    txtOut = Left$(Bas64buffer(I%), 16384)
                                    Debug.Print "SENDING BLOCK"
                                    sInBuffer = txtOut
                                    TLSEncrypt
                                    SendData AddRecHeader(sOutBuffer)
                                    IncSendSeqNum
                                    Sleep 100
                                    If Len(txtOut) >= 16384 Then Bas64buffer(I%) = Mid$(Bas64buffer(I%), 16385)
                                Loop Until Len(txtOut) < 16384
                                If I% = UBound(FileAttach) Then 'Last attachment
                                    'Send ending boundary string
                                    txtOut = "--" & cBoundaryA & "--" & vbCrLf
                                    OutArray(5) = OutArray(5) & txtOut & vbCrLf
                                    sInBuffer = txtOut & vbCrLf
                                    TLSEncrypt
                                    SendData AddRecHeader(sOutBuffer)
                                    IncSendSeqNum
                                End If
                            Next I%
                        Else 'plain text mail
                            txtOut = txtOut & "Content-Type: text/plain;" & vbCrLf
                            txtOut = txtOut & vbTab & sCharset
                            txtOut = txtOut & "Content-Transfer-Encoding: 7bit" & vbCrLf
                            txtOut = txtOut & "X-Mailer: AutoMail" & vbCrLf & vbCrLf
                            Debug.Print "SENDING: " & txtOut
                            sInBuffer = txtOut
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            sInBuffer = sBodyText
                            Debug.Print "SENDING MAIL BODY TEXT: " & sBodyText
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            OutArray(5) = sBodyText
                        End If
                        sInBuffer = "." & vbCrLf
                        TLSEncrypt
                        SendData AddRecHeader(sOutBuffer)
                        IncSendSeqNum
                        LCurState = Server_QUIT '6
                    Case 235 'Authentication response
                        txtOut = "MAIL FROM: " & Chr(32) & sUserName '"MAIL FROM: "
                        OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                        sInBuffer = txtOut & vbCrLf
                        Debug.Print "SENDING ENCRYPTED: " & txtOut
                        TLSEncrypt
                        SendData AddRecHeader(sOutBuffer)
                        IncSendSeqNum
                        LCurState = Server_FROM '3
                    Case 500 'Command not recognized
                        If LCurState = Server_AUTH Then 'AUTH command not recognized
                            txtOut = "MAIL FROM:" & Chr(32) & sUserName
                            OutArray(4) = OutArray(4) & sTemp & vbCrLf & txtOut & vbCrLf
                            sInBuffer = txtOut & vbCrLf
                            Debug.Print "SENDING ENCRYPTED: " & txtOut
                            TLSEncrypt
                            SendData AddRecHeader(sOutBuffer)
                            IncSendSeqNum
                            LCurState = Server_FROM '3
                        End If
                    Case Else
                        SysBusy = False
                        LCurState = Server_CONNECT '0
                        OutArray(4) = OutArray(4) & sTemp & vbCrLf & "Send Error" & vbCrLf '"Send Error"
                        OutArray(3) = False 'Sent flag set to false
                        Debug.Print "EXITING : Message did not send properly"
                        ExitConnection "Message did not Send Properly!"
                        Exit Sub
                End Select
            End If
    End Select
Done:
    RecLen = 0
    ReDim Header(0)
    If Len(InBuff) > 0 Then GoTo GetNextRecord
End Sub

Private Function ParseInput(a$, Sep$) As String
    Dim M%, N%, N1%
    ReDim ParseArray(100)
    a$ = Trim$(Replace(a$, ":", " "))
    M% = 1: N1% = 0
    Do Until M% = 0
        M% = N% + 1
        N% = InStr(M%, a$, Sep$)
        If N% = 0 Then
            ParseArray(N1%) = Trim$(Mid$(a$, M%))
            N1% = N1% + 1
            Exit Do
        End If
        ParseArray(N1%) = Trim$(Mid$(a$, M%, N% - M%))
        N1% = N1% + 1
        Do Until Mid$(a$, N% + 1, 1) <> Sep$
            N% = N% + 1
        Loop
    Loop
    ReDim Preserve ParseArray(N1% - 1)
    ParseInput = ParseArray(0)
End Function

Public Sub PrepareEmail()
 Dim s1 As String
 Dim sFil64, sHeader As String
 Dim Pt1 As Long
 Dim AFils As Variant
 Dim I As Integer, iPics As Integer
     
    ReDim FileAttach(0) As String
    ReDim FileType(0) As String
    ReDim FileName(0) As String
    
    Debug.Print "SEND.PrepareEmail"
    
    If sAttach <> LogFile Then
        sAttach = StrConv(sAttach, vbFromUnicode)
    End If
    
    FileAttach(0) = ""
    FileType(0) = ""
    iPics = 0
    If (Len(sAttach) > 0) Then
        AFils = Split(sAttach, cAST) 'multiple attachments: filenames are seperated with "*"
        For I = 0 To UBound(AFils)
            s1 = AFils(I)
            If (UCFileExists(s1) = True) Then
                Pt1 = InStrRev(s1, "\")
                ReDim Preserve FileAttach(0 To iPics) As String
                ReDim Preserve Bas64buffer(0 To iPics) As String
                ReDim Preserve FileName(0 To iPics) As String
                ReDim Preserve FileType(0 To iPics) As String
                FileName(iPics) = Right$(s1, (Len(s1) - Pt1))
                sHeader = GetAttachmentHeader(FileName(iPics))
                sFil64 = ConvertToBase64(s1, True, True)
                sHeader = sHeader & sFil64 & vbCrLf & vbCrLf
                FileAttach(iPics) = s1
                Bas64buffer(iPics) = sHeader
                Debug.Print Len(Bas64buffer(iPics))
                iPics = iPics + 1
            End If
        Next I
   End If
    
Recip:
    ReDim Recipient(0) As String
    Recipient(0) = ""
    
    If (Len(sRecip) > 0) Then
        AFils = Split(sRecip, cSEMIC) 'multiple Recipients: names are separated with ";"
        For I = 0 To UBound(AFils)
            s1 = "<" & AFils(I) & ">"
            Debug.Print "RECIPIENT : " & s1
            If InStr(1, s1, "@") > 2 Then
                ReDim Preserve Recipient(0 To I) As String
                Recipient(I) = s1
            Else
                MsgBox "Faulty Email Adress for " & s1 & vbCrLf & "Check the recipient string", vbCritical, "ERROR"
            End If
        Next I
    End If
    Let sToEmail = "<" & Recipient(0) & ">"
End Sub


'************************************************************************************************************************
'************************************************************************************************************************
'*******                                                                                                          *******
'*******                                        CRYPT CLS                                                         *******
'*******                                                                                                          *******
'************************************************************************************************************************
'************************************************************************************************************************

'******************************
'******************************
'******    PROPERTIES    ******
'******************************
'******************************

Private Property Let CipherToUse(sNewCipher As String)
    lCipher = Asc(Left$(sNewCipher, 1)) * 256 + Asc(Right$(sNewCipher, 1))
    Select Case lCipher
        Case 4
            Alg_Hash = CALG_MD5
        Case 5
            Alg_Hash = CALG_SHA1
        Case Else
            Alg_Hash = CALG_SHA1
    End Select
End Property

Private Property Let msgType(vNewValue As Byte)
    sMsgType = vNewValue
End Property

Public Property Let msgVerMaj(bNewValue As Byte)
    sMsgVerMaj = bNewValue
End Property

Public Property Let msgVerMin(bNewValue As Byte)
    sMsgVerMin = bNewValue
End Property

Private Property Let InBuffer(vNewValue As String)
    sInBuffer = vNewValue
End Property

Private Property Let PubKey(NewKey As String)
    sPublicKey = NewKey
End Property

Private Property Let ClientRandom(sClientRandom As String)
    ReDim bClientRandom(0)
    bClientRandom = StrConv(sClientRandom, vbFromUnicode)
End Property

Private Property Let ServerRandom(sServerRandom As String)
    ReDim bServerRandom(0)
    bServerRandom = StrConv(sServerRandom, vbFromUnicode)
End Property

Private Property Let PreMasterKey(sPreMasterKey As String)
    ReDim bPreMasterKey(0)
    bPreMasterKey = StrConv(sPreMasterKey, vbFromUnicode)
End Property

Private Property Let Signature(vNewValue As String)
    sSignature = vNewValue
End Property

Private Sub Hash_SHAMD5(fServer As Boolean)
    Const Routine As String = "Hash_SHAMD5"
    Dim hTLS1PRF As Long
    Dim DATA_BLOB As CRYPT_DATA_BLOB
    Dim strHash As String
    Dim bHandShakeHash() As Byte
    Dim lLen As Long
    Dim bFinish() As Byte
    Dim sBuffer As String
    If CryptCreateHash(hCryptProv, CALG_TLS1PRF, hMasterKey, 0, hTLS1PRF) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
        GoTo ReleaseHandles
    End If
    bHandShakeHash = StrConv(sInBuffer, vbFromUnicode)
    DATA_BLOB.pbData = VarPtr(bHandShakeHash(0))
    DATA_BLOB.cbData = Len(sInBuffer) '36
    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_SEED, DATA_BLOB, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSHP, Routine)
        GoTo ReleaseHandles
    End If
    If fServer Then
        bFinish = StrConv(abServerFinished, vbFromUnicode)
    Else
        bFinish = StrConv(abClientFinished, vbFromUnicode)
    End If
    DATA_BLOB.pbData = VarPtr(bFinish(0))
    DATA_BLOB.cbData = UBound(bFinish) + 1 '15
    If CryptSetHashParam(hTLS1PRF, HP_TLS1PRF_LABEL, DATA_BLOB, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSHP, Routine)
        GoTo ReleaseHandles
    End If
    lLen = 12
    'Prepare buffer
    sBuffer = String$(lLen, Chr$(0))
    If CryptGetHashParam(hTLS1PRF, HP_HASHVAL, sBuffer, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        GoTo ReleaseHandles
    End If
    sInBuffer = Chr$(&H14) & Chr$(0) & Chr$(0) & Chr$(Len(sBuffer) Mod 256) & sBuffer
    Call TLSEncrypt
    flgFinal = False
ReleaseHandles:
    If hTLS1PRF <> 0 Then CryptDestroyHash hTLS1PRF
End Sub

Private Function ExportMasterKey() As Boolean
    Const Routine As String = "ExportMasterKey"
    Dim hPubCertKey As Long 'temporary handle for Certificate Public key
    Dim hExchngKey As Long
    Dim hMasterHash As Long
    Dim lLen As Long
    Dim TempBlob As T_SIMPLEKEYBLOB
    Dim MasterKeyBlob As T_EXP_SIMPLEKEYBLOB
    Dim Algorithm As SCHANNEL_ALG
    Dim Data As CRYPT_DATA_BLOB
    Dim sTemp As String
    Dim bTemp() As Byte
    Dim sPreMaster As String
    lStatus = CFB_BUSY
    'Acquire Context
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, 0) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Get handle to Exchange key
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hExchngKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGUK, Routine)
        GoTo ReleaseHandles
    End If
    'Instead of creating the encrypted Pre-Master Key with the following call,
    'I chose to create the key, encrypt it with the exchange key, import it,
    'and then encrypt it using the server public key. Why you ask? The encrypted
    'Pre-Master key requires the server private key to decrypt.
    'If CryptGenKey(hCryptProv, CALG_TLS1_MASTER, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
    '    RaiseEvent Error(err.LastDllError, CGK, Routine)
    '    GoTo ReleaseHandles
    'End If
    lLen = 0
    If CryptEncrypt(hExchngKey, 0, True, 0, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    sTemp = sInBuffer & String$(lLen - Len(sInBuffer), Chr$(0))
    If CryptEncrypt(hExchngKey, 0, True, 0, sTemp, Len(sInBuffer), Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    Let PreMasterKey = sTemp 'converts to byte array in bPreMasterKey()
    'Prepare Master Key Blob
    TempBlob.bType = SIMPLEBLOB
    TempBlob.bVersion = CUR_BLOB_VERSION
    TempBlob.reserved = 0
    TempBlob.aiKeyAlg = CALG_TLS1_MASTER
    TempBlob.Algid = CALG_RSA_KEYX
    CopyMemory MasterKeyBlob.bSimpleKey(1), TempBlob.bType, 12
    CopyMemory MasterKeyBlob.bSimpleKey(13), bPreMasterKey(0), RSA_KEY_LEN
    'Import & Decrypt Pre-Master Key
    If CryptImportKey(hCryptProv, MasterKeyBlob, Len(MasterKeyBlob), hExchngKey, CRYPT_EXPORTABLE, hMasterKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    'Import Certificate Public Key as blob
    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    lLen = 0
    'Get length of encryption
    If CryptEncrypt(hPubCertKey, 0, True, 0, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    'Encrypt Pre-Master Key with Public Cert Key
    sTemp = sInBuffer & String$(lLen - Len(sInBuffer), Chr$(0))
    If CryptEncrypt(hPubCertKey, 0, True, 0, sTemp, Len(sInBuffer), Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    sOutBuffer = sTemp 'Save Public key encrypted Pre-Master Key
    'Set SCHANNEL_ENC_KEY parameters
    Algorithm.dwUse = SCHANNEL_ENC_KEY
    Algorithm.Algid = CALG_RC4
    Algorithm.cBits = 128
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Set SCHANNEL_MAC_KEY parameters
    Algorithm.dwUse = SCHANNEL_MAC_KEY
    Algorithm.Algid = Alg_Hash
    If Alg_Hash = CALG_MD5 Then
        Algorithm.cBits = 128
    Else
        Algorithm.cBits = 160
    End If
    Algorithm.dwFlags = 0
    Algorithm.dwReserved = 0
    If CryptSetKeyParam(hMasterKey, KP_SCHANNEL_ALG, Algorithm, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add CLIENT_RANDOM
    Data.cbData = UBound(bClientRandom) + 1
    Data.pbData = VarPtr(bClientRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_CLIENT_RANDOM, Data, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Add SERVER_RANDOM
    Data.cbData = UBound(bServerRandom) + 1
    Data.pbData = VarPtr(bServerRandom(0))
    If CryptSetKeyParam(hMasterKey, KP_SERVER_RANDOM, Data, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Create Master Hash
    If CryptCreateHash(hCryptProv, CALG_SCHANNEL_MASTER_HASH, hMasterKey, 0, hMasterHash) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CSKP, Routine)
        GoTo ReleaseHandles
    End If
    'Create the other SCHANNEL keys
    fClient = True
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_ENC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, CRYPT_SERVER, 0) + CRYPT_EXPORTABLE, hReadMAC) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    If CryptDeriveKey(hCryptProv, CALG_SCHANNEL_MAC_KEY, hMasterHash, IIf(fClient, 0, CRYPT_SERVER) + CRYPT_EXPORTABLE, hWriteMAC) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CDrK, Routine)
        GoTo ReleaseHandles
    End If
    lLen = 0
    'Get WriteMAC Key Length
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get WriteMAC Key (used in HMAC routine)
    If CryptExportKey(hWriteMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        WRITE_MAC_KEY = Left$(sTemp, 20)
    Else
        WRITE_MAC_KEY = Left$(sTemp, 16)
    End If
    lLen = 0
    'Get ReadMAC Key Length
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get ReadMAC Key (used in HMAC routine)
    If CryptExportKey(hReadMAC, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    If Alg_Hash = CALG_SHA1 Then
        READ_MAC_KEY = Left$(sTemp, 20)
    Else
        READ_MAC_KEY = Left$(sTemp, 16)
    End If
    lLen = 0
    'Get Write Key Length
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Write Key
    If CryptExportKey(hWriteKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    WRITE_KEY = Left$(sTemp, 16)
    lLen = 0
    'Get Read Key Length
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, ByVal 0&, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    'Get Read Key
    If CryptExportKey(hReadKey, hExchngKey, SIMPLEBLOB, 0, MasterKeyBlob, lLen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CEK, Routine)
        GoTo ReleaseHandles
    End If
    ReDim bTemp(Len(MasterKeyBlob) - 1)
    CopyMemory bTemp(0), MasterKeyBlob, Len(MasterKeyBlob)
    sTemp = Mid$(StrConv(bTemp, vbUnicode), 13)
    If CryptDecrypt(hExchngKey, 0, True, 0, sTemp, Len(sTemp)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CD, Routine)
        GoTo ReleaseHandles
    End If
    READ_KEY = Left$(sTemp, 16)
    ExportMasterKey = True 'Success
ReleaseHandles:
    If hPubCertKey Then CryptDestroyKey hPubCertKey
    If hMasterHash Then CryptDestroyHash hMasterHash
    flgFinal = True
    lStatus = CFB_READY
End Function

Private Sub IncSendSeqNum(Optional flgClear As Boolean)

    If flgClear Then
        SEND_SEQ_NUM(0) = 0
        SEND_SEQ_NUM(1) = 0
        SEND_SEQ_NUM(2) = 0
        Exit Sub
    End If
    If SEND_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
        SEND_SEQ_NUM(0) = SEND_SEQ_NUM(0) + 1
    ElseIf SEND_SEQ_NUM(1) < 16777215 Then
        SEND_SEQ_NUM(1) = SEND_SEQ_NUM(1) + 1
        SEND_SEQ_NUM(0) = 0
    ElseIf SEND_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
        SEND_SEQ_NUM(2) = SEND_SEQ_NUM(2) + 1
        SEND_SEQ_NUM(1) = 0
    Else
        SEND_SEQ_NUM(0) = 0
        SEND_SEQ_NUM(1) = 0
        SEND_SEQ_NUM(2) = 0
    End If

End Sub

Private Sub IncRecvSeqNum(Optional flgClear As Boolean)
    If flgClear Then
        RECV_SEQ_NUM(0) = 0
        RECV_SEQ_NUM(1) = 0
        RECV_SEQ_NUM(2) = 0
        Exit Sub
    End If
    If RECV_SEQ_NUM(0) < 16777215 Then '256 ^ 3 - 1
        RECV_SEQ_NUM(0) = RECV_SEQ_NUM(0) + 1
    ElseIf RECV_SEQ_NUM(1) < 16777215 Then
        RECV_SEQ_NUM(1) = RECV_SEQ_NUM(1) + 1
        RECV_SEQ_NUM(0) = 0
    ElseIf RECV_SEQ_NUM(2) < 65535 Then '256 ^ 2 - 1
        RECV_SEQ_NUM(2) = RECV_SEQ_NUM(2) + 1
        RECV_SEQ_NUM(1) = 0
    Else
        RECV_SEQ_NUM(0) = 0
        RECV_SEQ_NUM(1) = 0
        RECV_SEQ_NUM(2) = 0
    End If

End Sub

Private Function VerifySig() As Boolean

    Const Routine As String = "VerifySig"
    Dim hHash As Long
    Dim hPubCertKey As Long
    Dim sDescription As String
    'Create a hash object.
    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
        GoTo ReleaseHandles:
    End If
    'Hash the data
    If CryptHashData(hHash, sInBuffer, Len(sInBuffer), 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CHD, Routine)
        GoTo ReleaseHandles:
    End If
    'Import Certificate Public Key as blob
    If CryptImportKey(hCryptProv, PublicKey, Len(PublicKey), 0, CRYPT_EXPORTABLE, hPubCertKey) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CIK, Routine)
        GoTo ReleaseHandles
    End If
    sDescription = vbNullString
    If CryptVerifySignature(hHash, sSignature, Len(sSignature), hPubCertKey, sDescription, 0) = 0 Then
        VerifySig = False
    Else
        VerifySig = True
    End If
ReleaseHandles:
    If hPubCertKey Then CryptDestroyKey hPubCertKey
    If hHash Then CryptDestroyHash hHash
End Function

Private Function MD5_AddToHash(ByVal strToHash As String, hHash) As Long
    Const Routine As String = "MD5_AddToHash"
    If hHash = 0 Then
        If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
            Exit Function
        End If
    End If
    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
        Exit Function
    End If
    MD5_AddToHash = hHash
End Function

Private Function CopyHash(hHashToCopy) As Long
    Const Routine As String = "CopyHash"
    Dim phHash As Long
    If CryptDuplicateHash(hHashToCopy, ByVal 0&, 0&, phHash) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CDuH, Routine)
        CopyHash = 0
    Else
        CopyHash = phHash
    End If
End Function

Private Function MD5_GetHash(hHash As Long) As String
    Const Routine As String = "MD5_GetHash"
    Dim strHash As String
    Dim lLen As Long
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(lLen, vbNullChar)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then
        CryptDestroyHash hHash
        hHash = 0
    End If
    MD5_GetHash = strHash
End Function

Private Function MD5_Hash(ByVal TheString As String) As String
    Const Routine As String = "MD5_Hash"
    Dim strHash As String
    Dim hHash As Long
    Dim lLen As Long
    If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
        Exit Function
    End If
    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    'Get length of hash
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(lLen, vbNull)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then CryptDestroyHash hHash
    MD5_Hash = strHash
End Function

Private Function SHA1_AddToHash(ByVal strToHash As String, hHash) As Long
    Const Routine As String = "SHA1_AddToHash"
    If hHash = 0 Then
        If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
            Exit Function
        End If
    End If
    If CryptHashData(hHash, strToHash, Len(strToHash), 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    SHA1_AddToHash = hHash
End Function

Private Function SHA1_GetHash(hHash As Long) As String
    Const Routine As String = "SHA1_GetHash"
    Dim strHash As String
    Dim lLen As Long
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(lLen, vbNullChar)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then
        CryptDestroyHash hHash
        hHash = 0
    End If
    SHA1_GetHash = strHash
End Function

Private Function SHA1_Hash(ByVal TheString As String) As String
    Const Routine As String = "SHA1_Hash"
    Dim strHash As String
    Dim hHash As Long
    Dim lLen As Long
    If CryptCreateHash(hCryptProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CCH, Routine)
        Exit Function
    End If
    If CryptHashData(hHash, TheString, Len(TheString), 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CHD, Routine)
        Exit Function
    End If
    If CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    strHash = String(lLen, vbNull)
    If CryptGetHashParam(hHash, HP_HASHVAL, strHash, lLen, 0) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CGHP, Routine)
        Exit Function
    End If
    If hHash <> 0 Then CryptDestroyHash hHash
    SHA1_Hash = strHash
End Function

Private Function SendSeqNum() As String

    Dim bTmp(1 To 8) As Byte
    CopyMemory bTmp(1), SEND_SEQ_NUM(0), 3
    CopyMemory bTmp(4), SEND_SEQ_NUM(1), 3
    CopyMemory bTmp(7), SEND_SEQ_NUM(2), 2
    SendSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
End Function

Private Function RecvSeqNum() As String
    Dim bTmp(1 To 8) As Byte
    CopyMemory bTmp(1), RECV_SEQ_NUM(0), 3
    CopyMemory bTmp(4), RECV_SEQ_NUM(1), 3
    CopyMemory bTmp(7), RECV_SEQ_NUM(2), 2
    RecvSeqNum = StrReverse(StrConv(bTmp, vbUnicode))
End Function

Private Function ReadKeyFile(sFileName As String, bData() As Byte) As Boolean
    Dim lFileHandle As Long
    ' Reads a byte array from a file.
    ' If Data is fixed, it must be of the expected size.
    On Error GoTo ErrorHandler
    lFileHandle = FreeFile
    Open sFileName For Binary As lFileHandle
    Get lFileHandle, , bData
    Close lFileHandle
    ReadKeyFile = True
    Exit Function
ErrorHandler:
    ReadKeyFile = False
End Function

Private Function SetPubKey() As Boolean
    Dim bNewKey() As Byte
    bNewKey = StrConv(sPublicKey, vbFromUnicode)
    NewPubKey.bType = PUBLICKEYBLOB
    NewPubKey.bVersion = CUR_BLOB_VERSION
    NewPubKey.reserved = 0
    NewPubKey.aiKeyAlg = CALG_RSA_KEYX
    NewPubKey.magic = &H31415352 '"RSA1"
    NewPubKey.bitlen = Len(sPublicKey) * 8
    NewPubKey.pubexp = 65537 '0x00010001
    CopyMemory NewPubKey.modulus(1), bNewKey(0), Len(sPublicKey)
    CopyMemory PublicKey.bPublicKey(1), NewPubKey, Len(sPublicKey) + 20
    SetPubKey = True
End Function

Private Sub TLSDecrypt()
    Const Routine As String = "TLSDecrypt"
    Dim sBuffer As String
    lStatus = CFB_BUSY
    'Get handle to the provider
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, 0) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    sBuffer = sInBuffer
    'Decrypt data
    If CryptDecrypt(hReadKey, 0, False, 0, sBuffer, Len(sBuffer)) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
        GoTo ReleaseHandles
    End If
    If Not VerifyHMAC(sBuffer) Then
        RaiseEvent CryptError(&H80090002, "VerifyHMAC", Routine)
    End If
    sOutBuffer = sBuffer

ReleaseHandles:
    lStatus = CFB_READY
End Sub

Private Sub TLSEncrypt()
    Const Routine As String = "TLSEncrypt"
    Dim sBuffer As String
    Dim buflen As Long
    'switch Status property
    lStatus = CFB_BUSY
    'Get handle to the provider
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, 0) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    'Add HMAC to message
    'Choose either HMAC routine, as both should deliver the same result
    sBuffer = sInBuffer & HMAC(Alg_Hash, WRITE_MAC_KEY, sInBuffer, True)
    buflen = Len(sBuffer)
    'Null string returns required buffer length
    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, vbNullString, Len(sBuffer), buflen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    If buflen > Len(sBuffer) Then 'Adjust buffer length if necessary
        sBuffer = sBuffer & String$(buflen - Len(sBuffer), Chr$(0))
    End If
    'Encrypt data
    If CryptEncrypt(hWriteKey, 0, flgFinal, 0, sBuffer, Len(sBuffer), buflen) = 0 Then
        RaiseEvent CryptError(Err.LastDllError, CE, Routine)
        GoTo ReleaseHandles
    End If
    sOutBuffer = sBuffer
ReleaseHandles:
'    If hCryptProv Then CryptReleaseContext hCryptProv, 0
    lStatus = CFB_READY
End Sub

Private Function VerifyHMAC(ByRef DecryptedRecord As String) As Boolean
    Dim N%
    'Verify the Message Authentication Code
    Dim AppendedHMAC As String
    Dim CalculatedHMAC As String
    Dim lLen As Long
    If Alg_Hash = CALG_MD5 Then
        lLen = Len(DecryptedRecord) - 16
    Else
        lLen = Len(DecryptedRecord) - 20
    End If
    AppendedHMAC = Mid(DecryptedRecord, lLen + 1)
    DecryptedRecord = Mid(DecryptedRecord, 1, lLen)
    'Either HMAC routine should produce the correct result
    CalculatedHMAC = HMAC(Alg_Hash, READ_MAC_KEY, DecryptedRecord, False)
'    CalculatedHMAC = HMAC2(hReadMAC, DecryptedRecord, False)
    If CalculatedHMAC = AppendedHMAC Then
        VerifyHMAC = True
    Else
        VerifyHMAC = False
    End If
End Function

Private Function WriteFile(sFileName As String, bData() As Byte) As Boolean
    Dim lFileHandle As Long
    ' Writes a byte array out to a file.
    On Error GoTo ErrorHandler
    lFileHandle = FreeFile
    Open sFileName For Binary As lFileHandle
    Put lFileHandle, , bData
    Close lFileHandle
    WriteFile = True
    Exit Function
ErrorHandler:
    WriteFile = False
End Function

Private Function HMAC(HashAlg As Long, sKey As String, sMsg As String, flgSend As Boolean) As String
    
    'Note: if key is longer than the blocksize (64 for MD5/SHA1), the key should
    'be hashed first to reduce it's length
    Const Routine As String = "HMAC"
    Dim MAC_Header As String
    Dim o_key_pad As String
    Dim i_key_pad As String
    Dim N%
    If flgSend Then
        If Len(sMsg) >= 65536 Then
            MAC_Header = SendSeqNum & Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) _
            & Chr$(Int(Len(sMsg) / 65536)) & Chr$((Len(sMsg) Mod 65536) / 256) & Chr$(((Len(sMsg) Mod 65536) / 256) Mod 256)
        Else
            MAC_Header = SendSeqNum & Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) _
            & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
        End If
    Else
        If Len(sMsg) >= 65536 Then
            MAC_Header = RecvSeqNum & Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) _
            & Chr$(Int(Len(sMsg) / 65536)) & Chr$((Len(sMsg) Mod 65536) / 256) & Chr$(((Len(sMsg) Mod 65536) / 256) Mod 256)
        Else
            MAC_Header = RecvSeqNum & Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) _
            & Chr$(Int(Len(sMsg) / 256)) & Chr$(Len(sMsg) Mod 256)
        End If
    End If
    If Len(sKey) > 64 Then
        If HashAlg = CALG_MD5 Then
            sKey = MD5_Hash(sKey)
        ElseIf HashAlg = CALG_SHA1 Then
            sKey = SHA1_Hash(sKey)
        End If
    End If
    o_key_pad = String$(64, &H5C)
    i_key_pad = String$(64, &H36)
    For N% = 1 To Len(sKey)
        Mid$(o_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H5C)
        Mid$(i_key_pad, N%, 1) = Chr$(Asc(Mid$(sKey, N%, 1)) Xor &H36)
    Next N%
    If HashAlg = CALG_MD5 Then
        HMAC = MD5_Hash(o_key_pad & MD5_Hash(i_key_pad & MAC_Header & sMsg))
    ElseIf HashAlg = CALG_SHA1 Then
        HMAC = SHA1_Hash(o_key_pad & SHA1_Hash(i_key_pad & MAC_Header & sMsg))
    End If
End Function

Private Sub GenRandom(lLen As Long)
    Const Routine As String = "GenRandom"
    Dim sRandom As String
    If hCryptProv = 0 Then
        If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, 0) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
            GoTo ReleaseHandles
        End If
    End If
    sRandom = String$(lLen, Chr$(0))
    If CryptGenRandom(hCryptProv, lLen, sRandom) = 0 Then
        sRndBuffer = ""
    Else
        sRndBuffer = sRandom
    End If
ReleaseHandles:

End Sub

Private Function InitUser() As Long
    
    Const Routine As String = "InitUser"
    Dim hCryptKey As Long
    Dim avProviderData(1000) As Byte
    Dim lProviderDataLen As Long
    Dim lDataSize As Long
    Dim sUserName As String
    Dim lPoint As Long
    Dim sBuffer As String
    Dim strMsg As String
    sContainer = "JAC SSL Container" & vbNullChar
    sProvider = MS_DEF_RSA_SCHANNEL_PROV
    sProvType = PROV_RSA_SCHANNEL
    'Attempt to acquire a handle to the key container.
    If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, 0) = 0 Then
        'Create key container.
        If CryptAcquireContext(hCryptProv, ByVal sContainer, ByVal sProvider, sProvType, CRYPT_NEWKEYSET) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CAQ, Routine)
            Exit Function
        End If
        'Get name of key container.
        lProviderDataLen = Len(avProviderData(0)) * (UBound(avProviderData) + 1)
        If CryptGetProvParam(hCryptProv, PP_CONTAINER, avProviderData(0), lProviderDataLen, 0) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CGPP, Routine)
            avProviderData(0) = 0
        End If
        'Get sUserName from avProviderData()
        lPoint = LBound(avProviderData)
        While lPoint <= UBound(avProviderData)
            If avProviderData(lPoint) <> 0 Then
                sUserName = sUserName & Chr$(avProviderData(lPoint))
            Else
                lPoint = UBound(avProviderData)
            End If
            lPoint = lPoint + 1
        Wend
    End If
    'Attempt to get handle to exchange key
    If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hCryptKey) = 0 Then
        If Err.LastDllError = NTE_NO_KEY Then
            Debug.Print ("System does not have Key Exchange pair!" & vbCrLf & "One will be created.")
            ' The use of &H80000000 guarantees a 2048 bit key.
            If CryptGenKey(hCryptProv, AT_KEYEXCHANGE, &H8000000, hCryptKey) = 0 Then
                RaiseEvent CryptError(Err.LastDllError, CGK, Routine)
                Exit Function
            Else
                CryptDestroyKey hCryptProv
            End If
        Else
            RaiseEvent CryptError(Err.LastDllError, CGUK, Routine)
            Exit Function
        End If
    End If
    'Enumerate the supported algorithms.
    Dim flag As Long
    Dim Algid As Long
    Dim dwBits  As Long
    Dim namelen As Long
    Dim algType As String
    Dim LPTR As Long
    Dim Name As String
    flag = 1
    If CryptGetProvParam(hCryptProv, PP_ENUMALGS, ByVal vbNullString, LPTR, flag) = 0 Then
            RaiseEvent CryptError(Err.LastDllError, CGPP, Routine)
            Exit Function
        End If
    Dim b() As Byte
    Dim bName() As Byte
    ReDim b(LPTR - 1)
    strMsg = strMsg & vbCrLf & "ALGID" & vbTab & "dwBits" & vbTab & "algType" & String(2, vbTab) & "len" & vbTab & "Name" & vbCrLf _
        & String(5, "-") & vbTab & String(6, "-") & vbTab & String(12, "-") & vbTab & String(3, "-") & vbTab & String(3, "-") & vbCrLf
    While CryptGetProvParam(hCryptProv, PP_ENUMALGS, b(0), LPTR, flag) <> 0
        flag = 2
        CopyMemory Algid, b(0), LenB(Algid)
        CopyMemory dwBits, b(4), LenB(dwBits)
        CopyMemory namelen, b(8), LenB(namelen)
        ReDim bName(namelen - 2)
        CopyMemory bName(0), b(12), namelen - 1
        Name = StrConv(bName, vbUnicode)
        Select Case (Algid And &HF000)
            Case ALG_CLASS_ANY
                algType = "Any" & "         "
            Case ALG_CLASS_SIGNATURE
                algType = "Signature" & "   "
            Case ALG_CLASS_MSG_ENCRYPT
                algType = "Msg_Encrypt" & " "
            Case ALG_CLASS_DATA_ENCRYPT
                algType = "Data_Encrypt"
            Case ALG_CLASS_HASH
                algType = "Hash" & "        "
            Case ALG_CLASS_KEY_EXCHANGE
                algType = "Exchange" & "    "
            Case Else
                algType = "Unknwon" & "     "
        End Select
        strMsg = strMsg & Algid & vbTab & Right("     " & dwBits, 6) & vbTab & algType & vbTab & namelen & vbTab & Name & vbCrLf
    Wend
    InitUser = True
End Function

Private Function GetRandomString() As String
    Call GenRandom(28&)
    CLIENT_RANDOM = GetUnixTime & sRndBuffer
    '36 ciphers
    'Both TLS_RSA_WITH_RC4_128_MD5 & TLS_RSA_WITH_RC4_128_SHA
    GetRandomString = Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & CLIENT_RANDOM _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H5) _
        & Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(Len(SNI)) & SNI
    Mid$(GetRandomString, 3, 1) = Chr$(Int((Len(GetRandomString) - 4) / 256))
    Mid$(GetRandomString, 4, 1) = Chr$((Len(GetRandomString) - 4) Mod 256)
End Function

Private Function AddRecHeader(strRecord As String) As String
    'Prepend record with record header
    Dim TLSHeader As String
    If Len(strRecord) >= 65536 Then
        TLSHeader = Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) & _
        Chr$(Int(Len(strRecord) / 65536)) & Chr$((Len(strRecord) Mod 65536) / 256) & Chr$(((Len(strRecord) Mod 65536) / 256) Mod 256)
    Else
        TLSHeader = Chr$(sMsgType) & Chr$(sMsgVerMaj) & Chr$(sMsgVerMin) & _
        Chr$(Int(Len(strRecord) / 256)) & Chr$(Len(strRecord) Mod 256)
    End If
    AddRecHeader = TLSHeader & strRecord
End Function

Private Function GetUnixTime() As String
    Dim M%, N%
    Dim Temp$
    GetUnixTime = String$(4, vbNull)
    Temp$ = Hex(DateDiff("s", "01/01/1970", Date$))
    M% = 1
    For N% = 1 To Len(Temp$) Step 2
        Mid$(GetUnixTime, M%, 1) = Chr$(val("&H" & Mid$(Temp$, N%, 2)))
        M% = M% + 1
    Next N%
End Function

Private Function GetClientExchange() As String
    Dim sKey As String
    Dim sCipherSpec As String
    Dim sClientFinish As String
    Dim sClientFinishEnc As String
    'Create Master Key
    Let ClientRandom = CLIENT_RANDOM
    Let ServerRandom = SERVER_RANDOM
    'Get random Pre-Master Key
    Call GenRandom(46&)
    PRE_MASTER_KEY = Chr$(&H3) & Chr$(&H1) & sRndBuffer
    sInBuffer = PRE_MASTER_KEY
    'Set the server Public Key in clsCrypt
    Let PubKey = PUBLIC_KEY
    Call SetPubKey
    Call ExportMasterKey
    'Encrypted Pre-Master Key is now resident in OutBuffer
    '& decrypted Pre-Master Key is in InBuffer
    sKey = Chr$(TLS1_MT_CLIENT_KEY_EXCHANGE) & Chr$(&H0) & Chr$(&H1) & Chr$(&H2) _
        & Chr$(&H1) & Chr$(&H0) & StrReverse(sOutBuffer)
    sCipherSpec = Chr$(TLS1_RT_CHANGE_CIPHER_SPEC) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & Chr$(&H0) & Chr$(&H1) & Chr$(&H1)
    Call IncSendSeqNum(True)  'Reset Send Sequence Number
    'Hash data (Change Cipher Spec not included)
    hSESSION_MD5 = MD5_AddToHash(sKey, hSESSION_MD5)
    hSESSION_SHA1 = SHA1_AddToHash(sKey, hSESSION_SHA1)
    'Make a copy of the hashes
    hCOPY_MD5 = CopyHash(hSESSION_MD5)
    hCOPY_SHA1 = CopyHash(hSESSION_SHA1)
    'Concantenate copy of Hashes
    sInBuffer = MD5_GetHash(hCOPY_MD5) _
        & SHA1_GetHash(hCOPY_SHA1)
    'Hash Session Hashes
    Call Hash_SHAMD5(False)
    sClientFinish = sInBuffer
    sClientFinishEnc = sOutBuffer
    hSESSION_MD5 = MD5_AddToHash(sClientFinish, hSESSION_MD5)
    hSESSION_SHA1 = SHA1_AddToHash(sClientFinish, hSESSION_SHA1)
    GetClientExchange = AddRecHeader(sKey) & sCipherSpec & AddRecHeader(sClientFinishEnc)
End Function

Private Function VerifySignature(sBuff As String) As Boolean
    Dim CertData As String
    Dim CertKey As String
    Dim SigData As String
    Dim sSearch As String
    Dim lPntr1 As Long
    Dim lPntr2 As Long
    Dim lPntr3 As Long
    Dim lLen As Long
    'This routine returns the first part of the certificate only (w/o signature)
    lPntr1 = InStr(sBuff, Chr$(&H30) & Chr$(&H82)) + 4
    If lPntr1 < 5 Then Exit Function
    lPntr2 = InStr(lPntr1, sBuff, Chr$(&H30) & Chr$(&H82))
    lLen = Asc(Mid$(sBuff, lPntr2 + 2, 1)) * 256 + Asc(Mid$(sBuff, lPntr2 + 3, 1)) + 4
    CertData = Mid$(sBuff, lPntr2, lLen)
    lPntr2 = lPntr2 + lLen
    sInBuffer = CertData
    If lPntr1 > 0 Then
        RSA_KEY_LEN = 256
        sSearch = RSA_KEY_256V3
    Else
        lPntr1 = InStr(sBuff, RSA_KEY_128V3)
        If lPntr1 > 0 Then
            RSA_KEY_LEN = 128
            sSearch = RSA_KEY_128V3
        Else
            VerifySignature = False
            Exit Function
        End If
    End If
    'Save Server Public Key
    lPntr1 = InStr(lPntr1, sBuff, sSearch)
    If lPntr1 > 0 Then
        lPntr1 = lPntr1 + Len(sSearch)
        PUBLIC_KEY = StrReverse(Mid$(sBuff, lPntr1, RSA_KEY_LEN))
        'Find next certificate Public Key
        lPntr1 = InStr(lPntr1, sBuff, sSearch)
        If lPntr1 > 0 Then 'Use key from next Cert in chain
            CertKey = StrReverse(Mid$(sBuff, lPntr1 + Len(sSearch), RSA_KEY_LEN))
        Else 'Use Server Public key
            CertKey = PUBLIC_KEY
        End If
    Else
        VerifySignature = False
        Exit Function
    End If
    Let PubKey = CertKey
    If Not SetPubKey Then
        VerifySignature = False
        Exit Function
    End If
    If RSA_KEY_LEN = 256 Then
        sSearch = Chr$(&H3) & Chr$(&H82) & Chr$(&H1) & Chr$(&H1) & Chr$(&H0)
    ElseIf RSA_KEY_LEN = 128 Then
        sSearch = Chr$(&H3) & Chr$(&H81) & Chr$(&H81) & Chr$(&H0)
    End If
    lPntr3 = InStr(lPntr2, sBuff, sSearch)
    If lPntr3 > 0 Then
        lPntr3 = lPntr3 + Len(sSearch)
        SigData = StrReverse(Mid$(sBuff, lPntr3, RSA_KEY_LEN))
    Else
        VerifySignature = False
        Exit Function
    End If
    Let Signature = SigData
    VerifySignature = VerifySig
End Function

Private Function GetAttachmentHeader(ByVal sFil As String) As String
  Dim s As String, sExt1 As String, Sid As String, sCT1 As String
  Dim Pt1 As Long
  Dim Boo1 As Boolean
   On Error Resume Next
   Pt1 = InStrRev(sFil, cDOT)
   If (Pt1 > 0) Then
       sExt1 = Right$(sFil, (Len(sFil) - (Pt1 - 1)))  '-- get  ".xxx"
       Boo1 = RegGetString(HKEY_CLASSES_ROOT, sExt1, "Content Type", sCT1)
   End If
   If (Len(sCT1) = 0) Then sCT1 = "application/octet-stream"
   s = cDASH2 & cBoundaryA & vbCrLf
   s = s & "Content-Type: image/gif;" & sCT1 & cSEMIC & vbCrLf
   s = s & vbTab & "Name=" & Qt2 & sFil & Qt2 & vbCrLf
   s = s & "Content-Transfer-Encoding: base64" & vbCrLf
   s = s & "Content-Disposition: attachment;" & vbCrLf
   s = s & vbTab & "Filename=" & Qt2 & sFil & Qt2 & vbCrLf & vbCrLf
   Debug.Print s
   GetAttachmentHeader = s
End Function

Private Function RegGetString(ByVal KeyH As Long, sPath As String, sValName As String, sReturn As String) As Boolean
 Dim hKey As Long, lRet As Long, LBuf As Long, LType As Long
 Dim sBuf As String, sRet As String
  On Error Resume Next
     lRet = RegOpenKeyEx(ByVal KeyH, sPath, 0&, KEY_QUERY_VALUE, hKey)
        If (lRet <> 0) Then Exit Function
     LBuf = 255
     sBuf = String$(LBuf, 0)
       lRet = RegQueryValueExString(hKey, sValName, 0&, LType, sBuf, LBuf)
          If (lRet = 0) And (LBuf > 1) Then
             sReturn = Left$(sBuf, (LBuf - 1))
             RegGetString = True
          End If
    lRet = RegCloseKey(hKey)
End Function

Private Function GetDateString() As String
  Dim ST1 As SYSTEMTIME
  Dim sBuf As String
  Dim Pt1 As Long, lRet As Long
     On Error Resume Next
   GetSystemTime ST1
    sBuf = String$(30, 0)
  lRet = InternetTimeFromSystemTime(ST1, INTERNET_RFC1123_FORMAT, sBuf, INTERNET_RFC1123_BUFSIZE)
     If (lRet <> 0) Then
         Pt1 = InStr(1, sBuf, cNULL)
           If (Pt1 > 0) Then sBuf = Left$(sBuf, (Pt1 - 1))
        GetDateString = "Date: " & sBuf
     Else
        GetDateString = "Date: " & CStr(Now)
     End If
End Function

Private Function GetGUID() As String
 Dim ID1 As guid
 Dim lRet As Long, lLen As Long
 Dim Sid As String
     On Error Resume Next
   lRet = CoCreateGuid(ID1)
     If (lRet = 0) Then
       lLen = 255
       Sid = String$(lLen, 0)
         lRet = StringFromGUID2(ID1, StrPtr(Sid), lLen)
           If (lRet > 30) Then
             Sid = Left$(Sid, (lRet - 2))  '--snip off { and }
             Sid = Right$(Sid, (Len(Sid) - 1))
             Sid = Replace(Sid, cDASH, cDOT)
             GetGUID = Sid
           End If
     End If
End Function



