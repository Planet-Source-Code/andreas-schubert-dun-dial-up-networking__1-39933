Attribute VB_Name = "modDUN"
' **********************************************************************
' Modul : Dial-up networking capabilities for VB
' Author: Andreas Schubert
' Date  : 2002 / 09 / 15
'
' You have the royalty free right to use this module in your applications
' as long as you provide this copyright notice with it
' If you find bugs or change the code, please e-mail me a copy of it to
' Andy@andreas-schubert.net
' I would like to know where this was used, so if you integrate this in your applications,
' I would be glad if you give me a short note.
'
' Of course, I take no responsibility for any damage the use of this code could do.
' Use at your own risk!
'
' Visit also my new homepage at www.andreas-schubert.net!
' I am truly sorry for the -not so long - sample project, but I've got loads of work to do.
' So, If you have any questions, feel free to send me a mail, I will try to answer all questions!
' **********************************************************************

' Declares from RASAPI32.DLL
Public Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" _
    (ByVal lpStrNull As String, ByVal lpszPhonebook As String, _
    lpRasEntryName As Any, lpCb As Long, lpCEntries As Long) As Long

Public Declare Function RasRenameEntry Lib "rasapi32.dll" Alias "RasRenameEntryA" _
        (ByVal lpszPhonebook As String, ByVal lpszOldEntry As String, _
        ByVal lpszNewEntry As String) As Long

Public Declare Function RasDeleteEntry Lib "rasapi32.dll" Alias "RasDeleteEntryA" _
        (ByVal lpszPhonebook As String, ByVal lpszEntry As String) As Long

Public Declare Function RasValidateEntryName Lib "rasapi32.dll" Alias "RasValidateEntryNameA" _
        (ByVal lpszPhonebook As String, ByVal lpszEntry As String) As Long


Public Declare Function RasCreatePhonebookEntry Lib "rasapi32.dll" Alias "RasCreatePhonebookEntryA" _
        (ByVal hwnd As Long, ByVal lpszPhonebook As String) As Long

Public Declare Function RasEditPhonebookEntry Lib "rasapi32.dll" Alias "RasEditPhonebookEntryA" _
        (ByVal hwnd As Long, ByVal lpszPhonebook As String, _
        ByVal lpszEntryName As String) As Long


Public Declare Function RasGetEntryProperties Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" _
       (ByVal lpszPhonebook As String, ByVal lpszEntry As String, _
        lpRasEntry As Any, lpdwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, lpdwDeviceInfoSize As Long) As Long

Public Declare Function RasSetEntryProperties Lib "rasapi32.dll" Alias "RasSetEntryPropertiesA" _
        (ByVal lpszPhonebook As String, ByVal lpszEntry As String, _
        lpRasEntry As Any, ByVal dwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, ByVal dwDeviceInfoSize As Long) As Long


Public Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, lpRasDialParams As Any, _
        blnPasswordRetrieved As Long) As Long

Public Declare Function RasSetEntryDialParams Lib "rasapi32.dll" Alias "RasSetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, lpRasDialParams As Any, _
        ByVal blnRemovePassword As Long) As Long

Public Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" _
      (lpRasDialExtensions As Any, ByVal lpszPhonebook As String, _
       lpRasDialParams As Any, ByVal dwNotifierType As Long, _
       ByVal hwndNotifier As Long, lphRasConn As Long) As Long


Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
   


Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" _
         (lpRasconn As Any, lpCb As Long, lpcConnections As Long) As Long
          

Public Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" _
        (ByVal hRasConn As Long, lpRasConnStatus As Any) As Long


Public Declare Function RasGetProjectionInfo Lib "rasapi32.dll" Alias "RasGetProjectionInfoA" _
            (ByVal hRasConn As Long, ByVal rasProjectionType As Long, _
             lpProjection As Any, lpCb As Long) As Long

Public Declare Function RasEnumDevices Lib "rasapi32.dll" Alias "RasEnumDevicesA" ( _
        lpRasDevInfo As Any, lpCb As Long, _
        lpCDevices As Long) As Long

Public Declare Function RasGetErrorString Lib "rasapi32.dll" Alias "RasGetErrorStringA" _
      (ByVal uErrorValue As Long, ByVal lpszErrorString As String, _
       cBufSize As Long) As Long

' end declarations from rasapi32.dll
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' Declarations from kernel32.dll
Public Declare Function FormatMessage _
     Lib "kernel32" Alias "FormatMessageA" _
      (ByVal dwFlags As Long, lpSource As Any, _
       ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
       ByVal lpBuffer As String, ByVal nSize As Long, _
       Arguments As Long) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' end declarations from kernel32.dll
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Constants
Const ERROR_INVALID_HANDLE = 6&
' end constants
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Data types:

' RasEntryname structure getting filled by RasEnumEntries
Public Type tRasEntryName
   EntryName As String
   Win2000_SystemPhonebook As Boolean
   PhonebookPath As String
End Type

' RasDialParams is used in the RasDial function to specify the entry to dial
Public Type tRasDialParams
    EntryName As String
    PhoneNumber As String
    CallbackNumber As String
    UserName As String
    Password As String
    Domain As String
    SubEntryIndex As Long
    RasDialFunc2CallbackId As Long
End Type

' **********************************************************************
' Enums needed by RASEntry
Public Type RASIPADDR
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type

Public Enum RasEntryOptions
   RASEO_UseCountryAndAreaCodes = &H1
   RASEO_SpecificIpAddr = &H2
   RASEO_SpecificNameServers = &H4
   RASEO_IpHeaderCompression = &H8
   RASEO_RemoteDefaultGateway = &H10
   RASEO_DisableLcpExtensions = &H20
   RASEO_TerminalBeforeDial = &H40
   RASEO_TerminalAfterDial = &H80
   RASEO_ModemLights = &H100
   RASEO_SwCompression = &H200
   RASEO_RequireEncryptedPw = &H400
   RASEO_RequireMsEncryptedPw = &H800
   RASEO_RequireDataEncryption = &H1000
   RASEO_NetworkLogon = &H2000
   RASEO_UseLogonCredentials = &H4000
   RASEO_PromoteAlternates = &H8000
   RASEO_SecureLocalFiles = &H10000
   RASEO_RequireEAP = &H20000
   RASEO_RequirePAP = &H40000
   RASEO_RequireSPAP = &H80000
   RASEO_Custom = &H100000
   RASEO_PreviewPhoneNumber = &H200000
   RASEO_SharedPhoneNumbers = &H800000
   RASEO_PreviewUserPw = &H1000000
   RASEO_PreviewDomain = &H2000000
   RASEO_ShowDialingProgress = &H4000000
   RASEO_RequireCHAP = &H8000000
   RASEO_RequireMsCHAP = &H10000000
   RASEO_RequireMsCHAP2 = &H20000000
   RASEO_RequireW95MSCHAP = &H40000000
   RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
   RASNP_NetBEUI = &H1
   RASNP_Ipx = &H2
   RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
   RASFP_Ppp = &H1
   RASFP_Slip = &H2
   RASFP_Ras = &H4
End Enum
'end enums for RasEntry
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' translated RasEntry structure
' for details see Platform SDK
' the RasEntry structure is used by RASGetEntryProperties and RASSetEntryProperties
Public Type tRasEntry
   Options As RasEntryOptions
   CountryID As Long
   CountryCode As Long
   AreaCode As String
   LocalPhoneNumber As String
   AlternateNumbers As String
   ipAddr As RASIPADDR
   ipAddrDns As RASIPADDR
   ipAddrDnsAlt As RASIPADDR
   ipAddrWins As RASIPADDR
   ipAddrWinsAlt As RASIPADDR
   FrameSize As Long
   fNetProtocols As RASNetProtocols
   FramingProtocol As RasFramingProtocols
   ScriptName As String
   AutodialDll As String
   AutodialFunc As String
   DeviceType As String
   DeviceName As String
   X25PadType As String
   X25Address As String
   X25Facilities As String
   X25UserData As String
   Channels As Long
   NT4En_SubEntries As Long
   NT4En_DialMode As Long
   NT4En_DialExtraPercent As Long
   NT4En_DialExtraSampleSeconds As Long
   NT4En_HangUpExtraPercent As Long
   NT4En_HangUpExtraSampleSeconds As Long
   NT4En_IdleDisconnectSeconds As Long
   Win2000_Type As Long
   Win2000_EncryptionType As Long
   Win2000_CustomAuthKey As Long
   Win2000_guidId(0 To 15) As Byte
   Win2000_CustomDialDll As String
   Win2000_VpnStrategy As Long
End Type



' Type RasConnStatus
' is used by RASGetConnectStatus

Type tRASCONNSTATUS
      lRasConnState As RASCONNSTATE
      dwError As Long
      sDeviceType As String
      sDeviceName As String
      sNTPhoneNumber As String
End Type


' Enum RasConnstate
' states that may occur during a RAS connection operation
' when dialing asyncronous, the state is passed to the callback function
' can also be used by calling RASGetConnecStatus
Enum RASCONNSTATE
   RASCS_OpenPort = 0
   RASCS_PortOpened = 1
   RASCS_ConnectDevice = 2
   RASCS_DeviceConnected = 3
   RASCS_AllDevicesConnected = 4
   RASCS_Authenticate = 5
   RASCS_AuthNotify = 6
   RASCS_AuthRetry = 7
   RASCS_AuthCallback = 8
   RASCS_AuthChangePassword = 9
   RASCS_AuthProject = 10
   RASCS_AuthLinkSpeed = 11
   RASCS_AuthAck = 12
   RASCS_ReAuthenticate = 13
   RASCS_Authenticated = 14
   RASCS_PrepareForCallback = 15
   RASCS_WaitForModemReset = 16
   RASCS_WaitForCallback = 17
   RASCS_Projected = 18
   RASCS_StartAuthentication = 19
   RASCS_CallbackComplete = 20
   RASCS_LogonNetwork = 21
   RASCS_SubEntryConnected = 22
   RASCS_SubEntryDisconnected = 23
   RASCS_Interactive = &H1000
   RASCS_RetryAuthentication = &H1001
   RASCS_CallbackSetByCaller = &H1002
   RASCS_PasswordExpired = &H1003
   RASCS_InvokeEapUI = &H1004
   RASCS_Connected = &H2000
   RASCS_Disconnected = &H2001
End Enum

' end enum RasConnState
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' type RasConn
' used by RasEnumConnections
Type tRASCONN
   hRasConn As Long
   sEntryName As String
   sDeviceType As String
   sDeviceName As String
   sPhonebook  As String
   lngSubEntry As Long
   guidEntry(15) As Byte
End Type

'The RASAMB projection info describes the result of a RAS AMB (Authentication Message Block)  projection. This protocol is used with NT 3.1 and OS/2 1.3 downlevel RAS servers.
Type tRASAMB
      dwError As Long
      sNetBiosError As String
      bLana As Byte
End Type

' describes result of a PPP NBF (NetBEUI) projection
Type tRASPPPNBF
    dwError As Long
    dwNetBiosError As Long
    szNetBiosError As String
    szWorkstationName As String
    bLana As Byte
End Type

' describes results of a PPP IPX (Internetwork Packet Exchange) projection.
Type tRASPPPIPX
    dwError As Long
    szIpxAddress As String
End Type

' describes results of a PPP IP (Internet) projection
Type tRASPPPIP
    dwError As Long
    szIpAddress As String
    szServerIpAddress As String
End Type

' describes the results of a SLIP (Serial Line IP) projection
Type tRASSLIP
    dwError As Long
    szIpAddress As String
End Type


' The RASEnumDevices function returns a list of all RAS capable devices, with name and type
' The returned names and device types are stored in a RASDEVINFO structure
Public Type tRASDEVINFO
   DeviceType As String
   DeviceName As String
End Type


' end data types and enums
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasErrorHandler
' Purpose: returns a RAS errorcode as plaintext
' Input:   an Errorcode you received
' Return:  plaintext according to the errorcode
' usage:
'          retVal = yourRASFunction()
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
' **********************************************************************
Function fRASErrorHandler(retVal As Long) As String
   Dim strError As String, i As Long
   strError = String(512, 0)
   If retVal > 600 Then
      RasGetErrorString retVal, strError, 512&
   Else
      FormatMessage &H1000, ByVal 0&, retVal, 0&, strError, 512, ByVal 0&
   End If
   i = InStr(strError, Chr$(0))
   If i > 1 Then fRASErrorHandler = Left$(strError, i - 1)
End Function
'fRasErrorHandler ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasGetAllEntries
' Purpose: Retrieve a list of all dial-up connections on the machine
' Return:  Returns the number of entries
' usage:
'          Dim myEntries() as tRasEntryName
'          Dim lngCount as long
'          lngCount = fRasGetAllEntries(myEntries)
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasGetAllEntries(clsRasEntryName() As tRasEntryName, _
                  Optional strPhonebook As String) As Long
   
Dim retVal             As Long
Dim i               As Long
Dim lpCb            As Long   'count of bytes
Dim lpCEntries      As Long  'count of entries
Dim b()             As Byte
Dim strTemp         As String
Dim dwSize          As Long 'size of each entry
Dim lngLen          As Long
Dim lngBLen         As Variant
ReDim b(3)
   'determine appropiate size for b()
   lngBLen = Array(532&, 264&, 28&)
   For i = 0 To 2
        CopyMemory b(0), CLng(lngBLen(i)), 4
        retVal = RasEnumEntries(vbNullString, strPhonebook, b(0), lpCb, lpCEntries)
        If retVal <> 632 Then Exit For
   Next i
   
   fRasGetAllEntries = lpCEntries
   If lpCEntries = 0 Then Exit Function
   
   dwSize = lpCb \ lpCEntries
   
   ReDim b(lpCb - 1)
   CopyMemory b(0), dwSize, 4
   
   retVal = RasEnumEntries(vbNullString, strPhonebook, b(0), lpCb, lpCEntries)
   If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
   
   strTemp = String(dwSize - 4, 0)
   ReDim clsRasEntryName(lpCEntries - 1)
   
   If dwSize = 28 Then lngLen = 21 Else lngLen = 257
   
   For i = 0 To lpCEntries - 1
         CopyMemory ByVal strTemp, b((i * dwSize) + 4), lngLen
         clsRasEntryName(i).EntryName = _
         Left(strTemp, InStr(strTemp, Chr$(0)) - 1)
   Next i
   
   If dwSize > 264 Then
        For i = 0 To lpCEntries - 1
            CopyMemory clsRasEntryName(i).Win2000_SystemPhonebook, b((i * dwSize) + 264), 2&
            CopyMemory ByVal strTemp, b((i * dwSize) + 268), 260&
            clsRasEntryName(i).PhonebookPath = _
            Left(strTemp, InStr(strTemp, Chr$(0)) - 1)
        Next i
   Else
       For i = 0 To lpCEntries - 1
            clsRasEntryName(i).PhonebookPath = strPhonebook
       Next i
   End If
End Function
'fRasGetAllEntries ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fRasGetEntryProperties
' Purpose: Returns detailed information about an existing entry
' Return:  0 if the function succeded, elsewise the errorcode
'          a filled tRasEntry structure
' usage:
'          Dim clsRasEntry as tRasEntry
'          Dim retVal as long
'          retVal = fRasGetEntryProperties("MyConnection", clsRasEntry)
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
'
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasGetEntryProperties(strEntryName As String, clsRasEntry As tRasEntry, _
         Optional strPhonebook As String) As Long
   
Dim retVal         As Long
Dim lngCb       As Long
Dim lngBuffLen  As Long
Dim b()         As Byte
Dim lngPos      As Long
Dim lngStrLen   As Long
   retVal = RasGetEntryProperties(vbNullString, vbNullString, ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
   
   retVal = RasGetEntryProperties(strPhonebook, strEntryName, ByVal 0&, lngBuffLen, ByVal 0&, ByVal 0&)
   
   If retVal <> 603 Then fRasGetEntryProperties = retVal: Exit Function
   
   ReDim b(lngBuffLen - 1)
   CopyMemory b(0), lngCb, 4
   
   retVal = RasGetEntryProperties(strPhonebook, strEntryName, b(0), lngBuffLen, ByVal 0&, ByVal 0&)
   
   fRasGetEntryProperties = retVal
   If retVal <> 0 Then Exit Function
   
   CopyMemory clsRasEntry.Options, b(4), 4
   CopyMemory clsRasEntry.CountryID, b(8), 4
   CopyMemory clsRasEntry.CountryCode, b(12), 4
   CopyByteToTrimmedString clsRasEntry.AreaCode, b(16), 11
   CopyByteToTrimmedString clsRasEntry.LocalPhoneNumber, b(27), 129
   
   CopyMemory lngPos, b(156), 4
   If lngPos <> 0 Then
     lngStrLen = lngBuffLen - lngPos
     clsRasEntry.AlternateNumbers = String(lngStrLen, 0)
     CopyMemory ByVal clsRasEntry.AlternateNumbers, b(lngPos), lngStrLen
   End If
   
   CopyMemory clsRasEntry.ipAddr, b(160), 4
   CopyMemory clsRasEntry.ipAddrDns, b(164), 4
   CopyMemory clsRasEntry.ipAddrDnsAlt, b(168), 4
   CopyMemory clsRasEntry.ipAddrWins, b(172), 4
   CopyMemory clsRasEntry.ipAddrWinsAlt, b(176), 4
   CopyMemory clsRasEntry.FrameSize, b(180), 4
   CopyMemory clsRasEntry.fNetProtocols, b(184), 4
   CopyMemory clsRasEntry.FramingProtocol, b(188), 4
   CopyByteToTrimmedString clsRasEntry.ScriptName, b(192), 260
   CopyByteToTrimmedString clsRasEntry.AutodialDll, b(452), 260
   CopyByteToTrimmedString clsRasEntry.AutodialFunc, b(712), 260
   CopyByteToTrimmedString clsRasEntry.DeviceType, b(972), 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyByteToTrimmedString clsRasEntry.DeviceName, b(989), lngStrLen
      lngPos = 989 + lngStrLen
   CopyByteToTrimmedString clsRasEntry.X25PadType, b(lngPos), 33
      lngPos = lngPos + 33
   CopyByteToTrimmedString clsRasEntry.X25Address, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25Facilities, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25UserData, b(lngPos), 201
      lngPos = lngPos + 203
   CopyMemory clsRasEntry.Channels, b(lngPos), 4
   
   If lngCb > 1768 Then 'NT4 Enhancements & Win2000
      CopyMemory clsRasEntry.NT4En_SubEntries, b(1768), 4
      CopyMemory clsRasEntry.NT4En_DialMode, b(1772), 4
      CopyMemory clsRasEntry.NT4En_DialExtraPercent, b(1776), 4
      CopyMemory clsRasEntry.NT4En_DialExtraSampleSeconds, b(1780), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraPercent, b(1784), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraSampleSeconds, b(1788), 4
      CopyMemory clsRasEntry.NT4En_IdleDisconnectSeconds, b(1792), 4
      
      If lngCb > 1796 Then ' Win2000
         CopyMemory clsRasEntry.Win2000_Type, b(1796), 4
         CopyMemory clsRasEntry.Win2000_EncryptionType, b(1800), 4
         CopyMemory clsRasEntry.Win2000_CustomAuthKey, b(1804), 4
         CopyMemory clsRasEntry.Win2000_guidId(0), b(1808), 16
         CopyByteToTrimmedString _
                  clsRasEntry.Win2000_CustomDialDll, b(1824), 260
         CopyMemory clsRasEntry.Win2000_VpnStrategy, b(2084), 4
      End If
      
   End If
   
End Function



' **********************************************************************
' Function fRasSetEntryProperties
' Purpose: Set information for an existing entry or creates a new entry
' Return:  0 if the function succeded, elsewise the errorcode
' Input:   a valid tRasEntry structure
' usage:
'          Dim clsRasEntry as tRasEntry
'          Dim retVal as long
'          retVal = fRasSetEntryProperties("MyConnection", clsRasEntry)
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
'
' Note: you will probably have to specify some values of the tRasEntry structur
'       in order to successfully create an entry
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasSetEntryProperties(strEntryName As String, clsRasEntry As tRasEntry, _
         Optional strPhonebook As String) As Long
Dim retVal             As Long
Dim lngCb           As Long
Dim lngBuffLen      As Long
Dim b()             As Byte
Dim lngPos          As Long
Dim lngStrLen       As Long
   
   retVal = RasGetEntryProperties(vbNullString, vbNullString, ByVal 0&, lngCb, ByVal 0&, ByVal 0&)

   If retVal <> 603 Then fRasSetEntryProperties = retVal: Exit Function
   
   lngStrLen = Len(clsRasEntry.AlternateNumbers)
   lngBuffLen = lngCb + lngStrLen + 1
   ReDim b(lngBuffLen)
   
   CopyMemory b(0), lngCb, 4
   CopyMemory b(4), clsRasEntry.Options, 4
   CopyMemory b(8), clsRasEntry.CountryID, 4
   CopyMemory b(12), clsRasEntry.CountryCode, 4
   CopyStringToByte b(16), clsRasEntry.AreaCode, 11
   CopyStringToByte b(27), clsRasEntry.LocalPhoneNumber, 129
   
   If lngStrLen > 0 Then
     CopyMemory b(lngCb), ByVal clsRasEntry.AlternateNumbers, lngStrLen
     CopyMemory b(156), lngCb, 4
   End If

   CopyMemory b(160), clsRasEntry.ipAddr, 4
   CopyMemory b(164), clsRasEntry.ipAddrDns, 4
   CopyMemory b(168), clsRasEntry.ipAddrDnsAlt, 4
   CopyMemory b(172), clsRasEntry.ipAddrWins, 4
   CopyMemory b(176), clsRasEntry.ipAddrWinsAlt, 4
   CopyMemory b(180), clsRasEntry.FrameSize, 4
   CopyMemory b(184), clsRasEntry.fNetProtocols, 4
   CopyMemory b(188), clsRasEntry.FramingProtocol, 4
   CopyStringToByte b(192), clsRasEntry.ScriptName, 260
   CopyStringToByte b(452), clsRasEntry.AutodialDll, 260
   CopyStringToByte b(712), clsRasEntry.AutodialFunc, 260
   CopyStringToByte b(972), clsRasEntry.DeviceType, 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyStringToByte b(989), clsRasEntry.DeviceName, lngStrLen
      lngPos = 989 + lngStrLen
   CopyStringToByte b(lngPos), clsRasEntry.X25PadType, 33
      lngPos = lngPos + 33
   CopyStringToByte b(lngPos), clsRasEntry.X25Address, 201
      lngPos = lngPos + 201
   CopyStringToByte b(lngPos), clsRasEntry.X25Facilities, 201
      lngPos = lngPos + 201
   CopyStringToByte b(lngPos), clsRasEntry.X25UserData, 201
      lngPos = lngPos + 203
   CopyMemory b(lngPos), clsRasEntry.Channels, 4
   
   If lngCb > 1768 Then 'NT4 Enhancements & Win2000
      CopyMemory b(1768), clsRasEntry.NT4En_SubEntries, 4
      CopyMemory b(1772), clsRasEntry.NT4En_DialMode, 4
      CopyMemory b(1776), clsRasEntry.NT4En_DialExtraPercent, 4
      CopyMemory b(1780), clsRasEntry.NT4En_DialExtraSampleSeconds, 4
      CopyMemory b(1784), clsRasEntry.NT4En_HangUpExtraPercent, 4
      CopyMemory b(1788), clsRasEntry.NT4En_HangUpExtraSampleSeconds, 4
      CopyMemory b(1792), clsRasEntry.NT4En_IdleDisconnectSeconds, 4
      
      If lngCb > 1796 Then ' Win2000
         CopyMemory b(1796), clsRasEntry.Win2000_Type, 4
         CopyMemory b(1800), clsRasEntry.Win2000_EncryptionType, 4
         CopyMemory b(1804), clsRasEntry.Win2000_CustomAuthKey, 4
         CopyMemory b(1808), clsRasEntry.Win2000_guidId(0), 16
         CopyStringToByte b(1824), clsRasEntry.Win2000_CustomDialDll, 260
         CopyMemory b(2084), clsRasEntry.Win2000_VpnStrategy, 4
      End If
      
   End If
   
   retVal = RasSetEntryProperties(strPhonebook, strEntryName, b(0), lngCb, ByVal 0&, ByVal 0&)
   
   fRasSetEntryProperties = retVal

End Function
'fRasSetEntryProperties ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fRasGetEntryDialParams
' Purpose: Get dial information for an existing entry
' Return:  0 if the function succeded, elsewise the errorcode
'          a bytearray containing the dial information
' usage:
'          Dim b() as Byte
'          Dim retVal as long
'          retVal = fRasGetEntryDialParams(b, vbNullString, "MyConnection")
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
'          the bytearray can be transformed with fBytesToRasDialParams
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasGetEntryDialParams(bytesOut() As Byte, strPhonebook As String, strEntryName As String, _
               Optional blnPasswordRetrieved As Boolean) As Long
Dim retVal         As Long
Dim blnPsswrd   As Long
Dim bLens       As Variant
Dim lngLen      As Long
Dim i           As Long
   
   bLens = Array(1060&, 1052&, 816&)
   'try out the three different sizes for RasDialParams
   For i = 0 To 2
      lngLen = bLens(i)
      ReDim bytesOut(lngLen - 1)
      CopyMemory bytesOut(0), lngLen, 4
      If lngLen = 816& Then
         CopyStringToByte bytesOut(4), strEntryName, 20
      Else
         CopyStringToByte bytesOut(4), strEntryName, 257
      End If
      retVal = RasGetEntryDialParams(strPhonebook, bytesOut(0), blnPsswrd)
      If retVal = 0 Then Exit For
   Next i
   
   blnPasswordRetrieved = blnPsswrd
   fRasGetEntryDialParams = retVal
End Function
'fRasGetEntryDialParams ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Function fRasSetEntryDialParams
' Purpose: Set dial information for an existing entry
' Return:  0 if the function succeded, elsewise the errorcode
' Input    a bytearray containing the dial information
' usage:
'          Dim b() as Byte
'          Dim retVal as long
'          Dim tDialP As tRasDialParams
'          tdialp.UserName = "XXX"  ' set your information here
'          retval=fRasDialParamsToBytes(tdialp,b)
'          retVal = fRasSetEntryDialParams(vbNullString, b)
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasSetEntryDialParams(strPhonebook As String, bytesIn() As Byte, blnRemovePassword As Boolean) As Long
   
   fRasSetEntryDialParams = RasSetEntryDialParams _
               (strPhonebook, bytesIn(0), blnRemovePassword)
End Function
 



' **********************************************************************
' Function fBytesToRasDialParams
' Purpose: Transform the bytearray received by fRasGetEntryDialParams into a user-friendly structure
' Return:  True if the function succeded, elsewise false
'          a tRasdialParams structure
' Input:   a bytearray containing the dial information
' usage:
'          Dim b as Byte()
'          Dim retVal as long
'          Dim RasDialParams as tRasDialParams
'          retVal = fRasGetEntryDialParams(b, vbNullString, "MyConnection")
'          If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
'          if fBytesToRasDialParams(b, rasdialparams) then
'           ' do whatever you want here
'          else
'             msgbox "Structure could not be transformed!",vbError
'          endif
' **********************************************************************
Function fBytesToRasDialParams(bytesIn() As Byte, udtRasDialParamsOUT As tRasDialParams) As Boolean
   
Dim iPos        As Long
Dim lngLen      As Long
Dim dwSize      As Long
Dim a           As String

   On Error GoTo badBytes
   
   CopyMemory dwSize, bytesIn(0), 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   iPos = 4
 
   With udtRasDialParamsOUT
      CopyByteToTrimmedString .EntryName, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .PhoneNumber, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .CallbackNumber, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .UserName, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .Password, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 16
      CopyByteToTrimmedString .Domain, bytesIn(iPos), lngLen
      
      If dwSize > 1052& Then
         CopyMemory .SubEntryIndex, bytesIn(1052), 4&
         CopyMemory .RasDialFunc2CallbackId, bytesIn(1056), 4&
      End If
   End With
   fBytesToRasDialParams = True
   Exit Function
badBytes:
   'error handling goes here ??
   fBytesToRasDialParams = False
End Function
'fBytesToRasDialParams ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasDialParamsToBytes
' Purpose: transfer a tRasDialParams structure into a bytearray
' Return:  True if the function succeded, elsewise false
'          a bytearray
' Input:   a tRasDialParams structure containing the dial information

' **********************************************************************
Function fRasDialParamsToBytes(udtRasDialParamsIN As tRasDialParams, bytesOut() As Byte) As Boolean
   
Dim retVal         As Long
Dim blnPsswrd   As Long
Dim b()         As Byte
Dim bLens       As Variant
Dim dwSize      As Long
Dim i           As Long
Dim iPos        As Long
Dim lngLen      As Long
   
   bLens = Array(1060&, 1052&, 816&)
   For i = 0 To 2
      dwSize = bLens(i)
      ReDim b(dwSize - 1)
      CopyMemory b(0), dwSize, 4
      retVal = RasGetEntryDialParams(vbNullString, b(0), blnPsswrd)
      If retVal = 623& Then Exit For
   Next i
   
   If retVal <> 623& Then Exit Function
   
   On Error GoTo badBytes
   ReDim bytesOut(dwSize - 1)
   CopyMemory bytesOut(0), dwSize, 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   iPos = 4
   With udtRasDialParamsIN
      CopyStringToByte bytesOut(iPos), .EntryName, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .PhoneNumber, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .CallbackNumber, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .UserName, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .Password, lngLen
      iPos = iPos + lngLen: lngLen = 16
      CopyStringToByte bytesOut(iPos), .Domain, lngLen
      
      If dwSize > 1052& Then
         CopyMemory bytesOut(1052), .SubEntryIndex, 4&
         CopyMemory bytesOut(1056), .RasDialFunc2CallbackId, 4&
      End If
   End With
   fRasDialParamsToBytes = True
   Exit Function
badBytes:
   'error handling goes here ??
   fRasDialParamsToBytes = False
End Function
'fRasDialParamsToBytes ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fSyncronousDial
' Purpose: dial a entry syncronous (meaning wait until conn is established or failed
' Return:  the connection handle if the function succeedes, elsewise 0
' Input:   the name of the entry to dial
' Remember that if the method returns a non zero connection handle, you must hangup that connection even if the connection fails.
'
' Usage:
'       simple:
'           Dim hConn As Long
'           hConn = fSyncronousDial(vbNullString, "My Connection")
'           if hconn = 0 then
'               msgbox "Connection not established!"
'           endif
'       sample with altered data:
'           Dim retVal As Long
'           Dim b() As Byte
'           Dim myDialParams As tRasDialParams
'           Dim lngHConn As Long
'           Dim strPhonebook As String
'
'           With myDialParams
'           .EntryName = "My Connection"
'           .UserName = "Me"
'           .Password = "password"
'           .PhoneNumber = "0800 123456789"
'           End With
'           retVal = fRasDialParamsToBytes(myDialParams, b)
'           if retVal <>0 then
'               fRASErrorHandler (retVal)
'           else
'               code pauses on next line until connection established or fails
'               retVal = RasDial(ByVal 0&, strPhonebook, b(0), 0&, 0&, lngHConn)
'               if retVal <> 0 then
'                   fRasErrorHandler (retVal)
'               else
'                   msgbox "Connection handle: " & cstr(lngHconn)
'               endif
'           endif
'
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored

' **********************************************************************
Function fSyncronousDial(strPhonebook As String, strEntryName As String) As Long
Dim retVal As Long
Dim b() As Byte
Dim lngHConn As Long
   
   retVal = fRasGetEntryDialParams(b, strPhonebook, strEntryName)
   
   If retVal <> 0 Then
    fRASErrorHandler (retVal)
    fSyncronousDial = 0
    Exit Function
   End If
   
   ' code pauses on next line until connection established or fails
   retVal = RasDial(ByVal 0&, strPhonebook, b(0), 0&, 0&, lngHConn)
   If retVal <> 0 Then
    fRASErrorHandler (retVal)
    fSyncronousDial = 0
    Exit Function
   End If
  
   fSyncronousDial = lngHConn
End Function
'fSyncronousDial ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^






' **********************************************************************
' Function fASyncronousDial
' Purpose: dial a entry syncronous (meaning wait until conn is established or failed
' Return:  the connection handle if the function succeedes, elsewise 0
' Input:   the name of the entry to dial
'Remember you have to eventually Hangup the connection if the
'connection's handle value is non-zero .
'Dim hConn As Long
'hConn = fAsyncronousDial(vbNullString, "My Connection")
'The  rasConnectionState  parameter of the RasDialFunc1 is a long that
'tell us the connection state (see the RasConnState structure for details)
'
'Warning!: The asynchronous dial uses a callback function.  Your
'callback procedure (the RasDailFunc1 procedure) must stay in scope
'until the connection is established, disconnected or fails.
' **********************************************************************
Function fAsyncronousDial(strPhonebook As String, strEntryName As String) As Long
Dim retVal As Long
Dim b() As Byte
Dim lngHConn As Long

   retVal = fRasGetEntryDialParams(b, strPhonebook, strEntryName)
   If retVal <> 0 Then
        fRASErrorHandler (retVal)
        fAsyncronousDial = 0
        Exit Function
   End If
   
   retVal = RasDial(ByVal 0&, strPhonebook, b(0), _
                     1&, AddressOf RasDialFunc1, lngHConn)
   If retVal <> 0 Then
        fRASErrorHandler (retVal)
        fAsyncronousDial = 0
        Exit Function
   End If
   
   fAsyncronousDial = lngHConn
End Function
'fASyncronousDial ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
'   RasDialFunc1 is the callback procedure for fASyncronousDial
' **********************************************************************
Sub RasDialFunc1(ByVal hRasConn As Long, ByVal unMsg As Long, _
       ByVal rasConnectionState As Long, ByVal dwError As Long, _
       ByVal dwExtendedError As Long)
   
    ' do check for errors
    ' or check the rasConnectionState
    'Debug.Print hRasConn, Hex$(rasConnectionState), dwError
    If rasConnectionState = RASCS_Connected Then
    '  Debug.Print "connected"
    ElseIf rasConnectionState = RASCS_Disconnected Then
    '  Debug.Print "disconnected"
    End If
    
    If dwError <> 0 Then fRASErrorHandler (dwError)
End Sub
'RasDialFunc1 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Function fRasHangUp
' Purpose: finish an existing dial-up connection
' Return:  0, if disconnect was successfull, otherwise the error-code
' Input: connection handle to hang up
'        WaitForDisconnect  (true, if you want to wait 'til
'        connection is freed, false to return immediately
' Usage:
'           Dim hConn As Long
'           hConn = fSyncronousDial(vbNullString, "My Connection")
'           if hconn = 0 then
'               msgbox "Connection not established!"
'           else
'               if fRasHangUp(hConn) = 0 then
'                   msgbox "Connection closed successfully!"
'               else
'                   MsgBox "Connection not successfully closed!" & fRasErrorHandler(retVal)
'               endif
'           endif
' Note: it may take some time for RAS to actually disconnect and
' release ressources.
' It's up to you as to how you decide to handle this, but you should
' wait for Ras to complete it's hangup before closing your app or
' calling the RasDial or RasHangUp again. To do this, Pass true to the
' WaitForDisconnect parameter
' Note 2: If you are using asynchronous dialling you can call the fRasHangUp
' function at any time even before the connection is established.
' If your RasDailFunc1 callback procedure is still in scope
' (connection has not been established or disconnected yet) then you
' should NOT use the WaitForDisconnected = true as this would block your
' RasDialFunc1 procedure. Instead of this monitor the rasConnectionState
' parameter of your RasDailFunc1 function to be RASCS_disconnected!
' **********************************************************************
Function fRasHangUp(ByVal hRasConn As Long, WaitForDisconnect As Boolean) As Long
Dim rnt As Long
Dim lngError As Long
Dim myConnectionStatus As tRASCONNSTATUS
    retVal = RasHangUp(hRasConn)
    If retVal <> 0 Then
        fRasHangUp = retVal
        Exit Function
    End If
    If WaitForDisconnect Then
        Do
            Sleep 0&
            lngError = fRasGetConnectStatus(hRasConn, myConnectionStatus)
        Loop While lngError <> ERROR_INVALID_HANDLE
    End If
    fRasHangUp = 0
End Function
'fRasHangUp ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasEnumDevices
' Purpose: returns a list of all RAS capable devices installed on the computer giving their name and devicetype
' Return:  0, if no devices were detected, else the number of devices
' Input:   nothing
' Output:  a tRasDevInfo structure containing the devices
' Usage:
'          Dim retVal As Long
'          Dim clstRasDevInfo() As tRASDEVINFO
'          retVal = fRasEnumDevices(clstRasDevInfo)
'          For j = 0 To retVal - 1
'           Debug.Print clstRasDevInfo(j).DeviceName
'           Debug.Print clstRasDevInfo(j).DeviceType
'          Next j
' **********************************************************************
Function fRasEnumDevices(clsTRasDevInfo() As tRASDEVINFO) As Long
Dim retVal             As Long
Dim i               As Long
Dim lpCb            As Long
Dim lpCDevices      As Long
Dim b()             As Byte
Dim dwSize          As Long
   
   retVal = RasEnumDevices(ByVal 0&, lpCb, lpCDevices)

   If lpCDevices = 0 Then Exit Function
   
   dwSize = lpCb \ lpCDevices
   
   ReDim b(lpCb - 1)
   
   CopyMemory b(0), dwSize, 4
   
   retVal = RasEnumDevices(b(0), lpCb, lpCDevices)
   
   If lpCDevices = 0 Then Exit Function
   
   ReDim clsTRasDevInfo(lpCDevices - 1)
   
   For i = 0 To lpCDevices - 1
     CopyByteToTrimmedString clsTRasDevInfo(i).DeviceType, _
                                    b((i * dwSize) + 4), 17
     CopyByteToTrimmedString clsTRasDevInfo(i).DeviceName, _
                           b((i * dwSize) + 21), dwSize - 21
   Next i
   
   fRasEnumDevices = lpCDevices

End Function
'fRasEnumDevices ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasGetConnectStatus
' Purpose: Returning status information about an existing connection
' Return:  0 if the call succeeds, else an error code which can be used with fRasErrorHandler
' Input:   a valid handle to an existing connection
' Output:  a tRasConnStatus structure
' Usage:
'       Dim retVal As Long
'       Dim myConnStatus As tRASCONNSTATUS
'
'       retVal = tRasGetConnectStatus(hRasConn, myConnStatus)
'
'       If retVal <> 0 Then fRasErrorHandler(retVal)
' Note: the phone number of the connection is only returned on Windows NT4 windows 2000.
' **********************************************************************
Function fRasGetConnectStatus(hRasConn As Long, udttRasConnStatus As tRASCONNSTATUS) As Long
Dim i           As Long
Dim dwSize      As Long
Dim aVarLens    As Variant
Dim b()         As Byte

   aVarLens = Array(288&, 160&, 64&)
   
   For i = 0 To 2
      dwSize = aVarLens(i)
      ReDim b(dwSize - 1)
      CopyMemory b(0), dwSize, 4
      retVal = RasGetConnectStatus(hRasConn, b(0))
      If retVal <> 632 Then Exit For
   Next i
   
   fRasGetConnectStatus = retVal
   If retVal <> 0 Then Exit Function
      
   With udttRasConnStatus
      CopyMemory .lRasConnState, b(4), 4
      CopyMemory .dwError, b(8), 4
      CopyByteToTrimmedString .sDeviceType, b(12), 17&
      If dwSize = 64& Then
         CopyByteToTrimmedString .sDeviceName, b(29), 33&
      ElseIf dwSize = 160& Then
         CopyByteToTrimmedString .sDeviceName, b(29), 129&
      Else
         CopyByteToTrimmedString .sDeviceName, b(29), 129&
         CopyByteToTrimmedString .sNTPhoneNumber, b(158), 129&
      End If
   End With
   
End Function
'fRasGetConnectionStatus ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fRasEnumConnections
' Purpose: enumerate all connections and fill a tRasconn array with the information
' Return:  number of existing connections
' Input:   nothing
' Output:  an Array of tRasConn structure
' Usage:
'       Dim nConnections As Long
'       Dim myConnections() As tRASCONN
'       nConnections = fRasEnumConnections(myConnections)
' **********************************************************************
Function fRasEnumConnections(atRasConns() As tRASCONN) As Long
Dim retVal         As Long
Dim b()         As Byte
Dim aLens       As Variant
Dim dwSize      As Long
Dim lpCb        As Long
Dim lpConns     As Long
Dim i           As Long

   ReDim b(3)
   aLens = Array(692&, 676&, 412&, 32&)

   For i = 0 To 3
      dwSize = aLens(i)
      CopyMemory b(0), dwSize, 4
      lpCb = 4
      retVal = RasEnumConnections(b(0), lpCb, lpConns)
      If retVal <> 632 And retVal <> 610 Then Exit For
   Next i

   fRasEnumConnections = lpConns
   If lpConns = 0 Then Exit Function

   lpCb = dwSize * lpConns
   ReDim b(lpCb - 1)
   CopyMemory b(0), dwSize, 4
   retVal = RasEnumConnections(b(0), lpCb, lpConns)

   ' copy bytes to atRasConns
   ReDim atRasConns(lpConns - 1)
   For i = 0 To lpConns - 1
      With atRasConns(i)
         CopyMemory .hRasConn, b(i * dwSize + 4), 4
         If dwSize = 32& Then
            CopyByteToTrimmedString .sEntryName, b(i * dwSize + 8), 21&
         Else
            CopyByteToTrimmedString .sEntryName, b(i * dwSize + 8), 257&
            CopyByteToTrimmedString .sDeviceType, b(i * dwSize + 265), 17&
            CopyByteToTrimmedString .sDeviceName, b(i * dwSize + 282), 129&
            If dwSize > 412& Then
              CopyByteToTrimmedString .sPhonebook, b(i * dwSize + 411), 260&
              CopyMemory .lngSubEntry, b(i * dwSize + 672), 4
              If dwSize > 676& Then
                CopyMemory .guidEntry(0), b(i * dwSize + 676), 16
              End If
            End If
         End If
      End With
   Next i
End Function
'fRasEnumConnections ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Function fRasGetRASAMB
' Purpose: describes the result of a RAS Authentication Message Block
'          (which is used by NT 3 and OS/2 1.3 RAS servers)
' Return:  0 if the call succeded, else an error handle
' Input:   a valid connection handle
' Output:  an Array of tRASAMB structure
' Usage:
'       Dim nConn As Long
'       Dim vRasAmb As tRASAMB
'       nConn = fRasGetRASAMB(nConn,vRasAmb)
' **********************************************************************
Function fRasGetRASAMB(hRasConn As Long, udttRASAMB As tRASAMB) As Long
   
Dim b()     As Byte
Dim retVal     As Long
Dim lpCb    As Long

   lpCb = 28&
   ReDim b(lpCb - 1)
   CopyMemory b(0), lpCb, 4
   
   retVal = RasGetProjectionInfo(hRasConn, &H10000, b(0), lpCb)
   
   fRasGetRASAMB = retVal
   If retVal <> 0 Then Exit Function
   
   With udttRASAMB
      CopyMemory .dwError, b(4), 4
      CopyByteToTrimmedString .sNetBiosError, b(8), 17
      CopyMemory .bLana, b(25), 1
   End With
   
End Function
'fRasGetRASAMB ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasGetRASPPPNBF
' Purpose: describes the result of a NetBeui projection
' Return:  0 if the call succeded, else an error handle
' Input:   a valid connection handle
' Output:  an Array of tRASPPPNBF structure
' Usage:
'       Dim nConn As Long
'       Dim vRasPPP As tRASPPPNBF
'       nConn = fRasGetRASPPPNBF(nConn,vRasAPP)
' **********************************************************************
Function fRasGetRASPPPNBF(hRasConn As Long, udttRASPPPNBF As tRASPPPNBF) As Long
   
Dim b()     As Byte
Dim retVal     As Long
Dim lpCb    As Long
   
   lpCb = 48&
   ReDim b(lpCb - 1)
   CopyMemory b(0), lpCb, 4
   
   retVal = RasGetProjectionInfo(hRasConn, &H803F&, b(0), lpCb)
   
   fRasGetRASPPPNBF = retVal
   If retVal <> 0 Then Exit Function
   
   With udttRASPPPNBF
      CopyMemory .dwError, b(4), 4
      CopyMemory .dwNetBiosError, b(8), 4
      CopyByteToTrimmedString .szNetBiosError, b(12), 17
      CopyByteToTrimmedString .szWorkstationName, b(29), 17
      CopyMemory .bLana, b(46), 1
   End With
   
End Function
'fRasGetRASPPPNBF ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Function fRasGetRASPPPIPX
' Purpose: describes the result of a Internetwork Packet Exchange projection
' Return:  0 if the call succeded, else an error handle
' Input:   a valid connection handle
' Output:  an Array of tRASPPPIPX structure
' Usage:
'       Dim nConn As Long
'       Dim vRasPPP As tRASPPPIPX
'       nConn = fRasGetRASPPPIPX(nConn,vRasAPP)
' **********************************************************************
Function fRasGetRASPPPIPX(hRasConn As Long, udttRASPPPIPX As tRASPPPIPX) As Long
   
Dim b() As Byte
Dim retVal As Long
Dim lpCb As Long
   
   lpCb = 24&
   ReDim b(lpCb - 1)
   CopyMemory b(0), lpCb, 4
   
   retVal = RasGetProjectionInfo(hRasConn, &H802B&, b(0), lpCb)
   
   fRasGetRASPPPIPX = retVal
   If retVal <> 0 Then Exit Function
   
   With udttRASPPPIPX
      CopyMemory .dwError, b(4), 4
      CopyByteToTrimmedString .szIpxAddress, b(8), 16
   End With
   
End Function
'fRasGetRASPPPIPX ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

 
' **********************************************************************
' Function fRasGetRASPPPIP
' Purpose: describes the result of a IP projection
' Return:  0 if the call succeded, else an error handle
' Input:   a valid connection handle
' Output:  an Array of tRASPPPIP structure
' Usage:
'       Dim nConn As Long
'       Dim vRasPPP As tRASPPPIP
'       nConn = fRasGetRASPPPIP(nConn,vRasAPP)
' **********************************************************************
Function fRasGetRASPPPIP(hRasConn As Long, udttRASPPPIP As tRASPPPIP) As Long
   
Dim b()     As Byte
Dim retVal     As Long
Dim lpCb    As Long
   
   lpCb = 40&
   ReDim b(lpCb - 1)
   CopyMemory b(0), lpCb, 4
   
   retVal = RasGetProjectionInfo(hRasConn, &H8021&, b(0), lpCb)
   
   fRasGetRASPPPIP = retVal
   If retVal <> 0 Then Exit Function
   
   With udttRASPPPIP
      CopyMemory .dwError, b(4), 4
      CopyByteToTrimmedString .szIpAddress, b(8), 16
      CopyByteToTrimmedString .szServerIpAddress, b(24), 16
   End With
   
End Function
'fRasGetRASPPPIP ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRasGetRASSLIP
' Purpose: describes the result of a serial line ip projection
' Return:  0 if the call succeded, else an error handle
' Input:   a valid connection handle
' Output:  an Array of tRASSLIP structure
' Usage:
'       Dim nConn As Long
'       Dim vRasPPP As tRASSlip
'       nConn = fRasGetRASSlip(nConn,vRasAPP)
' **********************************************************************
Function fRasGetRASSLIP(hRasConn As Long, udttRASSLIP As tRASSLIP) As Long
   
Dim b()     As Byte
Dim retVal     As Long
Dim lpCb    As Long
   
   lpCb = 24&
   ReDim b(lpCb - 1)
   CopyMemory b(0), lpCb, 4
   
   retVal = RasGetProjectionInfo(hRasConn, &H20000, b(0), lpCb)
   
   fRasGetRASSLIP = retVal
   If retVal <> 0 Then Exit Function
   
   With udttRASSLIP
      CopyMemory .dwError, b(4), 4
      CopyByteToTrimmedString .szIpAddress, b(8), 16
   End With
   
End Function
'fRasGetRASSLIP ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Procedure sDLG_AddModem
' Purpose: displays the "Add New Modem" dialog
'          this is by far the easiest way to add a new modem
'          if you don't want to use the above methods
' Input:  n/a
' Return: n/a
' Usage:
'       Call sDLG_AddModem
' **********************************************************************
Sub sDLG_AddModem()
    Call Shell("RunDLL32 shell32.dll,Control_RunDLL MODEM.CPL,Modems,Add")
End Sub
'sDLG_AddModem ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' **********************************************************************
' Procedure sDLG_ConfigModem
' Purpose: displays the "Configure Modem" dialog
'          this is by far the easiest way to configure a modem
'          if you don't want to use the above methods
' Input:  n/a
' Return: n/a
' Usage:
'       Call sDLG_ConfigModem
' **********************************************************************
Sub sDLG_ConfigModem()
    Call Shell("RunDLL32 shell32.dll,Control_RunDLL MODEM.CPL,Modems")
End Sub
'sDLG_ConfigModem ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRASRenameEntry
' Purpose: renames an existing dial-up connection
' Input:  Name of the existing connection,
'         new name of the existing connection
' Return: 0, if the function succedes, else RasErrorcode
' Usage:
'       Dim retVal As Long
'       retVal = fRasRenameEntry("My Connection", "My New Name")
'       If retVal <> 0 Then
'           MsgBox fRASErrorHandler(retVal)
'       End If
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRASRenameEntry(ByVal OldName As String, ByVal NewName As String, Optional strPhonebook As String) As Long
Dim retVal     As Long
    fRASRenameEntry = RasRenameEntry(IIf((IsMissing(strPhonebook) = True Or stringphonebook = ""), vbNullString, strPhonebook), OldName, NewName)
End Function
'fRasRenameEntry ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fRASDeleteEntry
' Purpose: deletes an existing dial-up connection
' Input:  Name of the existing connection,
' Return: 0, if the function succedes, else RasErrorcode
' Usage:
'       Dim retVal As Long
'       retVal = fRasDeleteEntry("My Connection")
'       If retVal <> 0 Then
'           MsgBox fRASErrorHandler(retVal)
'       End If
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRASDeleteEntry(ByVal ConName As String, Optional strPhonebook As String) As Long
Dim retVal     As Long
    fRASDeleteEntry = RasDeleteEntry(IIf((IsMissing(strPhonebook) = True Or stringphonebook = ""), vbNullString, strPhonebook), ConName)
End Function
'fRasRenameConnection ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRASValidateEntryName
' Purpose: validating an entry
' Input:  Name of the entry to check
' Return: 0, if the name is valid and does not exist already
'         123, if the name syntax is invalid
'         183, if the entry name already exists
'         else: Errorcode
' Usage:
'       Dim retVal As Long
'       retVal = fRasValidateEntryName("My New Name")
'       If retVal <> 0 Then
'           Debug.Print retVal, fRASErrorHandler(retVal)
'       End If
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRASValidateEntryName(ByVal EntryName As String, Optional strPhonebook As String) As Long
Dim retVal     As Long
    fRASValidateEntryName = RasValidateEntryName(IIf((IsMissing(strPhonebook) = True Or stringphonebook = ""), vbNullString, strPhonebook), EntryName)
End Function
'fRasValidateEntryName ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' Function fRASCreatePhoneBookEntryDialog
' Purpose: shows the "create new entry" dialog
' Input:  handle of your form on which you want to display the dialog modal
'         pass 0 if you don't want to display the dialog modal
' Return: 0, if the function succeeds, else 621
'         Note: If the user hits the "Cancel" button of the dialog,
'               the function still returns 0.
'               To know if the dialog was cancelled, you could call
'               fRasEnumEntries before and after the dialog
' Usage:
'       Dim retVal As Long
'       retVal = fRasCreatePhoneBookEntryDialog(0)
'       If retVal =621 then
'           messagebox "Error"
'       elseif retVal=0 then
'           messagebox "Succeded"
'       else
'           messagebox "Unknown Error: " &  fRASErrorHandler(retVal)
'       End If
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Public Function fRasCreatePhoneBookEntryDialog(ByVal hwnd As Long, Optional strPhonebook As String) As Long
    fRasCreatePhoneBookEntryDialog = RasCreatePhonebookEntry(IIf(hwnd = 0, 0&, hwnd), IIf((IsMissing(strPhonebook) = True Or stringphonebook = ""), vbNullString, strPhonebook))
End Function
'fRasCreatePhoneBookEntryDialog ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



' **********************************************************************
' Function fRASEditPhoneBookEntryDialog
' Purpose: shows the "edit phonebook entry" dialog
' Input:  handle of your form on which you want to display the dialog modal
'         pass 0 if you don't want to display the dialog modal
'         the name of an existing entry
' Return: 0, if the function succeeds, else Errorcode
' Usage:
'       Dim retVal As Long
'       retVal = fRasEditPhoneBookEntryDialog(0, "My Entry")
'       if retVal=0 then
'           messagebox "Succeded"
'       else
'           messagebox "Unknown Error: " &  fRASErrorHandler(retVal)
'       End If
' on NT / 2000 you can specify path and filename of the phonebook to use
' if none is given, default will be used
' on Win9x, this parameter is ignored
' **********************************************************************
Function fRasEditPhoneBookEntryDialog(ByVal hwnd As Long, ByVal EntryName As String, Optional strPhonebook As String) As Long
Dim retVal     As Long
    fRasEditPhoneBookEntryDialog = RasEditPhonebookEntry(IIf(hwnd = 0, 0&, hwnd), IIf((IsMissing(strPhonebook) = True Or stringphonebook = ""), vbNullString, strPhonebook), EntryName)
End Function
'fRasEditPhoneBookEntryDialog ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


' **********************************************************************
' **********************************************************************
' **********************************************************************
' Helper functions
' No need to explain here
' **********************************************************************
' **********************************************************************
' **********************************************************************


Sub CopyByteToTrimmedString(strToCopyTo As String, bPos As Byte, lngMaxLen As Long)
Dim strTemp     As String
Dim lngLen      As Long
   strTemp = String(lngMaxLen + 1, 0)
   CopyMemory ByVal strTemp, bPos, lngMaxLen
   lngLen = InStr(strTemp, Chr$(0)) - 1
   strToCopyTo = Left$(strTemp, lngLen)
End Sub
 

 
Sub CopyStringToByte(bPos As Byte, strToCopy As String, lngMaxLen As Long)
Dim lngLen      As Long
   lngLen = Len(strToCopy)
   If lngLen = 0 Then
      Exit Sub
   ElseIf lngLen > lngMaxLen Then
      lngLen = lngMaxLen
   End If
   CopyMemory bPos, ByVal strToCopy, lngLen
End Sub




