Attribute VB_Name = "Module1"
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As ip_option_information, ReplyBuffer As icmp_echo_reply, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Const PING_TIMEOUT = 1200
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYSSTATUS_LEN = 256
Public Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Public Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Public Const SOCKET_ERROR = -1
Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Public Declare Function WSACleanup Lib "wsock32" () As Integer

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)

Public Type ip_option_information
 TTL             As Byte     'Time To Live [ Nb Max de sauts de routeurs ]
 Tos             As Byte     'Type Of Service [ Type de trame ]
 flags           As Byte     'IP header flags [ En-tête de la trame ]
 OptionsSize     As Byte     'Taille des trames
 OptionsData     As Long     'Options ( hops, TargetIP,... )
End Type

Public Type icmp_echo_reply
 Address         As Long                     'Replying address
 Status          As Long                     'Reply IP_STATUS, values as defined above
 RoundTripTime   As Long                     'RTT in milliseconds
 DataSize        As Integer                  'Reply data size in bytes
 Reserved        As Integer                  'Reserved for system use
 DataPointer     As Long                     'Pointer to the reply data
 Options         As ip_option_information    'Reply options
 Data            As String * 250             'Reply data which should be a copy of the string sent, NULL terminated
End Type                                     ' this field length should be large enough to contain the string sent

Public Type tagWSAData
 wVersion            As Integer
 wHighVersion        As Integer
 szDescription       As String * WSADESCRIPTION_LEN_1
 szSystemStatus      As String * WSASYSSTATUS_LEN_1
 iMaxSockets         As Integer
 iMaxUdpDg           As Integer
 lpVendorInfo        As String * 200
End Type

Public ReturnedIP$
Public RoundTime(3) As Integer
Public IgnoreFirstPing As Boolean
Public AutoUpdate As Boolean
Public Ancre As Boolean
Public SelectedIp$

Public Function PingIp(ByVal IP As String, TTL As Integer, Timeout As Integer)
 '   Va retourner  l'ip du routeur de distance TTL
 ' Si le TTL est suffisant pour atteindre la cible
 ' alors  on retourne 1.  Sinon on retourne 0 pour
 ' indiquer  qu'un nouveau  routeur à été atteind.
 ' Si une erreure se produit, on retourne -1
 
 Dim hFile       As Long
 Dim lRet        As Long
 Dim lIPAddress  As Long
 Dim strMessage  As String
 Dim pOptions    As ip_option_information
 Dim pReturn     As icmp_echo_reply
 Dim iVal        As Integer
 Dim lPingRet    As Long
 Dim pWsaData    As tagWSAData
    
 strMessage = "Echo this string of data"
 iVal = WSAStartup(&H101, pWsaData)
 lIPAddress = ConvertIPAddressToLong(IP)
 hFile = IcmpCreateFile()
    
 pOptions.TTL = TTL
    
 lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, Len(strMessage), pOptions, pReturn, Len(pReturn), Timeout)

 If lRet = 0 Then
  PingIp = -1
 Else
  If pReturn.Status <> 0 Then
   PingIp = 0
   RoundTime(0) = pReturn.RoundTripTime
   RoundTime(1) = PingTime(IP, TTL, Timeout)
   RoundTime(2) = PingTime(IP, TTL, Timeout)
   nIP$ = LongToIp$(Hex$(pReturn.Address))
   ReturnedIP$ = nIP$
  Else
   PingIp = 1
   RoundTime(0) = pReturn.RoundTripTime
   RoundTime(1) = PingTime(IP, TTL, Timeout)
   RoundTime(2) = PingTime(IP, TTL, Timeout)
   nIP$ = LongToIp$(Hex$(pReturn.Address))
   ReturnedIP$ = IP$
  End If
 End If
                        
 lRet = IcmpCloseHandle(hFile)
 iVal = WSACleanup()
End Function

Function ConvertIPAddressToLong(strAddress As String) As Long
 ' Convertion chaine IP en long
 
 Dim strTemp             As String
 Dim lAddress            As Long
 Dim iValCount           As Integer
 Dim lDotValues(1 To 4)  As String
    
 strTemp = strAddress
 iValCount = 0
    
 While InStr(strTemp, ".") > 0
  iValCount = iValCount + 1
  lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)
  strTemp = Mid(strTemp, InStr(strTemp, ".") + 1)
 Wend
        
 iValCount = iValCount + 1
 lDotValues(iValCount) = strTemp
    
 If iValCount <> 4 Then
  ConvertIPAddressToLong = 0
  Exit Function
 End If
        
 lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & Right("00" & Hex(lDotValues(3)), 2) & Right("00" & Hex(lDotValues(2)), 2) & Right("00" & Hex(lDotValues(1)), 2))
               
 ConvertIPAddressToLong = lAddress
End Function

Function LongToIp$(Value$)
 ' Convertion d'un long en addresse IP ( chaine )
 
 Value$ = "00000" + Value$
 Value$ = Right$(Value$, 8)
 op1$ = Right$(Value$, 2)
 op2$ = Mid$(Value$, 5, 2)
 op3$ = Mid$(Value$, 3, 2)
 op4$ = Left$(Value$, 2)
 LongToIp$ = HexDec$(op1$) + "." + HexDec$(op2$) + "." + HexDec$(op3$) + "." + HexDec$(op4$)
End Function

Function HexDec$(Value$)
 ' Convertion Hexa en Decimal
 
 Id = 0: Result = 0
 For i = Len(Value$) To 1 Step -1
  If Mid$(Value$, i, 1) = "0" Then Vl = 0
  If Mid$(Value$, i, 1) = "1" Then Vl = 1
  If Mid$(Value$, i, 1) = "2" Then Vl = 2
  If Mid$(Value$, i, 1) = "3" Then Vl = 3
  If Mid$(Value$, i, 1) = "4" Then Vl = 4
  If Mid$(Value$, i, 1) = "5" Then Vl = 5
  If Mid$(Value$, i, 1) = "6" Then Vl = 6
  If Mid$(Value$, i, 1) = "7" Then Vl = 7
  If Mid$(Value$, i, 1) = "8" Then Vl = 8
  If Mid$(Value$, i, 1) = "9" Then Vl = 9
  If Mid$(Value$, i, 1) = "A" Then Vl = 10
  If Mid$(Value$, i, 1) = "B" Then Vl = 11
  If Mid$(Value$, i, 1) = "C" Then Vl = 12
  If Mid$(Value$, i, 1) = "D" Then Vl = 13
  If Mid$(Value$, i, 1) = "E" Then Vl = 14
  If Mid$(Value$, i, 1) = "F" Then Vl = 15
  Result = Result + (Vl * 16 ^ Id)
  Id = Id + 1
 Next i
 HexDec$ = Str$(Result)
 HexDec$ = LTrim$(HexDec$)
End Function

Public Function PingTime(ByVal IP As String, TTL As Integer, Timeout As Integer)
 '   Permet de faire les pings 2 et 3 ( le code est
 ' le même que pour PingIP.
 
 Dim hFile       As Long
 Dim lRet        As Long
 Dim lIPAddress  As Long
 Dim strMessage  As String
 Dim pOptions    As ip_option_information
 Dim pReturn     As icmp_echo_reply
 Dim iVal        As Integer
 Dim lPingRet    As Long
 Dim pWsaData    As tagWSAData
    
 strMessage = "Echo this string of data"
 iVal = WSAStartup(&H101, pWsaData)
 lIPAddress = ConvertIPAddressToLong(IP)
 hFile = IcmpCreateFile()
    
 pOptions.TTL = TTL
    
 DoEvents
 lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, Len(strMessage), pOptions, pReturn, Len(pReturn), Timeout)

 If lRet = 0 Then
  PingTime = -1
 Else
  If pReturn.Status <> 0 Then
   PingTime = pReturn.RoundTripTime
  Else
   PingTime = pReturn.RoundTripTime
  End If
 End If
                        
 lRet = IcmpCloseHandle(hFile)
 iVal = WSACleanup()
End Function

