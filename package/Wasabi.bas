Attribute VB_Name = "Wasabi"
' ============================================================================
' Wasabi v1.8.0
' Copyright (c) 2026 UesleiDev
'
' Permission is hereby granted, free of charge, to any person obtaining a
' copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' ============================================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" ( _
        ByVal pszPrincipal As LongPtr, ByVal pszPackage As String, ByVal fCredentialUse As Long, _
        ByVal pvLogonID As LongPtr, ByRef pAuthData As Any, ByVal pGetKeyFn As LongPtr, _
        ByVal pvGetKeyArgument As LongPtr, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare PtrSafe Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" ( _
        ByRef phCredential As SecHandle, ByVal phContext As LongPtr, ByVal pszTargetName As String, _
        ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
        ByVal pInput As LongPtr, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, _
        ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" ( _
        ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, _
        ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
        ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, _
        ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare PtrSafe Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As LongPtr) As Long
    Private Declare PtrSafe Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" ( _
        ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long
    Private Declare PtrSafe Function EncryptMessage Lib "secur32.dll" ( _
        ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare PtrSafe Function DecryptMessage Lib "secur32.dll" ( _
        ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As LongPtr
    Private Declare PtrSafe Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
    Private Declare PtrSafe Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function sock_send Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function sock_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function sock_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
    Private Declare PtrSafe Function sock_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal hostname As String) As LongPtr
    Private Declare PtrSafe Function sock_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
    Private Declare PtrSafe Function sock_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function sock_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As TIMEVAL) As Long
    Private Declare PtrSafe Function sock_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByVal optlen As Long) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)
    Private Declare PtrSafe Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByVal src As LongPtr, ByVal size As Long)
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Const NULL_PTR As LongPtr = 0
#Else
    Private Declare Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" ( _
        ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, _
        ByVal pvLogonID As Long, ByRef pAuthData As Any, ByVal pGetKeyFn As Long, _
        ByVal pvGetKeyArgument As Long, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" ( _
        ByRef phCredential As SecHandle, ByVal phContext As Long, ByVal pszTargetName As String, _
        ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
        ByVal pInput As Long, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, _
        ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" ( _
        ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, _
        ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
        ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, _
        ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As Long) As Long
    Private Declare Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" ( _
        ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long
    Private Declare Function EncryptMessage Lib "secur32.dll" ( _
        ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare Function DecryptMessage Lib "secur32.dll" ( _
        ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long
    Private Declare Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
    Private Declare Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
    Private Declare Function sock_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function sock_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function sock_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
    Private Declare Function sock_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal hostname As String) As Long
    Private Declare Function sock_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
    Private Declare Function sock_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function sock_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As TIMEVAL) As Long
    Private Declare Function sock_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)
    Private Declare Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByVal src As Long, ByVal size As Long)
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Const NULL_PTR As Long = 0
#End If

#If VBA7 Then
    Private Const INVALID_SOCKET As LongPtr = -1
#Else
    Private Const INVALID_SOCKET As Long = -1
#End If

Private Type BinaryMessage
    data() As Byte
End Type

Private Type SecHandle
#If VBA7 Then
    dwLower As LongPtr
    dwUpper As LongPtr
#Else
    dwLower As Long
    dwUpper As Long
#End If
End Type

Private Type SECURITY_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type SecBuffer
    cbBuffer As Long
    BufferType As Long
#If VBA7 Then
    pvBuffer As LongPtr
#Else
    pvBuffer As Long
#End If
End Type

Private Type SecBufferDesc
    ulVersion As Long
    cBuffers As Long
#If VBA7 Then
    pBuffers As LongPtr
#Else
    pBuffers As Long
#End If
End Type

Private Type SecPkgContext_StreamSizes
    cbHeader As Long
    cbTrailer As Long
    cbMaximumMessage As Long
    cBuffers As Long
    cbBlockSize As Long
End Type

Private Type SCHANNEL_CRED
    dwVersion As Long
    cCreds As Long
#If VBA7 Then
    paCred As LongPtr
    hRootStore As LongPtr
#Else
    paCred As Long
    hRootStore As Long
#End If
    cMappers As Long
#If VBA7 Then
    aphMappers As LongPtr
#Else
    aphMappers As Long
#End If
    cSupportedAlgs As Long
#If VBA7 Then
    palgSupportedAlgs As LongPtr
#Else
    palgSupportedAlgs As Long
#End If
    grbitEnabledProtocols As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan As Long
    dwFlags As Long
    dwCredFormat As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
#If VBA7 Then
    lpVendorInfo As LongPtr
#Else
    lpVendorInfo As Long
#End If
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type TIMEVAL
    tv_sec As Long
    tv_usec As Long
End Type

Private Type FD_SET
    fd_count As Long
#If VBA7 Then
    fd_array(0 To 0) As LongPtr
#Else
    fd_array(0 To 0) As Long
#End If
End Type

Private Type HOSTENT32
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type HOSTENT64
    h_name As LongPtr
    h_aliases As LongPtr
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As LongPtr
End Type

Private Type WasabiStats
    BytesSent As Currency
    BytesReceived As Currency
    MessagesSent As Long
    MessagesReceived As Long
    ConnectedAt As Long
End Type

Private Type WasabiConnection
#If VBA7 Then
    Socket As LongPtr
#Else
    Socket As Long
#End If
    Connected As Boolean
    TLS As Boolean
    Host As String
    Port As Long
    Path As String
    hCred As SecHandle
    hContext As SecHandle
    Sizes As SecPkgContext_StreamSizes
    recvBuffer() As Byte
    recvLen As Long
    DecryptBuffer() As Byte
    DecryptLen As Long
    MsgQueue() As String
    MsgHead As Long
    MsgTail As Long
    MsgCount As Long
    BinaryQueue() As BinaryMessage
    BinaryHead As Long
    BinaryTail As Long
    BinaryCount As Long
    FragmentBuffer() As Byte
    FragmentLen As Long
    FragmentOpcode As Byte
    Fragmenting As Boolean
    LastError As Long
    LastErrorCode As Long
    TechnicalDetails As String
    OriginalUrl As String
    CustomHeaders() As String
    CustomHeaderCount As Long
    AutoReconnect As Boolean
    ReconnectMaxAttempts As Long
    ReconnectAttempts As Long
    ReconnectBaseDelayMs As Long
    PingIntervalMs As Long
    LastPingSentAt As Long
    ReceiveTimeoutMs As Long
    EnableErrorDialog As Boolean
    LogCallback As String
    Stats As WasabiStats
    NoDelay As Boolean
    proxyHost As String
    proxyPort As Long
    proxyUser As String
    proxyPass As String
    ProxyEnabled As Boolean
    InactivityTimeoutMs As Long
    LastActivityAt As Long
    SubProtocol As String
    CustomBufferSize As Long
    CustomFragmentSize As Long
End Type

Public Enum WasabiError
    ERR_NONE = 0
    ERR_WSA_STARTUP_FAILED = 1
    ERR_SOCKET_CREATE_FAILED = 2
    ERR_DNS_RESOLVE_FAILED = 3
    ERR_CONNECT_FAILED = 4
    ERR_TLS_ACQUIRE_CREDS_FAILED = 5
    ERR_TLS_HANDSHAKE_FAILED = 6
    ERR_TLS_HANDSHAKE_TIMEOUT = 7
    ERR_WEBSOCKET_HANDSHAKE_FAILED = 8
    ERR_WEBSOCKET_HANDSHAKE_TIMEOUT = 9
    ERR_SEND_FAILED = 10
    ERR_RECV_FAILED = 11
    ERR_NOT_CONNECTED = 12
    ERR_ALREADY_CONNECTED = 13
    ERR_TLS_ENCRYPT_FAILED = 14
    ERR_TLS_DECRYPT_FAILED = 15
    ERR_INVALID_URL = 16
    ERR_HANDSHAKE_REJECTED = 17
    ERR_CONNECTION_LOST = 18
    ERR_INVALID_HANDLE = 19
    ERR_MAX_CONNECTIONS = 20
    ERR_PROXY_CONNECT_FAILED = 21
    ERR_PROXY_AUTH_FAILED = 22
    ERR_PROXY_TUNNEL_FAILED = 23
    ERR_INACTIVITY_TIMEOUT = 24
End Enum

Private Const BUFFER_SIZE As Long = 262144
Private Const FRAGMENT_BUFFER_SIZE As Long = 262144
Private Const SOL_SOCKET   As Long = 65535
Private Const SO_KEEPALIVE As Long = 8
Private Const SO_RCVBUF As Long = &H1002
Private Const SO_SNDBUF As Long = &H1001
Private Const IPPROTO_TCP_SOL As Long = 6
Private Const TCP_NODELAY     As Long = 1
Private Const CP_UTF8 As Long = 65001
Private Const SECPKG_CRED_OUTBOUND As Long = &H2
Private Const SCHANNEL_CRED_VERSION As Long = &H4
Private Const SCH_CRED_NO_DEFAULT_CREDS As Long = &H10
Private Const SCH_CRED_MANUAL_CRED_VALIDATION As Long = &H8
Private Const SCH_CRED_IGNORE_NO_REVOCATION_CHECK As Long = &H800
Private Const SCH_CRED_IGNORE_REVOCATION_OFFLINE As Long = &H1000
Private Const SP_PROT_TLS1_2_CLIENT As Long = &H800
Private Const SP_PROT_TLS1_3_CLIENT As Long = &H2000
Private Const ISC_REQ_SEQUENCE_DETECT As Long = &H8
Private Const ISC_REQ_REPLAY_DETECT As Long = &H4
Private Const ISC_REQ_CONFIDENTIALITY As Long = &H10
Private Const ISC_REQ_ALLOCATE_MEMORY As Long = &H100
Private Const ISC_REQ_STREAM As Long = &H8000
Private Const SECBUFFER_TOKEN As Long = 2
Private Const SECBUFFER_DATA As Long = 1
Private Const SECBUFFER_EMPTY As Long = 0
Private Const SECBUFFER_STREAM_HEADER As Long = 7
Private Const SECBUFFER_STREAM_TRAILER As Long = 6
Private Const SECBUFFER_EXTRA As Long = 5
Private Const SECBUFFER_VERSION As Long = 0
Private Const SECPKG_ATTR_STREAM_SIZES As Long = 4
Private Const SEC_E_OK As Long = 0
Private Const SEC_I_CONTINUE_NEEDED As Long = &H90312
Private Const SEC_E_INCOMPLETE_MESSAGE As Long = &H80090318
Private Const SEC_I_RENEGOTIATE As Long = &H90321
Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const FIONBIO As Long = &H8004667E
Private Const FIONREAD As Long = &H4004667F
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MSG_QUEUE_SIZE As Long = 512
Private Const WS_OPCODE_CONTINUATION As Byte = 0
Private Const WS_OPCODE_TEXT As Byte = 1
Private Const WS_OPCODE_BINARY As Byte = 2
Private Const WS_OPCODE_CLOSE As Byte = 8
Private Const WS_OPCODE_PING As Byte = 9
Private Const WS_OPCODE_PONG As Byte = 10
Private Const MAX_CONNECTIONS As Long = 64
Private Const INVALID_CONN_HANDLE As Long = -1
Private Const DEFAULT_RECEIVE_TIMEOUT_MS As Long = 5000
Private Const DEFAULT_PING_INTERVAL_MS As Long = 0
Private Const DEFAULT_RECONNECT_BASE_DELAY_MS As Long = 1000
Private Const DEFAULT_RECONNECT_MAX_ATTEMPTS As Long = 5
Private Const MAX_RECONNECT_DELAY_MS As Long = 30000

Private m_WSAInitialized As Boolean
Private m_Connections() As WasabiConnection
Private m_ConnectionCount As Long
Private m_DefaultHandle As Long

Private m_LastError As WasabiError
Private m_LastErrorCode As Long
Private m_TechnicalDetails As String
Public EnableErrorDialog As Boolean

Private Sub InitConnectionPool()
    If m_ConnectionCount > 0 Then Exit Sub
    Randomize
    ReDim m_Connections(0 To MAX_CONNECTIONS - 1)
    Dim i As Long
    For i = 0 To MAX_CONNECTIONS - 1
        m_Connections(i).Socket = INVALID_SOCKET
        m_Connections(i).Connected = False
    Next i
    m_ConnectionCount = MAX_CONNECTIONS
End Sub

Private Function AllocConnection() As Long
    Dim i As Long
    Dim bufSize As Long
    Dim fragSize As Long
    
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If Not m_Connections(i).Connected And m_Connections(i).Socket = INVALID_SOCKET Then
            bufSize = IIf(m_Connections(i).CustomBufferSize > 0, m_Connections(i).CustomBufferSize, BUFFER_SIZE)
            fragSize = IIf(m_Connections(i).CustomFragmentSize > 0, m_Connections(i).CustomFragmentSize, FRAGMENT_BUFFER_SIZE)
            ReDim m_Connections(i).recvBuffer(0 To bufSize - 1)
            ReDim m_Connections(i).DecryptBuffer(0 To bufSize - 1)
            ReDim m_Connections(i).MsgQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).BinaryQueue(0 To MSG_QUEUE_SIZE - 1)
            ReDim m_Connections(i).FragmentBuffer(0 To fragSize - 1)
            ReDim m_Connections(i).CustomHeaders(0 To 31)
            m_Connections(i).CustomBufferSize = 0
            m_Connections(i).CustomFragmentSize = 0
            m_Connections(i).recvLen = 0
            m_Connections(i).DecryptLen = 0
            m_Connections(i).MsgHead = 0
            m_Connections(i).MsgTail = 0
            m_Connections(i).MsgCount = 0
            m_Connections(i).BinaryHead = 0
            m_Connections(i).BinaryTail = 0
            m_Connections(i).BinaryCount = 0
            m_Connections(i).FragmentLen = 0
            m_Connections(i).Fragmenting = False
            m_Connections(i).FragmentOpcode = 0
            m_Connections(i).LastError = ERR_NONE
            m_Connections(i).LastErrorCode = 0
            m_Connections(i).TechnicalDetails = ""
            m_Connections(i).OriginalUrl = ""
            m_Connections(i).CustomHeaderCount = 0
            m_Connections(i).AutoReconnect = False
            m_Connections(i).ReconnectMaxAttempts = DEFAULT_RECONNECT_MAX_ATTEMPTS
            m_Connections(i).ReconnectAttempts = 0
            m_Connections(i).ReconnectBaseDelayMs = DEFAULT_RECONNECT_BASE_DELAY_MS
            m_Connections(i).PingIntervalMs = DEFAULT_PING_INTERVAL_MS
            m_Connections(i).LastPingSentAt = 0
            m_Connections(i).ReceiveTimeoutMs = DEFAULT_RECEIVE_TIMEOUT_MS
            m_Connections(i).EnableErrorDialog = False
            m_Connections(i).LogCallback = ""
            m_Connections(i).Stats.BytesSent = 0
            m_Connections(i).Stats.BytesReceived = 0
            m_Connections(i).Stats.MessagesSent = 0
            m_Connections(i).Stats.MessagesReceived = 0
            m_Connections(i).Stats.ConnectedAt = 0
            m_Connections(i).NoDelay = False
            m_Connections(i).proxyHost = ""
            m_Connections(i).proxyPort = 0
            m_Connections(i).proxyUser = ""
            m_Connections(i).proxyPass = ""
            m_Connections(i).ProxyEnabled = False
            m_Connections(i).InactivityTimeoutMs = 0
            m_Connections(i).LastActivityAt = 0
            m_Connections(i).SubProtocol = ""
            AllocConnection = i
            Exit Function
        End If
    Next i
    AllocConnection = INVALID_CONN_HANDLE
End Function

Private Function ValidHandle(ByVal handle As Long) As Boolean
    If handle < 0 Or handle >= MAX_CONNECTIONS Then Exit Function
    InitConnectionPool
    ValidHandle = True
End Function

Private Sub WasabiLog(ByVal handle As Long, ByVal msg As String)
    Debug.Print "[WASABI] " & msg
    If ValidHandle(handle) Then
        If m_Connections(handle).LogCallback <> "" Then
            Application.Run m_Connections(handle).LogCallback, msg
        End If
    End If
End Sub

Private Sub LogError(ByVal errType As WasabiError, ByVal technicalMsg As String, ByVal userMsg As String, Optional ByVal ErrCode As Long = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim showDialog As Boolean

    m_LastError = errType
    m_LastErrorCode = ErrCode
    m_TechnicalDetails = technicalMsg

    showDialog = EnableErrorDialog

    If ValidHandle(handle) Then
        m_Connections(handle).LastError = errType
        m_Connections(handle).LastErrorCode = ErrCode
        m_Connections(handle).TechnicalDetails = technicalMsg
        showDialog = m_Connections(handle).EnableErrorDialog
    End If

    WasabiLog handle, "Error Code: " & errType
    WasabiLog handle, "Technical: " & technicalMsg
    If ErrCode <> 0 Then
        WasabiLog handle, "System Code: " & ErrCode & " (0x" & hex(ErrCode) & ")"
    End If
    WasabiLog handle, String(70, "-")

    If showDialog Then
        MsgBox userMsg, vbCritical, "WebSocket Connection Failed"
    End If
End Sub

Public Function WebSocketGetAllHandles() As Long()
    Dim result() As Long
    Dim i As Long
    Dim idx As Long
    Dim count As Long

    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Then count = count + 1
    Next i

    If count = 0 Then
        WebSocketGetAllHandles = result
        Exit Function
    End If

    ReDim result(0 To count - 1)
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Then
            result(idx) = i
            idx = idx + 1
        End If
    Next i

    WebSocketGetAllHandles = result
End Function

Public Function WebSocketSetDefaultHandle(ByVal handle As Long) As Boolean
    If Not ValidHandle(handle) Then Exit Function
    If Not m_Connections(handle).Connected Then Exit Function
    m_DefaultHandle = handle
    WebSocketSetDefaultHandle = True
End Function

Public Function WebSocketGetDefaultHandle() As Long
    WebSocketGetDefaultHandle = m_DefaultHandle
End Function

Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    With m_Connections(h).Stats
        .BytesSent = 0
        .BytesReceived = 0
        .MessagesSent = 0
        .MessagesReceived = 0
        .ConnectedAt = GetTickCount()
    End With
End Sub

Private Sub CleanupHandle(ByVal handle As Long)
    If Not ValidHandle(handle) Then Exit Sub

    With m_Connections(handle)
        If .Socket <> INVALID_SOCKET Then
            sock_closesocket .Socket
            .Socket = INVALID_SOCKET
        End If
        If .hContext.dwLower <> 0 Or .hContext.dwUpper <> 0 Then
            DeleteSecurityContext .hContext
            .hContext.dwLower = 0
            .hContext.dwUpper = 0
        End If
        If .hCred.dwLower <> 0 Or .hCred.dwUpper <> 0 Then
            FreeCredentialsHandle .hCred
            .hCred.dwLower = 0
            .hCred.dwUpper = 0
        End If
        .Connected = False
        .TLS = False
        .Host = ""
        .Port = 0
        .Path = ""
        .recvLen = 0
        .DecryptLen = 0
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
        .FragmentLen = 0
        .Fragmenting = False
        .FragmentOpcode = 0
        .LastPingSentAt = 0
        .ReconnectAttempts = 0
        .OriginalUrl = ""
        .AutoReconnect = False
        .ReconnectMaxAttempts = DEFAULT_RECONNECT_MAX_ATTEMPTS
        .ReconnectBaseDelayMs = DEFAULT_RECONNECT_BASE_DELAY_MS
        .PingIntervalMs = DEFAULT_PING_INTERVAL_MS
        .ReceiveTimeoutMs = DEFAULT_RECEIVE_TIMEOUT_MS
        .InactivityTimeoutMs = 0
        .LastActivityAt = 0
        .SubProtocol = ""
        .LogCallback = ""
        .EnableErrorDialog = False
        .CustomHeaderCount = 0
        .NoDelay = False
        .ProxyEnabled = False
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""
    End With
End Sub

Private Sub CleanupConnection()
    CleanupHandle m_DefaultHandle
End Sub

Private Function TickDiff(ByVal startTick As Long, ByVal endTick As Long) As Long
    If endTick >= startTick Then
        TickDiff = endTick - startTick
    Else
        TickDiff = (2147483647 - startTick) + endTick + 1
    End If
End Function

Private Function SafeArrayLen(ByRef arr() As Byte) As Long
    Dim lo As Long
    Dim hi As Long
    lo = LBound(arr)
    hi = UBound(arr)
    If hi >= lo Then
        SafeArrayLen = hi - lo + 1
    Else
        SafeArrayLen = 0
    End If
End Function

Private Function ApplyTCPNoDelayFor(ByVal handle As Long) As Boolean
    Dim optVal As Long
    Dim result As Long

    If Not ValidHandle(handle) Then Exit Function

    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Function
        optVal = IIf(.NoDelay, 1, 0)
        result = sock_setsockopt(.Socket, IPPROTO_TCP, TCP_NODELAY, optVal, 4)
        If result <> 0 Then
            Dim wsaErr As Long
            wsaErr = WSAGetLastError()
            WasabiLog handle, "Warning: setsockopt(TCP_NODELAY=" & optVal & ") failed with WSA error " & wsaErr & " (handle=" & handle & ")"
            ApplyTCPNoDelayFor = False
        Else
            WasabiLog handle, "TCP_NODELAY=" & optVal & " applied (handle=" & handle & ")"
            ApplyTCPNoDelayFor = True
        End If
    End With
End Function

Private Sub ApplyKeepAliveFor(ByVal handle As Long)
    Dim optVal As Long
    Dim result As Long

    If Not ValidHandle(handle) Then Exit Sub

    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub
        optVal = 1
        result = sock_setsockopt(.Socket, SOL_SOCKET, SO_KEEPALIVE, optVal, 4)
        If result <> 0 Then
            Dim wsaErr As Long
            wsaErr = WSAGetLastError()
            WasabiLog handle, "Warning: setsockopt(SO_KEEPALIVE) failed with WSA error " & wsaErr & " (handle=" & handle & ")"
        Else
            WasabiLog handle, "SO_KEEPALIVE applied (handle=" & handle & ")"
        End If
    End With
End Sub

Private Sub ApplySocketBuffersFor(ByVal handle As Long)
    Dim optVal As Long
    Dim result As Long

    If Not ValidHandle(handle) Then Exit Sub

    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub

        optVal = BUFFER_SIZE
        result = sock_setsockopt(.Socket, SOL_SOCKET, SO_RCVBUF, optVal, 4)
        If result <> 0 Then
            Dim wsaErr As Long
            wsaErr = WSAGetLastError()
            WasabiLog handle, "Warning: setsockopt(SO_RCVBUF) failed with WSA error " & wsaErr & " (handle=" & handle & ")"
        Else
            WasabiLog handle, "SO_RCVBUF=" & BUFFER_SIZE & " applied (handle=" & handle & ")"
        End If

        result = sock_setsockopt(.Socket, SOL_SOCKET, SO_SNDBUF, optVal, 4)
        If result <> 0 Then
            wsaErr = WSAGetLastError()
            WasabiLog handle, "Warning: setsockopt(SO_SNDBUF) failed with WSA error " & wsaErr & " (handle=" & handle & ")"
        Else
            WasabiLog handle, "SO_SNDBUF=" & BUFFER_SIZE & " applied (handle=" & handle & ")"
        End If
    End With
End Sub

Private Function DoProxyConnectFor(ByVal handle As Long) As Boolean
    Dim req As String
    Dim sendBuf() As Byte
    Dim recvBuf() As Byte
    Dim received As Long
    Dim response As String
    Dim sendResult As Long
    Dim wsaErr As Long

    With m_Connections(handle)
        req = "CONNECT " & .Host & ":" & .Port & " HTTP/1.1" & vbCrLf
        req = req & "Host: " & .Host & ":" & .Port & vbCrLf

        If .proxyUser <> "" Then
            Dim creds As String
            Dim credsB64 As String
            creds = .proxyUser & ":" & .proxyPass
            credsB64 = Base64Encode(StrConv(creds, vbFromUnicode))
            req = req & "Proxy-Authorization: Basic " & credsB64 & vbCrLf
        End If

        req = req & vbCrLf

        sendBuf = StrConv(req, vbFromUnicode)
        sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
        If sendResult <= 0 Then
            wsaErr = WSAGetLastError()
            LogError ERR_PROXY_CONNECT_FAILED, "send() to proxy failed with WSA error " & wsaErr, "Failed to send CONNECT to proxy.", wsaErr, handle
            Exit Function
        End If

        If Not WaitForDataOn(handle, 5000) Then
            LogError ERR_PROXY_CONNECT_FAILED, "Proxy CONNECT timeout", "Proxy did not respond to CONNECT request.", 0, handle
            Exit Function
        End If

        ReDim recvBuf(0 To 4095)
        received = sock_recv(.Socket, recvBuf(0), 4096, 0)
        If received <= 0 Then
            wsaErr = WSAGetLastError()
            LogError ERR_PROXY_CONNECT_FAILED, "recv() from proxy failed with WSA error " & wsaErr, "Failed to receive proxy response.", wsaErr, handle
            Exit Function
        End If

        response = Left(StrConv(recvBuf, vbUnicode), received)

        If InStr(response, "407") > 0 Then
            LogError ERR_PROXY_AUTH_FAILED, "Proxy returned 407 Proxy Authentication Required", "Proxy authentication failed." & vbCrLf & "Please check your proxy credentials.", 0, handle
            Exit Function
        End If

        If InStr(response, "200") = 0 Then
            Dim statusLine As String
            Dim lineEnd As Long
            lineEnd = InStr(response, vbCrLf)
            If lineEnd > 0 Then
                statusLine = Left(response, lineEnd - 1)
            Else
                statusLine = Left(response, 50)
            End If
            LogError ERR_PROXY_TUNNEL_FAILED, "Proxy CONNECT rejected: " & statusLine, "Proxy refused the tunnel connection." & vbCrLf & "Server: " & .Host & ":" & .Port, 0, handle
            Exit Function
        End If
    End With

    DoProxyConnectFor = True
End Function

Private Function Base64Encode(ByRef bytes() As Byte) As String
    Dim base64Chars As String
    Dim result() As String
    Dim i As Long
    Dim b1 As Long, b2 As Long, b3 As Long
    Dim chunk As Long
    Dim dataLen As Long
    Dim resultIdx As Long

    base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    dataLen = UBound(bytes) - LBound(bytes) + 1
    
    ReDim result(0 To ((dataLen + 2) \ 3) * 4)
    resultIdx = 0

    For i = LBound(bytes) To LBound(bytes) + dataLen - 1 Step 3
        b1 = CLng(bytes(i))
        If i + 1 <= LBound(bytes) + dataLen - 1 Then b2 = CLng(bytes(i + 1)) Else b2 = 0
        If i + 2 <= LBound(bytes) + dataLen - 1 Then b3 = CLng(bytes(i + 2)) Else b3 = 0

        chunk = b1 * 65536 + b2 * 256 + b3

        result(resultIdx) = Mid(base64Chars, (chunk \ 262144) + 1, 1)
        resultIdx = resultIdx + 1
        result(resultIdx) = Mid(base64Chars, ((chunk \ 4096) And 63) + 1, 1)
        resultIdx = resultIdx + 1

        If i + 1 <= LBound(bytes) + dataLen - 1 Then
            result(resultIdx) = Mid(base64Chars, ((chunk \ 64) And 63) + 1, 1)
        Else
            result(resultIdx) = "="
        End If
        resultIdx = resultIdx + 1

        If i + 2 <= LBound(bytes) + dataLen - 1 Then
            result(resultIdx) = Mid(base64Chars, (chunk And 63) + 1, 1)
        Else
            result(resultIdx) = "="
        End If
        resultIdx = resultIdx + 1
    Next i

    Base64Encode = Join(result, "")
End Function

Private Function ConnectHandle(ByVal handle As Long, ByVal url As String) As Boolean
    Dim addr As SOCKADDR_IN
    Dim schannelCred As SCHANNEL_CRED
    Dim tsExpiry As SECURITY_INTEGER
    Dim wsaErr As Long
    Dim connectHost As String
    Dim connectPort As Long
    Dim nbMode As Long
    Dim writeSet As FD_SET
    Dim exceptSet As FD_SET
    Dim tvConnect As TIMEVAL
    Dim selResult As Long
    Dim connectResult As Long

    With m_Connections(handle)
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""
        .OriginalUrl = url

        If Not ParseURLInto(.Host, .Port, .Path, .TLS, url) Then
            LogError ERR_INVALID_URL, "Invalid URL format: " & url, "Invalid WebSocket URL format." & vbCrLf & vbCrLf & "Expected format:" & vbCrLf & "ws://hostname/path or wss://hostname/path", 0, handle
            GoTo Fail
        End If

        .Socket = sock_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If .Socket = INVALID_SOCKET Then
            wsaErr = WSAGetLastError()
            LogError ERR_SOCKET_CREATE_FAILED, "socket() failed with WSA error " & wsaErr, "Failed to create network socket." & vbCrLf & "Please check your network configuration.", wsaErr, handle
            GoTo Fail
        End If

        If .ProxyEnabled And .proxyHost <> "" And .proxyPort > 0 Then
            connectHost = .proxyHost
            connectPort = .proxyPort
        Else
            connectHost = .Host
            connectPort = .Port
        End If

        addr.sin_family = AF_INET
        addr.sin_port = sock_htons(connectPort)
        addr.sin_addr = ResolveHostStr(connectHost, handle)
        If addr.sin_addr = 0 Then GoTo Fail

        nbMode = 1
        sock_ioctlsocket .Socket, FIONBIO, nbMode

        connectResult = sock_connect(.Socket, addr, LenB(addr))
        If connectResult <> 0 Then
            wsaErr = WSAGetLastError()
            If wsaErr <> 10035 Then
                LogError ERR_CONNECT_FAILED, "connect() to " & connectHost & ":" & connectPort & " failed with WSA error " & wsaErr, "Unable to connect to server." & vbCrLf & vbCrLf & "Server: " & connectHost & vbCrLf & "Port: " & connectPort & vbCrLf & vbCrLf & "Please verify the server address and your network connection.", wsaErr, handle
                GoTo Fail
            End If
        End If

        writeSet.fd_count = 1
        writeSet.fd_array(0) = .Socket
        exceptSet.fd_count = 1
        exceptSet.fd_array(0) = .Socket
        tvConnect.tv_sec = 10
        tvConnect.tv_usec = 0

        selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tvConnect)
        If selResult <= 0 Or exceptSet.fd_count > 0 Then
            LogError ERR_CONNECT_FAILED, "connect() timeout or refused connecting to " & connectHost & ":" & connectPort, "Unable to connect to server." & vbCrLf & vbCrLf & "Server: " & connectHost & vbCrLf & "Port: " & connectPort & vbCrLf & vbCrLf & "Connection timed out.", 0, handle
            GoTo Fail
        End If

        nbMode = 0
        sock_ioctlsocket .Socket, FIONBIO, nbMode

        ApplyTCPNoDelayFor handle
        ApplyKeepAliveFor handle
        ApplySocketBuffersFor handle

        If .ProxyEnabled And .proxyHost <> "" And .proxyPort > 0 Then
            If Not DoProxyConnectFor(handle) Then GoTo Fail
        End If

        If .TLS Then
            Dim credBytes() As Byte
            ReDim credBytes(0 To LenB(schannelCred) - 1)
            CopyMemory schannelCred, credBytes(0), LenB(schannelCred)
            schannelCred.dwVersion = SCHANNEL_CRED_VERSION
            schannelCred.grbitEnabledProtocols = SP_PROT_TLS1_2_CLIENT Or SP_PROT_TLS1_3_CLIENT
            schannelCred.dwFlags = SCH_CRED_NO_DEFAULT_CREDS Or SCH_CRED_MANUAL_CRED_VALIDATION Or SCH_CRED_IGNORE_NO_REVOCATION_CHECK Or SCH_CRED_IGNORE_REVOCATION_OFFLINE

            Dim acquireResult As Long
            acquireResult = AcquireCredentialsHandle(NULL_PTR, "Microsoft Unified Security Protocol Provider", SECPKG_CRED_OUTBOUND, NULL_PTR, schannelCred, NULL_PTR, NULL_PTR, .hCred, tsExpiry)
            If acquireResult <> 0 Then
                LogError ERR_TLS_ACQUIRE_CREDS_FAILED, "AcquireCredentialsHandle failed with SSPI error 0x" & hex(acquireResult), "SSL/TLS initialization failed." & vbCrLf & "Please check your system security settings.", acquireResult, handle
                GoTo Fail
            End If

            Dim tlsResult As Long
            tlsResult = DoTLSHandshakeFor(handle)
            If tlsResult <> 0 Then
                If tlsResult = -1 Then
                    LogError ERR_TLS_HANDSHAKE_TIMEOUT, "TLS handshake timeout with " & .Host, "SSL/TLS connection timeout." & vbCrLf & vbCrLf & "The server did not respond to the security handshake.", 0, handle
                Else
                    LogError ERR_TLS_HANDSHAKE_FAILED, "InitializeSecurityContext failed with SSPI error 0x" & hex(tlsResult), "SSL/TLS handshake failed." & vbCrLf & vbCrLf & "Unable to establish secure connection with server.", tlsResult, handle
                End If
                GoTo Fail
            End If

            QueryContextAttributes .hContext, SECPKG_ATTR_STREAM_SIZES, .Sizes
        End If

        If Not DoWebSocketHandshakeFor(handle) Then GoTo Fail

        .Connected = True
        .Stats.ConnectedAt = GetTickCount()
        .Stats.BytesSent = 0
        .Stats.BytesReceived = 0
        .Stats.MessagesSent = 0
        .Stats.MessagesReceived = 0
        .LastPingSentAt = GetTickCount()
    End With

    ConnectHandle = True
    Exit Function

Fail:
    CleanupHandle handle
End Function

Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1) As Boolean
    Dim wsa As WSADATA
    Dim wsaErr As Long
    Dim handle As Long

    m_LastError = ERR_NONE
    m_LastErrorCode = 0
    m_TechnicalDetails = ""

    InitConnectionPool

    If Not m_WSAInitialized Then
        wsaErr = WSAStartup(&H202, wsa)
        If wsaErr <> 0 Then
            LogError ERR_WSA_STARTUP_FAILED, "WSAStartup failed with code " & wsaErr, "Network initialization failed." & vbCrLf & "Please check your network settings.", wsaErr
            Exit Function
        End If
        m_WSAInitialized = True
    End If

    handle = AllocConnection()
    If handle = INVALID_CONN_HANDLE Then
        LogError ERR_MAX_CONNECTIONS, "Maximum concurrent connections reached (" & MAX_CONNECTIONS & ")", "Too many simultaneous connections."
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    If Not ConnectHandle(handle, url) Then
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If

    m_DefaultHandle = handle
    outHandle = handle
    WebSocketConnect = True
    WasabiLog handle, "Successfully connected to " & url & " (handle=" & handle & ")"
End Function

Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim reasonBytes() As Byte
    Dim reasonLen As Long
    Dim payloadLen As Long
    Dim totalSent As Long
    Dim toSend As Long
    Dim sendResult As Long
    Dim i As Long
    Dim h As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then Exit Function

        If Len(reason) > 0 Then
            reasonBytes = StringToUtf8Safe(reason)
            reasonLen = SafeArrayLen(reasonBytes)
        Else
            reasonLen = 0
        End If

        payloadLen = 2 + reasonLen

        ReDim frame(0 To 5 + payloadLen)

        For i = 0 To 3
            mask(i) = Int(Rnd * 256)
        Next i

        frame(0) = &H88
        frame(1) = &H80 Or CByte(payloadLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)

        Dim codeByte0 As Byte
        Dim codeByte1 As Byte
        codeByte0 = CByte((code \ 256) And &HFF)
        codeByte1 = CByte(code And &HFF)
        frame(6) = codeByte0 Xor mask(0)
        frame(7) = codeByte1 Xor mask(1)

        For i = 0 To reasonLen - 1
            frame(8 + i) = reasonBytes(i) Xor mask((i + 2) Mod 4)
        Next i

        If .TLS Then
            WebSocketSendClose = TLSSendFor(h, frame)
        Else
            totalSent = 0
            toSend = UBound(frame) + 1
            Do While totalSent < toSend
                sendResult = sock_send(.Socket, frame(totalSent), toSend - totalSent, 0)
                If sendResult <= 0 Then
                    Dim wsaErr As Long
                    wsaErr = WSAGetLastError()
                    LogError ERR_SEND_FAILED, "send() failed with WSA error " & wsaErr, "Failed to send CLOSE frame." & vbCrLf & "Connection may have been lost.", wsaErr, h
                    Exit Function
                End If
                totalSent = totalSent + sendResult
            Loop
            WebSocketSendClose = True
        End If

        .Connected = False
    End With
End Function

Public Sub WebSocketDisconnectAll()
    Dim i As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Or m_Connections(i).Socket <> INVALID_SOCKET Then
            WebSocketDisconnect i
        End If
    Next i
End Sub

Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        .AutoReconnect = False
        If .Connected Then WebSocketSendClose 1000, "", h
    End With

    CleanupHandle h

    If h = m_DefaultHandle Then m_DefaultHandle = 0

    Dim anyActive As Boolean
    Dim i As Long
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Or m_Connections(i).Socket <> INVALID_SOCKET Then
            anyActive = True
            Exit For
        End If
    Next i

    If Not anyActive And m_WSAInitialized Then
        WSACleanup
        m_WSAInitialized = False
    End If
End Sub

Public Function WebSocketGetQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetQueueCapacity = MSG_QUEUE_SIZE - m_Connections(h).MsgCount
End Function

Public Function WebSocketGetBinaryQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetBinaryQueueCapacity = MSG_QUEUE_SIZE - m_Connections(h).BinaryCount
End Function

Public Function WebSocketGetConnectionCount() As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Then count = count + 1
    Next i
    WebSocketGetConnectionCount = count
End Function

Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Or .Stats.ConnectedAt = 0 Then Exit Function
        WebSocketGetUptime = TickDiff(.Stats.ConnectedAt, GetTickCount()) \ 1000
    End With
End Function

Public Function WebSocketSend(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim msgBytes() As Byte
    Dim frame() As Byte
    Dim msgLen As Long
    Dim mask(0 To 3) As Byte
    Dim i As Long
    Dim headerLen As Long
    Dim h As Long
    Dim totalSent As Long
    Dim toSend As Long
    Dim sendResult As Long
    Dim sendOk As Boolean

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then
            LogError ERR_NOT_CONNECTED, "Send attempted while not connected", "Not connected to WebSocket server.", 0, h
            Exit Function
        End If

        msgBytes = StringToUtf8Safe(message)
        msgLen = SafeArrayLen(msgBytes)

        If msgLen = 0 Then
            WebSocketSend = True
            Exit Function
        End If

        For i = 0 To 3
            mask(i) = Int(Rnd * 256)
        Next i

        If msgLen < 126 Then
            headerLen = 6
            ReDim frame(0 To headerLen + msgLen - 1)
            frame(0) = &H81
            frame(1) = &H80 Or CByte(msgLen)
            frame(2) = mask(0)
            frame(3) = mask(1)
            frame(4) = mask(2)
            frame(5) = mask(3)
        ElseIf msgLen < 65536 Then
            headerLen = 8
            ReDim frame(0 To headerLen + msgLen - 1)
            frame(0) = &H81
            frame(1) = &HFE
            frame(2) = CByte((msgLen \ 256) And &HFF)
            frame(3) = CByte(msgLen And &HFF)
            frame(4) = mask(0)
            frame(5) = mask(1)
            frame(6) = mask(2)
            frame(7) = mask(3)
        Else
            headerLen = 14
            ReDim frame(0 To headerLen + msgLen - 1)
            frame(0) = &H81
            frame(1) = &HFF
            frame(2) = 0
            frame(3) = 0
            frame(4) = 0
            frame(5) = 0
            frame(6) = CByte((msgLen \ 16777216) And &HFF)
            frame(7) = CByte((msgLen \ 65536) And &HFF)
            frame(8) = CByte((msgLen \ 256) And &HFF)
            frame(9) = CByte(msgLen And &HFF)
            frame(10) = mask(0)
            frame(11) = mask(1)
            frame(12) = mask(2)
            frame(13) = mask(3)
        End If

        For i = 0 To msgLen - 1
            frame(headerLen + i) = msgBytes(i) Xor mask(i Mod 4)
        Next i

        If .TLS Then
            sendOk = TLSSendFor(h, frame)
        Else
            totalSent = 0
            toSend = UBound(frame) + 1
            Do While totalSent < toSend
                sendResult = sock_send(.Socket, frame(totalSent), toSend - totalSent, 0)
                If sendResult <= 0 Then
                    Dim wsaErr As Long
                    wsaErr = WSAGetLastError()
                    LogError ERR_SEND_FAILED, "send() failed with WSA error " & wsaErr, "Failed to send message." & vbCrLf & "Connection may have been lost.", wsaErr, h
                    .Connected = False
                    Exit Function
                End If
                totalSent = totalSent + sendResult
            Loop
            sendOk = True
        End If

        If sendOk Then
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            .Stats.MessagesSent = .Stats.MessagesSent + 1
        End If

        WebSocketSend = sendOk
    End With
End Function

Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim frame() As Byte
    Dim dataLen As Long
    Dim mask(0 To 3) As Byte
    Dim i As Long
    Dim headerLen As Long
    Dim totalSent As Long
    Dim toSend As Long
    Dim sendResult As Long
    Dim h As Long
    Dim sendOk As Boolean

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then
            LogError ERR_NOT_CONNECTED, "Send attempted while not connected", "Not connected to WebSocket server.", 0, h
            Exit Function
        End If

        dataLen = SafeArrayLen(data)

        If dataLen = 0 Then
            WebSocketSendBinary = True
            Exit Function
        End If

        For i = 0 To 3
            mask(i) = Int(Rnd * 256)
        Next i

        If dataLen < 126 Then
            headerLen = 6
            ReDim frame(0 To headerLen + dataLen - 1)
            frame(0) = &H82
            frame(1) = &H80 Or CByte(dataLen)
            frame(2) = mask(0)
            frame(3) = mask(1)
            frame(4) = mask(2)
            frame(5) = mask(3)
        ElseIf dataLen < 65536 Then
            headerLen = 8
            ReDim frame(0 To headerLen + dataLen - 1)
            frame(0) = &H82
            frame(1) = &HFE
            frame(2) = CByte((dataLen \ 256) And &HFF)
            frame(3) = CByte(dataLen And &HFF)
            frame(4) = mask(0)
            frame(5) = mask(1)
            frame(6) = mask(2)
            frame(7) = mask(3)
        Else
            headerLen = 14
            ReDim frame(0 To headerLen + dataLen - 1)
            frame(0) = &H82
            frame(1) = &HFF
            frame(2) = 0
            frame(3) = 0
            frame(4) = 0
            frame(5) = 0
            frame(6) = CByte((dataLen \ 16777216) And &HFF)
            frame(7) = CByte((dataLen \ 65536) And &HFF)
            frame(8) = CByte((dataLen \ 256) And &HFF)
            frame(9) = CByte(dataLen And &HFF)
            frame(10) = mask(0)
            frame(11) = mask(1)
            frame(12) = mask(2)
            frame(13) = mask(3)
        End If

        For i = 0 To dataLen - 1
            frame(headerLen + i) = data(LBound(data) + i) Xor mask(i Mod 4)
        Next i

        If .TLS Then
            sendOk = TLSSendFor(h, frame)
        Else
            totalSent = 0
            toSend = UBound(frame) + 1
            Do While totalSent < toSend
                sendResult = sock_send(.Socket, frame(totalSent), toSend - totalSent, 0)
                If sendResult <= 0 Then
                    Dim wsaErr As Long
                    wsaErr = WSAGetLastError()
                    LogError ERR_SEND_FAILED, "send() failed with WSA error " & wsaErr, "Failed to send binary data." & vbCrLf & "Connection may have been lost.", wsaErr, h
                    .Connected = False
                    Exit Function
                End If
                totalSent = totalSent + sendResult
            Loop
            sendOk = True
        End If

        If sendOk Then
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            .Stats.MessagesSent = .Stats.MessagesSent + 1
        End If

        WebSocketSendBinary = sendOk
    End With
End Function

Public Function WebSocketBroadcast(ByVal message As String) As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Then
            If WebSocketSend(message, i) Then count = count + 1
        End If
    Next i
    WebSocketBroadcast = count
End Function

Public Function WebSocketBroadcastBinary(ByRef data() As Byte) As Long
    Dim i As Long
    Dim count As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Then
            If WebSocketSendBinary(data, i) Then count = count + 1
        End If
    Next i
    WebSocketBroadcastBinary = count
End Function

Private Sub TickAndFeedBuffer(ByVal handle As Long)
    Dim available As Long
    Dim tempBuf() As Byte
    Dim received As Long
    Dim wsaErr As Long

    With m_Connections(handle)
        If sock_ioctlsocket(.Socket, FIONREAD, available) <> 0 Then
            wsaErr = WSAGetLastError()
            LogError ERR_CONNECTION_LOST, "ioctlsocket failed with WSA error " & wsaErr, "Connection lost.", wsaErr, handle
            .Connected = False
            If .AutoReconnect Then TryReconnectFor handle
            Exit Sub
        End If

        If available <= 0 Then Exit Sub
        If available > BUFFER_SIZE \ 2 Then available = BUFFER_SIZE \ 2

        ReDim tempBuf(0 To available - 1)
        received = sock_recv(.Socket, tempBuf(0), available, 0)

        If received > 0 Then
            .Stats.BytesReceived = .Stats.BytesReceived + received
            .LastActivityAt = GetTickCount()
            If .TLS Then
                Dim copyLenTLS As Long
                copyLenTLS = received
                If .recvLen + copyLenTLS > BUFFER_SIZE Then copyLenTLS = BUFFER_SIZE - .recvLen
                If copyLenTLS > 0 Then
                    CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLenTLS
                    .recvLen = .recvLen + copyLenTLS
                End If
                TLSDecryptFor handle
            Else
                Dim copyLenRaw As Long
                copyLenRaw = received
                If .DecryptLen + copyLenRaw > BUFFER_SIZE Then copyLenRaw = BUFFER_SIZE - .DecryptLen
                If copyLenRaw > 0 Then
                    CopyMemory .DecryptBuffer(.DecryptLen), tempBuf(0), copyLenRaw
                    .DecryptLen = .DecryptLen + copyLenRaw
                End If
            End If
            ProcessFramesFor handle
        ElseIf received = 0 Then
            LogError ERR_CONNECTION_LOST, "recv() returned 0 (connection closed by peer)", "Server closed the connection.", 0, handle
            .Connected = False
            If .AutoReconnect Then TryReconnectFor handle
        Else
            wsaErr = WSAGetLastError()
            LogError ERR_RECV_FAILED, "recv() failed with WSA error " & wsaErr, "Failed to receive data.", wsaErr, handle
            .Connected = False
            If .AutoReconnect Then TryReconnectFor handle
        End If
    End With
End Sub

Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
    Dim h As Long
    Dim results() As String
    Dim count As Long
    Dim i As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then
        ReDim results(0 To 0)
        WebSocketReceiveAll = results
        Exit Function
    End If

    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnectFor h
            ReDim results(0 To 0)
            WebSocketReceiveAll = results
            Exit Function
        End If

        TickMaintenanceFor h
        If .DecryptLen > 0 Then ProcessFramesFor h
        TickAndFeedBuffer h

        count = .MsgCount
        If count = 0 Then
            ReDim results(0 To 0)
            WebSocketReceiveAll = results
            Exit Function
        End If

        ReDim results(0 To count - 1)
        For i = 0 To count - 1
            results(i) = .MsgQueue(.MsgHead)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
        Next i
    End With

    WebSocketReceiveAll = results
End Function

Public Function WebSocketReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnectFor h
            Exit Function
        End If

        TickMaintenanceFor h
        If .DecryptLen > 0 Then ProcessFramesFor h
        TickAndFeedBuffer h

        If .MsgCount > 0 Then
            WebSocketReceive = .MsgQueue(.MsgHead)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
        End If
    End With
End Function

Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long

    WebSocketReceiveBinaryCheck = False

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnectFor h
            Exit Function
        End If

        TickMaintenanceFor h
        If .DecryptLen > 0 Then ProcessFramesFor h
        TickAndFeedBuffer h

        If .BinaryCount > 0 Then
            outData = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
            WebSocketReceiveBinaryCheck = True
        End If
    End With
End Function

Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
    Dim emptyArray() As Byte
    Dim h As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then
        WebSocketReceiveBinary = emptyArray
        Exit Function
    End If

    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnectFor h
            WebSocketReceiveBinary = emptyArray
            Exit Function
        End If

        TickMaintenanceFor h
        If .DecryptLen > 0 Then ProcessFramesFor h
        TickAndFeedBuffer h

        If .BinaryCount > 0 Then
            WebSocketReceiveBinary = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
        Else
            WebSocketReceiveBinary = emptyArray
        End If
    End With
End Function

Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketIsConnected = m_Connections(h).Connected
End Function

Public Function WebSocketGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidHandle(h) Then
        WebSocketGetLastError = m_Connections(h).LastError
    Else
        WebSocketGetLastError = m_LastError
    End If
End Function

Public Function WebSocketGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidHandle(h) Then
        WebSocketGetLastErrorCode = m_Connections(h).LastErrorCode
    Else
        WebSocketGetLastErrorCode = m_LastErrorCode
    End If
End Function

Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidHandle(h) Then
        WebSocketGetTechnicalDetails = m_Connections(h).TechnicalDetails
    Else
        WebSocketGetTechnicalDetails = m_TechnicalDetails
    End If
End Function

Public Function WebSocketGetPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetPendingCount = m_Connections(h).MsgCount
End Function

Public Function WebSocketGetBinaryPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetBinaryPendingCount = m_Connections(h).BinaryCount
End Function

Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    With m_Connections(h)
        If .MsgCount > 0 Then WebSocketPeek = .MsgQueue(.MsgHead)
    End With
End Function

Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    With m_Connections(h)
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
    End With
End Sub

Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim pingBytes() As Byte
    Dim pingLen As Long
    Dim i As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then Exit Function

        If Len(payload) > 0 Then
            pingBytes = StrConv(payload, vbFromUnicode)
            pingLen = UBound(pingBytes) + 1
        Else
            pingLen = 0
        End If

        For i = 0 To 3
            mask(i) = Int(Rnd * 256)
        Next i

        If pingLen = 0 Then
            ReDim frame(0 To 5)
            frame(0) = &H89
            frame(1) = &H80
        Else
            ReDim frame(0 To 5 + pingLen)
            frame(0) = &H89
            frame(1) = &H80 Or CByte(pingLen)
            For i = 0 To pingLen - 1
                frame(6 + i) = pingBytes(i) Xor mask(i Mod 4)
            Next i
        End If
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)

        If .TLS Then
            WebSocketSendPing = TLSSendFor(h, frame)
        Else
            Dim sendResult As Long
            sendResult = sock_send(.Socket, frame(0), UBound(frame) + 1, 0)
            WebSocketSendPing = (sendResult > 0)
        End If

        If WebSocketSendPing Then
            .LastPingSentAt = GetTickCount()
            WasabiLog h, "Ping sent (handle=" & h & ")"
        End If
    End With
End Function

Public Function WebSocketSendPong(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim pongBytes() As Byte
    Dim pongLen As Long
    Dim i As Long

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If Not .Connected Then Exit Function

        If Len(payload) > 0 Then
            pongBytes = StrConv(payload, vbFromUnicode)
            pongLen = UBound(pongBytes) + 1
        Else
            pongLen = 0
        End If

        For i = 0 To 3
            mask(i) = Int(Rnd * 256)
        Next i

        If pongLen = 0 Then
            ReDim frame(0 To 5)
            frame(0) = &H8A
            frame(1) = &H80
        Else
            ReDim frame(0 To 5 + pongLen)
            frame(0) = &H8A
            frame(1) = &H80 Or CByte(pongLen)
            For i = 0 To pongLen - 1
                frame(6 + i) = pongBytes(i) Xor mask(i Mod 4)
            Next i
        End If
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)

        If .TLS Then
            WebSocketSendPong = TLSSendFor(h, frame)
        Else
            Dim sendResult As Long
            sendResult = sock_send(.Socket, frame(0), UBound(frame) + 1, 0)
            WebSocketSendPong = (sendResult > 0)
        End If
    End With
End Function

Public Sub WebSocketSetBufferSizes(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        If .Connected Then
            WasabiLog h, "Cannot change buffer sizes while connected (handle=" & h & ")"
            Exit Sub
        End If

        If bufferSize >= 8192 And bufferSize <= 16777216 Then
            .CustomBufferSize = bufferSize
            WasabiLog h, "Buffer size set to " & bufferSize & " bytes (handle=" & h & ")"
        Else
            WasabiLog h, "Invalid buffer size " & bufferSize & " (must be 8KB-16MB) (handle=" & h & ")"
        End If

        If fragmentSize >= 4096 And fragmentSize <= 16777216 Then
            .CustomFragmentSize = fragmentSize
            WasabiLog h, "Fragment size set to " & fragmentSize & " bytes (handle=" & h & ")"
        Else
            WasabiLog h, "Invalid fragment size " & fragmentSize & " (must be 4KB-16MB) (handle=" & h & ")"
        End If
    End With
End Sub

Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    m_Connections(h).NoDelay = enabled

    If m_Connections(h).Socket <> 0 And m_Connections(h).Socket <> -1 Then
        WebSocketSetNoDelay = ApplyTCPNoDelayFor(h)
    Else
        WebSocketSetNoDelay = True
    End If
End Function

Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetReconnectInfo = "AutoReconnect=" & IIf(.AutoReconnect, "1", "0") & _
            "|Attempts=" & .ReconnectAttempts & _
            "|MaxAttempts=" & .ReconnectMaxAttempts & _
            "|BaseDelayMs=" & .ReconnectBaseDelayMs
    End With
End Function

Public Function WebSocketGetSubProtocol(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetSubProtocol = m_Connections(h).SubProtocol
End Function

Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).SubProtocol = protocol
End Sub

Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).InactivityTimeoutMs = timeoutMs
    m_Connections(h).LastActivityAt = GetTickCount()
End Sub

Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        .proxyHost = proxyHost
        .proxyPort = proxyPort
        .proxyUser = proxyUser
        .proxyPass = proxyPass
        .ProxyEnabled = (Len(proxyHost) > 0 And proxyPort > 0)
    End With
End Sub

Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .ProxyEnabled = False
    End With
End Sub

Public Function WebSocketGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If .ProxyEnabled Then
            WebSocketGetProxyInfo = "Host=" & .proxyHost & "|Port=" & .proxyPort & "|Auth=" & IIf(.proxyUser <> "", "Yes", "No")
        Else
            WebSocketGetProxyInfo = "Disabled"
        End If
    End With
End Function

Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        .AutoReconnect = enabled
        .ReconnectMaxAttempts = maxAttempts
        .ReconnectBaseDelayMs = baseDelayMs
        .ReconnectAttempts = 0
    End With
End Sub

Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).PingIntervalMs = intervalMs
    m_Connections(h).LastPingSentAt = GetTickCount()
End Sub

Public Sub WebSocketSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).ReceiveTimeoutMs = timeoutMs
End Sub

Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub

    With m_Connections(h)
        If .CustomHeaderCount > UBound(.CustomHeaders) Then
            ReDim Preserve .CustomHeaders(0 To .CustomHeaderCount)
        End If
        .CustomHeaders(.CustomHeaderCount) = headerName & ": " & headerValue
        .CustomHeaderCount = .CustomHeaderCount + 1
    End With
End Sub

Public Sub WebSocketClearHeaders(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).CustomHeaderCount = 0
End Sub

Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).LogCallback = callbackName
End Sub

Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Sub
    m_Connections(h).EnableErrorDialog = enabled
End Sub

Public Function WebSocketGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetHost = m_Connections(h).Host
End Function

Public Function WebSocketGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetPort = m_Connections(h).Port
End Function

Public Function WebSocketGetPath(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function
    WebSocketGetPath = m_Connections(h).Path
End Function

Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim uptime As Long
    Dim result As String

    h = ResolveHandle(handle)
    If Not ValidHandle(h) Then Exit Function

    With m_Connections(h)
        If .Stats.ConnectedAt > 0 Then
            uptime = TickDiff(.Stats.ConnectedAt, GetTickCount()) \ 1000
        Else
            uptime = 0
        End If
        result = "BytesSent=" & Format(.Stats.BytesSent, "0") & _
                 "|BytesReceived=" & Format(.Stats.BytesReceived, "0") & _
                 "|MessagesSent=" & .Stats.MessagesSent & _
                 "|MessagesReceived=" & .Stats.MessagesReceived & _
                 "|UptimeSeconds=" & uptime & _
                 "|Queued=" & .MsgCount & _
                 "|BinaryQueued=" & .BinaryCount & _
                 "|NoDelay=" & IIf(.NoDelay, "1", "0") & _
                 "|Proxy=" & IIf(.ProxyEnabled, .proxyHost & ":" & .proxyPort, "none")
    End With

    WebSocketGetStats = result
End Function

Private Sub TickMaintenanceFor(ByVal handle As Long)
    With m_Connections(handle)
        If Not .Connected Then Exit Sub

        Dim now As Long
        now = GetTickCount()

        If .PingIntervalMs > 0 Then
            If TickDiff(.LastPingSentAt, now) >= .PingIntervalMs Then
                WebSocketSendPing "", handle
            End If
        End If

        If .InactivityTimeoutMs > 0 And .LastActivityAt > 0 Then
            If TickDiff(.LastActivityAt, now) >= .InactivityTimeoutMs Then
                WasabiLog handle, "Inactivity timeout reached (handle=" & handle & ")"
                LogError ERR_INACTIVITY_TIMEOUT, "No data received for " & .InactivityTimeoutMs & "ms", "Connection timed out due to inactivity.", 0, handle
                .Connected = False
                If .AutoReconnect Then TryReconnectFor handle
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub TryReconnectFor(ByVal handle As Long)
    Dim delayMs As Long
    Dim attempt As Long
    Dim i As Long

    If Not m_Connections(handle).AutoReconnect Then Exit Sub
    If m_Connections(handle).ReconnectMaxAttempts > 0 And m_Connections(handle).ReconnectAttempts >= m_Connections(handle).ReconnectMaxAttempts Then
        WasabiLog handle, "Auto-reconnect giving up after " & m_Connections(handle).ReconnectAttempts & " attempts (handle=" & handle & ")"
        m_Connections(handle).AutoReconnect = False
        Exit Sub
    End If

    m_Connections(handle).ReconnectAttempts = m_Connections(handle).ReconnectAttempts + 1
    attempt = m_Connections(handle).ReconnectAttempts

    delayMs = m_Connections(handle).ReconnectBaseDelayMs
    For i = 1 To attempt - 1
        delayMs = delayMs * 2
        If delayMs > MAX_RECONNECT_DELAY_MS Then
            delayMs = MAX_RECONNECT_DELAY_MS
            Exit For
        End If
    Next i

    WasabiLog handle, "Auto-reconnect attempt " & attempt & " in " & delayMs & "ms (handle=" & handle & ")"

    Dim savedUrl               As String
    Dim savedAutoReconnect     As Boolean
    Dim savedMaxAttempts       As Long
    Dim savedBaseDelay         As Long
    Dim savedAttempts          As Long
    Dim savedPingInterval      As Long
    Dim savedReceiveTimeout    As Long
    Dim savedLogCallback       As String
    Dim savedErrorDialog       As Boolean
    Dim savedHeaders()         As String
    Dim savedHeaderCount       As Long
    Dim savedNoDelay           As Boolean
    Dim savedProxyHost         As String
    Dim savedProxyPort         As Long
    Dim savedProxyUser         As String
    Dim savedProxyPass         As String
    Dim savedProxyEnabled      As Boolean
    Dim savedInactivityTimeout As Long
    Dim savedSubProtocol       As String

    With m_Connections(handle)
        savedUrl = .OriginalUrl
        savedAutoReconnect = .AutoReconnect
        savedMaxAttempts = .ReconnectMaxAttempts
        savedBaseDelay = .ReconnectBaseDelayMs
        savedAttempts = .ReconnectAttempts
        savedPingInterval = .PingIntervalMs
        savedReceiveTimeout = .ReceiveTimeoutMs
        savedLogCallback = .LogCallback
        savedErrorDialog = .EnableErrorDialog
        savedHeaderCount = .CustomHeaderCount
        savedNoDelay = .NoDelay
        savedProxyHost = .proxyHost
        savedProxyPort = .proxyPort
        savedProxyUser = .proxyUser
        savedProxyPass = .proxyPass
        savedProxyEnabled = .ProxyEnabled
        savedInactivityTimeout = .InactivityTimeoutMs
        savedSubProtocol = .SubProtocol

        If savedHeaderCount > 0 Then
            ReDim savedHeaders(0 To savedHeaderCount - 1)
            For i = 0 To savedHeaderCount - 1
                savedHeaders(i) = .CustomHeaders(i)
            Next i
        End If
    End With

    CleanupHandle handle
    Dim startTick As Long
    startTick = GetTickCount()
    Do
        DoEvents
        If TickDiff(startTick, GetTickCount()) >= delayMs Then Exit Do
    Loop

    If Not m_WSAInitialized Then
        Dim wsa As WSADATA
        WSAStartup &H202, wsa
        m_WSAInitialized = True
    End If

    Dim bufSize  As Long
    Dim fragSize As Long
    bufSize = IIf(m_Connections(handle).CustomBufferSize > 0, m_Connections(handle).CustomBufferSize, BUFFER_SIZE)
    fragSize = IIf(m_Connections(handle).CustomFragmentSize > 0, m_Connections(handle).CustomFragmentSize, FRAGMENT_BUFFER_SIZE)

    ReDim m_Connections(handle).recvBuffer(0 To bufSize - 1)
    ReDim m_Connections(handle).DecryptBuffer(0 To bufSize - 1)
    ReDim m_Connections(handle).MsgQueue(0 To MSG_QUEUE_SIZE - 1)
    ReDim m_Connections(handle).BinaryQueue(0 To MSG_QUEUE_SIZE - 1)
    ReDim m_Connections(handle).FragmentBuffer(0 To fragSize - 1)
    ReDim m_Connections(handle).CustomHeaders(0 To 31)

    With m_Connections(handle)
        .OriginalUrl = savedUrl
        .AutoReconnect = savedAutoReconnect
        .ReconnectMaxAttempts = savedMaxAttempts
        .ReconnectBaseDelayMs = savedBaseDelay
        .ReconnectAttempts = savedAttempts
        .PingIntervalMs = savedPingInterval
        .ReceiveTimeoutMs = savedReceiveTimeout
        .LogCallback = savedLogCallback
        .EnableErrorDialog = savedErrorDialog
        .CustomHeaderCount = savedHeaderCount
        .NoDelay = savedNoDelay
        .proxyHost = savedProxyHost
        .proxyPort = savedProxyPort
        .proxyUser = savedProxyUser
        .proxyPass = savedProxyPass
        .ProxyEnabled = savedProxyEnabled
        .InactivityTimeoutMs = savedInactivityTimeout
        .SubProtocol = savedSubProtocol
        .recvLen = 0
        .DecryptLen = 0
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
        .FragmentLen = 0
        .Fragmenting = False
        .FragmentOpcode = 0
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""

        For i = 0 To savedHeaderCount - 1
            .CustomHeaders(i) = savedHeaders(i)
        Next i
    End With

    If Not ConnectHandle(handle, savedUrl) Then
        WasabiLog handle, "Auto-reconnect failed (handle=" & handle & ")"
        Exit Sub
    End If

    m_Connections(handle).ReconnectAttempts = 0
    WasabiLog handle, "Auto-reconnect succeeded (handle=" & handle & ")"
End Sub

Private Function ResolveHandle(ByVal handle As Long) As Long
    If handle = INVALID_CONN_HANDLE Then
        ResolveHandle = m_DefaultHandle
    Else
        ResolveHandle = handle
    End If
End Function

Private Function StringToUtf8Safe(ByVal str As String) As Byte()
    Dim length As Long
    Dim buffer() As Byte

    If Len(str) = 0 Then
        StringToUtf8Safe = buffer
        Exit Function
    End If

    length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, 0, 0, NULL_PTR, NULL_PTR)

    If length > 1 Then
        ReDim buffer(0 To length - 2)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(str), -1, buffer(0), length, NULL_PTR, NULL_PTR
    End If

    StringToUtf8Safe = buffer
End Function

Private Function Utf8ToString(ByRef utf8Bytes() As Byte, Optional ByVal length As Long = -1) As String
    Dim charCount As Long
    Dim result As String
    Dim actualLength As Long

    If length = -1 Then
        actualLength = SafeArrayLen(utf8Bytes)
    Else
        actualLength = length
    End If

    If actualLength <= 0 Then
        Utf8ToString = ""
        Exit Function
    End If

    charCount = MultiByteToWideChar(CP_UTF8, 0, utf8Bytes(LBound(utf8Bytes)), actualLength, NULL_PTR, 0)

    If charCount > 0 Then
        result = String$(charCount, vbNullChar)
        MultiByteToWideChar CP_UTF8, 0, utf8Bytes(LBound(utf8Bytes)), actualLength, StrPtr(result), charCount
    End If

    Utf8ToString = result
End Function

Private Function WaitForDataOn(ByVal handle As Long, ByVal timeoutMs As Long) As Boolean
    Dim readSet As FD_SET
    Dim timeout As TIMEVAL
    Dim result As Long
    Dim effectiveTimeout As Long

    effectiveTimeout = timeoutMs
    If ValidHandle(handle) Then
        If m_Connections(handle).ReceiveTimeoutMs > 0 Then
            effectiveTimeout = m_Connections(handle).ReceiveTimeoutMs
        End If
    End If

    readSet.fd_count = 1
    readSet.fd_array(0) = m_Connections(handle).Socket
    timeout.tv_sec = effectiveTimeout \ 1000
    timeout.tv_usec = (effectiveTimeout Mod 1000) * 1000

    result = sock_select(0, readSet, ByVal 0&, ByVal 0&, timeout)
    WaitForDataOn = (result > 0)
End Function

Private Function WaitForData(ByVal timeoutMs As Long) As Boolean
    WaitForData = WaitForDataOn(m_DefaultHandle, timeoutMs)
End Function

Private Function ParseURLInto(ByRef outHost As String, ByRef outPort As Long, ByRef outPath As String, ByRef outTLS As Boolean, ByVal url As String) As Boolean
    Dim work As String
    Dim slashPos As Long
    Dim colonPos As Long
    Dim portStr As String
    Dim portVal As Long
    Dim i As Long
    Dim c As String

    If Len(Trim(url)) = 0 Then Exit Function

    work = url
    outTLS = False
    outPort = 80

    If LCase(Left(work, 6)) = "wss://" Then
        work = Mid(work, 7)
        outTLS = True
        outPort = 443
    ElseIf LCase(Left(work, 5)) = "ws://" Then
        work = Mid(work, 6)
    Else
        Exit Function
    End If

    If Len(work) = 0 Then Exit Function

    slashPos = InStr(work, "/")
    If slashPos > 0 Then
        outPath = Mid(work, slashPos)
        work = Left(work, slashPos - 1)
    Else
        outPath = "/"
    End If

    colonPos = InStr(work, ":")
    If colonPos > 0 Then
        outHost = Left(work, colonPos - 1)
        portStr = Mid(work, colonPos + 1)

        If Len(portStr) = 0 Then Exit Function

        For i = 1 To Len(portStr)
            c = Mid(portStr, i, 1)
            If c < "0" Or c > "9" Then Exit Function
        Next i

        portVal = val(portStr)
        If portVal <= 0 Or portVal > 65535 Then Exit Function
        outPort = portVal
    Else
        outHost = work
    End If

    If Len(outHost) = 0 Then Exit Function

    ParseURLInto = True
End Function

Private Function ParseURL(ByVal url As String) As Boolean
    ParseURL = ParseURLInto(m_Connections(m_DefaultHandle).Host, m_Connections(m_DefaultHandle).Port, m_Connections(m_DefaultHandle).Path, m_Connections(m_DefaultHandle).TLS, url)
End Function

Private Function ResolveHostStr(ByVal hostname As String, ByVal handle As Long) As Long
    Dim addr As Long
    Dim wsaErr As Long

#If VBA7 Then
    Dim hostent As LongPtr
    Dim he64 As HOSTENT64
    Dim addrList As LongPtr
    Dim pAddr As LongPtr
#Else
    Dim hostent As Long
    Dim he32 As HOSTENT32
    Dim addrList As Long
    Dim pAddr As Long
#End If

    addr = sock_inet_addr(hostname)
    If addr <> INADDR_NONE Then
        ResolveHostStr = addr
        Exit Function
    End If

    hostent = sock_gethostbyname(hostname)
    If hostent = 0 Then
        wsaErr = WSAGetLastError()
        Select Case wsaErr
            Case 11001
                LogError ERR_DNS_RESOLVE_FAILED, "gethostbyname() failed for '" & hostname & "' with WSAHOST_NOT_FOUND (11001)", "Unable to resolve server address." & vbCrLf & vbCrLf & "Hostname: " & hostname & vbCrLf & vbCrLf & "Please verify the server address is correct.", wsaErr, handle
            Case 11002
                LogError ERR_DNS_RESOLVE_FAILED, "gethostbyname() failed for '" & hostname & "' with WSATRY_AGAIN (11002)", "Temporary DNS error." & vbCrLf & vbCrLf & "Please try again in a moment.", wsaErr, handle
            Case 11003
                LogError ERR_DNS_RESOLVE_FAILED, "gethostbyname() failed for '" & hostname & "' with WSANO_RECOVERY (11003)", "DNS lookup failed." & vbCrLf & vbCrLf & "Please check your DNS settings.", wsaErr, handle
            Case 11004
                LogError ERR_DNS_RESOLVE_FAILED, "gethostbyname() failed for '" & hostname & "' with WSANO_DATA (11004)", "Server address found but no IP available.", wsaErr, handle
            Case Else
                LogError ERR_DNS_RESOLVE_FAILED, "gethostbyname() failed for '" & hostname & "' with WSA error " & wsaErr, "DNS resolution failed.", wsaErr, handle
        End Select
        Exit Function
    End If

#If VBA7 Then
    CopyMemoryFromPtr he64, hostent, LenB(he64)
    addrList = he64.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 8
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#Else
    CopyMemoryFromPtr he32, hostent, LenB(he32)
    addrList = he32.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 4
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#End If

    ResolveHostStr = addr
End Function

Private Function ResolveHost(ByVal hostname As String) As Long
    ResolveHost = ResolveHostStr(hostname, m_DefaultHandle)
End Function

Private Function U32Shl1(ByVal v As Long) As Long
    Dim lo As Long
    lo = v And &H3FFFFFFF
    U32Shl1 = lo * 2
    If v And &H40000000 Then U32Shl1 = U32Shl1 Or &H80000000
End Function

Private Function AddU32(ByVal a As Long, ByVal b As Long) As Long
    Dim aLo As Long, bLo As Long
    Dim aHi As Long, bHi As Long
    Dim sLo As Long, sHi As Long

    aLo = a And &HFFFF&
    bLo = b And &HFFFF&
    aHi = (a And &H7FFF0000) \ &H10000
    bHi = (b And &H7FFF0000) \ &H10000

    sLo = aLo + bLo
    sHi = aHi + bHi + (sLo \ &H10000)

    If a And &H80000000 Then sHi = sHi + &H8000&
    If b And &H80000000 Then sHi = sHi + &H8000&

    AddU32 = (sLo And &HFFFF&) Or ((sHi And &H7FFF&) * &H10000)
    If sHi And &H8000& Then AddU32 = AddU32 Or &H80000000
End Function

Private Function RotL32(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long
    Dim result As Long
    Dim msb As Boolean

    result = v
    For i = 1 To n
        msb = (result And &H80000000) <> 0
        result = U32Shl1(result)
        If msb Then result = result Or 1
    Next i
    RotL32 = result
End Function

Private Function SHA1(ByRef data() As Byte) As Byte()
    Dim h0 As Long, h1 As Long, h2 As Long, h3 As Long, h4 As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim f As Long, k As Long, temp As Long
    Dim w(0 To 79) As Long
    Dim msg() As Byte
    Dim origLen As Long
    Dim totalLen As Long
    Dim padLen As Long
    Dim bitLenLo As Long
    Dim bitLenHi As Long
    Dim i As Long, chunk As Long
    Dim result(0 To 19) As Byte
    Dim hArr(0 To 4) As Long
    Dim i0 As Long, i1 As Long, i2 As Long, i3 As Long
    Dim wordVal As Long
    Dim v As Long

    origLen = UBound(data) - LBound(data) + 1
    padLen = 64 - ((origLen + 9) Mod 64)
    If padLen = 64 Then padLen = 0
    totalLen = origLen + 1 + padLen + 8

    ReDim msg(0 To totalLen - 1)
    For i = 0 To origLen - 1
        msg(i) = data(LBound(data) + i)
    Next i
    msg(origLen) = &H80

    bitLenLo = origLen And &H1FFFFFFF
    bitLenLo = bitLenLo * 8
    bitLenHi = (origLen \ &H20000000) And &H7

    msg(totalLen - 1) = CByte(bitLenLo And &HFF)
    msg(totalLen - 2) = CByte((bitLenLo \ 256) And &HFF)
    msg(totalLen - 3) = CByte((bitLenLo \ 65536) And &HFF)
    msg(totalLen - 4) = CByte((bitLenLo \ 16777216) And &HFF)
    msg(totalLen - 5) = CByte(bitLenHi And &HFF)
    msg(totalLen - 6) = 0
    msg(totalLen - 7) = 0
    msg(totalLen - 8) = 0

    h0 = &H67452301
    h1 = &HEFCDAB89
    h2 = &H98BADCFE
    h3 = &H10325476
    h4 = &HC3D2E1F0

    For chunk = 0 To (totalLen \ 64) - 1
        For i = 0 To 15
            i0 = CLng(msg(chunk * 64 + i * 4 + 0)) And &HFF&
            i1 = CLng(msg(chunk * 64 + i * 4 + 1)) And &HFF&
            i2 = CLng(msg(chunk * 64 + i * 4 + 2)) And &HFF&
            i3 = CLng(msg(chunk * 64 + i * 4 + 3)) And &HFF&
            wordVal = (i1 * &H10000) Or (i2 * &H100&) Or i3
            wordVal = wordVal Or ((i0 And &H7F&) * &H1000000)
            If i0 And &H80& Then wordVal = wordVal Or &H80000000
            w(i) = wordVal
        Next i

        For i = 16 To 79
            w(i) = RotL32(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16), 1)
        Next i

        a = h0: b = h1: c = h2: d = h3: e = h4

        For i = 0 To 79
            If i < 20 Then
                f = (b And c) Or ((Not b) And d)
                k = &H5A827999
            ElseIf i < 40 Then
                f = b Xor c Xor d
                k = &H6ED9EBA1
            ElseIf i < 60 Then
                f = (b And c) Or (b And d) Or (c And d)
                k = &H8F1BBCDC
            Else
                f = b Xor c Xor d
                k = &HCA62C1D6
            End If
            temp = AddU32(AddU32(AddU32(AddU32(RotL32(a, 5), f), e), k), w(i))
            e = d: d = c
            c = RotL32(b, 30)
            b = a: a = temp
        Next i

        h0 = AddU32(h0, a)
        h1 = AddU32(h1, b)
        h2 = AddU32(h2, c)
        h3 = AddU32(h3, d)
        h4 = AddU32(h4, e)
    Next chunk

    hArr(0) = h0: hArr(1) = h1: hArr(2) = h2: hArr(3) = h3: hArr(4) = h4

    For i = 0 To 4
        v = hArr(i)
        result(i * 4 + 0) = CByte(((v And &H7F000000) \ &H1000000) Or IIf((v And &H80000000) <> 0, &H80&, 0))
        result(i * 4 + 1) = CByte((v And &HFF0000) \ &H10000)
        result(i * 4 + 2) = CByte((v And &HFF00&) \ &H100&)
        result(i * 4 + 3) = CByte(v And &HFF&)
    Next i

    SHA1 = result
End Function

Private Function ComputeWebSocketAccept(ByVal wsKey As String) As String
    Dim inputBytes() As Byte
    Dim hashBytes() As Byte
    inputBytes = StrConv(wsKey & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11", vbFromUnicode)
    hashBytes = SHA1(inputBytes)
    ComputeWebSocketAccept = Base64Encode(hashBytes)
End Function

Private Function GenerateWebSocketKey() As String
    Dim bytes(0 To 15) As Byte
    Dim i As Long
    For i = 0 To 15
        bytes(i) = Int(Rnd * 256)
    Next i
    GenerateWebSocketKey = Base64Encode(bytes)
End Function

Private Function DoTLSHandshakeFor(ByVal handle As Long) As Long
    Dim outBuffer As SecBuffer
    Dim outBufferDesc As SecBufferDesc
    Dim inBuffer(0 To 1) As SecBuffer
    Dim inBufferDesc As SecBufferDesc
    Dim contextAttr As Long
    Dim tsExpiry As SECURITY_INTEGER
    Dim result As Long
    Dim contextFlags As Long
    Dim recvBuffer() As Byte
    Dim recvLen As Long
    Dim outData() As Byte
    Dim firstCall As Boolean
    Dim recv As Long
    Dim loopCount As Long
    Dim i As Long
    Dim extraData As Long

    contextFlags = ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM

    ReDim recvBuffer(0 To 32767)
    recvLen = 0
    firstCall = True
    loopCount = 0

    With m_Connections(handle)
        Do
            loopCount = loopCount + 1

            If firstCall Then
                outBufferDesc.ulVersion = SECBUFFER_VERSION
                outBufferDesc.cBuffers = 1
                outBufferDesc.pBuffers = VarPtr(outBuffer)
                outBuffer.cbBuffer = 0
                outBuffer.BufferType = SECBUFFER_TOKEN
                outBuffer.pvBuffer = 0

                result = InitializeSecurityContext(.hCred, NULL_PTR, .Host, contextFlags, 0, 0, NULL_PTR, 0, .hContext, outBufferDesc, contextAttr, tsExpiry)
                firstCall = False
            Else
                inBufferDesc.ulVersion = SECBUFFER_VERSION
                inBufferDesc.cBuffers = 2
                inBufferDesc.pBuffers = VarPtr(inBuffer(0))
                inBuffer(0).cbBuffer = recvLen
                inBuffer(0).BufferType = SECBUFFER_TOKEN
                inBuffer(0).pvBuffer = VarPtr(recvBuffer(0))
                inBuffer(1).cbBuffer = 0
                inBuffer(1).BufferType = SECBUFFER_EMPTY
                inBuffer(1).pvBuffer = 0

                outBufferDesc.ulVersion = SECBUFFER_VERSION
                outBufferDesc.cBuffers = 1
                outBufferDesc.pBuffers = VarPtr(outBuffer)
                outBuffer.cbBuffer = 0
                outBuffer.BufferType = SECBUFFER_TOKEN
                outBuffer.pvBuffer = 0

                result = InitializeSecurityContextContinue(.hCred, .hContext, .Host, contextFlags, 0, 0, inBufferDesc, 0, .hContext, outBufferDesc, contextAttr, tsExpiry)

                extraData = 0
                For i = 0 To 1
                    If inBuffer(i).BufferType = SECBUFFER_EXTRA And inBuffer(i).cbBuffer > 0 Then
                        extraData = inBuffer(i).cbBuffer
                        Exit For
                    End If
                Next i

                If extraData > 0 Then
                    For i = 0 To extraData - 1
                        recvBuffer(i) = recvBuffer(recvLen - extraData + i)
                    Next i
                    recvLen = extraData
                ElseIf result <> SEC_E_INCOMPLETE_MESSAGE Then
                    recvLen = 0
                End If
            End If

            If result < 0 And result <> SEC_E_INCOMPLETE_MESSAGE Then
                DoTLSHandshakeFor = result
                Exit Function
            End If

            If outBuffer.cbBuffer > 0 And outBuffer.pvBuffer <> 0 Then
                ReDim outData(0 To outBuffer.cbBuffer - 1)
                CopyMemoryFromPtr outData(0), outBuffer.pvBuffer, outBuffer.cbBuffer
                sock_send .Socket, outData(0), outBuffer.cbBuffer, 0
                FreeContextBuffer outBuffer.pvBuffer
            End If

            If result = SEC_E_OK Then
                DoTLSHandshakeFor = 0
                Exit Function
            End If

            If result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE Then
                If Not WaitForDataOn(handle, 10000) Then
                    DoTLSHandshakeFor = -1
                    Exit Function
                End If

                recv = sock_recv(.Socket, recvBuffer(recvLen), 32768 - recvLen, 0)

                If recv <= 0 Then
                    DoTLSHandshakeFor = -1
                    Exit Function
                End If
                recvLen = recvLen + recv
            End If

            If loopCount > 30 Then
                DoTLSHandshakeFor = -1
                Exit Function
            End If

        Loop While result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE
    End With

    DoTLSHandshakeFor = result
End Function

Private Function DoTLSHandshake() As Long
    DoTLSHandshake = DoTLSHandshakeFor(m_DefaultHandle)
End Function

Private Function DoWebSocketHandshakeFor(ByVal handle As Long) As Boolean
    Dim handshake As String
    Dim sendBuf() As Byte
    Dim response As String
    Dim wsKey As String
    Dim sendResult As Long
    Dim i As Long
    Dim hostHeader As String
    Dim originHeader As String
    Dim expectedAccept As String
    Dim actualAccept As String
    Dim acceptPos As Long
    Dim acceptLineEnd As Long
    Dim wsaErr As Long

    wsKey = GenerateWebSocketKey()

    With m_Connections(handle)
        If (.TLS And .Port <> 443) Or (Not .TLS And .Port <> 80) Then
            hostHeader = .Host & ":" & .Port
        Else
            hostHeader = .Host
        End If

        If .TLS Then
            If .Port <> 443 Then
                originHeader = "https://" & .Host & ":" & .Port
            Else
                originHeader = "https://" & .Host
            End If
        Else
            If .Port <> 80 Then
                originHeader = "http://" & .Host & ":" & .Port
            Else
                originHeader = "http://" & .Host
            End If
        End If

        handshake = "GET " & .Path & " HTTP/1.1" & vbCrLf
        handshake = handshake & "Host: " & hostHeader & vbCrLf
        handshake = handshake & "Upgrade: websocket" & vbCrLf
        handshake = handshake & "Connection: Upgrade" & vbCrLf
        handshake = handshake & "Sec-WebSocket-Key: " & wsKey & vbCrLf
        handshake = handshake & "Sec-WebSocket-Version: 13" & vbCrLf
        If .SubProtocol <> "" Then
            handshake = handshake & "Sec-WebSocket-Protocol: " & .SubProtocol & vbCrLf
        End If
        'handshake = handshake & "Sec-WebSocket-Extensions: permessage-deflate" & vbCrLf (TODO)
        handshake = handshake & "Origin: " & originHeader & vbCrLf
        handshake = handshake & "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" & vbCrLf

        For i = 0 To .CustomHeaderCount - 1
            handshake = handshake & .CustomHeaders(i) & vbCrLf
        Next i

        handshake = handshake & vbCrLf

        sendBuf = StrConv(handshake, vbFromUnicode)

        If .TLS Then
            If Not TLSSendFor(handle, sendBuf) Then
                LogError ERR_WEBSOCKET_HANDSHAKE_FAILED, "Failed to send WebSocket handshake via TLS", "WebSocket handshake failed." & vbCrLf & "Unable to send upgrade request.", 0, handle
                Exit Function
            End If
        Else
            sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
            If sendResult <= 0 Then
                wsaErr = WSAGetLastError()
                LogError ERR_WEBSOCKET_HANDSHAKE_FAILED, "send() failed during WebSocket handshake with WSA error " & wsaErr, "WebSocket handshake failed." & vbCrLf & "Unable to send upgrade request.", wsaErr, handle
                Exit Function
            End If
        End If

        If .TLS Then
            response = TLSReceiveHandshakeFor(handle)
        Else
            Dim recvBuf() As Byte
            Dim received As Long
            If Not WaitForDataOn(handle, 5000) Then
                LogError ERR_WEBSOCKET_HANDSHAKE_TIMEOUT, "No response received for WebSocket handshake within 5 seconds", "WebSocket handshake timeout." & vbCrLf & vbCrLf & "Server did not respond to the upgrade request.", 0, handle
                Exit Function
            End If
            ReDim recvBuf(0 To 4095)
            received = sock_recv(.Socket, recvBuf(0), 4096, 0)
            If received > 0 Then
                response = Left(StrConv(recvBuf, vbUnicode), received)
            Else
                wsaErr = WSAGetLastError()
                LogError ERR_WEBSOCKET_HANDSHAKE_FAILED, "recv() failed during WebSocket handshake with WSA error " & wsaErr, "WebSocket handshake failed." & vbCrLf & "Unable to receive server response.", wsaErr, handle
                Exit Function
            End If
        End If

        If InStr(response, "101") = 0 Then
            Dim statusLine As String
            Dim lineEnd As Long
            lineEnd = InStr(response, vbCrLf)
            If lineEnd > 0 Then
                statusLine = Left(response, lineEnd - 1)
            Else
                statusLine = Left(response, 50)
            End If
            LogError ERR_HANDSHAKE_REJECTED, "WebSocket upgrade rejected. Server response: " & statusLine, "WebSocket connection rejected by server." & vbCrLf & vbCrLf & "The server did not accept the WebSocket upgrade request.", 0, handle
            Exit Function
        End If

        expectedAccept = ComputeWebSocketAccept(wsKey)
        acceptPos = InStr(LCase(response), "sec-websocket-accept:")
        If acceptPos > 0 Then
            acceptLineEnd = InStr(acceptPos, response, vbCrLf)
            If acceptLineEnd > 0 Then
                actualAccept = Trim(Mid(response, acceptPos + 21, acceptLineEnd - acceptPos - 21))
            End If
        End If

        If actualAccept <> expectedAccept Then
            LogError ERR_HANDSHAKE_REJECTED, "Sec-WebSocket-Accept mismatch. Expected: " & expectedAccept & " Got: " & actualAccept, "WebSocket handshake failed." & vbCrLf & "Server returned invalid accept key.", 0, handle
            Exit Function
        End If
    End With

    DoWebSocketHandshakeFor = True
End Function

Private Function DoWebSocketHandshake() As Boolean
    DoWebSocketHandshake = DoWebSocketHandshakeFor(m_DefaultHandle)
End Function

Private Function TLSSendFor(ByVal handle As Long, ByRef data() As Byte) As Boolean
    Dim buffers(0 To 3) As SecBuffer
    Dim bufferDesc As SecBufferDesc
    Dim sendBuf() As Byte
    Dim dataLen As Long
    Dim totalLen As Long
    Dim totalSent As Long
    Dim toSend As Long
    Dim sendResult As Long
    Dim result As Long

    With m_Connections(handle)
        dataLen = SafeArrayLen(data)
        totalLen = .Sizes.cbHeader + dataLen + .Sizes.cbTrailer

        ReDim sendBuf(0 To totalLen - 1)
        CopyMemory sendBuf(.Sizes.cbHeader), data(LBound(data)), dataLen

        buffers(0).cbBuffer = .Sizes.cbHeader
        buffers(0).BufferType = SECBUFFER_STREAM_HEADER
        buffers(0).pvBuffer = VarPtr(sendBuf(0))
        buffers(1).cbBuffer = dataLen
        buffers(1).BufferType = SECBUFFER_DATA
        buffers(1).pvBuffer = VarPtr(sendBuf(.Sizes.cbHeader))
        buffers(2).cbBuffer = .Sizes.cbTrailer
        buffers(2).BufferType = SECBUFFER_STREAM_TRAILER
        buffers(2).pvBuffer = VarPtr(sendBuf(.Sizes.cbHeader + dataLen))
        buffers(3).cbBuffer = 0
        buffers(3).BufferType = SECBUFFER_EMPTY
        buffers(3).pvBuffer = 0

        bufferDesc.ulVersion = SECBUFFER_VERSION
        bufferDesc.cBuffers = 4
        bufferDesc.pBuffers = VarPtr(buffers(0))

        result = EncryptMessage(.hContext, 0, bufferDesc, 0)
        If result <> 0 Then
            .LastError = ERR_TLS_ENCRYPT_FAILED
            .LastErrorCode = result
            .TechnicalDetails = "EncryptMessage failed with SSPI error 0x" & hex(result)
            m_LastError = ERR_TLS_ENCRYPT_FAILED
            m_LastErrorCode = result
            m_TechnicalDetails = .TechnicalDetails
            WasabiLog handle, .TechnicalDetails
            Exit Function
        End If

        toSend = buffers(0).cbBuffer + buffers(1).cbBuffer + buffers(2).cbBuffer
        totalSent = 0
        Do While totalSent < toSend
            sendResult = sock_send(.Socket, sendBuf(totalSent), toSend - totalSent, 0)
            If sendResult <= 0 Then
                Dim wsaErr As Long
                wsaErr = WSAGetLastError()
                .LastError = ERR_SEND_FAILED
                .LastErrorCode = wsaErr
                .TechnicalDetails = "send() failed after TLS encryption with WSA error " & wsaErr
                m_LastError = ERR_SEND_FAILED
                m_LastErrorCode = wsaErr
                m_TechnicalDetails = .TechnicalDetails
                WasabiLog handle, .TechnicalDetails
                Exit Function
            End If
            totalSent = totalSent + sendResult
        Loop

        TLSSendFor = True
    End With
End Function

Private Function TLSSend(ByRef data() As Byte) As Boolean
    TLSSend = TLSSendFor(m_DefaultHandle, data)
End Function

Private Function TLSReceiveHandshakeFor(ByVal handle As Long) As String
    Dim tempBuf() As Byte
    Dim received As Long
    Dim i As Long
    Dim headerEnd As Long
    Dim httpResponse As String
    Dim headerBytes() As Byte
    Dim remainingLen As Long
    Dim copyLen As Long

    With m_Connections(handle)
        Do
            If Not WaitForDataOn(handle, 5000) Then Exit Do

            ReDim tempBuf(0 To 8191)
            received = sock_recv(.Socket, tempBuf(0), 8192, 0)

            If received > 0 Then
                copyLen = received
                If .recvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .recvLen
                If copyLen > 0 Then
                    CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                    .recvLen = .recvLen + copyLen
                End If

                TLSDecryptFor handle

                If .DecryptLen > 0 Then
                    headerEnd = -1
                    For i = 0 To .DecryptLen - 4
                        If .DecryptBuffer(i) = 13 And .DecryptBuffer(i + 1) = 10 And .DecryptBuffer(i + 2) = 13 And .DecryptBuffer(i + 3) = 10 Then
                            headerEnd = i + 4
                            Exit For
                        End If
                    Next i

                    If headerEnd > 0 Then
                        ReDim headerBytes(0 To headerEnd - 1)
                        CopyMemory headerBytes(0), .DecryptBuffer(0), headerEnd
                        httpResponse = StrConv(headerBytes, vbUnicode)

                        remainingLen = .DecryptLen - headerEnd
                        If remainingLen > 0 Then
                            CopyMemory .DecryptBuffer(0), .DecryptBuffer(headerEnd), remainingLen
                        End If
                        .DecryptLen = remainingLen

                        TLSReceiveHandshakeFor = httpResponse
                        Exit Function
                    End If
                End If
            Else
                Exit Do
            End If
        Loop

        If .DecryptLen > 0 Then
            ReDim headerBytes(0 To .DecryptLen - 1)
            CopyMemory headerBytes(0), .DecryptBuffer(0), .DecryptLen
            TLSReceiveHandshakeFor = StrConv(headerBytes, vbUnicode)
            .DecryptLen = 0
        End If
    End With
End Function

Private Function TLSReceiveHandshake() As String
    TLSReceiveHandshake = TLSReceiveHandshakeFor(m_DefaultHandle)
End Function

Private Sub TLSDecryptFor(ByVal handle As Long)
    Dim buffers(0 To 3) As SecBuffer
    Dim bufferDesc As SecBufferDesc
    Dim result As Long
    Dim qop As Long
    Dim i As Long
    Dim dataLen As Long
    Dim extraLen As Long

    With m_Connections(handle)
        Do While .recvLen > 0
            buffers(0).cbBuffer = .recvLen
            buffers(0).BufferType = SECBUFFER_DATA
            buffers(0).pvBuffer = VarPtr(.recvBuffer(0))
            buffers(1).cbBuffer = 0
            buffers(1).BufferType = SECBUFFER_EMPTY
            buffers(1).pvBuffer = 0
            buffers(2).cbBuffer = 0
            buffers(2).BufferType = SECBUFFER_EMPTY
            buffers(2).pvBuffer = 0
            buffers(3).cbBuffer = 0
            buffers(3).BufferType = SECBUFFER_EMPTY
            buffers(3).pvBuffer = 0

            bufferDesc.ulVersion = SECBUFFER_VERSION
            bufferDesc.cBuffers = 4
            bufferDesc.pBuffers = VarPtr(buffers(0))

            result = DecryptMessage(.hContext, bufferDesc, 0, qop)

            If result = SEC_E_INCOMPLETE_MESSAGE Then Exit Sub

            If result = SEC_I_RENEGOTIATE Then
                WasabiLog handle, "TLS renegotiation requested by server, closing connection (handle=" & handle & ")"
                LogError ERR_TLS_DECRYPT_FAILED, "SEC_I_RENEGOTIATE received from server", "Secure connection interrupted: server requested renegotiation.", SEC_I_RENEGOTIATE, handle
                .Connected = False
                If .AutoReconnect Then TryReconnectFor handle
                Exit Sub
            End If

            If result <> SEC_E_OK Then
                .LastError = ERR_TLS_DECRYPT_FAILED
                .LastErrorCode = result
                .TechnicalDetails = "DecryptMessage failed with SSPI error 0x" & hex(result)
                m_LastError = ERR_TLS_DECRYPT_FAILED
                m_LastErrorCode = result
                m_TechnicalDetails = .TechnicalDetails
                WasabiLog handle, .TechnicalDetails
                Exit Sub
            End If

            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_DATA Then
                    dataLen = buffers(i).cbBuffer
                    If dataLen > 0 And .DecryptLen + dataLen <= BUFFER_SIZE Then
                        CopyMemoryFromPtr .DecryptBuffer(.DecryptLen), buffers(i).pvBuffer, dataLen
                        .DecryptLen = .DecryptLen + dataLen
                    End If
                End If
            Next i

            extraLen = 0
            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_EXTRA Then
                    extraLen = buffers(i).cbBuffer
                    Exit For
                End If
            Next i

            If extraLen > 0 Then
                CopyMemory .recvBuffer(0), .recvBuffer(.recvLen - extraLen), extraLen
                .recvLen = extraLen
            Else
                .recvLen = 0
            End If
        Loop
    End With
End Sub

Private Sub TLSDecrypt()
    TLSDecryptFor m_DefaultHandle
End Sub

Private Sub ProcessFramesFor(ByVal handle As Long)
    Dim fin As Boolean
    Dim opcode As Byte
    Dim masked As Boolean
    Dim payloadLen As Long
    Dim headerLen As Long
    Dim frameLen As Long
    Dim i As Long
    Dim payload() As Byte
    Dim maskKey(0 To 3) As Byte
    Dim completePayload() As Byte
    Dim copyFrag As Long
    Dim remaining As Long

    With m_Connections(handle)
        Do While .DecryptLen >= 2
            fin = (.DecryptBuffer(0) And &H80) <> 0
            opcode = .DecryptBuffer(0) And &HF
            masked = (.DecryptBuffer(1) And &H80) <> 0
            payloadLen = .DecryptBuffer(1) And &H7F
            headerLen = 2

            If payloadLen = 126 Then
                If .DecryptLen < 4 Then Exit Sub
                payloadLen = CLng(.DecryptBuffer(2)) * 256 + .DecryptBuffer(3)
                headerLen = 4
            ElseIf payloadLen = 127 Then
                If .DecryptLen < 10 Then Exit Sub
                payloadLen = CLng(.DecryptBuffer(6)) * 16777216 + CLng(.DecryptBuffer(7)) * 65536 + CLng(.DecryptBuffer(8)) * 256 + .DecryptBuffer(9)
                headerLen = 10
            End If

            If payloadLen > BUFFER_SIZE Then
                WasabiLog handle, "Frame payload too large (" & payloadLen & " bytes), closing connection (handle=" & handle & ")"
                LogError ERR_CONNECTION_LOST, "Received oversized frame payload: " & payloadLen & " bytes", "Connection closed: server sent oversized frame.", 0, handle
                .Connected = False
                If .AutoReconnect Then TryReconnectFor handle
                Exit Sub
            End If

            If masked Then
                If .DecryptLen < headerLen + 4 Then Exit Sub
                For i = 0 To 3
                    maskKey(i) = .DecryptBuffer(headerLen + i)
                Next i
                headerLen = headerLen + 4
            End If

            frameLen = headerLen + payloadLen
            If .DecryptLen < frameLen Then Exit Sub

            If payloadLen > 0 Then
                ReDim payload(0 To payloadLen - 1)
                CopyMemory payload(0), .DecryptBuffer(headerLen), payloadLen
                If masked Then
                    For i = 0 To payloadLen - 1
                        payload(i) = payload(i) Xor maskKey(i Mod 4)
                    Next i
                End If
            End If

            Select Case opcode
                Case WS_OPCODE_CONTINUATION
                    If .Fragmenting And payloadLen > 0 Then
                        copyFrag = payloadLen
                        If .FragmentLen + copyFrag > FRAGMENT_BUFFER_SIZE Then copyFrag = FRAGMENT_BUFFER_SIZE - .FragmentLen
                        If copyFrag > 0 Then
                            CopyMemory .FragmentBuffer(.FragmentLen), payload(0), copyFrag
                            .FragmentLen = .FragmentLen + copyFrag
                        End If
                    End If
                    If fin And .Fragmenting Then
                        If .FragmentLen > 0 Then
                            ReDim completePayload(0 To .FragmentLen - 1)
                            CopyMemory completePayload(0), .FragmentBuffer(0), .FragmentLen
                            If .FragmentOpcode = WS_OPCODE_TEXT Then
                                EnqueueFor handle, Utf8ToString(completePayload)
                            Else
                                EnqueueBinaryFor handle, completePayload
                            End If
                            .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                        End If
                        .FragmentLen = 0
                        .Fragmenting = False
                        .FragmentOpcode = 0
                    End If

                Case WS_OPCODE_TEXT
                    If fin Then
                        If payloadLen > 0 Then
                            EnqueueFor handle, Utf8ToString(payload)
                            .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                        End If
                    Else
                        .Fragmenting = True
                        .FragmentOpcode = opcode
                        .FragmentLen = 0
                        If payloadLen > 0 Then
                            copyFrag = payloadLen
                            If copyFrag > FRAGMENT_BUFFER_SIZE Then copyFrag = FRAGMENT_BUFFER_SIZE
                            CopyMemory .FragmentBuffer(0), payload(0), copyFrag
                            .FragmentLen = copyFrag
                        End If
                    End If

                Case WS_OPCODE_BINARY
                    If fin Then
                        If payloadLen > 0 Then
                            EnqueueBinaryFor handle, payload
                            .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                        End If
                    Else
                        .Fragmenting = True
                        .FragmentOpcode = opcode
                        .FragmentLen = 0
                        If payloadLen > 0 Then
                            copyFrag = payloadLen
                            If copyFrag > FRAGMENT_BUFFER_SIZE Then copyFrag = FRAGMENT_BUFFER_SIZE
                            CopyMemory .FragmentBuffer(0), payload(0), copyFrag
                            .FragmentLen = copyFrag
                        End If
                    End If

                Case WS_OPCODE_CLOSE
                    WasabiLog handle, "Received CLOSE frame from server (handle=" & handle & ")"
                    LogError ERR_CONNECTION_LOST, "Received WebSocket CLOSE opcode", "Server closed the WebSocket connection.", 0, handle
                    WebSocketSendClose 1000, "", handle
                    .Connected = False
                    Exit Sub

                Case WS_OPCODE_PING
                    SendPongFor handle, payload, payloadLen

            End Select

            remaining = .DecryptLen - frameLen
            If remaining > 0 Then
                CopyMemory .DecryptBuffer(0), .DecryptBuffer(frameLen), remaining
            End If
            .DecryptLen = remaining
        Loop
    End With
End Sub

Private Sub ProcessFrames()
    ProcessFramesFor m_DefaultHandle
End Sub

Private Sub SendPongFor(ByVal handle As Long, ByRef pingPayload() As Byte, ByVal pingLen As Long)
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim i As Long

    For i = 0 To 3
        mask(i) = Int(Rnd * 256)
    Next i

    If pingLen = 0 Then
        ReDim frame(0 To 5)
        frame(0) = &H8A
        frame(1) = &H80
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
    Else
        ReDim frame(0 To 5 + pingLen)
        frame(0) = &H8A
        frame(1) = &H80 Or CByte(pingLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
        For i = 0 To pingLen - 1
            frame(6 + i) = pingPayload(i) Xor mask(i Mod 4)
        Next i
    End If

    If m_Connections(handle).TLS Then
        TLSSendFor handle, frame
    Else
        sock_send m_Connections(handle).Socket, frame(0), UBound(frame) + 1, 0
    End If
End Sub

Private Sub SendPong(ByRef pingPayload() As Byte, ByVal pingLen As Long)
    SendPongFor m_DefaultHandle, pingPayload, pingLen
End Sub

Private Sub EnqueueFor(ByVal handle As Long, ByVal message As String)
    With m_Connections(handle)
        If .MsgCount >= MSG_QUEUE_SIZE Then
            WasabiLog handle, "Warning: message queue full, dropping message (handle=" & handle & ")"
            Exit Sub
        End If
        .MsgQueue(.MsgTail) = message
        .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
        .MsgCount = .MsgCount + 1
    End With
End Sub

Private Sub Enqueue(ByVal message As String)
    EnqueueFor m_DefaultHandle, message
End Sub

Private Sub EnqueueBinaryFor(ByVal handle As Long, ByRef data() As Byte)
    With m_Connections(handle)
        If .BinaryCount >= MSG_QUEUE_SIZE Then
            WasabiLog handle, "Warning: binary queue full, dropping message (handle=" & handle & ")"
            Exit Sub
        End If
        .BinaryQueue(.BinaryTail).data = data
        .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
        .BinaryCount = .BinaryCount + 1
    End With
End Sub

Private Sub EnqueueBinary(ByRef data() As Byte)
    EnqueueBinaryFor m_DefaultHandle, data
End Sub
