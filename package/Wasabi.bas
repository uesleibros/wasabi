Attribute VB_Name = "Wasabi"
' ============================================================================
' Wasabi v2.1.1-vNext
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
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function zlib_deflateInit2 Lib "zlib1.dll" Alias "deflateInit2_" (ByRef strm As ZStream, ByVal level As Long, ByVal method As Long, ByVal windowBits As Long, ByVal memLevel As Long, ByVal strategy As Long, ByVal version As String, ByVal stream_size As Long) As Long
    Private Declare PtrSafe Function zlib_deflate Lib "zlib1.dll" Alias "deflate" (ByRef strm As ZStream, ByVal flush As Long) As Long
    Private Declare PtrSafe Function zlib_deflateEnd Lib "zlib1.dll" Alias "deflateEnd" (ByRef strm As ZStream) As Long
    Private Declare PtrSafe Function zlib_inflateInit2 Lib "zlib1.dll" Alias "inflateInit2_" (ByRef strm As ZStream, ByVal windowBits As Long, ByVal version As String, ByVal stream_size As Long) As Long
    Private Declare PtrSafe Function zlib_inflate Lib "zlib1.dll" Alias "inflate" (ByRef strm As ZStream, ByVal flush As Long) As Long
    Private Declare PtrSafe Function zlib_inflateEnd Lib "zlib1.dll" Alias "inflateEnd" (ByRef strm As ZStream) As Long
    Private Declare PtrSafe Function CertGetCertificateChain Lib "crypt32.dll" (ByVal hChainEngine As LongPtr, ByVal pCertContext As LongPtr, ByVal pTime As LongPtr, ByVal hAdditionalStore As LongPtr, ByRef pChainPara As CERT_CHAIN_PARA, ByVal dwFlags As Long, ByVal pvReserved As LongPtr, ByRef ppChainContext As LongPtr) As Long
    Private Declare PtrSafe Function CertVerifyCertificateChainPolicy Lib "crypt32.dll" (ByVal pszPolicyOID As LongPtr, ByVal pChainContext As LongPtr, ByRef pPolicyPara As CERT_CHAIN_POLICY_PARA, ByRef pPolicyStatus As CERT_CHAIN_POLICY_STATUS) As Long
    Private Declare PtrSafe Sub CertFreeCertificateChain Lib "crypt32.dll" (ByVal pChainContext As LongPtr)
    Private Declare PtrSafe Function CertOpenStore Lib "crypt32.dll" (ByVal lpszStoreProvider As LongPtr, ByVal dwEncodingType As Long, ByVal hCryptProv As LongPtr, ByVal dwFlags As Long, ByVal pvPara As LongPtr) As LongPtr
    Private Declare PtrSafe Function CertFindCertificateInStore Lib "crypt32.dll" (ByVal hCertStore As LongPtr, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, ByRef pvFindPara As Any, ByVal pPrevCertContext As LongPtr) As LongPtr
    Private Declare PtrSafe Function CertCloseStore Lib "crypt32.dll" (ByVal hCertStore As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function PFXImportCertStore Lib "crypt32.dll" (ByRef pPFX As CRYPT_DATA_BLOB, ByVal szPassword As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function CertFreeCertificateContext Lib "crypt32.dll" (ByVal pCertContext As LongPtr) As Long
    Private Declare PtrSafe Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As LongPtr, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonID As LongPtr, ByRef pAuthData As Any, ByVal pGetKeyFn As LongPtr, ByVal pvGetKeyArgument As LongPtr, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare PtrSafe Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByVal phContext As LongPtr, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByVal pInput As LongPtr, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare PtrSafe Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare PtrSafe Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As LongPtr) As Long
    Private Declare PtrSafe Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" (ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long
    Private Declare PtrSafe Function EncryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare PtrSafe Function DecryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function sock_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByRef optlen As Long) As Long
    Private Declare PtrSafe Function sock_getaddrinfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, ByVal pHints As LongPtr, ByRef ppResult As LongPtr) As Long
    Private Declare PtrSafe Sub sock_freeaddrinfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal pAddrInfo As LongPtr)
    Private Declare PtrSafe Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As LongPtr
    Private Declare PtrSafe Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
    Private Declare PtrSafe Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByVal name As LongPtr, ByVal namelen As Long) As Long
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
    Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwLen As Long, ByRef pbBuffer As Byte) As Long
    Private Const NULL_PTR As LongPtr = 0
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function zlib_deflateInit2 Lib "zlib1.dll" Alias "deflateInit2_" (ByRef strm As ZStream, ByVal level As Long, ByVal method As Long, ByVal windowBits As Long, ByVal memLevel As Long, ByVal strategy As Long, ByVal version As String, ByVal stream_size As Long) As Long
    Private Declare Function zlib_deflate Lib "zlib1.dll" Alias "deflate" (ByRef strm As ZStream, ByVal flush As Long) As Long
    Private Declare Function zlib_deflateEnd Lib "zlib1.dll" Alias "deflateEnd" (ByRef strm As ZStream) As Long
    Private Declare Function zlib_inflateInit2 Lib "zlib1.dll" Alias "inflateInit2_" (ByRef strm As ZStream, ByVal windowBits As Long, ByVal version As String, ByVal stream_size As Long) As Long
    Private Declare Function zlib_inflate Lib "zlib1.dll" Alias "inflate" (ByRef strm As ZStream, ByVal flush As Long) As Long
    Private Declare Function zlib_inflateEnd Lib "zlib1.dll" Alias "inflateEnd" (ByRef strm As ZStream) As Long
    Private Declare Function CertGetCertificateChain Lib "crypt32.dll" (ByVal hChainEngine As Long, ByVal pCertContext As Long, ByVal pTime As Long, ByVal hAdditionalStore As Long, ByRef pChainPara As CERT_CHAIN_PARA, ByVal dwFlags As Long, ByVal pvReserved As Long, ByRef ppChainContext As Long) As Long
    Private Declare Function CertVerifyCertificateChainPolicy Lib "crypt32.dll" (ByVal pszPolicyOID As Long, ByVal pChainContext As Long, ByRef pPolicyPara As CERT_CHAIN_POLICY_PARA, ByRef pPolicyStatus As CERT_CHAIN_POLICY_STATUS) As Long
    Private Declare Sub CertFreeCertificateChain Lib "crypt32.dll" (ByVal pChainContext As Long)
    Private Declare Function CertOpenStore Lib "crypt32.dll" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
    Private Declare Function CertFindCertificateInStore Lib "crypt32.dll" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, ByRef pvFindPara As Any, ByVal pPrevCertContext As Long) As Long
    Private Declare Function CertCloseStore Lib "crypt32.dll" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
    Private Declare Function PFXImportCertStore Lib "crypt32.dll" (ByRef pPFX As CRYPT_DATA_BLOB, ByVal szPassword As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CertFreeCertificateContext Lib "crypt32.dll" (ByVal pCertContext As Long) As Long
    Private Declare Function AcquireCredentialsHandle Lib "secur32.dll" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonID As Long, ByRef pAuthData As Any, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, ByRef phCredential As SecHandle, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function FreeCredentialsHandle Lib "secur32.dll" (ByRef phCredential As SecHandle) As Long
    Private Declare Function InitializeSecurityContext Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByVal phContext As Long, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByVal pInput As Long, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function InitializeSecurityContextContinue Lib "secur32.dll" Alias "InitializeSecurityContextA" (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, ByVal pszTargetName As String, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, ByRef pfContextAttr As Long, ByRef ptsExpiry As SECURITY_INTEGER) As Long
    Private Declare Function DeleteSecurityContext Lib "secur32.dll" (ByRef phContext As SecHandle) As Long
    Private Declare Function FreeContextBuffer Lib "secur32.dll" (ByVal pvContextBuffer As Long) As Long
    Private Declare Function QueryContextAttributes Lib "secur32.dll" Alias "QueryContextAttributesA" (ByRef phContext As SecHandle, ByVal ulAttribute As Long, ByRef pBuffer As Any) As Long
    Private Declare Function EncryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByVal fQOP As Long, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
    Private Declare Function DecryptMessage Lib "secur32.dll" (ByRef phContext As SecHandle, ByRef pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByRef pfQOP As Long) As Long
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
    Private Declare Function sock_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optVal As Long, ByRef optlen As Long) As Long
    Private Declare Function sock_getaddrinfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, ByVal pHints As Long, ByRef ppResult As Long) As Long
    Private Declare Sub sock_freeaddrinfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal pAddrInfo As Long)
    Private Declare Function sock_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long
    Private Declare Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
    Private Declare Function sock_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByVal name As Long, ByVal namelen As Long) As Long
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
    Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByRef pbBuffer As Byte) As Long
    Private Const NULL_PTR As Long = 0
#End If

Private Const TCP_MAXSEG As Long = 4

#If VBA7 Then
    Private Const INVALID_SOCKET As LongPtr = -1
#Else
    Private Const INVALID_SOCKET As Long = -1
#End If

Private Type ZStream
#If VBA7 Then
    next_in   As LongPtr
    avail_in  As Long
    total_in  As Long
    next_out  As LongPtr
    avail_out As Long
    total_out As Long
    msg       As LongPtr
    state     As LongPtr
    zalloc    As LongPtr
    zfree     As LongPtr
    opaque    As LongPtr
#Else
    next_in   As Long
    avail_in  As Long
    total_in  As Long
    next_out  As Long
    avail_out As Long
    total_out As Long
    msg       As Long
    state     As Long
    zalloc    As Long
    zfree     As Long
    opaque    As Long
#End If
    data_type As Long
    adler     As Long
    reserved  As Long
End Type

Private Type CRYPT_DATA_BLOB
#If VBA7 Then
    cbData As Long
    pbData As LongPtr
#Else
    cbData As Long
    pbData As Long
#End If
End Type

Private Type CERT_ENHKEY_USAGE
    cUsageIdentifier As Long
#If VBA7 Then
    rgpszUsageIdentifier As LongPtr
#Else
    rgpszUsageIdentifier As Long
#End If
End Type

Private Type CERT_USAGE_MATCH
    dwType As Long
    Usage As Long
End Type

Private Type CERT_CHAIN_PARA
    cbSize As Long
    RequestedUsage_dwType As Long
    RequestedUsage_cUsage As Long
#If VBA7 Then
    RequestedUsage_rgpOID As LongPtr
#Else
    RequestedUsage_rgpOID As Long
#End If
End Type

Private Type SSL_EXTRA_CERT_CHAIN_POLICY_PARA
    cbSize As Long
    dwAuthType As Long
    fdwChecks As Long
#If VBA7 Then
    pwszServerName As LongPtr
#Else
    pwszServerName As Long
#End If
End Type

Private Type CERT_CHAIN_POLICY_PARA
    cbSize As Long
    dwFlags As Long
#If VBA7 Then
    pvExtraPolicyPara As LongPtr
#Else
    pvExtraPolicyPara As Long
#End If
End Type

Private Type CERT_CHAIN_POLICY_STATUS
    cbSize As Long
    dwError As Long
    lChainIndex As Long
    lElementIndex As Long
End Type

Private Type BatchBuffer
    Frames() As Byte
    FrameCount As Long
    totalLen As Long
    MaxFrames As Long
End Type

Private Type SOCKADDR_IN6
    sin6_family   As Integer
    sin6_port     As Integer
    sin6_flowinfo As Long
    sin6_addr(0 To 15) As Byte
    sin6_scope_id As Long
End Type

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

Private Type MTUInfo
    CurrentMTU As Long
    MaxSegmentSize As Long
    OptimalFrameSize As Long
    LastProbeTime As Long
    ProbeEnabled As Boolean
    UseTLSFragmentation As Boolean
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
    hClientCertStore As LongPtr
    pClientCertCtx As LongPtr
    hNtlmCred As SecHandle
#Else
    Socket As Long
    hClientCertStore As Long
    pClientCertCtx As Long
    hNtlmCred As SecHandle
#End If
    Connected As Boolean
    TLS As Boolean
    Host As String
    port As Long
    path As String
    OriginalUrl As String
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
    proxyType As Long
    ProxyEnabled As Boolean
    InactivityTimeoutMs As Long
    LastActivityAt As Long
    SubProtocol As String
    CustomBufferSize As Long
    CustomFragmentSize As Long
    mtu As MTUInfo
    AutoMTU As Boolean
    ZeroCopyEnabled As Boolean
    closeCode As Integer
    closeReason As String
    CloseInitiatedByUs As Boolean
    PreferIPv6 As Boolean
    ValidateServerCert As Boolean
    EnableRevocationCheck As Boolean
    ClientCertThumb As String
    ClientCertPfxPath As String
    ClientCertPfxPass As String
    UseHttp2 As Boolean
    ProxyUseNtlm As Boolean
    LastRttMs As Long
    LastPingTimestamp As Long
    MqttParserStage As Long
    MqttBuffer() As Byte
    MqttBufLen As Long
    MqttExpectedRemaining As Long
    MqttCurrentPacketType As Byte
    MqttCurrentFlags As Byte
    DeflateEnabled          As Boolean
    DeflateContextTakeover  As Boolean
    InflateContextTakeover  As Boolean
    DeflateWindowBits       As Long
    InflateWindowBits       As Long
    DeflateStream           As ZStream
    InflateStream           As ZStream
    DeflateReady            As Boolean
    InflateReady            As Boolean
    DeflateActive           As Boolean
    FragmentIsCompressed As Boolean
    ClientMaxWindowBits As Long
    ServerMaxWindowBits As Long
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
    ERR_CERT_LOAD_FAILED = 25
    ERR_CERT_VALIDATE_FAILED = 26
    ERR_FRAGMENT_OVERFLOW = 27
    ERR_TLS_RENEGOTIATE = 28
End Enum

Private Const BUFFER_SIZE As Long = 262144
Private Const FRAGMENT_BUFFER_SIZE As Long = 262144
Private Const MSG_QUEUE_SIZE As Long = 512
Private Const MAX_CONNECTIONS As Long = 64
Private Const INVALID_CONN_HANDLE As Long = -1
Private Const DEFAULT_RECEIVE_TIMEOUT_MS As Long = 5000
Private Const DEFAULT_PING_INTERVAL_MS As Long = 0
Private Const DEFAULT_RECONNECT_BASE_DELAY_MS As Long = 1000
Private Const DEFAULT_RECONNECT_MAX_ATTEMPTS As Long = 5
Private Const MAX_RECONNECT_DELAY_MS As Long = 30000
Private Const DEFAULT_MTU As Long = 1500
Private Const HAPPY_EYEBALLS_DELAY_MS As Long = 250
Private Const PMTU_DISCOVERY_INTERVAL_MS As Long = 60000
Private Const SOL_SOCKET As Long = 65535
Private Const SO_KEEPALIVE As Long = 8
Private Const SO_RCVBUF As Long = &H1002
Private Const SO_SNDBUF As Long = &H1001
Private Const IPPROTO_TCP_SOL As Long = 6
Private Const TCP_NODELAY As Long = 1
Private Const CP_UTF8 As Long = 65001
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const FIONBIO As Long = &H8004667E
Private Const FIONREAD As Long = &H4004667F
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const PROXY_TYPE_HTTP As Long = 0
Private Const PROXY_TYPE_SOCKS5 As Long = 1
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
Private Const CERT_CHAIN_POLICY_SSL As Long = 4
Private Const CERT_FIND_ANY As Long = 0
Private Const CERT_FIND_SUBJECT_STR_A As Long = &H80007
Private Const SECPKG_ATTR_REMOTE_CERT_CONTEXT As Long = &H53
Private Const SECPKG_ATTR_STREAM_SIZES As Long = 4
Private Const AUTHTYPE_SERVER As Long = 1
Private Const CERT_STORE_PROV_SYSTEM As Long = 10
Private Const CERT_SYSTEM_STORE_CURRENT_USER As Long = &H10000
Private Const CERT_CHAIN_REVOCATION_CHECK_CHAIN As Long = &H20000000
Private Const X509_ASN_ENCODING As Long = 1
Private Const PKCS_7_ASN_ENCODING As Long = &H10000
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const PKCS12_ALLOW_OVERWRITE_KEY As Long = &H4000
Private Const SECBUFFER_VERSION As Long = 0
Private Const SECBUFFER_EMPTY As Long = 0
Private Const SECBUFFER_DATA As Long = 1
Private Const SECBUFFER_TOKEN As Long = 2
Private Const SECBUFFER_EXTRA As Long = 5
Private Const SECBUFFER_STREAM_HEADER As Long = 7
Private Const SECBUFFER_STREAM_TRAILER As Long = 6
Private Const SEC_E_OK As Long = 0
Private Const SEC_I_CONTINUE_NEEDED As Long = &H90312
Private Const SEC_E_INCOMPLETE_MESSAGE As Long = &H80090318
Private Const SEC_I_RENEGOTIATE As Long = &H90321
Private Const ETHERNET_HEADER As Long = 14
Private Const IP_HEADER_MIN As Long = 20
Private Const TCP_HEADER_MIN As Long = 20
Private Const TLS_RECORD_HEADER As Long = 5
Private Const WEBSOCKET_HEADER_MAX As Long = 14
Private Const WS_OPCODE_CONTINUATION As Byte = 0
Private Const WS_OPCODE_TEXT As Byte = 1
Private Const WS_OPCODE_BINARY As Byte = 2
Private Const WS_OPCODE_CLOSE As Byte = 8
Private Const WS_OPCODE_PING As Byte = 9
Private Const WS_OPCODE_PONG As Byte = 10

Private Const ZLIB_VERSION          As String = "1.2.11"
Private Const Z_OK                  As Long = 0
Private Const Z_STREAM_END          As Long = 1
Private Const Z_SYNC_FLUSH          As Long = 2
Private Const Z_FINISH              As Long = 4
Private Const Z_DEFLATED            As Long = 8
Private Const Z_DEFAULT_COMPRESSION As Long = -1
Private Const Z_DEFAULT_STRATEGY    As Long = 0
Private Const ZLIB_WBITS_RAW       As Long = -15
Private Const ZLIB_MEM_LEVEL       As Long = 8

Private Const SECPKG_CRED_OUTBOUND_NTLM As Long = &H2
Private Const SEC_I_COMPLETE_NEEDED As Long = &H90313

Private Enum MqttPacketType
    MQTT_CONNECT = 1
    MQTT_CONNACK = 2
    MQTT_PUBLISH = 3
    MQTT_PUBACK = 4
    MQTT_PUBREC = 5
    MQTT_PUBREL = 6
    MQTT_PUBCOMP = 7
    MQTT_SUBSCRIBE = 8
    MQTT_SUBACK = 9
    MQTT_UNSUBSCRIBE = 10
    MQTT_UNSUBACK = 11
    MQTT_PINGREQ = 12
    MQTT_PINGRESP = 13
    MQTT_DISCONNECT = 14
End Enum

Private m_WSAInitialized As Boolean
Private m_Connections() As WasabiConnection
Private m_ConnectionCount As Long
Private m_DefaultHandle As Long
Private m_LastError As WasabiError
Private m_LastErrorCode As Long
Private m_TechnicalDetails As String
Private m_Utf8Buf() As Byte
Private m_Utf8BufSize As Long
Private m_ZeroCopyText As String
Private m_ZeroCopyBinary() As Byte
#If VBA7 Then
    Private m_ClientCertContextPtrs(0 To MAX_CONNECTIONS - 1) As LongPtr
#Else
    Private m_ClientCertContextPtrs(0 To MAX_CONNECTIONS - 1) As Long
#End If
Public EnableErrorDialog As Boolean

Private Sub FillRandomBytes(ByRef buf() As Byte, ByVal count As Long)
    Dim i As Long
    If CryptGenRandom(0, count, buf(0)) = 0 Then
        For i = 0 To count - 1
            buf(i) = CByte(Int(Rnd * 256))
        Next i
    End If
End Sub

Private Function TickDiff(ByVal startTick As Long, ByVal endTick As Long) As Long
    Dim s As Currency
    Dim e As Currency
    If startTick >= 0 Then
        s = startTick
    Else
        s = startTick + 4294967296@
    End If
    If endTick >= 0 Then
        e = endTick
    Else
        e = endTick + 4294967296@
    End If
    If e >= s Then
        TickDiff = CLng(e - s)
    Else
        TickDiff = CLng(e - s + 4294967296@)
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

Private Function GetProjectPath() As String
    Dim path As String
    path = Application.VBE.ActiveVBProject.FileName
    path = Left(path, InStrRev(path, "\"))
    GetProjectPath = path
End Function

Private Function ResolveHandle(ByVal handle As Long) As Long
    If handle = INVALID_CONN_HANDLE Then
        ResolveHandle = m_DefaultHandle
    Else
        ResolveHandle = handle
    End If
End Function

Private Function ValidIndex(ByVal handle As Long) As Boolean
    If handle < 0 Or handle >= MAX_CONNECTIONS Then Exit Function
    InitConnectionPool
    ValidIndex = True
End Function

Private Sub WasabiLog(ByVal handle As Long, ByVal msg As String)
    Debug.Print "[WASABI] " & msg
    If ValidIndex(handle) Then
        If m_Connections(handle).LogCallback <> "" Then
            Application.Run m_Connections(handle).LogCallback, msg
        End If
    End If
End Sub

Private Sub SetError(ByVal errType As WasabiError, ByVal techMsg As String, ByVal userMsg As String, ByVal handle As Long, Optional ByVal errCode As Long = 0)
    Static lastErr As Long
    Static lastHandle As Long
    If errType = ERR_NONE Then Exit Sub
    m_LastError = errType
    m_LastErrorCode = errCode
    m_TechnicalDetails = techMsg
    WasabiLog handle, "ERR " & errType & " | " & techMsg
    If errCode <> 0 Then WasabiLog handle, "SysCode: " & errCode & " (0x" & hex(errCode) & ")"
    If ValidIndex(handle) Then
        m_Connections(handle).LastError = errType
        m_Connections(handle).LastErrorCode = errCode
        m_Connections(handle).TechnicalDetails = techMsg
        If m_Connections(handle).EnableErrorDialog Then
            If errType <> lastErr Or handle <> lastHandle Then
                lastErr = errType
                lastHandle = handle
                MsgBox userMsg, vbCritical, "Wasabi WebSocket Error"
            End If
        End If
    ElseIf EnableErrorDialog Then
        MsgBox userMsg, vbCritical, "Wasabi WebSocket Error"
    End If
End Sub

Private Function WSAErrDesc(ByVal code As Long) As String
    Select Case code
        Case 10004: WSAErrDesc = "WSAEINTR - Interrupted"
        Case 10013: WSAErrDesc = "WSAEACCES - Permission denied"
        Case 10014: WSAErrDesc = "WSAEFAULT - Bad address"
        Case 10022: WSAErrDesc = "WSAEINVAL - Invalid argument"
        Case 10024: WSAErrDesc = "WSAEMFILE - Too many sockets"
        Case 10035: WSAErrDesc = "WSAEWOULDBLOCK - Operation would block"
        Case 10036: WSAErrDesc = "WSAEINPROGRESS - Operation in progress"
        Case 10037: WSAErrDesc = "WSAEALREADY - Already in progress"
        Case 10038: WSAErrDesc = "WSAENOTSOCK - Not a socket"
        Case 10039: WSAErrDesc = "WSAEDESTADDRREQ - Destination address required"
        Case 10040: WSAErrDesc = "WSAEMSGSIZE - Message too long"
        Case 10047: WSAErrDesc = "WSAEAFNOSUPPORT - Address family not supported"
        Case 10048: WSAErrDesc = "WSAEADDRINUSE - Address in use"
        Case 10049: WSAErrDesc = "WSAEADDRNOTAVAIL - Address not available"
        Case 10050: WSAErrDesc = "WSAENETDOWN - Network is down"
        Case 10051: WSAErrDesc = "WSAENETUNREACH - Network unreachable"
        Case 10052: WSAErrDesc = "WSAENETRESET - Network dropped connection"
        Case 10053: WSAErrDesc = "WSAECONNABORTED - Connection aborted"
        Case 10054: WSAErrDesc = "WSAECONNRESET - Connection reset by peer"
        Case 10055: WSAErrDesc = "WSAENOBUFS - No buffer space"
        Case 10056: WSAErrDesc = "WSAEISCONN - Socket already connected"
        Case 10057: WSAErrDesc = "WSAENOTCONN - Socket not connected"
        Case 10058: WSAErrDesc = "WSAESHUTDOWN - Shutdown"
        Case 10060: WSAErrDesc = "WSAETIMEDOUT - Connection timed out"
        Case 10061: WSAErrDesc = "WSAECONNREFUSED - Connection refused"
        Case 10064: WSAErrDesc = "WSAEHOSTDOWN - Host is down"
        Case 10065: WSAErrDesc = "WSAEHOSTUNREACH - Host unreachable"
        Case 11001: WSAErrDesc = "WSAHOST_NOT_FOUND - Host not found"
        Case 11002: WSAErrDesc = "WSATRY_AGAIN - Non-authoritative host not found"
        Case 11003: WSAErrDesc = "WSANO_RECOVERY - Non-recoverable DNS error"
        Case 11004: WSAErrDesc = "WSANO_DATA - No address for hostname"
        Case Else: WSAErrDesc = "WSA error " & code
    End Select
End Function

Private Function GetCloseCodeDesc(ByVal code As Integer) As String
    Select Case code
        Case 1000: GetCloseCodeDesc = "Normal Closure"
        Case 1001: GetCloseCodeDesc = "Going Away"
        Case 1002: GetCloseCodeDesc = "Protocol Error"
        Case 1003: GetCloseCodeDesc = "Unsupported Data"
        Case 1004: GetCloseCodeDesc = "Reserved"
        Case 1005: GetCloseCodeDesc = "No Status Received"
        Case 1006: GetCloseCodeDesc = "Abnormal Closure"
        Case 1007: GetCloseCodeDesc = "Invalid Frame Payload"
        Case 1008: GetCloseCodeDesc = "Policy Violation"
        Case 1009: GetCloseCodeDesc = "Message Too Big"
        Case 1010: GetCloseCodeDesc = "Mandatory Extension"
        Case 1011: GetCloseCodeDesc = "Internal Server Error"
        Case 1012: GetCloseCodeDesc = "Service Restart"
        Case 1013: GetCloseCodeDesc = "Try Again Later"
        Case 1014: GetCloseCodeDesc = "Bad Gateway"
        Case 1015: GetCloseCodeDesc = "TLS Handshake Failure"
        Case Else: GetCloseCodeDesc = "Unknown (" & code & ")"
    End Select
End Function

Private Sub InitConnectionPool()
    Dim i As Long
    If m_ConnectionCount > 0 Then Exit Sub
    Randomize
    ReDim m_Connections(0 To MAX_CONNECTIONS - 1)
    For i = 0 To MAX_CONNECTIONS - 1
        m_Connections(i).Socket = INVALID_SOCKET
        m_Connections(i).Connected = False
        m_Connections(i).hNtlmCred.dwLower = 0
        m_Connections(i).hNtlmCred.dwUpper = 0
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
            ReDim m_Connections(i).MqttBuffer(0 To 4095)
            ResetConnectionState i
            InitializeMTU i
            AllocConnection = i
            Exit Function
        End If
    Next i
    AllocConnection = INVALID_CONN_HANDLE
End Function

Private Sub ResetConnectionState(ByVal handle As Long)
    With m_Connections(handle)
        .Connected = False
        .TLS = False
        .Host = ""
        .port = 0
        .path = ""
        .OriginalUrl = ""
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
        .CustomHeaderCount = 0
        .AutoReconnect = False
        .ReconnectMaxAttempts = DEFAULT_RECONNECT_MAX_ATTEMPTS
        .ReconnectAttempts = 0
        .ReconnectBaseDelayMs = DEFAULT_RECONNECT_BASE_DELAY_MS
        .PingIntervalMs = DEFAULT_PING_INTERVAL_MS
        .LastPingSentAt = 0
        .ReceiveTimeoutMs = DEFAULT_RECEIVE_TIMEOUT_MS
        .EnableErrorDialog = False
        .LogCallback = ""
        .NoDelay = False
        .proxyHost = ""
        .proxyPort = 0
        .proxyUser = ""
        .proxyPass = ""
        .proxyType = PROXY_TYPE_HTTP
        .ProxyEnabled = False
        .InactivityTimeoutMs = 0
        .LastActivityAt = 0
        .SubProtocol = ""
        .CustomBufferSize = 0
        .CustomFragmentSize = 0
        .AutoMTU = True
        .ZeroCopyEnabled = False
        .closeCode = 0
        .closeReason = ""
        .CloseInitiatedByUs = False
        .PreferIPv6 = False
        .ValidateServerCert = False
        .EnableRevocationCheck = False
        .ClientCertThumb = ""
        .ClientCertPfxPath = ""
        .ClientCertPfxPass = ""
        .UseHttp2 = False
        .ProxyUseNtlm = False
        .LastRttMs = 0
        .LastPingTimestamp = 0
        .MqttParserStage = 0
        .MqttBufLen = 0
        .MqttExpectedRemaining = 0
        .MqttCurrentPacketType = 0
        .MqttCurrentFlags = 0
        With .Stats
            .BytesSent = 0
            .BytesReceived = 0
            .MessagesSent = 0
            .MessagesReceived = 0
            .ConnectedAt = 0
        End With
        .mtu.CurrentMTU = DEFAULT_MTU
        .mtu.MaxSegmentSize = 1460
        .mtu.OptimalFrameSize = 1024
        .mtu.LastProbeTime = 0
        .mtu.ProbeEnabled = True
        .mtu.UseTLSFragmentation = .TLS
        .DeflateEnabled = False
        .DeflateContextTakeover = True
        .InflateContextTakeover = True
        .DeflateWindowBits = ZLIB_WBITS_RAW
        .InflateWindowBits = ZLIB_WBITS_RAW
        .DeflateReady = False
        .InflateReady = False
        .DeflateActive = False
        .FragmentIsCompressed = False
        .ClientMaxWindowBits = 15
        .ServerMaxWindowBits = 15
    End With
End Sub

Private Sub FreeSecurityHandles(ByVal handle As Long)
    FreeDeflateStreams handle
    With m_Connections(handle)
        If .pClientCertCtx <> 0 Then
            CertFreeCertificateContext .pClientCertCtx
            .pClientCertCtx = 0
        End If
        If .hClientCertStore <> 0 Then
            CertCloseStore .hClientCertStore, 0
            .hClientCertStore = 0
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
        If .hNtlmCred.dwLower <> 0 Or .hNtlmCred.dwUpper <> 0 Then
            FreeCredentialsHandle .hNtlmCred
            .hNtlmCred.dwLower = 0
            .hNtlmCred.dwUpper = 0
        End If
    End With
End Sub

Private Sub CleanupHandle(ByVal handle As Long)
    If Not ValidIndex(handle) Then Exit Sub
    With m_Connections(handle)
        If .Socket <> INVALID_SOCKET Then
            sock_closesocket .Socket
            .Socket = INVALID_SOCKET
        End If
    End With
    FreeSecurityHandles handle
    If handle >= 0 And handle < MAX_CONNECTIONS Then
        m_ClientCertContextPtrs(handle) = 0
    End If
    ResetConnectionState handle
End Sub

Private Sub InitializeMTU(ByVal handle As Long)
    With m_Connections(handle)
        .mtu.CurrentMTU = DEFAULT_MTU
        .mtu.LastProbeTime = 0
        .mtu.ProbeEnabled = True
        CalculateOptimalFrameSize handle
    End With
End Sub

Private Sub CalculateOptimalFrameSize(ByVal handle As Long)
    Dim ipOverhead As Long
    Dim tlsOverhead As Long
    Dim available As Long
    With m_Connections(handle)
        ipOverhead = IIf(.PreferIPv6, 40, IP_HEADER_MIN)
        If .TLS Then
            tlsOverhead = TLS_RECORD_HEADER + .Sizes.cbHeader + .Sizes.cbTrailer
        Else
            tlsOverhead = 0
        End If
        available = .mtu.CurrentMTU - ETHERNET_HEADER - ipOverhead - TCP_HEADER_MIN - tlsOverhead - WEBSOCKET_HEADER_MAX
        If available < 125 Then
            available = 125
        End If
        If available > 65535 Then
            available = 65535
        End If
        .mtu.MaxSegmentSize = .mtu.CurrentMTU - ETHERNET_HEADER - ipOverhead - TCP_HEADER_MIN
        .mtu.OptimalFrameSize = available
    End With
End Sub

Private Sub probeMTU(ByVal handle As Long)
    Dim mss As Long
    Dim optVal As Long
    Dim optlen As Long
    Dim probeMTU As Long
    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub
        optlen = 4
        If sock_getsockopt(.Socket, IPPROTO_TCP_SOL, TCP_MAXSEG, optVal, optlen) = 0 And optVal > 0 Then
            mss = optVal
        Else
            mss = 1460
        End If
        probeMTU = mss + TCP_HEADER_MIN + IIf(.PreferIPv6, 40, IP_HEADER_MIN) + ETHERNET_HEADER
        If probeMTU <> .mtu.CurrentMTU Then
            .mtu.CurrentMTU = probeMTU
            CalculateOptimalFrameSize handle
            WasabiLog handle, "MTU updated: " & .mtu.CurrentMTU & " MSS=" & mss & " OptFrame=" & .mtu.OptimalFrameSize & " (handle=" & handle & ")"
        End If
        .mtu.LastProbeTime = GetTickCount()
    End With
End Sub

Private Sub ApplySocketOptions(ByVal handle As Long)
    Dim optVal As Long
    Dim wsaErr As Long
    With m_Connections(handle)
        If .Socket = INVALID_SOCKET Then Exit Sub
        optVal = IIf(.NoDelay, 1, 0)
        If sock_setsockopt(.Socket, IPPROTO_TCP_SOL, TCP_NODELAY, optVal, 4) <> 0 Then
            wsaErr = WSAGetLastError()
            WasabiLog handle, "TCP_NODELAY failed: " & WSAErrDesc(wsaErr)
        End If
        optVal = 1
        If sock_setsockopt(.Socket, SOL_SOCKET, SO_KEEPALIVE, optVal, 4) <> 0 Then
            wsaErr = WSAGetLastError()
            WasabiLog handle, "SO_KEEPALIVE failed: " & WSAErrDesc(wsaErr)
        End If
        optVal = BUFFER_SIZE
        sock_setsockopt .Socket, SOL_SOCKET, SO_RCVBUF, optVal, 4
        sock_setsockopt .Socket, SOL_SOCKET, SO_SNDBUF, optVal, 4
    End With
End Sub

Private Function WaitForDataOn(ByVal handle As Long, ByVal timeoutMs As Long) As Boolean
    Dim readSet As FD_SET
    Dim timeout As TIMEVAL
    Dim effective As Long
    effective = timeoutMs
    If effective = 0 And ValidIndex(handle) Then
        If m_Connections(handle).ReceiveTimeoutMs > 0 Then
            effective = m_Connections(handle).ReceiveTimeoutMs
        End If
    End If
    readSet.fd_count = 1
    readSet.fd_array(0) = m_Connections(handle).Socket
    timeout.tv_sec = effective \ 1000
    timeout.tv_usec = (effective Mod 1000) * 1000
    WaitForDataOn = (sock_select(0, readSet, ByVal 0&, ByVal 0&, timeout) > 0)
End Function

Private Function RawSendFor(ByVal handle As Long, ByRef frame() As Byte) As Boolean
    Dim totalSent As Long
    Dim toSend As Long
    Dim sent As Long
    Dim wsaErr As Long
    toSend = UBound(frame) + 1
    totalSent = 0
    With m_Connections(handle)
        Do While totalSent < toSend
            sent = sock_send(.Socket, frame(totalSent), toSend - totalSent, 0)
            If sent <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_SEND_FAILED, "send() failed: " & WSAErrDesc(wsaErr), "Failed to send data to server.", handle, wsaErr
                .Connected = False
                Exit Function
            End If
            totalSent = totalSent + sent
        Loop
    End With
    RawSendFor = True
End Function

Private Function BuildWSFrame(ByRef payload() As Byte, ByVal payloadLen As Long, ByVal opcode As Byte, ByVal isFinal As Boolean, Optional ByVal setRSV1 As Boolean = False) As Byte()
    Dim mask(0 To 3) As Byte
    Dim headerLen As Long
    Dim frame() As Byte
    Dim finBit As Byte
    Dim rsv1 As Byte
    Dim i As Long
    rsv1 = IIf(setRSV1, &H40, 0)
    FillRandomBytes mask, 4
    finBit = IIf(isFinal, &H80, 0)
    If payloadLen < 126 Then
        headerLen = 6
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &H80 Or CByte(payloadLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
    ElseIf payloadLen < 65536 Then
        headerLen = 8
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &HFE
        frame(2) = CByte((payloadLen \ 256) And &HFF)
        frame(3) = CByte(payloadLen And &HFF)
        frame(4) = mask(0)
        frame(5) = mask(1)
        frame(6) = mask(2)
        frame(7) = mask(3)
    Else
        headerLen = 14
        ReDim frame(0 To headerLen + payloadLen - 1)
        frame(0) = finBit Or rsv1 Or opcode
        frame(1) = &HFF
        frame(2) = 0
        frame(3) = 0
        frame(4) = 0
        frame(5) = 0
        frame(6) = CByte((payloadLen \ 16777216) And &HFF)
        frame(7) = CByte((payloadLen \ 65536) And &HFF)
        frame(8) = CByte((payloadLen \ 256) And &HFF)
        frame(9) = CByte(payloadLen And &HFF)
        frame(10) = mask(0)
        frame(11) = mask(1)
        frame(12) = mask(2)
        frame(13) = mask(3)
    End If
    For i = 0 To payloadLen - 1
        frame(headerLen + i) = payload(LBound(payload) + i) Xor mask(i Mod 4)
    Next i
    BuildWSFrame = frame
End Function

Private Function StringToUtf8(ByVal str As String) As Byte()
    Dim need As Long
    Dim written As Long
    Dim result() As Byte
    If Len(str) = 0 Then
        StringToUtf8 = result
        Exit Function
    End If
    need = Len(str) * 4
    If need > m_Utf8BufSize Then
        ReDim m_Utf8Buf(0 To need - 1)
        m_Utf8BufSize = need
    End If
    written = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), Len(str), m_Utf8Buf(0), need, NULL_PTR, NULL_PTR)
    If written > 0 Then
        ReDim result(0 To written - 1)
        CopyMemory result(0), m_Utf8Buf(0), written
    End If
    StringToUtf8 = result
End Function

Private Function Utf8ToString(ByRef utf8() As Byte, ByVal length As Long) As String
    Dim charCount As Long
    Dim result As String
    If length <= 0 Then
        Utf8ToString = ""
        Exit Function
    End If
    charCount = MultiByteToWideChar(CP_UTF8, 0, utf8(LBound(utf8)), length, NULL_PTR, 0)
    If charCount > 0 Then
        result = String$(charCount, vbNullChar)
        MultiByteToWideChar CP_UTF8, 0, utf8(LBound(utf8)), length, StrPtr(result), charCount
    End If
    Utf8ToString = result
End Function

Private Function Base64Encode(ByRef Bytes() As Byte) As String
    Dim b64 As String
    Dim dataLen As Long
    Dim result() As String
    Dim i As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim chunk As Long
    Dim idx As Long
    b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    dataLen = UBound(Bytes) - LBound(Bytes) + 1
    ReDim result(0 To ((dataLen + 2) \ 3) * 4 - 1)
    idx = 0
    For i = LBound(Bytes) To LBound(Bytes) + dataLen - 1 Step 3
        b1 = CLng(Bytes(i))
        If i + 1 <= LBound(Bytes) + dataLen - 1 Then
            b2 = CLng(Bytes(i + 1))
        Else
            b2 = 0
        End If
        If i + 2 <= LBound(Bytes) + dataLen - 1 Then
            b3 = CLng(Bytes(i + 2))
        Else
            b3 = 0
        End If
        chunk = b1 * 65536 + b2 * 256 + b3
        result(idx) = Mid(b64, (chunk \ 262144) + 1, 1)
        idx = idx + 1
        result(idx) = Mid(b64, ((chunk \ 4096) And 63) + 1, 1)
        idx = idx + 1
        If i + 1 <= LBound(Bytes) + dataLen - 1 Then
            result(idx) = Mid(b64, ((chunk \ 64) And 63) + 1, 1)
        Else
            result(idx) = "="
        End If
        idx = idx + 1
        If i + 2 <= LBound(Bytes) + dataLen - 1 Then
            result(idx) = Mid(b64, (chunk And 63) + 1, 1)
        Else
            result(idx) = "="
        End If
        idx = idx + 1
    Next i
    ReDim Preserve result(0 To idx - 1)
    Base64Encode = Join(result, "")
End Function

Private Function ParseURL(ByVal url As String, ByRef outHost As String, ByRef outPort As Long, ByRef outPath As String, ByRef outTLS As Boolean) As Boolean
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
        portVal = Val(portStr)
        If portVal <= 0 Or portVal > 65535 Then Exit Function
        outPort = portVal
    Else
        outHost = work
    End If
    If Len(outHost) = 0 Then Exit Function
    ParseURL = True
End Function

Private Function ResolveHostname(ByVal hostname As String, ByVal handle As Long) As Long
    Dim addr As Long
    Dim wsaErr As Long
#If VBA7 Then
    Dim hostent As LongPtr
    Dim he As HOSTENT64
    Dim addrList As LongPtr
    Dim pAddr As LongPtr
#Else
    Dim hostent As Long
    Dim he As HOSTENT32
    Dim addrList As Long
    Dim pAddr As Long
#End If
    addr = sock_inet_addr(hostname)
    If addr <> INADDR_NONE Then
        ResolveHostname = addr
        Exit Function
    End If
    hostent = sock_gethostbyname(hostname)
    If hostent = 0 Then
        wsaErr = WSAGetLastError()
        SetError ERR_DNS_RESOLVE_FAILED, "gethostbyname failed for '" & hostname & "': " & WSAErrDesc(wsaErr), "Cannot resolve address: " & hostname & vbCrLf & WSAErrDesc(wsaErr), handle, wsaErr
        Exit Function
    End If
#If VBA7 Then
    CopyMemoryFromPtr he, hostent, LenB(he)
    addrList = he.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 8
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#Else
    CopyMemoryFromPtr he, hostent, LenB(he)
    addrList = he.h_addr_list
    If addrList = 0 Then Exit Function
    CopyMemoryFromPtr pAddr, addrList, 4
    If pAddr = 0 Then Exit Function
    CopyMemoryFromPtr addr, pAddr, 4
#End If
    ResolveHostname = addr
End Function

Private Function ResolveAndConnect(ByVal handle As Long, ByVal hostname As String, ByVal port As Long) As Boolean
#If VBA7 Then
    Dim ppResult As LongPtr
    Dim pCur As LongPtr
    Dim pNext As LongPtr
    Dim pSockaddr As LongPtr
    Dim sock6 As LongPtr
    Dim sock4 As LongPtr
    Dim aiAddrLenFull As LongPtr
#Else
    Dim ppResult As Long
    Dim pCur As Long
    Dim pNext As Long
    Dim pSockaddr As Long
    Dim sock6 As Long
    Dim sock4 As Long
#End If
    Dim hints() As Byte
    Dim aiFamily As Long
    Dim aiSocktype As Long
    Dim aiProtocol As Long
    Dim aiAddrLen As Long
    Dim nbMode As Long
    Dim writeSet As FD_SET
    Dim exceptSet As FD_SET
    Dim tv As TIMEVAL
    Dim selResult As Long
    Dim wsaErr As Long
    Dim sa6() As Byte
    Dim sa4() As Byte
    Dim sa6Len As Long
    Dim sa4Len As Long
    Dim found6 As Boolean
    Dim found4 As Boolean
    Dim startTick As Long
    Dim elapsed As Long
    Dim sin4 As SOCKADDR_IN

    sock6 = INVALID_SOCKET
    sock4 = INVALID_SOCKET

#If VBA7 Then
    ReDim hints(0 To 47)
#Else
    ReDim hints(0 To 31)
#End If
    aiSocktype = SOCK_STREAM
    CopyMemory hints(8), aiSocktype, 4
    aiProtocol = IPPROTO_TCP
    CopyMemory hints(12), aiProtocol, 4

    ppResult = 0
    If sock_getaddrinfo(hostname, CStr(port), VarPtr(hints(0)), ppResult) = 0 And ppResult <> 0 Then
        pCur = ppResult
        Do While pCur <> 0
#If VBA7 Then
            CopyMemoryFromPtr aiFamily, pCur + 4, 4
            CopyMemoryFromPtr aiSocktype, pCur + 8, 4
            CopyMemoryFromPtr aiAddrLenFull, pCur + 16, 8
            aiAddrLen = CLng(aiAddrLenFull And &H7FFFFFFF)
            CopyMemoryFromPtr pSockaddr, pCur + 32, 8
            CopyMemoryFromPtr pNext, pCur + 40, 8
#Else
            CopyMemoryFromPtr aiFamily, pCur + 4, 4
            CopyMemoryFromPtr aiSocktype, pCur + 8, 4
            CopyMemoryFromPtr aiAddrLen, pCur + 16, 4
            CopyMemoryFromPtr pSockaddr, pCur + 24, 4
            CopyMemoryFromPtr pNext, pCur + 28, 4
#End If
            If aiSocktype = SOCK_STREAM And aiAddrLen > 0 And pSockaddr <> 0 Then
                If aiFamily = AF_INET6 And Not found6 Then
                    ReDim sa6(0 To aiAddrLen - 1)
                    CopyMemoryFromPtr sa6(0), pSockaddr, aiAddrLen
                    sa6Len = aiAddrLen
                    found6 = True
                ElseIf aiFamily = AF_INET And Not found4 Then
                    ReDim sa4(0 To aiAddrLen - 1)
                    CopyMemoryFromPtr sa4(0), pSockaddr, aiAddrLen
                    sa4Len = aiAddrLen
                    found4 = True
                End If
            End If
            If found6 And found4 Then Exit Do
            pCur = pNext
        Loop
        sock_freeaddrinfo ppResult
    End If

    If Not found6 And Not found4 Then
        sin4.sin_family = AF_INET
        sin4.sin_port = sock_htons(port)
        sin4.sin_addr = ResolveHostname(hostname, handle)
        If sin4.sin_addr = 0 Then Exit Function
        sa4Len = LenB(sin4)
        ReDim sa4(0 To sa4Len - 1)
        CopyMemory sa4(0), sin4, sa4Len
        found4 = True
    End If

    If found6 Then
        sock6 = sock_socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP)
        If sock6 <> INVALID_SOCKET Then
            nbMode = 1
            sock_ioctlsocket sock6, FIONBIO, nbMode
            sock_connect sock6, VarPtr(sa6(0)), sa6Len
        Else
            found6 = False
        End If
    End If

    If found6 And found4 Then
        startTick = GetTickCount()
        Do
            writeSet.fd_count = 1
            writeSet.fd_array(0) = sock6
            exceptSet.fd_count = 1
            exceptSet.fd_array(0) = sock6
            tv.tv_sec = 0
            tv.tv_usec = 50000
            selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
            If selResult > 0 And exceptSet.fd_count = 0 Then
                nbMode = 0
                sock_ioctlsocket sock6, FIONBIO, nbMode
                m_Connections(handle).Socket = sock6
                ResolveAndConnect = True
                Exit Function
            End If
            If selResult > 0 And exceptSet.fd_count > 0 Then
                sock_closesocket sock6
                sock6 = INVALID_SOCKET
                found6 = False
                Exit Do
            End If
            elapsed = TickDiff(startTick, GetTickCount())
            If elapsed >= HAPPY_EYEBALLS_DELAY_MS Then Exit Do
            DoEvents
        Loop
    End If

    If Not ResolveAndConnect And found4 Then
        sock4 = sock_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If sock4 <> INVALID_SOCKET Then
            nbMode = 1
            sock_ioctlsocket sock4, FIONBIO, nbMode
            sock_connect sock4, VarPtr(sa4(0)), sa4Len
        End If
    End If

    If Not ResolveAndConnect Then
        Dim raceTimeout As Long
        raceTimeout = 10000
        startTick = GetTickCount()
        Do
            If sock6 <> INVALID_SOCKET Then
                writeSet.fd_count = 1
                writeSet.fd_array(0) = sock6
                exceptSet.fd_count = 1
                exceptSet.fd_array(0) = sock6
                tv.tv_sec = 0
                tv.tv_usec = 50000
                selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
                If selResult > 0 And exceptSet.fd_count = 0 Then
                    nbMode = 0
                    sock_ioctlsocket sock6, FIONBIO, nbMode
                    m_Connections(handle).Socket = sock6
                    If sock4 <> INVALID_SOCKET Then sock_closesocket sock4
                    ResolveAndConnect = True
                    Exit Function
                End If
                If selResult > 0 And exceptSet.fd_count > 0 Then
                    sock_closesocket sock6
                    sock6 = INVALID_SOCKET
                End If
            End If
            If sock4 <> INVALID_SOCKET Then
                writeSet.fd_count = 1
                writeSet.fd_array(0) = sock4
                exceptSet.fd_count = 1
                exceptSet.fd_array(0) = sock4
                tv.tv_sec = 0
                tv.tv_usec = 50000
                selResult = sock_select(0, ByVal 0&, writeSet, exceptSet, tv)
                If selResult > 0 And exceptSet.fd_count = 0 Then
                    nbMode = 0
                    sock_ioctlsocket sock4, FIONBIO, nbMode
                    m_Connections(handle).Socket = sock4
                    If sock6 <> INVALID_SOCKET Then sock_closesocket sock6
                    ResolveAndConnect = True
                    Exit Function
                End If
                If selResult > 0 And exceptSet.fd_count > 0 Then
                    sock_closesocket sock4
                    sock4 = INVALID_SOCKET
                End If
            End If
            If sock6 = INVALID_SOCKET And sock4 = INVALID_SOCKET Then Exit Do
            elapsed = TickDiff(startTick, GetTickCount())
            If elapsed >= raceTimeout Then Exit Do
            DoEvents
        Loop
    End If

    If Not ResolveAndConnect Then
        If sock6 <> INVALID_SOCKET Then sock_closesocket sock6
        If sock4 <> INVALID_SOCKET Then sock_closesocket sock4
        wsaErr = WSAGetLastError()
        SetError ERR_CONNECT_FAILED, "Connect failed: " & WSAErrDesc(wsaErr), "Could not connect to server." & vbCrLf & WSAErrDesc(wsaErr), handle, wsaErr
    End If
End Function

Private Function DoProxyHTTP(ByVal handle As Long) As Boolean
    Dim req As String
    Dim sendBuf() As Byte
    Dim recvBuf() As Byte
    Dim received As Long
    Dim response As String
    Dim sendResult As Long
    Dim wsaErr As Long
    With m_Connections(handle)
        req = "CONNECT " & .Host & ":" & .port & " HTTP/1.1" & vbCrLf
        req = req & "Host: " & .Host & ":" & .port & vbCrLf
        If .proxyUser <> "" And Not .ProxyUseNtlm Then
            req = req & "Proxy-Authorization: Basic " & Base64Encode(StrConv(.proxyUser & ":" & .proxyPass, vbFromUnicode)) & vbCrLf
        End If
        If .ProxyUseNtlm Then
            req = req & "Proxy-Authorization: NTLM TlRMTVNTUAABAAAAB4IIogAAAAAAAAAAAAAAAAAAAAAFASgKAAAADw==" & vbCrLf
        End If
        req = req & vbCrLf
        sendBuf = StrConv(req, vbFromUnicode)
        sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
        If sendResult <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "send() to proxy failed with WSA error " & wsaErr, "Failed to send CONNECT to proxy.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "Proxy CONNECT timeout", "Proxy did not respond to CONNECT request.", handle
            Exit Function
        End If
        ReDim recvBuf(0 To 4095)
        received = sock_recv(.Socket, recvBuf(0), 4096, 0)
        If received <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "recv() from proxy failed with WSA error " & wsaErr, "Failed to receive proxy response.", handle, wsaErr
            Exit Function
        End If
        response = Left(StrConv(recvBuf, vbUnicode), received)
        If InStr(response, "407") > 0 Then
            If .ProxyUseNtlm Then
                Dim ntlmHeader As String
                Dim hPos As Long
                Dim lf As Long
                Dim ntlmToken As String
                hPos = InStr(LCase(response), "proxy-authenticate: ntlm")
                If hPos > 0 Then
                    ntlmHeader = Mid(response, hPos)
                    lf = InStr(ntlmHeader, vbCrLf)
                    If lf > 0 Then ntlmHeader = Left(ntlmHeader, lf - 1)
                    ntlmToken = GenerateNtlmToken(handle, ntlmHeader, .proxyHost)
                    If ntlmToken <> "" Then
                        req = "CONNECT " & .Host & ":" & .port & " HTTP/1.1" & vbCrLf
                        req = req & "Host: " & .Host & ":" & .port & vbCrLf
                        req = req & "Proxy-Authorization: " & ntlmToken & vbCrLf
                        req = req & vbCrLf
                        sendBuf = StrConv(req, vbFromUnicode)
                        sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
                        If sendResult <= 0 Then Exit Function
                        If Not WaitForDataOn(handle, 5000) Then Exit Function
                        received = sock_recv(.Socket, recvBuf(0), 4096, 0)
                        If received <= 0 Then Exit Function
                        response = Left(StrConv(recvBuf, vbUnicode), received)
                    End If
                End If
            Else
                SetError ERR_PROXY_AUTH_FAILED, "Proxy returned 407 Proxy Authentication Required", "Proxy authentication failed." & vbCrLf & "Please check your proxy credentials.", handle
                Exit Function
            End If
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
            SetError ERR_PROXY_TUNNEL_FAILED, "Proxy CONNECT rejected: " & statusLine, "Proxy refused the tunnel connection." & vbCrLf & "Server: " & .Host & ":" & .port, handle
            Exit Function
        End If
    End With
    DoProxyHTTP = True
End Function

Private Function DoProxySOCKS5(ByVal handle As Long) As Boolean
    Dim sendBuf() As Byte
    Dim recvBuf(0 To 255) As Byte
    Dim received As Long
    Dim wsaErr As Long
    Dim hostBytes() As Byte
    Dim hostLen As Byte
    Dim i As Long
    With m_Connections(handle)
        If .proxyUser <> "" Then
            ReDim sendBuf(0 To 3)
            sendBuf(0) = 5
            sendBuf(1) = 2
            sendBuf(2) = 0
            sendBuf(3) = 2
        Else
            ReDim sendBuf(0 To 2)
            sendBuf(0) = 5
            sendBuf(1) = 1
            sendBuf(2) = 0
        End If
        If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 greeting failed: " & wsaErr, "SOCKS5 handshake failed.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 greeting timeout", "SOCKS5 server did not respond.", handle
            Exit Function
        End If
        received = sock_recv(.Socket, recvBuf(0), 256, 0)
        If received < 2 Or recvBuf(0) <> 5 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 invalid greeting response", "SOCKS5 handshake failed.", handle
            Exit Function
        End If
        If recvBuf(1) = 255 Then
            SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 no acceptable auth method", "SOCKS5 authentication failed.", handle
            Exit Function
        End If
        If recvBuf(1) = 2 Then
            Dim userB() As Byte
            Dim passB() As Byte
            Dim uLen As Byte
            Dim pLen As Byte
            userB = StrConv(.proxyUser, vbFromUnicode)
            passB = StrConv(.proxyPass, vbFromUnicode)
            uLen = CByte(UBound(userB) + 1)
            pLen = CByte(UBound(passB) + 1)
            ReDim sendBuf(0 To 3 + uLen + pLen)
            sendBuf(0) = 1
            sendBuf(1) = uLen
            For i = 0 To uLen - 1
                sendBuf(2 + i) = userB(i)
            Next i
            sendBuf(2 + uLen) = pLen
            For i = 0 To pLen - 1
                sendBuf(3 + uLen + i) = passB(i)
            Next i
            If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth send failed: " & wsaErr, "SOCKS5 authentication failed.", handle, wsaErr
                Exit Function
            End If
            If Not WaitForDataOn(handle, 5000) Then
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth timeout", "SOCKS5 authentication timed out.", handle
                Exit Function
            End If
            received = sock_recv(.Socket, recvBuf(0), 256, 0)
            If received < 2 Or recvBuf(1) <> 0 Then
                SetError ERR_PROXY_AUTH_FAILED, "SOCKS5 auth rejected", "SOCKS5 authentication failed. Check credentials.", handle
                Exit Function
            End If
        End If
        hostBytes = StrConv(.Host, vbFromUnicode)
        If UBound(hostBytes) + 1 > 255 Then
            SetError ERR_PROXY_CONNECT_FAILED, "Hostname too long for SOCKS5: " & Len(.Host) & " chars", "Proxy hostname exceeds SOCKS5 limit.", handle
            Exit Function
        End If
        hostLen = CByte(UBound(hostBytes) + 1)
        ReDim sendBuf(0 To 6 + hostLen + 1)
        sendBuf(0) = 5
        sendBuf(1) = 1
        sendBuf(2) = 0
        sendBuf(3) = 3
        sendBuf(4) = hostLen
        For i = 0 To hostLen - 1
            sendBuf(5 + i) = hostBytes(i)
        Next i
        sendBuf(5 + hostLen) = CByte((.port \ 256) And &HFF)
        sendBuf(6 + hostLen) = CByte(.port And &HFF)
        If sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0) <= 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT send failed: " & wsaErr, "SOCKS5 connect request failed.", handle, wsaErr
            Exit Function
        End If
        If Not WaitForDataOn(handle, 5000) Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT timeout", "SOCKS5 server did not respond to CONNECT.", handle
            Exit Function
        End If
        received = sock_recv(.Socket, recvBuf(0), 256, 0)
        If received < 4 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT response too short", "SOCKS5 connect failed.", handle
            Exit Function
        End If
        If recvBuf(0) <> 5 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT wrong version: " & recvBuf(0), "SOCKS5 connect failed.", handle
            Exit Function
        End If
        If recvBuf(1) <> 0 Then
            SetError ERR_PROXY_CONNECT_FAILED, "SOCKS5 CONNECT rejected, code: " & recvBuf(1), "SOCKS5 server rejected connection, code: " & recvBuf(1), recvBuf(1), handle
            Exit Function
        End If
    End With
    DoProxySOCKS5 = True
End Function

Private Function GenerateNtlmToken(ByVal handle As Long, ByVal proxyAuthHeader As String, ByVal proxyHost As String) As String
    Dim hCred As SecHandle
    Dim hContext As SecHandle
    Dim tsExpiry As SECURITY_INTEGER
    Dim result As Long
    Dim outBuf As SecBuffer
    Dim outBufDesc As SecBufferDesc
    Dim inBuf(0 To 1) As SecBuffer
    Dim inBufDesc As SecBufferDesc
    Dim contextAttr As Long
    Dim targetName As String
    Dim serverToken() As Byte
    Dim b64token As String
    Dim outData() As Byte
    Dim recvLen As Long
    Dim i As Long
    targetName = "HTTP/" & proxyHost
    result = AcquireCredentialsHandle(NULL_PTR, "NTLM", SECPKG_CRED_OUTBOUND, NULL_PTR, ByVal 0&, 0, 0, hCred, tsExpiry)
    If result <> 0 Then Exit Function
    If InStr(proxyAuthHeader, "NTLM ") > 0 Then
        b64token = Mid(proxyAuthHeader, InStr(proxyAuthHeader, "NTLM ") + 5)
        serverToken = StrConv(b64token, vbFromUnicode)
    End If
    recvLen = UBound(serverToken) - LBound(serverToken) + 1
    Dim recvBuffer() As Byte
    If recvLen > 0 And Not IsEmpty(serverToken) Then
        ReDim recvBuffer(0 To recvLen - 1)
        CopyMemory recvBuffer(0), serverToken(0), recvLen
    End If
    outBufDesc.ulVersion = SECBUFFER_VERSION
    outBufDesc.cBuffers = 1
    outBufDesc.pBuffers = VarPtr(outBuf)
    outBuf.cbBuffer = 0
    outBuf.BufferType = SECBUFFER_TOKEN
    outBuf.pvBuffer = 0
    If recvLen = 0 Then
        result = InitializeSecurityContext(hCred, NULL_PTR, targetName, ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM, 0, 0, NULL_PTR, 0, hContext, outBufDesc, contextAttr, tsExpiry)
    Else
        inBufDesc.ulVersion = SECBUFFER_VERSION
        inBufDesc.cBuffers = 2
        inBufDesc.pBuffers = VarPtr(inBuf(0))
        inBuf(0).cbBuffer = recvLen
        inBuf(0).BufferType = SECBUFFER_TOKEN
        inBuf(0).pvBuffer = VarPtr(recvBuffer(0))
        inBuf(1).cbBuffer = 0
        inBuf(1).BufferType = SECBUFFER_EMPTY
        inBuf(1).pvBuffer = 0
        result = InitializeSecurityContextContinue(hCred, hContext, targetName, ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM, 0, 0, inBufDesc, 0, hContext, outBufDesc, contextAttr, tsExpiry)
    End If
    If outBuf.cbBuffer > 0 Then
        ReDim outData(0 To outBuf.cbBuffer - 1)
        CopyMemoryFromPtr outData(0), outBuf.pvBuffer, outBuf.cbBuffer
        GenerateNtlmToken = "NTLM " & Base64Encode(outData)
        FreeContextBuffer outBuf.pvBuffer
    End If
    DeleteSecurityContext hContext
    FreeCredentialsHandle hCred
End Function

Private Function LoadClientCert(ByVal handle As Long) As Boolean
#If VBA7 Then
    Dim outCtx As LongPtr
    Dim outStore As LongPtr
#Else
    Dim outCtx As Long
    Dim outStore As Long
#End If
    Dim fileNum As Integer
    Dim pfxBytes() As Byte
    Dim blob As CRYPT_DATA_BLOB
    Dim fileLen As Long
#If VBA7 Then
    Dim pwPtr As LongPtr
#Else
    Dim pwPtr As Long
#End If
    outCtx = 0
    outStore = 0
    With m_Connections(handle)
        If .ClientCertPfxPath <> "" Then
            If Dir(.ClientCertPfxPath) = "" Then
                SetError ERR_CERT_LOAD_FAILED, "PFX file not found: " & .ClientCertPfxPath, "Client certificate file not found.", handle
                Exit Function
            End If
            fileNum = FreeFile
            Open .ClientCertPfxPath For Binary Access Read As #fileNum
            fileLen = LOF(fileNum)
            If fileLen = 0 Then
                Close #fileNum
                SetError ERR_CERT_LOAD_FAILED, "PFX file is empty", "Client certificate file is empty.", handle
                Exit Function
            End If
            ReDim pfxBytes(0 To fileLen - 1)
            Get #fileNum, , pfxBytes
            Close #fileNum
            blob.cbData = fileLen
            blob.pbData = VarPtr(pfxBytes(0))
            pwPtr = IIf(Len(.ClientCertPfxPass) > 0, StrPtr(.ClientCertPfxPass), NULL_PTR)
            outStore = PFXImportCertStore(blob, pwPtr, CRYPT_EXPORTABLE Or PKCS12_ALLOW_OVERWRITE_KEY)
            If outStore = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "PFXImportCertStore failed: 0x" & hex(Err.LastDllError), "Failed to import client certificate PFX.", handle, Err.LastDllError
                Exit Function
            End If
            outCtx = CertFindCertificateInStore(outStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_ANY, ByVal NULL_PTR, 0)
            If outCtx = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "CertFindCertificateInStore (ANY) failed", "No certificate found in PFX.", handle
                CertCloseStore outStore, 0
                Exit Function
            End If
        ElseIf .ClientCertThumb <> "" Then
            outStore = CertOpenStore(CERT_STORE_PROV_SYSTEM, 0, NULL_PTR, CERT_SYSTEM_STORE_CURRENT_USER, StrPtr("MY"))
            If outStore = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "CertOpenStore (MY) failed: 0x" & hex(Err.LastDllError), "Cannot open Windows certificate store.", handle, Err.LastDllError
                Exit Function
            End If
            outCtx = CertFindCertificateInStore(outStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_SUBJECT_STR_A, ByVal StrPtr(.ClientCertThumb), 0)
            If outCtx = 0 Then
                SetError ERR_CERT_LOAD_FAILED, "Certificate not found for subject: " & .ClientCertThumb, "Client certificate not found in store.", handle
                CertCloseStore outStore, 0
                Exit Function
            End If
        Else
            Exit Function
        End If
        .pClientCertCtx = outCtx
        .hClientCertStore = outStore
        m_ClientCertContextPtrs(handle) = outCtx
    End With
    LoadClientCert = True
End Function

Private Function ValidateServerCert(ByVal handle As Long) As Boolean
#If VBA7 Then
    Dim pRemoteCert As LongPtr
    Dim pChainCtx As LongPtr
#Else
    Dim pRemoteCert As Long
    Dim pChainCtx As Long
#End If
    Dim chainPara As CERT_CHAIN_PARA
    Dim sslExtra As SSL_EXTRA_CERT_CHAIN_POLICY_PARA
    Dim policyPara As CERT_CHAIN_POLICY_PARA
    Dim policyStatus As CERT_CHAIN_POLICY_STATUS
    Dim result As Long
    Dim chainFlags As Long
    With m_Connections(handle)
        pRemoteCert = 0
        result = QueryContextAttributes(.hContext, SECPKG_ATTR_REMOTE_CERT_CONTEXT, pRemoteCert)
        If result <> 0 Or pRemoteCert = 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "QueryContextAttributes(REMOTE_CERT) failed: 0x" & hex(result), "Cannot retrieve server certificate.", handle, result
            Exit Function
        End If
        chainPara.cbSize = LenB(chainPara)
        pChainCtx = 0
        chainFlags = 0
        If .EnableRevocationCheck Then
            chainFlags = CERT_CHAIN_REVOCATION_CHECK_CHAIN
        End If
        result = CertGetCertificateChain(NULL_PTR, pRemoteCert, 0, 0, chainPara, chainFlags, NULL_PTR, pChainCtx)
        If result = 0 Or pChainCtx = 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "CertGetCertificateChain failed: 0x" & hex(Err.LastDllError), "Cannot build certificate chain.", handle
            CertFreeCertificateContext pRemoteCert
            Exit Function
        End If
        sslExtra.cbSize = LenB(sslExtra)
        sslExtra.dwAuthType = AUTHTYPE_SERVER
        sslExtra.fdwChecks = 0
        sslExtra.pwszServerName = StrPtr(.Host)
        policyPara.cbSize = LenB(policyPara)
        policyPara.dwFlags = 0
        policyPara.pvExtraPolicyPara = VarPtr(sslExtra)
        policyStatus.cbSize = LenB(policyStatus)
        result = CertVerifyCertificateChainPolicy(CERT_CHAIN_POLICY_SSL, pChainCtx, policyPara, policyStatus)
        CertFreeCertificateChain pChainCtx
        CertFreeCertificateContext pRemoteCert
        If result = 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "CertVerifyCertificateChainPolicy failed: 0x" & hex(Err.LastDllError), "Certificate policy check failed.", handle
            Exit Function
        End If
        If policyStatus.dwError <> 0 Then
            SetError ERR_CERT_VALIDATE_FAILED, "Cert validation error 0x" & hex(policyStatus.dwError) & " chain=" & policyStatus.lChainIndex & " elem=" & policyStatus.lElementIndex, "Server certificate is not trusted (0x" & hex(policyStatus.dwError) & ").", handle, policyStatus.dwError
            Exit Function
        End If
    End With
    ValidateServerCert = True
End Function

Private Function DoTLSHandshake(ByVal handle As Long) As Long
    Dim outBuf As SecBuffer
    Dim outBufDesc As SecBufferDesc
    Dim inBuf(0 To 1) As SecBuffer
    Dim inBufDesc As SecBufferDesc
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
                outBufDesc.ulVersion = SECBUFFER_VERSION
                outBufDesc.cBuffers = 1
                outBufDesc.pBuffers = VarPtr(outBuf)
                outBuf.cbBuffer = 0
                outBuf.BufferType = SECBUFFER_TOKEN
                outBuf.pvBuffer = 0
                result = InitializeSecurityContext(.hCred, NULL_PTR, .Host, contextFlags, 0, 0, NULL_PTR, 0, .hContext, outBufDesc, contextAttr, tsExpiry)
                firstCall = False
            Else
                inBufDesc.ulVersion = SECBUFFER_VERSION
                inBufDesc.cBuffers = 2
                inBufDesc.pBuffers = VarPtr(inBuf(0))
                inBuf(0).cbBuffer = recvLen
                inBuf(0).BufferType = SECBUFFER_TOKEN
                inBuf(0).pvBuffer = VarPtr(recvBuffer(0))
                inBuf(1).cbBuffer = 0
                inBuf(1).BufferType = SECBUFFER_EMPTY
                inBuf(1).pvBuffer = 0
                outBufDesc.ulVersion = SECBUFFER_VERSION
                outBufDesc.cBuffers = 1
                outBufDesc.pBuffers = VarPtr(outBuf)
                outBuf.cbBuffer = 0
                outBuf.BufferType = SECBUFFER_TOKEN
                outBuf.pvBuffer = 0
                result = InitializeSecurityContextContinue(.hCred, .hContext, .Host, contextFlags, 0, 0, inBufDesc, 0, .hContext, outBufDesc, contextAttr, tsExpiry)
                extraData = 0
                For i = 0 To 1
                    If inBuf(i).BufferType = SECBUFFER_EXTRA And inBuf(i).cbBuffer > 0 Then
                        extraData = inBuf(i).cbBuffer
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
                DoTLSHandshake = result
                Exit Function
            End If
            If outBuf.cbBuffer > 0 And outBuf.pvBuffer <> 0 Then
                ReDim outData(0 To outBuf.cbBuffer - 1)
                CopyMemoryFromPtr outData(0), outBuf.pvBuffer, outBuf.cbBuffer
                sock_send .Socket, outData(0), outBuf.cbBuffer, 0
                FreeContextBuffer outBuf.pvBuffer
            End If
            If result = SEC_E_OK Then
                DoTLSHandshake = 0
                Exit Function
            End If
            If result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE Then
                If Not WaitForDataOn(handle, 10000) Then
                    DoTLSHandshake = -1
                    Exit Function
                End If
                recv = sock_recv(.Socket, recvBuffer(recvLen), 32768 - recvLen, 0)
                If recv <= 0 Then
                    DoTLSHandshake = -1
                    Exit Function
                End If
                recvLen = recvLen + recv
            End If
            If loopCount > 30 Then
                DoTLSHandshake = -1
                Exit Function
            End If
        Loop While result = SEC_I_CONTINUE_NEEDED Or result = SEC_E_INCOMPLETE_MESSAGE
    End With
    DoTLSHandshake = result
End Function

Private Function TLSSend(ByVal handle As Long, ByRef data() As Byte) As Boolean
    Dim buffers(0 To 3) As SecBuffer
    Dim bufferDesc As SecBufferDesc
    Dim sendBuf() As Byte
    Dim dataLen As Long
    Dim totalLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim maxChunk As Long
    Dim result As Long
    Dim toSend As Long
    Dim totalSent As Long
    Dim sent As Long
    Dim wsaErr As Long
    Dim i As Long
    With m_Connections(handle)
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            TLSSend = True
            Exit Function
        End If
        maxChunk = .Sizes.cbMaximumMessage
        If maxChunk <= 0 Then
            maxChunk = 16384
        End If
        offset = 0
        Do While offset < dataLen
            chunkSize = maxChunk
            If offset + chunkSize > dataLen Then
                chunkSize = dataLen - offset
            End If
            totalLen = .Sizes.cbHeader + chunkSize + .Sizes.cbTrailer
            ReDim sendBuf(0 To totalLen - 1)
            For i = 0 To chunkSize - 1
                sendBuf(.Sizes.cbHeader + i) = data(LBound(data) + offset + i)
            Next i
            buffers(0).cbBuffer = .Sizes.cbHeader
            buffers(0).BufferType = SECBUFFER_STREAM_HEADER
            buffers(0).pvBuffer = VarPtr(sendBuf(0))
            buffers(1).cbBuffer = chunkSize
            buffers(1).BufferType = SECBUFFER_DATA
            buffers(1).pvBuffer = VarPtr(sendBuf(.Sizes.cbHeader))
            buffers(2).cbBuffer = .Sizes.cbTrailer
            buffers(2).BufferType = SECBUFFER_STREAM_TRAILER
            buffers(2).pvBuffer = VarPtr(sendBuf(.Sizes.cbHeader + chunkSize))
            buffers(3).cbBuffer = 0
            buffers(3).BufferType = SECBUFFER_EMPTY
            buffers(3).pvBuffer = 0
            bufferDesc.ulVersion = SECBUFFER_VERSION
            bufferDesc.cBuffers = 4
            bufferDesc.pBuffers = VarPtr(buffers(0))
            result = EncryptMessage(.hContext, 0, bufferDesc, 0)
            If result <> 0 Then
                SetError ERR_TLS_ENCRYPT_FAILED, "EncryptMessage failed: 0x" & hex(result), "TLS encryption error.", handle, result
                Exit Function
            End If
            toSend = buffers(0).cbBuffer + buffers(1).cbBuffer + buffers(2).cbBuffer
            totalSent = 0
            Do While totalSent < toSend
                sent = sock_send(.Socket, sendBuf(totalSent), toSend - totalSent, 0)
                If sent <= 0 Then
                    wsaErr = WSAGetLastError()
                    SetError ERR_SEND_FAILED, "send() after TLS encrypt failed: " & WSAErrDesc(wsaErr), "Failed to send encrypted data.", handle, wsaErr
                    .Connected = False
                    Exit Function
                End If
                totalSent = totalSent + sent
            Loop
            offset = offset + chunkSize
        Loop
    End With
    TLSSend = True
End Function

Private Sub TLSDecrypt(ByVal handle As Long)
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
                SetError ERR_TLS_RENEGOTIATE, "TLS renegotiation requested - closing", "Secure connection interrupted (renegotiation).", handle, SEC_I_RENEGOTIATE
                .Connected = False
                If .AutoReconnect Then TryReconnect handle
                Exit Sub
            End If
            If result <> SEC_E_OK Then
                SetError ERR_TLS_DECRYPT_FAILED, "DecryptMessage failed: 0x" & hex(result), "TLS decryption error.", handle, result
                Exit Sub
            End If
            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_DATA Then
                    dataLen = buffers(i).cbBuffer
                    If dataLen > 0 Then
                        If .DecryptLen + dataLen > UBound(.DecryptBuffer) + 1 Then
                            Dim newSize As Long
                            newSize = .DecryptLen + dataLen + BUFFER_SIZE
                            ReDim Preserve .DecryptBuffer(0 To newSize - 1)
                        End If
                        CopyMemoryFromPtr .DecryptBuffer(.DecryptLen), buffers(i).pvBuffer, dataLen
                        .DecryptLen = .DecryptLen + dataLen
                    End If
                End If
            Next i
            extraLen = 0
            For i = 0 To 3
                If buffers(i).BufferType = SECBUFFER_EXTRA Then
                    extraLen = buffers(i).cbBuffer
                    If extraLen > 0 Then
                        CopyMemoryFromPtr .recvBuffer(0), buffers(i).pvBuffer, extraLen
                    End If
                    Exit For
                End If
            Next i
            .recvLen = extraLen
        Loop
    End With
End Sub

Private Function ReceiveHTTPResponse(ByVal handle As Long) As String
    Dim tempBuf() As Byte
    Dim received As Long
    Dim headerEnd As Long
    Dim i As Long
    Dim headerBytes() As Byte
    Dim copyLen As Long
    Dim remainingLen As Long
    With m_Connections(handle)
        Do
            If Not WaitForDataOn(handle, 5000) Then Exit Do
            ReDim tempBuf(0 To 8191)
            received = sock_recv(.Socket, tempBuf(0), 8192, 0)
            If received <= 0 Then Exit Do
            copyLen = received
            If .recvLen + copyLen > BUFFER_SIZE Then
                copyLen = BUFFER_SIZE - .recvLen
            End If
            If copyLen > 0 Then
                CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                .recvLen = .recvLen + copyLen
            End If
            If .TLS Then TLSDecrypt handle
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
                    ReceiveHTTPResponse = StrConv(headerBytes, vbUnicode)
                    remainingLen = .DecryptLen - headerEnd
                    If remainingLen > 0 Then
                        CopyMemory .DecryptBuffer(0), .DecryptBuffer(headerEnd), remainingLen
                    End If
                    .DecryptLen = remainingLen
                    Exit Function
                End If
            End If
        Loop
        If .DecryptLen > 0 Then
            ReDim headerBytes(0 To .DecryptLen - 1)
            CopyMemory headerBytes(0), .DecryptBuffer(0), .DecryptLen
            ReceiveHTTPResponse = StrConv(headerBytes, vbUnicode)
            .DecryptLen = 0
        End If
    End With
End Function

Private Function U32Shl1(ByVal v As Long) As Long
    U32Shl1 = (v And &H3FFFFFFF) * 2
    If v And &H40000000 Then U32Shl1 = U32Shl1 Or &H80000000
End Function

Private Function SHR32(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long
    Dim result As Long
    result = v
    For i = 1 To n
        If result And &H80000000 Then
            result = (result And &H7FFFFFFF) \ 2 Or &H40000000
        Else
            result = result \ 2
        End If
    Next i
    SHR32 = result
End Function

Private Function ROTL32(ByVal v As Long, ByVal n As Long) As Long
    Dim hi As Long
    Dim i As Long
    n = n Mod 32
    If n = 0 Then
        ROTL32 = v
        Exit Function
    End If
    hi = U32Shl1(v)
    For i = 2 To n
        hi = U32Shl1(hi)
    Next i
    ROTL32 = hi Or SHR32(v, 32 - n)
End Function

Private Function ADD32(ByVal a As Long, ByVal b As Long) As Long
    Dim aLo As Long
    Dim bLo As Long
    Dim aHi As Long
    Dim bHi As Long
    Dim sLo As Long
    Dim sHi As Long
    aLo = a And &HFFFF&
    bLo = b And &HFFFF&
    aHi = (a And &H7FFF0000) \ &H10000
    bHi = (b And &H7FFF0000) \ &H10000
    sLo = aLo + bLo
    sHi = aHi + bHi + (sLo \ &H10000)
    If a And &H80000000 Then sHi = sHi + &H8000&
    If b And &H80000000 Then sHi = sHi + &H8000&
    ADD32 = (sLo And &HFFFF&) Or ((sHi And &H7FFF&) * &H10000)
    If sHi And &H8000& Then ADD32 = ADD32 Or &H80000000
End Function

Private Function SHA1(ByRef data() As Byte) As Byte()
    Dim h0 As Long
    Dim h1 As Long
    Dim h2 As Long
    Dim h3 As Long
    Dim h4 As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim k As Long
    Dim temp As Long
    Dim w(0 To 79) As Long
    Dim msg() As Byte
    Dim origLen As Long
    Dim totalLen As Long
    Dim padLen As Long
    Dim i As Long
    Dim chunk As Long
    Dim result(0 To 19) As Byte
    Dim hArr(0 To 4) As Long
    Dim v As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim bitLenLo As Long
    Dim wv As Long
    origLen = UBound(data) - LBound(data) + 1
    padLen = 64 - ((origLen + 9) Mod 64)
    If padLen = 64 Then padLen = 0
    totalLen = origLen + 1 + padLen + 8
    ReDim msg(0 To totalLen - 1)
    For i = 0 To origLen - 1
        msg(i) = data(LBound(data) + i)
    Next i
    msg(origLen) = &H80
    bitLenLo = origLen * 8
    msg(totalLen - 4) = CByte((bitLenLo \ &H1000000) And &HFF)
    msg(totalLen - 3) = CByte((bitLenLo \ &H10000) And &HFF)
    msg(totalLen - 2) = CByte((bitLenLo \ &H100) And &HFF)
    msg(totalLen - 1) = CByte(bitLenLo And &HFF)
    h0 = &H67452301
    h1 = &HEFCDAB89
    h2 = &H98BADCFE
    h3 = &H10325476
    h4 = &HC3D2E1F0
    For chunk = 0 To (totalLen \ 64) - 1
        For i = 0 To 15
            b0 = CLng(msg(chunk * 64 + i * 4)) And &HFF&
            b1 = CLng(msg(chunk * 64 + i * 4 + 1)) And &HFF&
            b2 = CLng(msg(chunk * 64 + i * 4 + 2)) And &HFF&
            b3 = CLng(msg(chunk * 64 + i * 4 + 3)) And &HFF&
            wv = b3 Or (b2 * &H100&) Or (b1 * &H10000)
            wv = wv Or ((b0 And &H7F&) * &H1000000)
            If b0 And &H80& Then wv = wv Or &H80000000
            w(i) = wv
        Next i
        For i = 16 To 79
            w(i) = ROTL32(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16), 1)
        Next i
        a = h0
        b = h1
        c = h2
        d = h3
        e = h4
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
            temp = ADD32(ADD32(ADD32(ADD32(ROTL32(a, 5), f), e), k), w(i))
            e = d
            d = c
            c = ROTL32(b, 30)
            b = a
            a = temp
        Next i
        h0 = ADD32(h0, a)
        h1 = ADD32(h1, b)
        h2 = ADD32(h2, c)
        h3 = ADD32(h3, d)
        h4 = ADD32(h4, e)
    Next chunk
    hArr(0) = h0
    hArr(1) = h1
    hArr(2) = h2
    hArr(3) = h3
    hArr(4) = h4
    For i = 0 To 4
        v = hArr(i)
        result(i * 4) = CByte(((v And &H7F000000) \ &H1000000) Or IIf((v And &H80000000) <> 0, &H80&, 0))
        result(i * 4 + 1) = CByte((v And &HFF0000) \ &H10000)
        result(i * 4 + 2) = CByte((v And &HFF00&) \ &H100&)
        result(i * 4 + 3) = CByte(v And &HFF&)
    Next i
    SHA1 = result
End Function

Private Function GenerateWSKey() As String
    Dim Bytes(0 To 15) As Byte
    FillRandomBytes Bytes, 16
    GenerateWSKey = Base64Encode(Bytes)
End Function

Private Sub ParseDeflateResponse(ByVal handle As Long, ByVal response As String)
    Dim extStart As Long
    Dim extLine As String
    Dim lf As Long
    Dim swbPos As Long
    Dim swbVal As Long
    Dim cwbPos As Long
    Dim cwbVal As Long
    
    extStart = InStr(LCase(response), "sec-websocket-extensions: permessage-deflate")
    If extStart = 0 Then
        m_Connections(handle).DeflateEnabled = False
        m_Connections(handle).DeflateActive = False
        Exit Sub
    End If
    
    extLine = Mid(response, extStart)
    lf = InStr(extLine, vbCrLf)
    If lf > 0 Then
        extLine = Left(extLine, lf - 1)
    End If
    
    With m_Connections(handle)
        If InStr(LCase(extLine), "client_no_context_takeover") > 0 Then
            .DeflateContextTakeover = False
        Else
            .DeflateContextTakeover = True
        End If
        
        If InStr(LCase(extLine), "server_no_context_takeover") > 0 Then
            .InflateContextTakeover = False
        Else
            .InflateContextTakeover = True
        End If
        
        swbPos = InStr(LCase(extLine), "server_max_window_bits=")
        If swbPos > 0 Then
            swbVal = Val(Mid(extLine, swbPos + 22))
            If swbVal >= 8 And swbVal <= 15 Then
                .DeflateWindowBits = -swbVal
                .ServerMaxWindowBits = swbVal
            End If
        End If
        
        cwbPos = InStr(LCase(extLine), "client_max_window_bits=")
        If cwbPos > 0 Then
            cwbVal = Val(Mid(extLine, cwbPos + 22))
            If cwbVal >= 8 And cwbVal <= 15 Then
                .InflateWindowBits = -cwbVal
                .ClientMaxWindowBits = cwbVal
            End If
        End If
        .DeflateActive = True
    End With
End Sub

Private Function DoWebSocketHandshake(ByVal handle As Long) As Boolean
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
    Dim recvBuf() As Byte
    Dim received As Long
    wsKey = GenerateWSKey()
    
    With m_Connections(handle)
        hostHeader = IIf((.TLS And .port <> 443) Or (Not .TLS And .port <> 80), .Host & ":" & .port, .Host)
        If .TLS Then
            originHeader = "https://" & IIf(.port <> 443, .Host & ":" & .port, .Host)
        Else
            originHeader = "http://" & IIf(.port <> 80, .Host & ":" & .port, .Host)
        End If
        handshake = "GET " & .path & " HTTP/1.1" & vbCrLf
        handshake = handshake & "Host: " & hostHeader & vbCrLf
        handshake = handshake & "Upgrade: websocket" & vbCrLf
        handshake = handshake & "Connection: Upgrade" & vbCrLf
        handshake = handshake & "Sec-WebSocket-Key: " & wsKey & vbCrLf
        handshake = handshake & "Sec-WebSocket-Version: 13" & vbCrLf
        If .DeflateEnabled Then
            Dim deflateExt As String
            deflateExt = "permessage-deflate"
            If Not .DeflateContextTakeover Then
                deflateExt = deflateExt & "; client_no_context_takeover"
            End If
            If Not .InflateContextTakeover Then
                deflateExt = deflateExt & "; server_no_context_takeover"
            End If
            If .ClientMaxWindowBits <> 15 Then
                deflateExt = deflateExt & "; client_max_window_bits=" & .ClientMaxWindowBits
            End If
            handshake = handshake & "Sec-WebSocket-Extensions: " & deflateExt & vbCrLf
        End If
        If .SubProtocol <> "" Then
            handshake = handshake & "Sec-WebSocket-Protocol: " & .SubProtocol & vbCrLf
        End If
        handshake = handshake & "Origin: " & originHeader & vbCrLf
        handshake = handshake & "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" & vbCrLf
        For i = 0 To .CustomHeaderCount - 1
            handshake = handshake & .CustomHeaders(i) & vbCrLf
        Next i
        handshake = handshake & vbCrLf
        sendBuf = StrConv(handshake, vbFromUnicode)
        If .TLS Then
            If Not TLSSend(handle, sendBuf) Then
                SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "TLS send of WS handshake failed", "WebSocket upgrade request failed.", handle
                Exit Function
            End If
            response = ReceiveHTTPResponse(handle)
        Else
            sendResult = sock_send(.Socket, sendBuf(0), UBound(sendBuf) + 1, 0)
            If sendResult <= 0 Then
                wsaErr = WSAGetLastError()
                SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "send() WS handshake failed: " & WSAErrDesc(wsaErr), "WebSocket upgrade request failed.", handle, wsaErr
                Exit Function
            End If
            If Not WaitForDataOn(handle, 5000) Then
                SetError ERR_WEBSOCKET_HANDSHAKE_TIMEOUT, "No WS handshake response within 5s", "WebSocket handshake timed out.", handle
                Exit Function
            End If
            ReDim recvBuf(0 To 4095)
            received = sock_recv(.Socket, recvBuf(0), 4096, 0)
            If received > 0 Then
                response = Left(StrConv(recvBuf, vbUnicode), received)
            Else
                wsaErr = WSAGetLastError()
                SetError ERR_WEBSOCKET_HANDSHAKE_FAILED, "recv() WS handshake failed: " & WSAErrDesc(wsaErr), "WebSocket handshake failed.", handle, wsaErr
                Exit Function
            End If
        End If
        If InStr(response, "101") = 0 Then
            Dim lineEnd As Long
            Dim statusLine As String
            lineEnd = InStr(response, vbCrLf)
            If lineEnd > 0 Then
                statusLine = Left(response, lineEnd - 1)
            Else
                statusLine = Left(response, 50)
            End If
            SetError ERR_HANDSHAKE_REJECTED, "Server rejected WS upgrade: " & statusLine, "WebSocket connection rejected: " & statusLine, handle
            Exit Function
        End If
        If .DeflateEnabled Then
            ParseDeflateResponse handle, response
        End If
        expectedAccept = Base64Encode(SHA1(StrConv(wsKey & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11", vbFromUnicode)))
        acceptPos = InStr(LCase(response), "sec-websocket-accept:")
        If acceptPos > 0 Then
            acceptLineEnd = InStr(acceptPos, response, vbCrLf)
            If acceptLineEnd > 0 Then
                actualAccept = Trim(Mid(response, acceptPos + 21, acceptLineEnd - acceptPos - 21))
            End If
        End If
        If actualAccept <> expectedAccept Then
            SetError ERR_HANDSHAKE_REJECTED, "Sec-WebSocket-Accept mismatch. Expected: " & expectedAccept & " Got: " & actualAccept, "WebSocket handshake failed: invalid accept key.", handle
            Exit Function
        End If
    End With
    DoWebSocketHandshake = True
End Function

Private Sub ProcessFrames(ByVal handle As Long)
    Dim opcode As Byte
    Dim fin As Boolean
    Dim isCompressed As Boolean
    Dim payloadLen As Long
    Dim wirePayloadLen As Long
    Dim maskFlag As Boolean
    Dim frameStart As Long
    Dim i As Long
    Dim payload() As Byte
    Dim textMsg As String
    Dim binaryData() As Byte
    Dim totalFrameLen As Long
    Dim inflLen As Long
    Dim inflBytes() As Byte
    With m_Connections(handle)
        Do While .DecryptLen >= 2
            fin = (.DecryptBuffer(0) And &H80) <> 0
            isCompressed = (.DecryptBuffer(0) And &H40) <> 0
            opcode = .DecryptBuffer(0) And &HF
            maskFlag = (.DecryptBuffer(1) And &H80) <> 0
            payloadLen = .DecryptBuffer(1) And &H7F
            frameStart = 2
            If payloadLen = 126 Then
                If .DecryptLen < 4 Then Exit Do
                payloadLen = CLng(.DecryptBuffer(2)) * 256 + CLng(.DecryptBuffer(3))
                frameStart = 4
            ElseIf payloadLen = 127 Then
                If .DecryptLen < 10 Then Exit Do
                payloadLen = 0
                For i = 2 To 9
                    payloadLen = payloadLen * 256 + CLng(.DecryptBuffer(i))
                Next i
                frameStart = 10
            End If
            If maskFlag Then frameStart = frameStart + 4
            If .DecryptLen < frameStart + payloadLen Then Exit Do
            wirePayloadLen = payloadLen
            If payloadLen > 0 Then
                ReDim payload(0 To payloadLen - 1)
                CopyMemory payload(0), .DecryptBuffer(frameStart), payloadLen
            End If
            Select Case opcode
                Case WS_OPCODE_TEXT
                    If Not fin Then
                        .Fragmenting = True
                        .FragmentOpcode = WS_OPCODE_TEXT
                        .FragmentIsCompressed = isCompressed
                        .FragmentLen = 0
                        If payloadLen > 0 Then
                            CopyMemory .FragmentBuffer(0), payload(0), payloadLen
                            .FragmentLen = payloadLen
                        End If
                    Else
                        Dim textPayload() As Byte
                        Dim textPayloadLen As Long
                        If .Fragmenting Then
                            If .FragmentLen + payloadLen > UBound(.FragmentBuffer) + 1 Then
                                SetError ERR_FRAGMENT_OVERFLOW, "Fragment buffer overflow on TEXT frame", "Received message too large.", handle
                                .Connected = False
                                Exit Sub
                            End If
                            If payloadLen > 0 Then
                                CopyMemory .FragmentBuffer(.FragmentLen), payload(0), payloadLen
                                .FragmentLen = .FragmentLen + payloadLen
                            End If
                            If .FragmentIsCompressed And .DeflateActive Then
                                Dim inflTextBytes() As Byte
                                Dim inflTextLen As Long
                                inflTextBytes = InflatePayload(handle, .FragmentBuffer, .FragmentLen, inflTextLen)
                                If inflTextLen = 0 Then
                                    WebSocketSendClose 1007, "Decompression failed", handle
                                    .Connected = False
                                    .Fragmenting = False
                                    .FragmentLen = 0
                                    Exit Sub
                                End If
                                textPayload = inflTextBytes
                                textPayloadLen = inflTextLen
                            Else
                                textPayload = .FragmentBuffer
                                textPayloadLen = .FragmentLen
                            End If
                            textMsg = Utf8ToString(textPayload, textPayloadLen)
                            .Fragmenting = False
                            .FragmentLen = 0
                        Else
                            If isCompressed And .DeflateActive Then
                                Dim inflTextSingle() As Byte
                                Dim inflTextSingleLen As Long
                                inflTextSingle = InflatePayload(handle, payload, payloadLen, inflTextSingleLen)
                                If inflTextSingleLen = 0 Then
                                    WebSocketSendClose 1007, "Decompression failed", handle
                                    .Connected = False
                                    Exit Sub
                                End If
                                textMsg = Utf8ToString(inflTextSingle, inflTextSingleLen)
                            Else
                                If payloadLen > 0 Then
                                    textMsg = Utf8ToString(payload, payloadLen)
                                Else
                                    textMsg = ""
                                End If
                            End If
                        End If
                        If .MsgCount < MSG_QUEUE_SIZE Then
                            .MsgQueue(.MsgTail) = textMsg
                            .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
                            .MsgCount = .MsgCount + 1
                            .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                        Else
                            WasabiLog handle, "Warning: message queue full, dropping message (handle=" & handle & ")"
                        End If
                    End If

                Case WS_OPCODE_BINARY
                    If Not fin Then
                        .Fragmenting = True
                        .FragmentOpcode = WS_OPCODE_BINARY
                        .FragmentIsCompressed = isCompressed
                        .FragmentLen = 0
                        If payloadLen > 0 Then
                            CopyMemory .FragmentBuffer(0), payload(0), payloadLen
                            .FragmentLen = payloadLen
                        End If
                    Else
                        If .Fragmenting Then
                            If .FragmentLen + payloadLen > UBound(.FragmentBuffer) + 1 Then
                                SetError ERR_FRAGMENT_OVERFLOW, "Fragment buffer overflow on BINARY frame", "Received binary message too large.", handle
                                .Connected = False
                                Exit Sub
                            End If
                            If payloadLen > 0 Then
                                CopyMemory .FragmentBuffer(.FragmentLen), payload(0), payloadLen
                                .FragmentLen = .FragmentLen + payloadLen
                            End If
                            If .FragmentIsCompressed And .DeflateActive Then
                                Dim inflBinBytes() As Byte
                                Dim inflBinLen As Long
                                inflBinBytes = InflatePayload(handle, .FragmentBuffer, .FragmentLen, inflBinLen)
                                If inflBinLen = 0 Then
                                    WebSocketSendClose 1007, "Decompression failed", handle
                                    .Connected = False
                                    .Fragmenting = False
                                    .FragmentLen = 0
                                    Exit Sub
                                End If
                                binaryData = inflBinBytes
                            Else
                                ReDim binaryData(0 To .FragmentLen - 1)
                                CopyMemory binaryData(0), .FragmentBuffer(0), .FragmentLen
                            End If
                            .Fragmenting = False
                            .FragmentLen = 0
                        Else
                            If isCompressed And .DeflateActive Then
                                Dim inflBinSingle() As Byte
                                Dim inflBinSingleLen As Long
                                inflBinSingle = InflatePayload(handle, payload, payloadLen, inflBinSingleLen)
                                If inflBinSingleLen = 0 Then
                                    WebSocketSendClose 1007, "Decompression failed", handle
                                    .Connected = False
                                    Exit Sub
                                End If
                                binaryData = inflBinSingle
                            Else
                                If payloadLen > 0 Then
                                    ReDim binaryData(0 To payloadLen - 1)
                                    CopyMemory binaryData(0), payload(0), payloadLen
                                End If
                            End If
                        End If
                        If .BinaryCount < MSG_QUEUE_SIZE Then
                            .BinaryQueue(.BinaryTail).data = binaryData
                            .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
                            .BinaryCount = .BinaryCount + 1
                            .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                        Else
                            WasabiLog handle, "Warning: binary queue full, dropping message (handle=" & handle & ")"
                        End If
                    End If

                Case WS_OPCODE_CONTINUATION
                    If .Fragmenting Then
                        If .FragmentLen + payloadLen > UBound(.FragmentBuffer) + 1 Then
                            SetError ERR_FRAGMENT_OVERFLOW, "Fragment buffer overflow on CONTINUATION frame", "Received message too large.", handle
                            .Connected = False
                            Exit Sub
                        End If
                        If payloadLen > 0 Then
                            CopyMemory .FragmentBuffer(.FragmentLen), payload(0), payloadLen
                            .FragmentLen = .FragmentLen + payloadLen
                        End If
                        If fin Then
                            Dim contPayload() As Byte
                            Dim contPayloadLen As Long
                            If .FragmentIsCompressed And .DeflateActive Then
                                Dim inflContBytes() As Byte
                                Dim inflContLen As Long
                                inflContBytes = InflatePayload(handle, .FragmentBuffer, .FragmentLen, inflContLen)
                                If inflContLen = 0 Then
                                    WebSocketSendClose 1007, "Decompression failed", handle
                                    .Connected = False
                                    .Fragmenting = False
                                    .FragmentLen = 0
                                    Exit Sub
                                End If
                                contPayload = inflContBytes
                                contPayloadLen = inflContLen
                            Else
                                contPayload = .FragmentBuffer
                                contPayloadLen = .FragmentLen
                            End If
                            If .FragmentOpcode = WS_OPCODE_TEXT Then
                                textMsg = Utf8ToString(contPayload, contPayloadLen)
                                If .MsgCount < MSG_QUEUE_SIZE Then
                                    .MsgQueue(.MsgTail) = textMsg
                                    .MsgTail = (.MsgTail + 1) Mod MSG_QUEUE_SIZE
                                    .MsgCount = .MsgCount + 1
                                    .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                                End If
                            ElseIf .FragmentOpcode = WS_OPCODE_BINARY Then
                                ReDim binaryData(0 To contPayloadLen - 1)
                                CopyMemory binaryData(0), contPayload(0), contPayloadLen
                                If .BinaryCount < MSG_QUEUE_SIZE Then
                                    .BinaryQueue(.BinaryTail).data = binaryData
                                    .BinaryTail = (.BinaryTail + 1) Mod MSG_QUEUE_SIZE
                                    .BinaryCount = .BinaryCount + 1
                                    .Stats.MessagesReceived = .Stats.MessagesReceived + 1
                                End If
                            End If
                            .Fragmenting = False
                            .FragmentLen = 0
                        End If
                    End If

                Case WS_OPCODE_CLOSE
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .Connected = False
                        Exit Sub
                    End If
                    ProcessCloseFrame handle, payload, payloadLen
                    Exit Sub

                Case WS_OPCODE_PING
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .Connected = False
                        Exit Sub
                    End If
                    SendPongFrame handle, payload, payloadLen

                Case WS_OPCODE_PONG
                    If isCompressed Then
                        WebSocketSendClose 1002, "RSV1 on control frame", handle
                        .Connected = False
                        Exit Sub
                    End If
                    ProcessPongForLatency handle
                    WasabiLog handle, "PONG received (handle=" & handle & ")"
            End Select
            totalFrameLen = frameStart + wirePayloadLen
            If .DecryptLen > totalFrameLen Then
                CopyMemory .DecryptBuffer(0), .DecryptBuffer(totalFrameLen), .DecryptLen - totalFrameLen
            End If
            .DecryptLen = .DecryptLen - totalFrameLen
        Loop
    End With
End Sub

Private Sub ProcessCloseFrame(ByVal handle As Long, ByRef payload() As Byte, ByVal payloadLen As Long)
    Dim closeCode As Integer
    Dim closeReason As String
    Dim replyFrame(0 To 7) As Byte
    Dim mask(0 To 3) As Byte
    Dim reasonBytes() As Byte
    Dim i As Long
    With m_Connections(handle)
        closeCode = 1005
        closeReason = ""
        If payloadLen >= 2 Then
            closeCode = CInt(payload(0)) * 256 + CInt(payload(1))
            If payloadLen > 2 Then
                ReDim reasonBytes(0 To payloadLen - 3)
                For i = 0 To payloadLen - 3
                    reasonBytes(i) = payload(2 + i)
                Next i
                closeReason = Utf8ToString(reasonBytes, payloadLen - 2)
            End If
        End If
        .closeCode = closeCode
        .closeReason = closeReason
        WasabiLog handle, "CLOSE received: " & closeCode & " (" & GetCloseCodeDesc(closeCode) & ") reason=""" & closeReason & """ (handle=" & handle & ")"
        If Not .CloseInitiatedByUs Then
            FillRandomBytes mask, 4
            replyFrame(0) = &H88
            replyFrame(1) = &H82
            replyFrame(2) = mask(0)
            replyFrame(3) = mask(1)
            replyFrame(4) = mask(2)
            replyFrame(5) = mask(3)
            If payloadLen >= 2 Then
                replyFrame(6) = payload(0) Xor mask(0)
                replyFrame(7) = payload(1) Xor mask(1)
            Else
                replyFrame(6) = CByte((1000 \ 256) And &HFF) Xor mask(0)
                replyFrame(7) = CByte(1000 And &HFF) Xor mask(1)
            End If
            Dim rf() As Byte
            ReDim rf(0 To 7)
            For i = 0 To 7
                rf(i) = replyFrame(i)
            Next i
            If .TLS Then
                TLSSend handle, rf
            Else
                sock_send .Socket, rf(0), 8, 0
            End If
        End If
        .Connected = False
    End With
End Sub

Private Sub SendPongFrame(ByVal handle As Long, ByRef pingPayload() As Byte, ByVal pingLen As Long)
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim i As Long
    FillRandomBytes mask, 4
    If pingLen = 0 Then
        ReDim frame(0 To 5)
        frame(0) = &H8A
        frame(1) = &H80
    Else
        ReDim frame(0 To 5 + pingLen)
        frame(0) = &H8A
        frame(1) = &H80 Or CByte(pingLen)
        For i = 0 To pingLen - 1
            frame(6 + i) = pingPayload(i) Xor mask(i Mod 4)
        Next i
    End If
    frame(2) = mask(0)
    frame(3) = mask(1)
    frame(4) = mask(2)
    frame(5) = mask(3)
    With m_Connections(handle)
        If .TLS Then
            TLSSend handle, frame
        Else
            sock_send .Socket, frame(0), UBound(frame) + 1, 0
        End If
    End With
End Sub

Private Sub ProcessPongForLatency(ByVal handle As Long)
    With m_Connections(handle)
        If .LastPingTimestamp > 0 Then
            .LastRttMs = TickDiff(.LastPingTimestamp, GetTickCount())
            .LastPingTimestamp = 0
        End If
    End With
End Sub

Private Sub FeedBuffer(ByVal handle As Long)
    Dim available As Long
    Dim tempBuf() As Byte
    Dim received As Long
    Dim wsaErr As Long
    Dim copyLen As Long
    With m_Connections(handle)
        If sock_ioctlsocket(.Socket, FIONREAD, available) <> 0 Then
            wsaErr = WSAGetLastError()
            SetError ERR_CONNECTION_LOST, "ioctlsocket(FIONREAD) failed: " & WSAErrDesc(wsaErr), "Connection lost.", handle, wsaErr
            .Connected = False
            If .AutoReconnect Then TryReconnect handle
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
                copyLen = received
                If .recvLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .recvLen
                If copyLen > 0 Then
                    CopyMemory .recvBuffer(.recvLen), tempBuf(0), copyLen
                    .recvLen = .recvLen + copyLen
                End If
                TLSDecrypt handle
            Else
                copyLen = received
                If .DecryptLen + copyLen > BUFFER_SIZE Then copyLen = BUFFER_SIZE - .DecryptLen
                If copyLen > 0 Then
                    CopyMemory .DecryptBuffer(.DecryptLen), tempBuf(0), copyLen
                    .DecryptLen = .DecryptLen + copyLen
                End If
            End If
            ProcessFrames handle
        ElseIf received = 0 Then
            SetError ERR_CONNECTION_LOST, "recv() returned 0 - server closed connection", "Server closed the connection.", handle
            .Connected = False
            If .AutoReconnect Then TryReconnect handle
        Else
            wsaErr = WSAGetLastError()
            If wsaErr <> 10035 Then
                SetError ERR_RECV_FAILED, "recv() failed: " & WSAErrDesc(wsaErr), "Failed to receive data.", handle, wsaErr
                .Connected = False
                If .AutoReconnect Then TryReconnect handle
            End If
        End If
    End With
End Sub

Private Sub TickMaintenance(ByVal handle As Long)
    Dim now As Long
    With m_Connections(handle)
        If Not .Connected Then Exit Sub
        now = GetTickCount()
        If .PingIntervalMs > 0 Then
            If TickDiff(.LastPingSentAt, now) >= .PingIntervalMs Then
                WebSocketSendPing "", handle
            End If
        End If
        If .InactivityTimeoutMs > 0 And .LastActivityAt > 0 Then
            If TickDiff(.LastActivityAt, now) >= .InactivityTimeoutMs Then
                SetError ERR_INACTIVITY_TIMEOUT, "Inactivity timeout after " & .InactivityTimeoutMs & "ms", "Connection timed out (inactivity).", handle
                .Connected = False
                If .AutoReconnect Then TryReconnect handle
                Exit Sub
            End If
        End If
        If .AutoMTU And .mtu.ProbeEnabled Then
            If TickDiff(.mtu.LastProbeTime, now) >= PMTU_DISCOVERY_INTERVAL_MS Then
                probeMTU handle
            End If
        End If
    End With
End Sub

Private Sub TryReconnect(ByVal handle As Long)
    Dim delayMs As Long
    Dim attempt As Long
    Dim i As Long
    Dim savedUrl As String
    Dim savedAutoReconnect As Boolean
    Dim savedMaxAttempts As Long
    Dim savedBaseDelay As Long
    Dim savedAttempts As Long
    Dim savedPingInterval As Long
    Dim savedReceiveTimeout As Long
    Dim savedLogCallback As String
    Dim savedErrorDialog As Boolean
    Dim savedHeaders() As String
    Dim savedHeaderCount As Long
    Dim savedNoDelay As Boolean
    Dim savedProxyHost As String
    Dim savedProxyPort As Long
    Dim savedProxyUser As String
    Dim savedProxyPass As String
    Dim savedProxyEnabled As Boolean
    Dim savedProxyType As Long
    Dim savedInactivityTimeout As Long
    Dim savedSubProtocol As String
    Dim savedBufferSize As Long
    Dim savedFragmentSize As Long
    Dim savedDeflateEnabled As Boolean
    Dim savedDeflateCtx As Boolean
    Dim savedInflateCtx As Boolean
    Dim startTick As Long
    If Not m_Connections(handle).AutoReconnect Then Exit Sub
    If m_Connections(handle).ReconnectMaxAttempts > 0 And m_Connections(handle).ReconnectAttempts >= m_Connections(handle).ReconnectMaxAttempts Then
        WasabiLog handle, "Auto-reconnect exhausted after " & m_Connections(handle).ReconnectAttempts & " attempts (handle=" & handle & ")"
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
    WasabiLog handle, "Reconnect attempt " & attempt & " in " & delayMs & "ms (handle=" & handle & ")"
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
        savedProxyType = .proxyType
        savedInactivityTimeout = .InactivityTimeoutMs
        savedSubProtocol = .SubProtocol
        savedBufferSize = .CustomBufferSize
        savedFragmentSize = .CustomFragmentSize
        savedDeflateEnabled = .DeflateEnabled
        savedDeflateCtx = .DeflateContextTakeover
        savedInflateCtx = .InflateContextTakeover
        If savedHeaderCount > 0 Then
            ReDim savedHeaders(0 To savedHeaderCount - 1)
            For i = 0 To savedHeaderCount - 1
                savedHeaders(i) = .CustomHeaders(i)
            Next i
        End If
    End With
    CleanupHandle handle
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
    Dim bufSize As Long
    Dim fragSize As Long
    bufSize = IIf(savedBufferSize > 0, savedBufferSize, BUFFER_SIZE)
    fragSize = IIf(savedFragmentSize > 0, savedFragmentSize, FRAGMENT_BUFFER_SIZE)
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
        .proxyType = savedProxyType
        .InactivityTimeoutMs = savedInactivityTimeout
        .SubProtocol = savedSubProtocol
        .CustomBufferSize = savedBufferSize
        .CustomFragmentSize = savedFragmentSize
        .DeflateEnabled = savedDeflateEnabled
        .DeflateContextTakeover = savedDeflateCtx
        .InflateContextTakeover = savedInflateCtx
        For i = 0 To savedHeaderCount - 1
            .CustomHeaders(i) = savedHeaders(i)
        Next i
    End With
    If Not ConnectHandle(handle, savedUrl) Then
        WasabiLog handle, "Reconnect attempt " & attempt & " failed (handle=" & handle & ")"
    Else
        m_Connections(handle).ReconnectAttempts = 0
        WasabiLog handle, "Reconnect succeeded (handle=" & handle & ")"
    End If
End Sub

Private Function ConnectHandle(ByVal handle As Long, ByVal url As String) As Boolean
    Dim schannelCred As SCHANNEL_CRED
    Dim tsExpiry As SECURITY_INTEGER
    Dim connectHost As String
    Dim connectPort As Long
    Dim zeroBytes() As Byte
    Dim acquireResult As Long
    Dim tlsResult As Long
    With m_Connections(handle)
        .LastError = ERR_NONE
        .LastErrorCode = 0
        .TechnicalDetails = ""
        .OriginalUrl = url
        If Not ParseURL(url, .Host, .port, .path, .TLS) Then
            SetError ERR_INVALID_URL, "Invalid URL: " & url, "Invalid WebSocket URL. Use ws:// or wss://", handle
            GoTo Fail
        End If
        connectHost = IIf(.ProxyEnabled And .proxyHost <> "", .proxyHost, .Host)
        connectPort = IIf(.ProxyEnabled And .proxyPort > 0, .proxyPort, .port)
        If Not ResolveAndConnect(handle, connectHost, connectPort) Then GoTo Fail
        InitializeMTU handle
        If .AutoMTU Then probeMTU handle
        ApplySocketOptions handle
        If .ProxyEnabled Then
            If .proxyType = PROXY_TYPE_SOCKS5 Then
                If Not DoProxySOCKS5(handle) Then GoTo Fail
            Else
                If Not DoProxyHTTP(handle) Then GoTo Fail
            End If
        End If
        If .TLS Then
            ReDim zeroBytes(0 To LenB(schannelCred) - 1)
            CopyMemory schannelCred, zeroBytes(0), LenB(schannelCred)
            schannelCred.dwVersion = SCHANNEL_CRED_VERSION
            schannelCred.grbitEnabledProtocols = SP_PROT_TLS1_2_CLIENT Or SP_PROT_TLS1_3_CLIENT
            schannelCred.dwFlags = SCH_CRED_NO_DEFAULT_CREDS Or SCH_CRED_MANUAL_CRED_VALIDATION Or SCH_CRED_IGNORE_NO_REVOCATION_CHECK Or SCH_CRED_IGNORE_REVOCATION_OFFLINE
            If .ClientCertThumb <> "" Or .ClientCertPfxPath <> "" Then
                If LoadClientCert(handle) Then
                    m_ClientCertContextPtrs(handle) = .pClientCertCtx
                    schannelCred.cCreds = 1
                    schannelCred.paCred = VarPtr(m_ClientCertContextPtrs(handle))
                Else
                    WasabiLog handle, "WARNING: client cert load failed, proceeding without it (handle=" & handle & ")"
                End If
            End If
            acquireResult = AcquireCredentialsHandle(NULL_PTR, "Microsoft Unified Security Protocol Provider", SECPKG_CRED_OUTBOUND, NULL_PTR, schannelCred, NULL_PTR, NULL_PTR, .hCred, tsExpiry)
            If acquireResult <> 0 Then
                SetError ERR_TLS_ACQUIRE_CREDS_FAILED, "AcquireCredentialsHandle failed: 0x" & hex(acquireResult), "TLS initialization failed (0x" & hex(acquireResult) & ").", handle, acquireResult
                GoTo Fail
            End If
            tlsResult = DoTLSHandshake(handle)
            If tlsResult <> 0 Then
                If tlsResult = -1 Then
                    SetError ERR_TLS_HANDSHAKE_TIMEOUT, "TLS handshake timed out with " & .Host, "TLS handshake timed out.", handle
                Else
                    SetError ERR_TLS_HANDSHAKE_FAILED, "TLS handshake failed: 0x" & hex(tlsResult), "TLS handshake failed (0x" & hex(tlsResult) & ").", handle, tlsResult
                End If
                GoTo Fail
            End If
            QueryContextAttributes .hContext, SECPKG_ATTR_STREAM_SIZES, .Sizes
            CalculateOptimalFrameSize handle
            If .ValidateServerCert Then
                If Not ValidateServerCert(handle) Then GoTo Fail
            End If
        End If
        If Not DoWebSocketHandshake(handle) Then GoTo Fail
        .Connected = True
        .Stats.ConnectedAt = GetTickCount()
        .Stats.BytesSent = 0
        .Stats.BytesReceived = 0
        .Stats.MessagesSent = 0
        .Stats.MessagesReceived = 0
        .LastPingSentAt = GetTickCount()
        .LastActivityAt = GetTickCount()
    End With
    ConnectHandle = True
    Exit Function
Fail:
    CleanupHandle handle
End Function

Private Function SendFrameFor(ByVal handle As Long, ByRef frame() As Byte) As Boolean
    If m_Connections(handle).TLS Then
        SendFrameFor = TLSSend(handle, frame)
    Else
        SendFrameFor = RawSendFor(handle, frame)
    End If
End Function

Private Function MqttEncodeRemainingLength(ByVal length As Long, ByRef buf() As Byte) As Long
    Dim encodedByte As Byte
    Dim idx As Long
    idx = 0
    Do
        encodedByte = CByte(length Mod 128)
        length = length \ 128
        If length > 0 Then
            encodedByte = encodedByte Or &H80
        End If
        buf(0 + idx) = encodedByte
        idx = idx + 1
    Loop While length > 0
    MqttEncodeRemainingLength = idx
End Function

Private Function BuildMqttConnectPacket(ByVal clientId As String, Optional ByVal username As String, Optional ByVal password As String, Optional ByVal keepAlive As Integer = 60) As Byte()
    Dim varHeader(0 To 9) As Byte
    Dim flags As Byte
    Dim clientBytes() As Byte
    Dim userBytes() As Byte
    Dim passBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim remaining As Long
    Dim rlBuf(0 To 3) As Byte
    Dim rlLen As Long
    Dim packet() As Byte
    Dim pos As Long
    varHeader(0) = 0
    varHeader(1) = 4
    varHeader(2) = Asc("M")
    varHeader(3) = Asc("Q")
    varHeader(4) = Asc("T")
    varHeader(5) = Asc("T")
    varHeader(6) = 4
    varHeader(7) = CByte((keepAlive \ 256) And &HFF)
    varHeader(8) = CByte(keepAlive And &HFF)
    flags = 0
    If username <> "" Then
        flags = flags Or &H80
    End If
    If password <> "" Then
        flags = flags Or &H40
    End If
    varHeader(9) = flags
    clientBytes = StringToUtf8(clientId)
    If username <> "" Then
        userBytes = StringToUtf8(username)
    End If
    If password <> "" Then
        passBytes = StringToUtf8(password)
    End If
    payloadLen = 2 + UBound(clientBytes) + 1
    If username <> "" Then
        payloadLen = payloadLen + 2 + UBound(userBytes) + 1
    End If
    If password <> "" Then
        payloadLen = payloadLen + 2 + UBound(passBytes) + 1
    End If
    ReDim payload(0 To payloadLen - 1)
    pos = 0
    payload(pos) = CByte((Len(clientId) \ 256) And &HFF)
    payload(pos + 1) = CByte(Len(clientId) And &HFF)
    pos = pos + 2
    CopyMemory payload(pos), clientBytes(0), UBound(clientBytes) + 1
    pos = pos + UBound(clientBytes) + 1
    If username <> "" Then
        payload(pos) = CByte((Len(username) \ 256) And &HFF)
        payload(pos + 1) = CByte(Len(username) And &HFF)
        pos = pos + 2
        CopyMemory payload(pos), userBytes(0), UBound(userBytes) + 1
        pos = pos + UBound(userBytes) + 1
    End If
    If password <> "" Then
        payload(pos) = CByte((Len(password) \ 256) And &HFF)
        payload(pos + 1) = CByte(Len(password) And &HFF)
        pos = pos + 2
        CopyMemory payload(pos), passBytes(0), UBound(passBytes) + 1
    End If
    remaining = UBound(varHeader) + 1 + payloadLen
    rlLen = MqttEncodeRemainingLength(remaining, rlBuf)
    ReDim packet(0 To 0 + rlLen + remaining - 1)
    packet(0) = MQTT_CONNECT * 16
    CopyMemory packet(1), rlBuf(0), rlLen
    CopyMemory packet(1 + rlLen), varHeader(0), 10
    CopyMemory packet(1 + rlLen + 10), payload(0), payloadLen
    BuildMqttConnectPacket = packet
End Function

Private Function MqttBuildPacket(ByVal ptype As Byte, ByVal flags As Byte, ByRef payload() As Byte, ByVal payloadLen As Long) As Byte()
    Dim remaining As Long
    Dim rlBuf(0 To 3) As Byte
    Dim rlLen As Long
    Dim packet() As Byte
    remaining = payloadLen
    rlLen = MqttEncodeRemainingLength(remaining, rlBuf)
    ReDim packet(0 To 0 + rlLen + remaining - 1)
    packet(0) = ptype * 16 Or flags
    CopyMemory packet(1), rlBuf(0), rlLen
    If payloadLen > 0 Then
        CopyMemory packet(1 + rlLen), payload(0), payloadLen
    End If
    MqttBuildPacket = packet
End Function

Private Sub MqttParseByte(ByVal handle As Long, ByVal b As Byte)
    With m_Connections(handle)
        Select Case .MqttParserStage
            Case 0
                .MqttCurrentPacketType = b \ 16
                .MqttCurrentFlags = b And &HF
                .MqttParserStage = 1
                .MqttExpectedRemaining = 0
                .MqttBufLen = 0
            Case 1
                .MqttExpectedRemaining = .MqttExpectedRemaining + (b And &H7F) * (256 ^ (.MqttBufLen))
                .MqttBufLen = .MqttBufLen + 1
                If (b And &H80) = 0 Then
                    .MqttParserStage = 2
                    .MqttBufLen = 0
                End If
            Case 2
                .MqttBuffer(.MqttBufLen) = b
                .MqttBufLen = .MqttBufLen + 1
                If .MqttBufLen >= .MqttExpectedRemaining Then
                    .MqttParserStage = 3
                End If
        End Select
    End With
End Sub

Private Function MqttHasPacket(ByVal handle As Long) As Boolean
    MqttHasPacket = (m_Connections(handle).MqttParserStage = 3)
End Function

Private Sub MqttResetParser(ByVal handle As Long)
    m_Connections(handle).MqttParserStage = 0
    m_Connections(handle).MqttBufLen = 0
End Sub

Private Function GetZlibName() As String
    #If Win64 Then
        GetZlibName = "zlib1_x64.dll"
    #Else
        GetZlibName = "zlib1_x86.dll"
    #End If
End Function

Private Function FindZlibPath() As String
    Dim searchPaths As Variant
    Dim i As Long
    Dim base As String
    Dim dllName As String
    
    searchPaths = Array( _
        GetProjectPath(), _
        GetProjectPath() & "\lib", _
        GetProjectPath() & "\deps", _
        GetProjectPath() & "\dlls", _
        GetProjectPath() & "\zlib", _
        GetProjectPath() & "\bin", _
        GetProjectPath() & "\x64", _
        GetProjectPath() & "\x86", _
        GetProjectPath() & "\native", _
        Environ$("SystemRoot") & "\System32", _
        Environ$("SystemRoot") & "\SysWOW64" _
    )
    
    dllName = GetZlibName()
    
    For i = LBound(searchPaths) To UBound(searchPaths)
        base = searchPaths(i)
        If Len(base) = 0 Then GoTo Continue
        
        If Dir$(base & "\" & dllName) <> "" Then
            FindZlibPath = base
            Exit Function
        End If
        
        If Dir$(base & "\zlib1.dll") <> "" Then
            FindZlibPath = base
            Exit Function
        End If
        
Continue:
    Next i
End Function

Private Sub LoadZlib()
    Static loaded As Boolean
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim dllName As String
    Dim fullPath As String
    
    If loaded Then Exit Sub
    
    dllName = GetZlibName()
    Dim projectPath As String
    projectPath = GetProjectPath()
    
    If projectPath <> "" Then
        fullPath = projectPath & "\" & dllName
        If Dir$(fullPath) <> "" Then
            h = LoadLibrary(fullPath)
            If h <> 0 Then loaded = True: Exit Sub
        End If
        
        fullPath = projectPath & "\zlib1.dll"
        If Dir$(fullPath) <> "" Then
            h = LoadLibrary(fullPath)
            If h <> 0 Then loaded = True: Exit Sub
        End If
    End If
    
    Dim foundPath As String
    foundPath = FindZlibPath()
    If foundPath <> "" Then
        fullPath = foundPath & "\" & dllName
        If Dir$(fullPath) <> "" Then
            h = LoadLibrary(fullPath)
            If h <> 0 Then loaded = True: Exit Sub
        End If
        
        fullPath = foundPath & "\zlib1.dll"
        If Dir$(fullPath) <> "" Then
            h = LoadLibrary(fullPath)
            If h <> 0 Then loaded = True: Exit Sub
        End If
    End If
    
    h = LoadLibrary("zlib1.dll")
    If h <> 0 Then loaded = True: Exit Sub
    
    WasabiLog INVALID_CONN_HANDLE, "LoadZlib: zlib1.dll not found - deflate unavailable"
End Sub

Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1, Optional ByVal DeflateEnabled As Boolean = False, Optional ByVal DeflateContextTakeover As Boolean = True) As Boolean
    Dim wsa As WSADATA
    Dim wsaErr As Long
    Dim handle As Long
    m_LastError = ERR_NONE
    m_LastErrorCode = 0
    m_TechnicalDetails = ""
    InitConnectionPool
    If DeflateEnabled Then LoadZlib
    If Not m_WSAInitialized Then
        wsaErr = WSAStartup(&H202, wsa)
        If wsaErr <> 0 Then
            SetError ERR_WSA_STARTUP_FAILED, "WSAStartup failed: " & wsaErr, "Network initialization failed. Code: " & wsaErr, INVALID_CONN_HANDLE, wsaErr
            outHandle = INVALID_CONN_HANDLE
            Exit Function
        End If
        m_WSAInitialized = True
    End If
    handle = AllocConnection()
    If handle = INVALID_CONN_HANDLE Then
        SetError ERR_MAX_CONNECTIONS, "Max connections (" & MAX_CONNECTIONS & ") reached", "Too many simultaneous connections.", INVALID_CONN_HANDLE
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If
    m_Connections(handle).DeflateEnabled = DeflateEnabled
    m_Connections(handle).DeflateContextTakeover = DeflateContextTakeover
    m_Connections(handle).InflateContextTakeover = DeflateContextTakeover
    If Not ConnectHandle(handle, url) Then
        outHandle = INVALID_CONN_HANDLE
        Exit Function
    End If
    m_DefaultHandle = handle
    outHandle = handle
    WebSocketConnect = True
    WasabiLog handle, "Connected to " & url & " (handle=" & handle & ")"
End Function

Public Sub WebSocketSetDeflate(ByVal enabled As Boolean, Optional ByVal contextTakeover As Boolean = True, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        If .Connected Then
            .DeflateEnabled = enabled
            .DeflateContextTakeover = contextTakeover
            .InflateContextTakeover = contextTakeover
            WasabiLog h, "DeflateEnabled set to " & enabled & " - will apply on next reconnect (handle=" & h & ")"
            Exit Sub
        End If
        .DeflateEnabled = enabled
        .DeflateContextTakeover = contextTakeover
        .InflateContextTakeover = contextTakeover
    End With
End Sub

Public Function WebSocketGetDeflateEnabled(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetDeflateEnabled = m_Connections(h).DeflateActive
End Function

Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    Dim i As Long
    Dim anyActive As Boolean
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).AutoReconnect = False
    If m_Connections(h).Connected Then WebSocketSendClose 1000, "", h
    CleanupHandle h
    If h = m_DefaultHandle Then
        m_DefaultHandle = 0
        For i = 0 To MAX_CONNECTIONS - 1
            If m_Connections(i).Connected Then
                m_DefaultHandle = i
                Exit For
            End If
        Next i
    End If
    anyActive = False
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

Public Sub WebSocketDisconnectAll()
    Dim i As Long
    InitConnectionPool
    For i = 0 To MAX_CONNECTIONS - 1
        If m_Connections(i).Connected Or m_Connections(i).Socket <> INVALID_SOCKET Then
            WebSocketDisconnect i
        End If
    Next i
End Sub

Public Function WebSocketSend(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim frame() As Byte
    Dim useDeflate As Boolean
    Dim compLen As Long
    Dim compBytes() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            SetError ERR_NOT_CONNECTED, "Send on disconnected handle=" & h, "Not connected to WebSocket server.", h
            Exit Function
        End If
        msgBytes = StringToUtf8(message)
        msgLen = SafeArrayLen(msgBytes)
        If msgLen = 0 Then
            WebSocketSend = True
            Exit Function
        End If
        useDeflate = .DeflateActive
        If useDeflate Then
            compBytes = DeflatePayload(h, msgBytes, msgLen, compLen)
            msgBytes = compBytes
            msgLen = compLen
        End If
        frame = BuildWSFrame(msgBytes, msgLen, WS_OPCODE_TEXT, True, useDeflate)
        If SendFrameFor(h, frame) Then
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            .Stats.MessagesSent = .Stats.MessagesSent + 1
            WebSocketSend = True
        End If
    End With
End Function

Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim dataLen As Long
    Dim frame() As Byte
    Dim useDeflate As Boolean
    Dim compLen As Long
    Dim compBytes() As Byte
    Dim sendData() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            SetError ERR_NOT_CONNECTED, "SendBinary on disconnected handle=" & h, "Not connected to WebSocket server.", h
            Exit Function
        End If
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            WebSocketSendBinary = True
            Exit Function
        End If
        useDeflate = .DeflateActive
        If useDeflate Then
            compBytes = DeflatePayload(h, data, dataLen, compLen)
            sendData = compBytes
            dataLen = compLen
        Else
            sendData = data
        End If
        frame = BuildWSFrame(sendData, dataLen, WS_OPCODE_BINARY, True, useDeflate)
        If SendFrameFor(h, frame) Then
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            .Stats.MessagesSent = .Stats.MessagesSent + 1
            WebSocketSendBinary = True
        End If
    End With
End Function

Public Function WebSocketSendMTUAware(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim opcode As Byte
    Dim isLast As Boolean
    Dim chunkBytes() As Byte
    Dim frame() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            SetError ERR_NOT_CONNECTED, "SendMTUAware on disconnected handle=" & h, "Not connected.", h
            Exit Function
        End If
        msgBytes = StringToUtf8(message)
        msgLen = SafeArrayLen(msgBytes)
        If msgLen = 0 Then
            WebSocketSendMTUAware = True
            Exit Function
        End If
        If Not .AutoMTU Or msgLen <= .mtu.OptimalFrameSize Then
            WebSocketSendMTUAware = WebSocketSend(message, h)
            Exit Function
        End If
        offset = 0
        opcode = WS_OPCODE_TEXT
        Do While offset < msgLen
            chunkSize = .mtu.OptimalFrameSize
            If offset + chunkSize > msgLen Then chunkSize = msgLen - offset
            isLast = (offset + chunkSize >= msgLen)
            ReDim chunkBytes(0 To chunkSize - 1)
            CopyMemory chunkBytes(0), msgBytes(offset), chunkSize
            frame = BuildWSFrame(chunkBytes, chunkSize, opcode, isLast)
            If Not SendFrameFor(h, frame) Then Exit Function
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            offset = offset + chunkSize
            opcode = WS_OPCODE_CONTINUATION
        Loop
        .Stats.MessagesSent = .Stats.MessagesSent + 1
    End With
    WebSocketSendMTUAware = True
End Function

Public Function WebSocketSendBinaryMTUAware(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim dataLen As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim opcode As Byte
    Dim isLast As Boolean
    Dim chunkBytes() As Byte
    Dim frame() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            SetError ERR_NOT_CONNECTED, "SendBinaryMTUAware on disconnected handle=" & h, "Not connected.", h
            Exit Function
        End If
        dataLen = SafeArrayLen(data)
        If dataLen = 0 Then
            WebSocketSendBinaryMTUAware = True
            Exit Function
        End If
        If Not .AutoMTU Or dataLen <= .mtu.OptimalFrameSize Then
            WebSocketSendBinaryMTUAware = WebSocketSendBinary(data, h)
            Exit Function
        End If
        offset = 0
        opcode = WS_OPCODE_BINARY
        Do While offset < dataLen
            chunkSize = .mtu.OptimalFrameSize
            If offset + chunkSize > dataLen Then chunkSize = dataLen - offset
            isLast = (offset + chunkSize >= dataLen)
            ReDim chunkBytes(0 To chunkSize - 1)
            CopyMemory chunkBytes(0), data(offset), chunkSize
            frame = BuildWSFrame(chunkBytes, chunkSize, opcode, isLast)
            If Not SendFrameFor(h, frame) Then Exit Function
            .Stats.BytesSent = .Stats.BytesSent + (UBound(frame) + 1)
            offset = offset + chunkSize
            opcode = WS_OPCODE_CONTINUATION
        Loop
        .Stats.MessagesSent = .Stats.MessagesSent + 1
    End With
    WebSocketSendBinaryMTUAware = True
End Function

Public Function WebSocketSendBatch(ByRef messages() As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim i As Long
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim frame() As Byte
    Dim frameSize As Long
    Dim batchBuf() As Byte
    Dim batchLen As Long
    Dim batchCount As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then Exit Function
        batchLen = 0
        batchCount = 0
        ReDim batchBuf(0 To 65535)
        For i = LBound(messages) To UBound(messages)
            msgBytes = StringToUtf8(messages(i))
            msgLen = SafeArrayLen(msgBytes)
            If msgLen = 0 Then GoTo NextMsg
            frame = BuildWSFrame(msgBytes, msgLen, WS_OPCODE_TEXT, True)
            frameSize = UBound(frame) + 1
            If batchLen + frameSize > 65536 Then
                Dim flushBuf() As Byte
                ReDim flushBuf(0 To batchLen - 1)
                CopyMemory flushBuf(0), batchBuf(0), batchLen
                If .TLS Then
                    If Not TLSSend(h, flushBuf) Then Exit Function
                Else
                    If Not RawSendFor(h, flushBuf) Then Exit Function
                End If
                .Stats.BytesSent = .Stats.BytesSent + batchLen
                .Stats.MessagesSent = .Stats.MessagesSent + batchCount
                batchLen = 0
                batchCount = 0
            End If
            CopyMemory batchBuf(batchLen), frame(0), frameSize
            batchLen = batchLen + frameSize
            batchCount = batchCount + 1
NextMsg:
        Next i
        If batchLen > 0 Then
            ReDim flushBuf(0 To batchLen - 1)
            CopyMemory flushBuf(0), batchBuf(0), batchLen
            If .TLS Then
                If Not TLSSend(h, flushBuf) Then Exit Function
            Else
                If Not RawSendFor(h, flushBuf) Then Exit Function
            End If
            .Stats.BytesSent = .Stats.BytesSent + batchLen
            .Stats.MessagesSent = .Stats.MessagesSent + batchCount
        End If
    End With
    WebSocketSendBatch = True
End Function

Public Function WebSocketSendBatchBinary(ByRef messages() As Variant, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim i As Long
    Dim bdata() As Byte
    Dim dataLen As Long
    Dim frame() As Byte
    Dim frameSize As Long
    Dim batchBuf() As Byte
    Dim batchLen As Long
    Dim batchCount As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then Exit Function
        batchLen = 0
        batchCount = 0
        ReDim batchBuf(0 To 65535)
        For i = LBound(messages) To UBound(messages)
            If IsArray(messages(i)) Then
                bdata = messages(i)
                dataLen = SafeArrayLen(bdata)
                If dataLen = 0 Then GoTo NextMsgBin
                frame = BuildWSFrame(bdata, dataLen, WS_OPCODE_BINARY, True)
                frameSize = UBound(frame) + 1
                If batchLen + frameSize > 65536 Then
                    Dim flushBuf() As Byte
                    ReDim flushBuf(0 To batchLen - 1)
                    CopyMemory flushBuf(0), batchBuf(0), batchLen
                    If .TLS Then
                        If Not TLSSend(h, flushBuf) Then Exit Function
                    Else
                        If Not RawSendFor(h, flushBuf) Then Exit Function
                    End If
                    .Stats.BytesSent = .Stats.BytesSent + batchLen
                    .Stats.MessagesSent = .Stats.MessagesSent + batchCount
                    batchLen = 0
                    batchCount = 0
                End If
                CopyMemory batchBuf(batchLen), frame(0), frameSize
                batchLen = batchLen + frameSize
                batchCount = batchCount + 1
            End If
NextMsgBin:
        Next i
        If batchLen > 0 Then
            ReDim flushBuf(0 To batchLen - 1)
            CopyMemory flushBuf(0), batchBuf(0), batchLen
            If .TLS Then
                If Not TLSSend(h, flushBuf) Then Exit Function
            Else
                If Not RawSendFor(h, flushBuf) Then Exit Function
            End If
            .Stats.BytesSent = .Stats.BytesSent + batchLen
            .Stats.MessagesSent = .Stats.MessagesSent + batchCount
        End If
    End With
    WebSocketSendBatchBinary = True
End Function

Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim reasonBytes() As Byte
    Dim reasonLen As Long
    Dim payloadLen As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then Exit Function
        .CloseInitiatedByUs = True
        .closeCode = code
        .closeReason = reason
        If Len(reason) > 0 Then
            reasonBytes = StringToUtf8(reason)
            reasonLen = SafeArrayLen(reasonBytes)
            If reasonLen > 123 Then reasonLen = 123
        End If
        payloadLen = 2 + reasonLen
        ReDim frame(0 To 5 + payloadLen)
        FillRandomBytes mask, 4
        frame(0) = &H88
        frame(1) = &H80 Or CByte(payloadLen)
        frame(2) = mask(0)
        frame(3) = mask(1)
        frame(4) = mask(2)
        frame(5) = mask(3)
        frame(6) = CByte((code \ 256) And &HFF) Xor mask(0)
        frame(7) = CByte(code And &HFF) Xor mask(1)
        For i = 0 To reasonLen - 1
            frame(8 + i) = reasonBytes(i) Xor mask((i + 2) Mod 4)
        Next i
        WasabiLog h, "Sending CLOSE: " & code & " (" & GetCloseCodeDesc(code) & ") reason=""" & reason & """ (handle=" & h & ")"
        WebSocketSendClose = SendFrameFor(h, frame)
        .Connected = False
    End With
End Function

Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim frame() As Byte
    Dim mask(0 To 3) As Byte
    Dim pingBytes() As Byte
    Dim pingLen As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then Exit Function
        If Len(payload) > 0 Then
            pingBytes = StringToUtf8(payload)
            pingLen = SafeArrayLen(pingBytes)
        End If
        FillRandomBytes mask, 4
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
        WebSocketSendPing = SendFrameFor(h, frame)
        If WebSocketSendPing Then
            .LastPingSentAt = GetTickCount()
            .LastPingTimestamp = GetTickCount()
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
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then Exit Function
        If Len(payload) > 0 Then
            pongBytes = StringToUtf8(payload)
            pongLen = SafeArrayLen(pongBytes)
        End If
        FillRandomBytes mask, 4
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
        WebSocketSendPong = SendFrameFor(h, frame)
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

Public Function WebSocketReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnect h
            Exit Function
        End If
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        If .MsgCount > 0 Then
            WebSocketReceive = .MsgQueue(.MsgHead)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
        End If
    End With
End Function

Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
    Dim h As Long
    Dim results() As String
    Dim count As Long
    Dim i As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then
        ReDim results(0)
        WebSocketReceiveAll = results
        Exit Function
    End If
    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnect h
            ReDim results(0)
            WebSocketReceiveAll = results
            Exit Function
        End If
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        count = .MsgCount
        If count = 0 Then
            ReDim results(0)
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

Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then
        WebSocketReceiveBinary = Empty
        Exit Function
    End If
    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnect h
            WebSocketReceiveBinary = Empty
            Exit Function
        End If
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        If .BinaryCount > 0 Then
            WebSocketReceiveBinary = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
        Else
            WebSocketReceiveBinary = Empty
        End If
    End With
End Function

Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnect h
            Exit Function
        End If
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        If .BinaryCount > 0 Then
            outData = .BinaryQueue(.BinaryHead).data
            Erase .BinaryQueue(.BinaryHead).data
            .BinaryHead = (.BinaryHead + 1) Mod MSG_QUEUE_SIZE
            .BinaryCount = .BinaryCount - 1
            WebSocketReceiveBinaryCheck = True
        End If
    End With
End Function

#If VBA7 Then
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As LongPtr, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#Else
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#End If
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If Not .Connected Then
            If .AutoReconnect Then TryReconnect h
            Exit Function
        End If
        If Not .ZeroCopyEnabled Then Exit Function
        TickMaintenance h
        If .DecryptLen > 0 Then ProcessFrames h
        FeedBuffer h
        If .MsgCount > 0 Then
            m_ZeroCopyText = .MsgQueue(.MsgHead)
            outPtr = StrPtr(m_ZeroCopyText)
            outLen = Len(m_ZeroCopyText)
            .MsgQueue(.MsgHead) = ""
            .MsgHead = (.MsgHead + 1) Mod MSG_QUEUE_SIZE
            .MsgCount = .MsgCount - 1
            WebSocketReceiveZeroCopy = True
        End If
    End With
End Function

Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .MsgCount > 0 Then WebSocketPeek = .MsgQueue(.MsgHead)
    End With
End Function

Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .MsgHead = 0
        .MsgTail = 0
        .MsgCount = 0
        .BinaryHead = 0
        .BinaryTail = 0
        .BinaryCount = 0
    End With
End Sub

Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketIsConnected = m_Connections(h).Connected
End Function

Public Function WebSocketGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetLastError = m_Connections(h).LastError
    Else
        WebSocketGetLastError = m_LastError
    End If
End Function

Public Function WebSocketGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetLastErrorCode = m_Connections(h).LastErrorCode
    Else
        WebSocketGetLastErrorCode = m_LastErrorCode
    End If
End Function

Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        WebSocketGetTechnicalDetails = m_Connections(h).TechnicalDetails
    Else
        WebSocketGetTechnicalDetails = m_TechnicalDetails
    End If
End Function

Public Function WebSocketGetErrorDescription(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim errType As WasabiError
    Dim errCode As Long
    Dim tech As String
    Dim desc As String
    h = ResolveHandle(handle)
    If ValidIndex(h) Then
        errType = m_Connections(h).LastError
        errCode = m_Connections(h).LastErrorCode
        tech = m_Connections(h).TechnicalDetails
    Else
        errType = m_LastError
        errCode = m_LastErrorCode
        tech = m_TechnicalDetails
    End If
    Select Case errType
        Case ERR_NONE: desc = "No error"
        Case ERR_WSA_STARTUP_FAILED: desc = "Winsock initialization failed"
        Case ERR_SOCKET_CREATE_FAILED: desc = "Failed to create socket"
        Case ERR_DNS_RESOLVE_FAILED: desc = "DNS resolution failed"
        Case ERR_CONNECT_FAILED: desc = "TCP connection failed"
        Case ERR_TLS_ACQUIRE_CREDS_FAILED: desc = "TLS credentials initialization failed"
        Case ERR_TLS_HANDSHAKE_FAILED: desc = "TLS handshake failed"
        Case ERR_TLS_HANDSHAKE_TIMEOUT: desc = "TLS handshake timed out"
        Case ERR_WEBSOCKET_HANDSHAKE_FAILED: desc = "WebSocket upgrade rejected"
        Case ERR_WEBSOCKET_HANDSHAKE_TIMEOUT: desc = "WebSocket handshake timed out"
        Case ERR_SEND_FAILED: desc = "Send failed"
        Case ERR_RECV_FAILED: desc = "Receive failed"
        Case ERR_NOT_CONNECTED: desc = "Not connected"
        Case ERR_ALREADY_CONNECTED: desc = "Already connected"
        Case ERR_TLS_ENCRYPT_FAILED: desc = "TLS encryption failed"
        Case ERR_TLS_DECRYPT_FAILED: desc = "TLS decryption failed"
        Case ERR_INVALID_URL: desc = "Invalid URL"
        Case ERR_HANDSHAKE_REJECTED: desc = "WebSocket handshake rejected by server"
        Case ERR_CONNECTION_LOST: desc = "Connection lost"
        Case ERR_INVALID_HANDLE: desc = "Invalid connection handle"
        Case ERR_MAX_CONNECTIONS: desc = "Maximum connections reached"
        Case ERR_PROXY_CONNECT_FAILED: desc = "Proxy connection failed"
        Case ERR_PROXY_AUTH_FAILED: desc = "Proxy authentication failed"
        Case ERR_PROXY_TUNNEL_FAILED: desc = "Proxy tunnel failed"
        Case ERR_INACTIVITY_TIMEOUT: desc = "Inactivity timeout"
        Case ERR_CERT_LOAD_FAILED: desc = "Client certificate load failed"
        Case ERR_CERT_VALIDATE_FAILED: desc = "Server certificate validation failed"
        Case ERR_FRAGMENT_OVERFLOW: desc = "Fragment buffer overflow"
        Case ERR_TLS_RENEGOTIATE: desc = "TLS renegotiation not supported"
        Case Else: desc = "Unknown error (" & errType & ")"
    End Select
    If errCode <> 0 Then desc = desc & " [0x" & hex(errCode) & "]"
    If Len(tech) > 0 Then desc = desc & " - " & tech
    WebSocketGetErrorDescription = desc
End Function

Public Function WebSocketGetPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPendingCount = m_Connections(h).MsgCount
End Function

Public Function WebSocketGetBinaryPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetBinaryPendingCount = m_Connections(h).BinaryCount
End Function

Public Function WebSocketGetQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetQueueCapacity = MSG_QUEUE_SIZE - m_Connections(h).MsgCount
End Function

Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim uptime As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .Stats.ConnectedAt > 0 Then uptime = TickDiff(.Stats.ConnectedAt, GetTickCount()) \ 1000
        WebSocketGetStats = "BytesSent=" & Format(.Stats.BytesSent, "0") & _
            "|BytesReceived=" & Format(.Stats.BytesReceived, "0") & _
            "|MessagesSent=" & .Stats.MessagesSent & _
            "|MessagesReceived=" & .Stats.MessagesReceived & _
            "|UptimeSeconds=" & uptime & _
            "|Queued=" & .MsgCount & _
            "|BinaryQueued=" & .BinaryCount & _
            "|NoDelay=" & IIf(.NoDelay, "1", "0") & _
            "|Proxy=" & IIf(.ProxyEnabled, .proxyHost & ":" & .proxyPort, "none")
    End With
End Function

Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .Connected And .Stats.ConnectedAt > 0 Then
            WebSocketGetUptime = TickDiff(.Stats.ConnectedAt, GetTickCount()) \ 1000
        End If
    End With
End Function

Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h).Stats
        .BytesSent = 0
        .BytesReceived = 0
        .MessagesSent = 0
        .MessagesReceived = 0
        .ConnectedAt = GetTickCount()
    End With
End Sub

Public Function WebSocketGetCloseCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Integer
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetCloseCode = m_Connections(h).closeCode
End Function

Public Function WebSocketGetCloseReason(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetCloseReason = m_Connections(h).closeReason
End Function

Public Function WebSocketGetCloseInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetCloseInfo = "Code=" & .closeCode & _
            "|Description=" & GetCloseCodeDesc(.closeCode) & _
            "|Reason=" & IIf(Len(.closeReason) > 0, .closeReason, "(empty)") & _
            "|InitiatedByUs=" & IIf(.CloseInitiatedByUs, "Yes", "No")
    End With
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
    If Not ValidIndex(handle) Then Exit Function
    If Not m_Connections(handle).Connected Then Exit Function
    m_DefaultHandle = handle
    WebSocketSetDefaultHandle = True
End Function

Public Function WebSocketGetDefaultHandle() As Long
    WebSocketGetDefaultHandle = m_DefaultHandle
End Function

Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .AutoReconnect = enabled
        .ReconnectMaxAttempts = maxAttempts
        .ReconnectBaseDelayMs = baseDelayMs
        .ReconnectAttempts = 0
    End With
End Sub

Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetReconnectInfo = "AutoReconnect=" & IIf(.AutoReconnect, "1", "0") & _
            "|Attempts=" & .ReconnectAttempts & _
            "|MaxAttempts=" & .ReconnectMaxAttempts & _
            "|BaseDelayMs=" & .ReconnectBaseDelayMs
    End With
End Function

Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).PingIntervalMs = intervalMs
    m_Connections(h).LastPingSentAt = GetTickCount()
End Sub

Public Sub WebSocketSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ReceiveTimeoutMs = timeoutMs
End Sub

Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).InactivityTimeoutMs = timeoutMs
    m_Connections(h).LastActivityAt = GetTickCount()
End Sub

Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        .proxyHost = proxyHost
        .proxyPort = proxyPort
        .proxyUser = proxyUser
        .proxyPass = proxyPass
        .proxyType = proxyType
        .ProxyEnabled = (Len(proxyHost) > 0 And proxyPort > 0)
    End With
End Sub

Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
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
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        If .ProxyEnabled Then
            WebSocketGetProxyInfo = "Type=" & IIf(.proxyType = PROXY_TYPE_SOCKS5, "SOCKS5", "HTTP") & _
                "|Host=" & .proxyHost & _
                "|Port=" & .proxyPort & _
                "|Auth=" & IIf(.proxyUser <> "", "Yes", "No")
        Else
            WebSocketGetProxyInfo = "Disabled"
        End If
    End With
End Function

Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).SubProtocol = protocol
End Sub

Public Function WebSocketGetSubProtocol(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetSubProtocol = m_Connections(h).SubProtocol
End Function

Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
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
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).CustomHeaderCount = 0
End Sub

Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).LogCallback = callbackName
End Sub

Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).EnableErrorDialog = enabled
End Sub

Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim optVal As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    m_Connections(h).NoDelay = enabled
    If m_Connections(h).Socket <> INVALID_SOCKET Then
        optVal = IIf(enabled, 1, 0)
        WebSocketSetNoDelay = (sock_setsockopt(m_Connections(h).Socket, IPPROTO_TCP_SOL, TCP_NODELAY, optVal, 4) = 0)
    Else
        WebSocketSetNoDelay = True
    End If
End Function

Public Sub WebSocketSetPreferIPv6(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).PreferIPv6 = enabled
End Sub

Public Sub WebSocketSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ValidateServerCert = enabled
End Sub

Public Sub WebSocketSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).EnableRevocationCheck = enabled
End Sub

Public Sub WebSocketSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ClientCertThumb = thumbprintOrSubject
    m_Connections(h).ClientCertPfxPath = ""
End Sub

Public Sub WebSocketSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ClientCertPfxPath = pfxPath
    m_Connections(h).ClientCertPfxPass = pfxPassword
    m_Connections(h).ClientCertThumb = ""
End Sub

Public Sub WebSocketSetBufferSizes(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    With m_Connections(h)
        If .Connected Then
            WasabiLog h, "Cannot change buffer sizes while connected (handle=" & h & ")"
            Exit Sub
        End If
        If bufferSize >= 8192 And bufferSize <= 16777216 Then
            .CustomBufferSize = bufferSize
        End If
        If fragmentSize >= 4096 And fragmentSize <= 16777216 Then
            .CustomFragmentSize = fragmentSize
        End If
    End With
End Sub

Public Sub WebSocketSetZeroCopy(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ZeroCopyEnabled = enabled
End Sub

Public Sub WebSocketSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If mtu < 576 Or mtu > 9000 Then
        mtu = DEFAULT_MTU
    End If
    m_Connections(h).mtu.CurrentMTU = mtu
    CalculateOptimalFrameSize h
End Sub

Public Sub WebSocketSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).AutoMTU = enabled
End Sub

Public Function WebSocketGetMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetMTU = m_Connections(h).mtu.CurrentMTU
End Function

Public Function WebSocketGetOptimalFrameSize(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetOptimalFrameSize = m_Connections(h).mtu.OptimalFrameSize
End Function

Public Function WebSocketGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    With m_Connections(h)
        WebSocketGetMTUInfo = "MTU=" & .mtu.CurrentMTU & _
            "|MSS=" & .mtu.MaxSegmentSize & _
            "|OptimalFrame=" & .mtu.OptimalFrameSize & _
            "|AutoMTU=" & IIf(.AutoMTU, "Yes", "No") & _
            "|ProbeEnabled=" & IIf(.mtu.ProbeEnabled, "Yes", "No")
    End With
End Function

Public Sub WebSocketProbeMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    If m_Connections(h).Connected Then
        probeMTU h
    End If
End Sub

Public Function WebSocketGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetHost = m_Connections(h).Host
End Function

Public Function WebSocketGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPort = m_Connections(h).port
End Function

Public Function WebSocketGetPath(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetPath = m_Connections(h).path
End Function

Public Sub WebSocketSetHttp2(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).UseHttp2 = enabled
End Sub

Public Sub WebSocketSetProxyNtlm(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Sub
    m_Connections(h).ProxyUseNtlm = enabled
End Sub

Public Function WebSocketGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
    Dim h As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    WebSocketGetLatency = m_Connections(h).LastRttMs
End Function

Public Function MqttConnect(ByVal clientId As String, Optional ByVal username As String = "", Optional ByVal password As String = "", Optional ByVal keepAlive As Integer = 60, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = BuildMqttConnectPacket(clientId, username, password, keepAlive)
    MqttConnect = WebSocketSendBinary(packet, h)
    If MqttConnect Then
        MqttResetParser h
    End If
End Function

Public Function MqttPublish(ByVal topic As String, ByVal message As String, Optional ByVal qos As Byte = 0, Optional ByVal retained As Boolean = False, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte
    Dim msgBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim pos As Long
    Dim flags As Byte
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    topicBytes = StringToUtf8(topic)
    msgBytes = StringToUtf8(message)
    payloadLen = 2 + UBound(topicBytes) + 1 + IIf(qos > 0, 2, 0) + UBound(msgBytes) + 1
    ReDim payload(0 To payloadLen - 1)
    pos = 0
    payload(pos) = CByte((Len(topic) \ 256) And &HFF)
    payload(pos + 1) = CByte(Len(topic) And &HFF)
    pos = pos + 2
    CopyMemory payload(pos), topicBytes(0), UBound(topicBytes) + 1
    pos = pos + UBound(topicBytes) + 1
    If qos > 0 Then
        payload(pos) = 0
        payload(pos + 1) = 1
        pos = pos + 2
    End If
    CopyMemory payload(pos), msgBytes(0), UBound(msgBytes) + 1
    flags = IIf(retained, 1, 0) Or (qos * 8)
    packet = MqttBuildPacket(MQTT_PUBLISH, flags, payload, payloadLen)
    MqttPublish = WebSocketSendBinary(packet, h)
End Function

Public Function MqttSubscribe(ByVal topic As String, Optional ByVal qos As Byte = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    topicBytes = StringToUtf8(topic)
    payloadLen = 2 + 2 + UBound(topicBytes) + 1 + 1
    ReDim payload(0 To payloadLen - 1)
    payload(0) = 0
    payload(1) = 10
    payload(2) = CByte((Len(topic) \ 256) And &HFF)
    payload(3) = CByte(Len(topic) And &HFF)
    CopyMemory payload(4), topicBytes(0), UBound(topicBytes) + 1
    payload(4 + UBound(topicBytes) + 1) = qos
    packet = MqttBuildPacket(MQTT_SUBSCRIBE, 2, payload, payloadLen)
    MqttSubscribe = WebSocketSendBinary(packet, h)
End Function

Public Function MqttUnsubscribe(ByVal topic As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim topicBytes() As Byte
    Dim payload() As Byte
    Dim payloadLen As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    topicBytes = StringToUtf8(topic)
    payloadLen = 2 + 2 + UBound(topicBytes) + 1
    ReDim payload(0 To payloadLen - 1)
    payload(0) = 0
    payload(1) = 10
    payload(2) = CByte((Len(topic) \ 256) And &HFF)
    payload(3) = CByte(Len(topic) And &HFF)
    CopyMemory payload(4), topicBytes(0), UBound(topicBytes) + 1
    packet = MqttBuildPacket(MQTT_UNSUBSCRIBE, 2, payload, payloadLen)
    MqttUnsubscribe = WebSocketSendBinary(packet, h)
End Function

Public Function MqttDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = MqttBuildPacket(MQTT_DISCONNECT, 0, NullByteArray(), 0)
    MqttDisconnect = WebSocketSendBinary(packet, h)
End Function

Public Function MqttPingReq(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
    Dim h As Long
    Dim packet() As Byte
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    packet = MqttBuildPacket(MQTT_PINGREQ, 0, NullByteArray(), 0)
    MqttPingReq = WebSocketSendBinary(packet, h)
End Function

Public Function MqttReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
    Dim h As Long
    Dim data() As Byte
    Dim i As Long
    Dim topicLen As Integer
    Dim topic As String
    Dim msgBytes() As Byte
    Dim msgLen As Long
    Dim flags As Byte
    Dim qos As Long
    Dim packetId As Integer
    Dim skipLen As Long
    h = ResolveHandle(handle)
    If Not ValidIndex(h) Then Exit Function
    If Not m_Connections(h).Connected Then Exit Function
    Do
        WebSocketReceive h
        If WebSocketReceiveBinaryCheck(data, h) Then
            For i = LBound(data) To UBound(data)
                MqttParseByte h, data(i)
            Next i
            If MqttHasPacket(h) Then
                Select Case m_Connections(h).MqttCurrentPacketType
                    Case MQTT_CONNACK
                        MqttResetParser h
                    Case MQTT_PUBLISH
                        flags = m_Connections(h).MqttCurrentFlags
                        qos = (flags And &H6) \ 2
                        topicLen = CInt(m_Connections(h).MqttBuffer(0)) * 256 + CInt(m_Connections(h).MqttBuffer(1))
                        topic = StrConv(m_Connections(h).MqttBuffer(), vbUnicode)
                        topic = Left(topic, topicLen)
                        skipLen = 2 + topicLen
                        If qos > 0 Then
                            packetId = CInt(m_Connections(h).MqttBuffer(skipLen)) * 256 + CInt(m_Connections(h).MqttBuffer(skipLen + 1))
                            skipLen = skipLen + 2
                        Else
                            packetId = 0
                        End If
                        msgLen = m_Connections(h).MqttBufLen - skipLen
                        If msgLen > 0 Then
                            ReDim msgBytes(0 To msgLen - 1)
                            CopyMemory msgBytes(0), m_Connections(h).MqttBuffer(skipLen), msgLen
                            MqttReceive = topic & "|" & StrConv(msgBytes, vbUnicode)
                        Else
                            MqttReceive = topic & "|"
                        End If
                        MqttResetParser h
                        Exit Function
                    Case MQTT_PINGRESP
                        MqttResetParser h
                    Case MQTT_UNSUBACK, MQTT_SUBACK
                        MqttResetParser h
                    Case Else
                        MqttResetParser h
                End Select
            End If
        Else
            Exit Do
        End If
        DoEvents
    Loop
End Function

Private Function NullByteArray() As Byte()
    Dim b() As Byte
    NullByteArray = b
End Function

Private Function DeflatePayload(ByVal handle As Long, ByRef data() As Byte, ByVal dataLen As Long, ByRef outLen As Long) As Byte()
    Dim outBuf()    As Byte
    Dim strm        As ZStream
    Dim ret         As Long
    Dim pIn()       As Byte
    Dim windowBits  As Long

    windowBits = m_Connections(handle).DeflateWindowBits
    If windowBits = 0 Then windowBits = ZLIB_WBITS_RAW

    ReDim outBuf(0 To dataLen + 256)
    ReDim pIn(0 To dataLen - 1)
    CopyMemory pIn(0), data(LBound(data)), dataLen

    If m_Connections(handle).DeflateContextTakeover And m_Connections(handle).DeflateReady Then
        strm = m_Connections(handle).DeflateStream
    Else
        zlib_deflateInit2 strm, Z_DEFAULT_COMPRESSION, Z_DEFLATED, windowBits, ZLIB_MEM_LEVEL, Z_DEFAULT_STRATEGY, ZLIB_VERSION, LenB(strm)
        m_Connections(handle).DeflateReady = True
    End If

    strm.next_in = VarPtr(pIn(0))
    strm.avail_in = dataLen
    strm.next_out = VarPtr(outBuf(0))
    strm.avail_out = UBound(outBuf) + 1

    zlib_deflate strm, Z_SYNC_FLUSH

    outLen = (UBound(outBuf) + 1) - strm.avail_out
    
    If m_Connections(handle).DeflateContextTakeover Then
        m_Connections(handle).DeflateStream = strm
    Else
        zlib_deflateEnd strm
        m_Connections(handle).DeflateReady = False
    End If

    ReDim Preserve outBuf(0 To outLen - 1)
    DeflatePayload = outBuf
End Function

Private Function InflatePayload(ByVal handle As Long, ByRef data() As Byte, ByVal dataLen As Long, ByRef outLen As Long) As Byte()
    Dim inBuf()    As Byte
    Dim outBuf()   As Byte
    Dim strm       As ZStream
    Dim ret        As Long
    Dim windowBits As Long
    Dim trailer(0 To 3) As Byte

    windowBits = m_Connections(handle).InflateWindowBits
    If windowBits = 0 Then windowBits = ZLIB_WBITS_RAW

    ReDim inBuf(0 To dataLen + 3)
    CopyMemory inBuf(0), data(LBound(data)), dataLen
    trailer(0) = &H0: trailer(1) = &H0: trailer(2) = &HFF: trailer(3) = &HFF
    CopyMemory inBuf(dataLen), trailer(0), 4

    ReDim outBuf(0 To dataLen * 8 + 4096)

    If m_Connections(handle).InflateContextTakeover And m_Connections(handle).InflateReady Then
        strm = m_Connections(handle).InflateStream
    Else
        zlib_inflateInit2 strm, windowBits, ZLIB_VERSION, LenB(strm)
        m_Connections(handle).InflateReady = True
    End If

    strm.next_in = VarPtr(inBuf(0))
    strm.avail_in = dataLen + 4
    strm.next_out = VarPtr(outBuf(0))
    strm.avail_out = UBound(outBuf) + 1

    ret = zlib_inflate(strm, Z_SYNC_FLUSH)
    
    If ret <> Z_OK And ret <> Z_STREAM_END Then
        zlib_inflateEnd strm
        m_Connections(handle).InflateReady = False
        outLen = 0
        InflatePayload = NullByteArray()
        Exit Function
    End If

    outLen = (UBound(outBuf) + 1) - strm.avail_out

    If m_Connections(handle).InflateContextTakeover Then
        m_Connections(handle).InflateStream = strm
    Else
        zlib_inflateEnd strm
        m_Connections(handle).InflateReady = False
    End If

    ReDim Preserve outBuf(0 To outLen - 1)
    InflatePayload = outBuf
End Function

Private Sub FreeDeflateStreams(ByVal handle As Long)
    With m_Connections(handle)
        If .DeflateReady Then
            zlib_deflateEnd .DeflateStream
            .DeflateReady = False
        End If
        If .InflateReady Then
            zlib_inflateEnd .InflateStream
            .InflateReady = False
        End If
    End With
End Sub
