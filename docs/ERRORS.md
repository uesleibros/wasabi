# Error Reference

This document describes every error code in the `WasabiError` enumeration,
including what triggers each error, what the associated system codes mean,
and how to diagnose and resolve common failures.

## Reading Wasabi Errors

Wasabi exposes three levels of error information per connection.

```vb
Dim errType As WasabiError
Dim sysCode As Long
Dim details As String

errType = WebSocketGetLastError(handle)
sysCode = WebSocketGetLastErrorCode(handle)
details = WebSocketGetTechnicalDetails(handle)

Debug.Print "Error:", errType
Debug.Print "System code:", sysCode
Debug.Print "Details:", details
```

**errType** is the high-level Wasabi error category. It tells you what failed.

**sysCode** is the raw error code from the underlying system call. For Winsock
errors, this is a WSA error code. For TLS errors, this is an SSPI/HRESULT
status code. For successful operations, this is zero.

**details** is a human-readable string describing the specific failure, including
the function name and the numeric code. This is the most useful field for
debugging.

> [!TIP]
> Always log all three values when diagnosing connection issues. The
> `WebSocketGetErrorDescription` function combines them into a single
> diagnostic string.

## Error Pattern

A typical error handling pattern looks like this.

```vb
Sub ConnectSafely()
    Dim h As Long

    If Not WebSocketConnect("wss://echo.websocket.org", h) Then
        Select Case WebSocketGetLastError(h)
            Case ERR_DNS_RESOLVE_FAILED
                Debug.Print "Cannot resolve hostname"
            Case ERR_CONNECT_FAILED
                Debug.Print "Server unreachable"
            Case ERR_TLS_HANDSHAKE_FAILED
                Debug.Print "TLS negotiation failed"
            Case ERR_HANDSHAKE_REJECTED
                Debug.Print "Server rejected WebSocket upgrade"
            Case ERR_PROXY_AUTH_FAILED
                Debug.Print "Proxy credentials rejected"
            Case ERR_MAX_CONNECTIONS
                Debug.Print "Connection pool full"
            Case Else
                Debug.Print "Unexpected error:", WebSocketGetLastError(h)
        End Select

        Debug.Print "System code:", WebSocketGetLastErrorCode(h)
        Debug.Print "Details:", WebSocketGetTechnicalDetails(h)
        Exit Sub
    End If

    Debug.Print "Connected successfully"
End Sub
```

## Complete Error Reference

### ERR_NONE (0)

No error occurred. The operation completed successfully.

### ERR_WSA_STARTUP_FAILED (1)

**What happened:** `WSAStartup` returned a non-zero value. The Winsock
subsystem could not be initialized.

**Common causes:**
- Corrupted Winsock installation
- Antivirus or security software blocking socket initialization
- Extremely rare on modern Windows

**What to check:**
- Run `netsh winsock reset` from an elevated command prompt and reboot
- Verify no security software is interfering with network initialization

```vb
' Example error output
' Error: 1
' System code: 10091
' Details: WSAStartup failed with code 10091
```

### ERR_SOCKET_CREATE_FAILED (2)

**What happened:** The `socket()` function returned `INVALID_SOCKET`. A TCP
socket could not be allocated.

**Common causes:**
- System ran out of socket handles
- Firewall blocking socket creation at the OS level
- VPN software intercepting socket calls

**What to check:**
- Close unused network connections
- Verify firewall and VPN configuration
- Check `WebSocketGetConnectionCount` to see if the pool is near capacity

```vb
' Example error output
' Error: 2
' System code: 10024
' Details: socket() failed with WSA error 10024
```

> [!NOTE]
> WSA error 10024 means `WSAEMFILE` (too many open sockets).

### ERR_DNS_RESOLVE_FAILED (3)

**What happened:** `gethostbyname()` could not resolve the hostname to an IP
address.

**Common causes:**
- Hostname is misspelled
- DNS server is unreachable
- Corporate proxy requires direct IP or different DNS
- Machine has no internet connectivity

**System codes and their meaning:**

| WSA Code | Name | Meaning |
|:---|:---|:---|
| 11001 | WSAHOST_NOT_FOUND | The hostname does not exist in DNS |
| 11002 | WSATRY_AGAIN | Temporary DNS failure, try again later |
| 11003 | WSANO_RECOVERY | Non-recoverable DNS error |
| 11004 | WSANO_DATA | Hostname exists but has no IP address record |

**What to check:**
- Verify the hostname is correct
- Try pinging the hostname from a command prompt
- Check if a proxy is required for external access
- Try using an IP address directly to isolate DNS issues

```vb
' Example error output
' Error: 3
' System code: 11001
' Details: gethostbyname() failed for 'bad.hostname.example' with WSAHOST_NOT_FOUND (11001)
```

> [!TIP]
> If you see error 11002, wait a moment and try again. This is typically a transient DNS issue that resolves itself.

### ERR_CONNECT_FAILED (4)

**What happened:** The TCP connection could not be established. Either
`connect()` failed immediately or `select()` timed out waiting for the
connection to complete. This includes failures from the Happy Eyeballs
dual‑stack connection process.

**Common causes:**
- Server is down or not listening on the specified port
- Firewall blocking outbound connections on the target port
- Corporate proxy required but not configured
- Wrong port number

**What to check:**
- Verify the server is running and accepting connections
- Try connecting to the same host and port with a browser or telnet
- Check if a proxy is needed via `WebSocketSetProxy`
- Verify the port is correct (80 for `ws://`, 443 for `wss://`)

```vb
' Example error output
' Error: 4
' System code: 10060
' Details: Connect failed: WSAETIMEDOUT - Connection timed out
```

> [!NOTE]
> WSA error 10060 means `WSAETIMEDOUT`. WSA error 10061 means `WSAECONNREFUSED` (server actively refused the connection).

### ERR_TLS_ACQUIRE_CREDS_FAILED (5)

**What happened:** `AcquireCredentialsHandle` could not initialize the
Schannel security provider.

**Common causes:**
- Schannel provider is disabled or corrupted in the Windows registry
- System-level security policy is blocking TLS
- Extremely rare on properly configured systems

**What to check:**
- Verify TLS is enabled in Windows Internet Options
- Check the Windows Event Log for Schannel errors
- Ensure the system clock is correct (certificate validation depends on it)

```vb
' Example error output
' Error: 5
' System code: -2146893043
' Details: AcquireCredentialsHandle failed: 0x8009030D
```

### ERR_TLS_HANDSHAKE_FAILED (6)

**What happened:** `InitializeSecurityContext` returned a fatal error during
the TLS handshake.

**Common causes:**
- Server does not support TLS 1.2 or 1.3
- Server requires a cipher suite that Schannel does not offer
- Server certificate is expired or untrusted
- Network device intercepting and modifying TLS traffic (SSL inspection)

**What to check:**
- Test the server with an external tool like `openssl s_client`
- Check the server's supported TLS versions and cipher suites
- Verify no corporate SSL inspection proxy is interfering
- Check the system clock

```vb
' Example error output
' Error: 6
' System code: -2146893018
' Details: TLS handshake failed: 0x80090326
```

> [!WARNING]
> SSPI error `0x80090326` means `SEC_E_ILLEGAL_MESSAGE`. This often indicates a middlebox (corporate proxy or firewall) is intercepting and corrupting TLS traffic.

### ERR_TLS_HANDSHAKE_TIMEOUT (7)

**What happened:** The TLS handshake did not complete within the allowed
time or iteration limit.

**Common causes:**
- Server is extremely slow to respond
- Network latency is very high
- Server accepted the TCP connection but is not responding to TLS
- Firewall is silently dropping TLS packets

**What to check:**
- Increase the receive timeout via `WebSocketSetReceiveTimeout`
- Test connectivity to the server from a browser
- Check if the server is behind a load balancer that accepted TCP but is not routing TLS

```vb
' Example error output
' Error: 7
' System code: 0
' Details: TLS handshake timed out with api.example.com
```

### ERR_WEBSOCKET_HANDSHAKE_FAILED (8)

**What happened:** Wasabi could not send or receive the HTTP upgrade request
that initiates the WebSocket connection.

**Common causes:**
- TLS was established but the HTTP request failed to send
- Server closed the connection before responding
- Network interruption between TLS completion and HTTP exchange

**What to check:**
- Verify the URL path is correct
- Check server logs for rejected requests
- Ensure custom headers are not malformed

```vb
' Example error output
' Error: 8
' System code: 10054
' Details: recv() WS handshake failed: WSAECONNRESET - Connection reset by peer
```

> [!NOTE]
> WSA error 10054 means `WSAECONNRESET` (connection reset by peer).

### ERR_WEBSOCKET_HANDSHAKE_TIMEOUT (9)

**What happened:** The server did not respond to the HTTP upgrade request
within the timeout period.

**Common causes:**
- Server is overloaded
- Wrong endpoint (server exists but does not handle WebSocket)
- Proxy or firewall silently consuming the upgrade request

**What to check:**
- Verify the URL path supports WebSocket
- Increase the receive timeout
- Check if the server requires specific headers or subprotocol

```vb
' Example error output
' Error: 9
' System code: 0
' Details: No WS handshake response within 5s
```

### ERR_SEND_FAILED (10)

**What happened:** A `send()` call returned zero or negative after attempting
to write data to the socket.

**Common causes:**
- Server closed the connection unexpectedly
- Network cable disconnected
- VPN dropped

**What to check:**
- Check `WebSocketIsConnected` before sending
- Enable auto reconnect for resilient applications
- Log the technical details for the specific WSA error

```vb
' Example error output
' Error: 10
' System code: 10054
' Details: send() failed: WSAECONNRESET - Connection reset by peer
```

### ERR_RECV_FAILED (11)

**What happened:** A `recv()` call returned a negative value.

**Common causes:**
- Same as `ERR_SEND_FAILED`
- Server forcibly closed the connection
- OS-level socket error

**What to check:**
- Same diagnostics as send failure
- Check if the server has connection duration limits

### ERR_NOT_CONNECTED (12)

**What happened:** A send operation was attempted on a handle that is not
currently connected.

**Common causes:**
- Connection was never established
- Connection was already closed or lost
- Wrong handle was passed

**What to check:**
- Verify the return value of `WebSocketConnect` before sending
- Check `WebSocketIsConnected` before each send in long-running loops
- Verify you are using the correct handle

```vb
' Safe send pattern
If WebSocketIsConnected(h) Then
    WebSocketSend "data", h
Else
    Debug.Print "Not connected"
End If
```

### ERR_ALREADY_CONNECTED (13)

Reserved for future use.

### ERR_TLS_ENCRYPT_FAILED (14)

**What happened:** `EncryptMessage` returned a non-zero SSPI status code.
The outgoing data could not be encrypted.

**Common causes:**
- TLS context was invalidated
- Internal state corruption after a partial send
- Extremely rare in normal operation

**What to check:**
- Disconnect and reconnect
- Log the SSPI error code from `WebSocketGetLastErrorCode`

### ERR_TLS_DECRYPT_FAILED (15)

**What happened:** `DecryptMessage` returned a fatal error other than
`SEC_I_RENEGOTIATE` (which is handled separately by `ERR_TLS_RENEGOTIATE`).

**Common causes:**
- Corrupted TLS record received
- Network device modifying encrypted traffic
- Internal SSPI failure

**What to check:**
- Disconnect and reconnect
- Check for SSL inspection proxies

### ERR_INVALID_URL (16)

**What happened:** The URL could not be parsed.

**Common causes:**
- URL does not start with `ws://` or `wss://`
- Empty URL string
- Missing hostname
- Port number out of range (must be 1-65535)
- Non-numeric characters in port

**What to check:**
- Verify the URL format: `ws://host/path` or `wss://host:port/path`

```vb
' Valid URLs
WebSocketConnect "ws://localhost/chat"
WebSocketConnect "wss://api.example.com/ws"
WebSocketConnect "wss://api.example.com:8443/stream"

' Invalid URLs
WebSocketConnect "http://example.com"       ' wrong scheme
WebSocketConnect "wss://"                    ' missing host
WebSocketConnect "wss://host:abc/path"       ' non-numeric port
WebSocketConnect "wss://host:99999/path"     ' port out of range
```

### ERR_HANDSHAKE_REJECTED (17)

**What happened:** The server responded to the HTTP upgrade request with a
status code other than 101, or the `Sec-WebSocket-Accept` header value did
not match the expected SHA-1 hash.

**Common causes:**
- Server returned 403 (forbidden) or 401 (unauthorized)
- Server returned 404 (wrong path)
- Server does not support WebSocket on this endpoint
- Missing required authentication headers
- Load balancer or CDN intercepting the upgrade

**What to check:**
- Verify the URL path is a WebSocket endpoint
- Add authentication headers via `WebSocketAddHeader` if required
- Check the technical details for the server's actual response line
- Test the endpoint with a browser-based WebSocket client

```vb
' Example error output
' Error: 17
' System code: 0
' Details: WebSocket upgrade rejected. Server response: HTTP/1.1 403 Forbidden
```

> [!TIP]
> The technical details string includes the server's HTTP status line, which is often enough to diagnose the issue without external tools.

### ERR_CONNECTION_LOST (18)

**What happened:** The connection was lost during normal operation. This
can be triggered by `recv()` returning zero (clean server close),
`ioctlsocket` failing, or an oversized frame being received.

**Common causes:**
- Server closed the connection normally
- Network interruption
- Server sent a frame larger than the configured buffer size
- Idle connection timed out on the server side

**What to check:**
- Enable auto reconnect for resilient applications
- Check if the server has idle timeout settings
- Use `WebSocketSetPingInterval` to keep the connection alive
- If the error mentions "oversized frame", increase buffer sizes via
  `WebSocketSetBufferSizes`

```vb
' Resilient connection pattern
WebSocketSetAutoReconnect True, 10, 2000, h
WebSocketSetPingInterval 25000, h
```

### ERR_INVALID_HANDLE (19)

**What happened:** The handle passed to a function is outside the valid
range (0 to 63).

**Common causes:**
- Using an uninitialized handle variable
- Using a handle after it was cleaned up
- Arithmetic error producing an out-of-range value

**What to check:**
- Verify that `WebSocketConnect` returned `True` before using the handle
- Do not reuse handles after `WebSocketDisconnect`

### ERR_MAX_CONNECTIONS (20)

**What happened:** All 64 slots in the connection pool are occupied.

**Common causes:**
- Opening connections without closing them
- Leaked handles from failed error handling paths
- Genuinely needing more than 64 simultaneous connections

**What to check:**
- Call `WebSocketDisconnect` on handles you no longer need
- Use `WebSocketGetConnectionCount` to monitor pool usage
- Use `WebSocketGetAllHandles` to find and audit active connections

```vb
' Audit active connections
Dim handles() As Long
Dim i As Long

handles = WebSocketGetAllHandles()

Debug.Print "Active connections:", WebSocketGetConnectionCount()
For i = LBound(handles) To UBound(handles)
    Debug.Print "Handle", handles(i), _
                "Host:", WebSocketGetHost(handles(i)), _
                "Uptime:", WebSocketGetUptime(handles(i)), "s"
Next i
```

### ERR_PROXY_CONNECT_FAILED (21)

**What happened:** The HTTP CONNECT request to the proxy server failed.
Either the `send()` to the proxy failed, the proxy did not respond, or the
proxy response could not be read.

**Common causes:**
- Wrong proxy host or port
- Proxy server is down
- Firewall blocking the proxy port

**What to check:**
- Verify proxy host and port with `WebSocketGetProxyInfo`
- Test proxy connectivity independently
- Check if the proxy requires authentication

### ERR_PROXY_AUTH_FAILED (22)

**What happened:** The proxy returned HTTP 407 (Proxy Authentication
Required).

**Common causes:**
- Wrong proxy username or password
- Proxy requires a different authentication scheme (NTLM, Kerberos)
- Proxy credentials expired

**What to check:**
- Verify credentials in `WebSocketSetProxy`
- Check with your network administrator for correct credentials
- Note that Wasabi only supports HTTP Basic proxy authentication

> [!WARNING]
> Wasabi uses HTTP Basic authentication for proxies, which sends credentials in Base64 encoding. If your proxy requires NTLM or Kerberos authentication, Wasabi cannot authenticate with it in the current version.

### ERR_PROXY_TUNNEL_FAILED (23)

**What happened:** The proxy accepted the connection but returned a non-200
status for the CONNECT tunnel request.

**Common causes:**
- Proxy policy blocks the target host or port
- Proxy does not allow CONNECT to non-443 ports
- Target hostname is blacklisted by the proxy

**What to check:**
- Verify the target host and port are allowed through the proxy
- Check the technical details for the proxy's HTTP status line
- Contact your network administrator if the proxy blocks WebSocket traffic

```vb
' Example error output
' Error: 23
' System code: 0
' Details: Proxy CONNECT rejected: HTTP/1.1 403 Forbidden
```

### ERR_INACTIVITY_TIMEOUT (24)

**What happened:** No data was received from the server within the
configured inactivity timeout period.

**Common causes:**
- Server stopped sending data
- Network interruption that did not fully close the socket
- Inactivity timeout is set too short for the application protocol
- Server expects the client to send periodic messages to stay alive

**What to check:**
- Increase the timeout via `WebSocketSetInactivityTimeout`
- Enable heartbeat via `WebSocketSetPingInterval`
- Check if the server requires periodic client messages
- Enable auto reconnect for automatic recovery

```vb
' Recommended resilient configuration
WebSocketSetInactivityTimeout 60000, h
WebSocketSetPingInterval 25000, h
WebSocketSetAutoReconnect True, 5, 2000, h
```

> [!TIP]
> Combining `WebSocketSetInactivityTimeout` with `WebSocketSetPingInterval` and `WebSocketSetAutoReconnect` creates the most resilient connection configuration available in Wasabi.

### ERR_CERT_LOAD_FAILED (25)

**What happened:** Wasabi failed to load a client certificate from a PFX file
or the Windows certificate store. The certificate was configured via
`WebSocketSetClientCert` or `WebSocketSetClientCertPfx` but could not be
found, imported, or parsed.

**Common causes:**
- Path to the PFX file is incorrect or the file is missing
- PFX file is empty or password-protected with the wrong password
- The specified subject or thumbprint does not match any certificate in the store
- The user account lacks permission to read the certificate store or file

**What to check:**
- Verify the file path and that the PFX file exists
- Confirm the PFX password is correct
- When using `WebSocketSetClientCert`, use a valid certificate subject or thumbprint
- Check that the certificate is installed in the correct store (Current User\My)

```vb
' Example error output
' Error: 25
' System code: 0
' Details: PFX file not found: C:\certs\client.pfx
```

> [!NOTE]
> If client certificate loading fails, Wasabi will continue the connection without a client certificate and log a warning. The server may then reject the TLS handshake if mTLS is required.

### ERR_CERT_VALIDATE_FAILED (26)

**What happened:** Server certificate validation was enabled
(`WebSocketSetCertValidation True`) and the chain verification failed.

**Common causes:**
- The server certificate is self-signed or issued by an untrusted CA
- The certificate has expired or is not yet valid
- The certificate's Common Name (CN) does not match the hostname
- A required intermediate CA certificate is missing on the client machine

**What to check:**
- Verify that the server certificate is trusted by opening the URL in a browser
- If using a self-signed certificate, disable validation or add the certificate to the Trusted Root store
- Ensure the system clock is correct

### ERR_FRAGMENT_OVERFLOW (27)

**What happened:** A fragmented WebSocket message grew larger than the
configured `FragmentBuffer` size.

The received continuation frames accumulated a payload that exceeded the
buffer capacity. The connection is closed to prevent memory corruption.

**Common causes:**
- Server is sending messages larger than the default 256 KB fragment buffer
- A sender is sending continuous fragments without a final FIN frame
- Buffer size is too small for the expected message size

**What to check:**
- Increase the fragment buffer size via `WebSocketSetBufferSizes` before connecting
- If the sender is an API, verify the maximum message size it may send
- Ensure the sender properly terminates fragmented messages

```vb
' Increasing the fragment buffer to 1 MB
WebSocketSetBufferSizes 262144, 1048576, h
```

### ERR_TLS_RENEGOTIATE (28)

**What happened:** The server requested a TLS renegotiation after the initial
handshake was complete. Wasabi does not support renegotiation and closes the
connection.

**Common causes:**
- The server is configured to require periodic re-authentication
- A security policy triggers renegotiation after a certain amount of data is transferred
- Some older servers may renegotiate by default

**What to check:**
- If possible, disable server-initiated TLS renegotiation
- Enable auto reconnect so the connection is automatically re-established
- Adjust server configuration to avoid renegotiation

> [!NOTE]
> Wasabi intentionally does not implement TLS renegotiation due to the complexity of handling it correctly in single-threaded VBA. Auto reconnect is the recommended recovery mechanism.

## Quick Diagnostic Checklist

When a connection fails, run through this list in order.

| Step | Check | How |
|:---|:---|:---|
| 1 | Is the URL valid? | Verify scheme, host, port, and path |
| 2 | Can you reach the host? | Ping the hostname from a command prompt |
| 3 | Is DNS resolving? | Try using an IP address instead of hostname |
| 4 | Is a proxy required? | Check with your network administrator |
| 5 | Is TLS the issue? | Try `ws://` instead of `wss://` to isolate |
| 6 | Is the path correct? | Verify the WebSocket endpoint with a browser tool |
| 7 | Are headers required? | Check if the server needs `Authorization` or other headers |
| 8 | What did the server say? | Read `WebSocketGetTechnicalDetails` for the server response |
| 9 | Is the pool full? | Check `WebSocketGetConnectionCount` |
| 10 | Is auto reconnect working? | Check `WebSocketGetReconnectInfo` |

## Common WSA Error Codes

These are the most frequently encountered Winsock error codes in Wasabi
diagnostic output.

| Code | Name | Meaning |
|:---|:---|:---|
| 10035 | WSAEWOULDBLOCK | Operation would block (normal for non-blocking sockets) |
| 10038 | WSAENOTSOCK | Socket handle is not valid |
| 10053 | WSAECONNABORTED | Connection aborted by local software |
| 10054 | WSAECONNRESET | Connection reset by remote host |
| 10060 | WSAETIMEDOUT | Connection timed out |
| 10061 | WSAECONNREFUSED | Connection actively refused by target |
| 11001 | WSAHOST_NOT_FOUND | Hostname does not exist |
| 11002 | WSATRY_AGAIN | Temporary DNS failure |
| 11003 | WSANO_RECOVERY | Non-recoverable DNS error |
| 11004 | WSANO_DATA | Hostname valid but no IP address available |

## Common SSPI Error Codes

These are the most frequently encountered Schannel/SSPI error codes.

| Code (hex) | Name | Meaning |
|:---|:---|:---|
| 0x80090300 | SEC_E_INSUFFICIENT_MEMORY | Not enough memory for security operation |
| 0x80090304 | SEC_E_INTERNAL_ERROR | Internal Schannel error |
| 0x80090305 | SEC_E_NOT_OWNER | Caller does not own the credentials |
| 0x8009030D | SEC_E_UNKNOWN_CREDENTIALS | Credentials not recognized |
| 0x80090311 | SEC_E_NO_AUTHENTICATING_AUTHORITY | No authority could be contacted for authentication |
| 0x80090318 | SEC_E_INCOMPLETE_MESSAGE | Received TLS record is incomplete (internal, handled by Wasabi) |
| 0x80090326 | SEC_E_ILLEGAL_MESSAGE | Received message is corrupted or unexpected |
| 0x00090312 | SEC_I_CONTINUE_NEEDED | Handshake needs more data (internal, handled by Wasabi) |
| 0x00090321 | SEC_I_RENEGOTIATE | Server requested TLS renegotiation (now handled as `ERR_TLS_RENEGOTIATE`) |
