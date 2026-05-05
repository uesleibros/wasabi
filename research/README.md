# Wasabi Research Archive

This folder contains all reference material, notes, test scripts, and design
decisions used during the development of Wasabi.

It is intended for project maintainers and advanced users who want to
understand the *why* behind implementation choices.

Feel free to browse, but the canonical documentation is in `docs/`.

## Reference URLs by topic

> [!IMPORTANT]
> Not all references mentioned here have been properly documented in this folder, some because the documentation is already sufficient and others because they are too simple and don't require it. Additionally, some information may be outdated with respect to the current version of the project.

### tls-verification
- Microsoft Schannel documentation: https://learn.microsoft.com/en-us/windows/win32/secauthn/sspi-ssl
- Certificate Chain Validation: https://learn.microsoft.com/en-us/windows/win32/seccrypto/certificate-chain-validation
- `CertGetCertificateChain`: https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-certgetcertificatechain
- `CertVerifyCertificateChainPolicy`: https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-certverifycertificatechainpolicy
- `SSL_EXTRA_CERT_CHAIN_POLICY_PARA`: https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/ns-wincrypt-ssl_extra_cert_chain_policy_para
- SSPI status codes: https://learn.microsoft.com/en-us/windows/win32/secauthn/sspi-status-codes
- Schannel constants (GitHub mirror of official docs): https://github.com/MicrosoftDocs/windows-driver-docs
- Certificate revocation checking: https://learn.microsoft.com/en-us/windows/win32/seccrypto/certificate-revocation-status-checking
- Client certificate authentication: https://learn.microsoft.com/en-us/windows/win32/secauthn/client-authentication-certificate

### 32bit-64bit-pointers
- Compiler constants in VBA: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/compiler-constants
- `addrinfo` structure: https://learn.microsoft.com/en-us/windows/win32/api/ws2def/ns-ws2def-addrinfo
- `getaddrinfo`: https://learn.microsoft.com/en-us/windows/win32/api/ws2tcpip/nf-ws2tcpip-getaddrinfo
- `hostent` structure: https://learn.microsoft.com/en-us/windows/win32/api/winsock/ns-winsock-hostent
- `gethostbyname`: https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-gethostbyname
- RFC 8305 "Happy Eyeballs v2": https://www.rfc-editor.org/rfc/rfc8305

### zlib-stdcall
- zlib homepage: https://zlib.net
- Gilles Vollant’s old stdcall zlib builds: http://www.winimage.com/zLibDll/ (HTTP only, frequently broken)
- Joveler.Compression.ZLib NuGet package: https://www.nuget.org/packages/Joveler.Compression.ZLib
- zlib 1.2.11 source code: https://zlib.net/zlib-1.2.11.tar.gz
- zlib manual (windowBits, Z_SYNC_FLUSH): https://zlib.net/manual.html

### permessage-deflate
- RFC 7692 "WebSocket Per-Message Deflate": https://www.rfc-editor.org/rfc/rfc7692
- RFC 6455 "The WebSocket Protocol": https://www.rfc-editor.org/rfc/rfc6455
- Compression Extensions for WebSocket (IETF draft history): https://tools.ietf.org/wg/hybi/

### happy-eyeballs-dual-stack
- RFC 8305 "Happy Eyeballs Version 2": https://www.rfc-editor.org/rfc/rfc8305
- Winsock `select()`: https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-select
- `ioctlsocket` (FIONBIO): https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-ioctlsocket

### path-mtu-discovery
- RFC 1191 "Path MTU Discovery": https://www.rfc-editor.org/rfc/rfc1191
- RFC 4821 "Packetization Layer Path MTU Discovery": https://www.rfc-editor.org/rfc/rfc4821
- `getsockopt` (TCP_MAXSEG): https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-getsockopt
- TCP/IP overhead breakdown: https://en.wikipedia.org/wiki/Maximum_segment_size

### mqtt-over-websocket
- MQTT 3.1.1 specification: https://docs.oasis-open.org/mqtt/mqtt/v3.1.1/os/mqtt-v3.1.1-os.html
- RFC 6455 (WebSocket as transport): https://www.rfc-editor.org/rfc/rfc6455
- MQTT variable-length encoding: https://docs.oasis-open.org/mqtt/mqtt/v3.1.1/errata01/os/mqtt-v3.1.1-errata01-os-complete.html#_Toc442180848

### proxy-http-socks5-ntlm
- RFC 1928 "SOCKS Protocol Version 5": https://www.rfc-editor.org/rfc/rfc1928
- RFC 2617 "HTTP Authentication: Basic and Digest": https://www.rfc-editor.org/rfc/rfc2617
- Microsoft NTLM (SSPI): https://learn.microsoft.com/en-us/windows/win32/secauthn/microsoft-ntlm
- HTTP CONNECT method: https://developer.mozilla.org/en-US/docs/Web/HTTP/Methods/CONNECT
- `InitializeSecurityContext` (NTLM): https://learn.microsoft.com/en-us/windows/win32/api/sspi/nf-sspi-initializesecuritycontexta

### batch-send-optimization
- Nagle’s algorithm: https://en.wikipedia.org/wiki/Nagle%27s_algorithm
- TCP send coalescing: https://learn.microsoft.com/en-us/windows/win32/winsock/performance-considerations-2

### zero-copy-receive
- VBA `StrPtr` documentation: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strptr
- Direct memory access in VBA: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-data-types-efficiently

### tcp-tuning
- `TCP_NODELAY`: https://learn.microsoft.com/en-us/windows/win32/winsock/ipproto-tcp-socket-options
- `SO_KEEPALIVE`: https://learn.microsoft.com/en-us/windows/win32/winsock/so-keepalive
- `SO_RCVBUF` / `SO_SNDBUF`: https://learn.microsoft.com/en-us/windows/win32/winsock/so-rcvbuf-and-so-sndbuf

### fragmentation-buffer
- RFC 6455 (Message fragmentation): https://www.rfc-editor.org/rfc/rfc6455#section-5.4
- WebSocket control frames and fragmentation: https://www.rfc-editor.org/rfc/rfc6455#section-5.5
- Compressed fragmented messages: RFC 7692 Section 5.1.2

### reconnect-backoff
- Exponential backoff and jitter: https://aws.amazon.com/blogs/architecture/exponential-backoff-and-jitter/
- RFC 8305 (Happy Eyeballs also covers reconnect backoff): https://www.rfc-editor.org/rfc/rfc8305

### tls-renegotiation-rejection
- RFC 5746 "TLS Renegotiation Indication Extension": https://www.rfc-editor.org/rfc/rfc5746
- CVE-2009-3555: https://cve.mitre.org/cgi-bin/cvename.cgi?name=CVE-2009-3555

### cryptographic-random
- `CryptGenRandom`: https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptgenrandom
- VBA `Rnd` fallback: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rnd-function

### sha1-from-scratch
- RFC 3174 "US Secure Hash Algorithm 1 (SHA1)": https://www.rfc-editor.org/rfc/rfc3174
- FIPS 180-4 (SHA-1 specification): https://csrc.nist.gov/publications/detail/fips/180/4/final

### base64-encoder
- RFC 4648 "The Base16, Base32 and Base64 Data Encodings": https://www.rfc-editor.org/rfc/rfc4648

### structure-alignment-vba
- Structure alignment in VB/VBA: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/alignment-constants
- `#If` compiler constants: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/compiler-constants

### vba6-compatibility
- VBA7 migration guide: https://learn.microsoft.com/en-us/office/client-developer/access/migration/32-bit-and-64-bit-access
- `Declare PtrSafe`: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/declare-statement

### utf8-handling
- UTF-8 in Windows: https://learn.microsoft.com/en-us/windows/win32/intl/utf-8-support
- `WideCharToMultiByte` / `MultiByteToWideChar`: https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar

### error-handling-philosophy
- VBA error handling: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-error-handling
- Winsock error codes: https://learn.microsoft.com/en-us/windows/win32/winsock/windows-sockets-error-codes-2
