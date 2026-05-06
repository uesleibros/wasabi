# Security Policy

## Supported Versions

| Version | Supported |
|---|---|
| Latest release | ![](resources/svg/thumbsup.svg) |
| Older releases | ![](resources/svg/thumbsdown.svg) |

Only the latest version of Wasabi receives security updates. If you are using
an older version, please update to the latest release before reporting an issue.

## Reporting a Vulnerability

If you discover a security vulnerability in Wasabi, **please do not open a
public GitHub issue**.

Instead, use **GitHub Private Vulnerability Reporting** for this repository.

Please include as much detail as possible:

- Description of the vulnerability
- Steps to reproduce
- Affected version(s)
- Potential impact
- Proof of concept, if available
- Suggested fix, if available

## Scope

The following areas are in scope for security reports:

- TLS and Schannel handling
- WebSocket frame parsing
- Buffer and memory safety
- Input validation
- Proxy handling and authentication
- DNS resolution and connection logic
- Internal SHA-1 handshake validation

The following are out of scope:

- Vulnerabilities in Windows itself
- Vulnerabilities in the Office or VBA runtime
- Social engineering attacks
- Misconfiguration of third-party servers

## Disclosure Policy

Wasabi follows a coordinated disclosure model.

Security issues will be investigated privately first. Once a fix is available,
the vulnerability may be disclosed publicly in the release notes or changelog.

## Security Notes

Wasabi uses native Windows APIs such as `ws2_32.dll`, `secur32.dll`, and
`kernel32.dll`. Security-related behavior may therefore depend partly on the
host Windows version and configuration.

The internal SHA-1 implementation is used only for the RFC 6455 WebSocket
handshake (`Sec-WebSocket-Accept`) and not for encryption or signing.

TLS encryption is handled by Schannel through Windows SSPI.
