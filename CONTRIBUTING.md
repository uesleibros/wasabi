# Contributing to Wasabi

Thank you for taking the time to contribute. Wasabi is a low-level networking module built entirely on raw Windows APIs, which means even small changes can affect stability, compatibility, or wire-level behavior across different Office and Windows versions. These guidelines exist to keep things reliable for everyone using it.

## Before You Open Anything

- Search existing issues before opening a new one. Chances are someone hit the same thing.
- Make sure the problem is reproducible and is not caused by the server or network you are connecting to.
- Remember that Wasabi supports both 32-bit and 64-bit Office, and both VBA6 and VBA7. Something that only manifests in one configuration still matters.
- Do not introduce external dependencies. One `.bas` file with no setup is the whole point.

## Reporting Bugs

Bug reports are always welcome. A good bug report saves a lot of back-and-forth.

Please include:

- Office application and version (e.g. Excel 2016, Excel 365)
- 32-bit or 64-bit Office
- Windows version
- Exact Wasabi version or commit hash
- A minimal reproducible example — the smaller the better
- Expected behavior
- Actual behavior
- Any error code returned by `WebSocketGetLastError` or `TcpGetLastError`
- Output of `WebSocketGetTechnicalDetails` or `TcpGetTechnicalDetails` if available

The more precise the report, the faster it gets fixed.

## Suggesting Features

Feature requests are welcome. When suggesting one, please explain:

- What problem it solves that the current API does not address
- Why the existing functions are insufficient
- Whether it would affect compatibility, public behavior, or the module's dependency-free nature

Features with clear real-world use cases in Office environments get prioritized.

## Pull Requests

For significant changes, please open an issue first so the direction can be discussed before you invest time implementing it. This matters most for:

- Changes to the public API surface
- Architectural refactors
- TLS or Schannel logic
- Buffer and memory management
- Socket I/O behavior
- Frame parsing (WebSocket, TCP framing, MQTT)
- Reconnect and reliability logic
- The async thunk or window subclassing mechanism

For small and isolated changes — typo fixes, documentation improvements, comment corrections, straightforward bug fixes — you can open a pull request directly.

## Code Style

- Write comments, names, and documentation in English
- Keep the style consistent with the existing code
- Prefer clarity over cleverness. This module will be read by people debugging production issues at 2am.
- Avoid unnecessary abstraction layers
- Do not rename public functions without prior discussion — existing users will break silently
- Do not introduce dependencies on external DLLs, COM libraries, or Tools → References
- Preserve VBA6/VBA7 and 32-bit/64-bit compatibility using `#If VBA7 Then` and `#If Win64 Then` conditional blocks, consistent with the existing patterns in the module

## Compatibility Requirements

Wasabi is designed to be:

- A single `.bas` file with no setup
- Dependency-free and native to Windows
- Compatible with old and modern Office environments, not just Excel 365
- Usable in any VBA host (Excel, Word, Access, Outlook, and so on), not just Excel

Anything that weakens one of these properties is unlikely to be accepted regardless of how useful the feature itself is.

Concretely, this means:

- No additional setup steps beyond importing the file
- No Tools → References changes
- No DLL registration or installation
- No assumption that only Excel is available

## Testing

Before submitting a pull request, verify at minimum that:

- The module compiles without errors in both 32-bit and 64-bit Office if you have access to both
- `ws://` connections still work
- `wss://` connections still work
- Sending and receiving text and binary messages still work
- Disconnect and cleanup still work without leaving orphaned sockets or memory

If your change touches a specific subsystem, test it directly:

- TLS or Schannel changes → test secure connections end to end
- Proxy changes → test HTTP CONNECT and SOCKS5 proxy paths
- Frame parsing changes → test text frames, binary frames, fragmented messages, and control frames (ping/pong/close)
- Reconnect changes → test connection loss and recovery under auto-reconnect
- Async changes → test that the handler fires correctly and that cleanup does not crash the host on disconnect and on VBE reset

## Documentation

If your contribution changes behavior, adds a function, or modifies anything in the public API, please update the relevant documentation in the same pull request:

- `README.md` if the change affects setup or high-level usage
- `docs/API_REFERENCE.md` for any new or modified public functions
- Inline comments in the module if the change is non-obvious
- `SECURITY.md` if the change touches anything security-related

## Security Issues

If you find a security vulnerability, **do not open a public GitHub issue**.

Use **GitHub Private Vulnerability Reporting** for this repository instead. Details on what to include are in [SECURITY.md](SECURITY.md).

## What Gets Rejected and Why

A pull request may be closed if it:

- Adds complexity without a proportional benefit
- Breaks compatibility with any supported configuration
- Changes public behavior without prior discussion
- Reduces portability or introduces new setup requirements
- Conflicts with the project's design goals

This is not a judgment on the quality of the work — it may simply not fit where Wasabi is going.

## A Note on the Project

Wasabi exists because native real-time networking in VBA used to mean either shelling out to a DLL, pulling in a heavyweight COM reference, or giving up entirely. The goal is to make WebSocket, TCP, TLS, and MQTT available in any VBA project on any Windows machine with nothing more than a file import.

If your contribution moves the project in that direction, it belongs here.
