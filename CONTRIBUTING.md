# Contributing to Wasabi

First of all, thank you for your interest in contributing to Wasabi.

Wasabi is a low-level WebSocket and WSS module for VBA built directly on native
Windows APIs. Because of that, even small changes can affect compatibility,
stability, or behavior across different Office and Windows versions.

To keep the project reliable, please read the guidelines below before opening an
issue or pull request.

## Before You Contribute

Please check the following first:

- Search existing issues before opening a new one
- Make sure the problem is reproducible
- Keep in mind that Wasabi supports both 32-bit and 64-bit Office
- Keep in mind that Wasabi supports both VBA6 and VBA7
- Do not introduce external dependencies

## Reporting Bugs

Bug reports are always welcome.

When opening a bug report, please include:

- Office application and version
- 32-bit or 64-bit Office
- Windows version
- The exact Wasabi version or commit
- A minimal reproducible example
- Expected behavior
- Actual behavior
- Any error codes returned by Wasabi
- Any technical details returned by `WebSocketGetTechnicalDetails`

The more precise the report, the easier it is to fix.

## Suggesting Features

Feature requests are welcome, especially when they improve:

- protocol compliance
- compatibility
- stability
- performance
- diagnostics
- real-world usability in Office environments

Please explain:

- what problem the feature solves
- why the current API is not enough
- whether the feature affects compatibility or public behavior

## Pull Requests

Pull requests are welcome, but for **significant changes** please open an issue
first so the change can be discussed before implementation.

This is especially important for:

- changes to the public API
- architectural refactors
- TLS or Schannel logic
- buffer and memory management
- socket I/O behavior
- reconnect logic
- protocol parsing

For **small fixes**, such as typo corrections, documentation improvements, or
isolated bug fixes, you may open a pull request directly.

## Code Style

Please follow these guidelines:

- Use **English** for comments, names, and documentation
- Keep the style consistent with the existing codebase
- Prefer clarity over cleverness
- Avoid unnecessary abstraction
- Do not rename public functions without discussion
- Do not introduce dependencies on external DLLs, COM libraries, or references
- Preserve compatibility with both VBA6/VBA7 and 32-bit/64-bit Office whenever possible

## Compatibility Requirements

Wasabi is designed to be:

- dependency-free
- native to Windows
- compatible with old and modern Office environments
- portable as a single `.bas` module

Any contribution that weakens one of these goals is unlikely to be accepted.

In particular:

- Do not require additional setup steps
- Do not require Tools → References changes
- Do not require installation or registration of external components
- Do not assume only Excel is used; Wasabi should remain usable in other VBA hosts

## Testing

Before submitting a pull request, please test your changes as much as possible.

At minimum, verify that:

- the module compiles without errors
- connection still works over `ws://`
- connection still works over `wss://`
- sending still works
- receiving still works
- disconnect still works
- no obvious regressions were introduced

If your change affects a specific subsystem, test that subsystem directly:

- TLS changes → test secure connections
- proxy changes → test proxy connections
- framing changes → test text and binary frames
- reconnect changes → test connection loss and recovery

## Documentation

If your contribution changes behavior, adds a feature, or modifies the public API,
please update the relevant documentation as part of the same pull request.

This may include:

- `README.md`
- examples
- comments in the module
- `SECURITY.md` if the change is security-related

## Security Issues

If you discover a security vulnerability, **do not open a public GitHub issue**.

Please use **GitHub Private Vulnerability Reporting** instead.

## Maintainer Review

All contributions are reviewed manually.

A pull request may be rejected if it:

- adds unnecessary complexity
- breaks compatibility
- changes behavior without discussion
- reduces stability
- conflicts with the project's design goals

This does not mean the contribution is bad only that it may not fit Wasabi's
scope or priorities.

## Final Note

Wasabi exists to make real-time communication in VBA possible without sacrificing
portability, compatibility, or control.

If your contribution helps move the project in that direction, it is welcome.
