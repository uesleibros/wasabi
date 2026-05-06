# Development Suite for Wasabi

> [!NOTE]
> The Wasabi framework and its development tools are currently in an **experimental** phase. Technical specifications are subject to frequent updates.

This directory contains the internal toolset required to build, test, and maintain the Wasabi networking engine. These resources are intended for contributors and are not required for end-users of the production modules.

## Directory Structure

### [/asm](./asm)
Contains low-level assembly source code and compiled binaries for the Wasabi Thunk Library.
* **Core Stability**: Safe Thunks to prevent Excel crashes during VBE Resets.
* **Performance**: Optimized routines for WebSocket masking, endianness swapping, and memory zeroing.

### [/scripts](./scripts)
Automation utilities to bridge the gap between low-level code and the VBA environment.
* **Build Automation**: Batch scripts for NASM compilation.
* **Extraction**: Python utilities to convert binary files into VBA-compatible hexadecimal arrays.

## Prerequisites for Development

To utilize the full development suite, the following tools must be installed and configured in your system PATH:

> [!IMPORTANT]
> * **NASM (Netwide Assembler)**: Required for compiling `.asm` files into raw binaries.
> * **Python 3.x**: Required for execution of automation scripts and mock servers.

## Developer Workflow

1. **Architecture Modification**: Update assembly logic in `/asm` to improve performance or stability.
2. **Compilation**: Run `compile_asm.bat` within `/scripts` to generate updated binaries.
3. **VBA Integration**: Use `bin_to_vba.py` to generate the hex opcodes for injection into the `Wasabi` core.

## Safety and Stability

> [!CAUTION]
> Direct manipulation of memory via Assembly thunks bypasses the standard VBA safety checks. Developers must ensure that all code injected into the `PAGE_EXECUTE_READWRITE` segments is properly audited.

### VBE Reset Protection
The primary mission of the development suite is to maintain the "Safe Thunk" system. This mechanism ensures that even if the VBA project is reset, the Windows message queue does not dispatch events to invalid memory addresses, thereby preventing host application crashes.
