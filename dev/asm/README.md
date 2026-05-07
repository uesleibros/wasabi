# Assembly Thunks for Wasabi

> [!NOTE]
> This system is still in an **experimental** phase, so there may be several problems.

> [!IMPORTANT]
> **Technical Purpose:** These files provide the low-level "Firewall" mechanism that allows Wasabi to use high-performance, event-driven networking without crashing Excel during development.

### The VBE Reset Problem

In standard VBA networking, using `WSAAsyncSelect` (Event-driven IO) is dangerous. When you click the **Reset (Blue Square)** button in the Visual Basic Editor:
1. VBA clears all variables and pointers from memory.
2. The Windows OS, however, still holds the network socket and attempts to send messages to the now-deleted memory address (`AddressOf`).
3. This "Jump to Nowhere" causes an immediate and fatal **Excel Crash**.

### The Safe Thunk Solution

The files in this directory contain the **Assembly (x86/x64)** source for a "Safe Thunk." This thunk acts as an intermediary between Windows and VBA.

1. **Heartbeat Check:** Before forwarding a network event to the VBA code, the Thunk checks a specific memory address (the "Flag").
2. **Safe Routing:** * **If Flag = 1:** The VBA project is active; the Thunk forwards the message to `Wasabi_WndProc`.
   * **If Flag = 0:** The user clicked Reset; the Thunk ignores the VBA code and safely redirects the message to the default Windows handler (`DefWindowProcW`).

### Utility Thunks

Beyond stability, the Wasabi project utilizes specialized thunks to handle data processing tasks that are inefficient in native VBA:

* **WebSocket Masking**: Implements the high-speed XOR bitwise operations required by the WebSocket protocol for all client-to-server data frames.
* **Fast Memory Zero**: A lightweight alternative to `RtlZeroMemory` for clearing buffers or network structures with minimal overhead.
* **High-Speed Memory Search**: An ultra-fast implementation using hardware-level byte comparison (`repe cmpsb`) to find byte patterns (needle in a haystack) within large TCP buffers, bypassing slow VBA loops.
* **Fast Tick Difference**: A micro-optimized subtraction routine that safely handles 32-bit unsigned integer wraparound for precise time calculations (e.g., ping intervals and timeouts), bypassing slow native VBA workarounds.

### Files in this Directory

| File | Architecture | Description |
| :--- | :--- | :--- |
| ![](../../resources/svg/assembly.svg) `safe_thunk_x64.asm` | 64-bit | Core Reset protection using RAX and FastCall convention. |
| ![](../../resources/svg/assembly.svg) `safe_thunk_x86.asm` | 32-bit | Core Reset protection using stack-based arguments and EAX. |
| ![](../../resources/svg/assembly.svg) `ws_mask_x64.asm` | 64-bit | High-speed XOR masking for WebSocket protocol frames. |
| ![](../../resources/svg/assembly.svg) `ws_mask_x86.asm` | 32-bit | High-speed XOR masking for WebSocket protocol frames. |
| ![](../../resources/svg/assembly.svg) `mem_zero_x64.asm` | 64-bit | Optimized memory zeroing for buffers. |
| ![](../../resources/svg/assembly.svg) `mem_zero_x86.asm` | 32-bit | Optimized memory zeroing for buffers. |
| ![](../../resources/svg/assembly.svg) `mem_find_x64.asm` | 64-bit | High-performance memory block search (Needle in a Haystack). |
| ![](../../resources/svg/assembly.svg) `mem_find_x86.asm` | 32-bit | High-performance memory block search (Needle in a Haystack). |
| ![](../../resources/svg/assembly.svg) `tick_diff_x64.asm` | 64-bit | High-performance tick/time difference calculation. |
| ![](../../resources/svg/assembly.svg) `tick_diff_x86.asm` | 32-bit | High-performance tick/time difference calculation. |

### Implementation Details

These thunks are injected into executable memory at runtime using the `VirtualAlloc` API with `PAGE_EXECUTE_READWRITE` permissions. 

#### Memory Lifecycle Management

To prevent memory leaks (Zombies) after a Reset, Wasabi uses **Window Properties (`SetProp`)** to tag the allocated memory addresses on the invisible event window. Upon the next initialization, the system:
1. Scans for existing windows named `WasabiEvents`.
2. Recovers the memory addresses from the previous session.
3. Frees the "Zombie" memory before allocating a fresh Thunk.

### Performance Impact

The Thunk adds approximately **5-10 CPU cycles** of overhead per network event. This is negligible compared to the thousands of cycles saved by eliminating the standard `DoEvents` polling loops. 

## Compilation and Verification

To test and verify the Assembly thunks, you need to assemble the `.asm` source files into raw machine code (binary format). This allows you to extract the exact opcodes used in the `RtlMoveMemory` operations within the VBA module.

### 1. Required Tooling: NASM

The Netwide Assembler (NASM) is the industry standard for this task. It is lightweight and can output raw binary files without the overhead of headers (like PE or ELF).

1. Download the NASM executable from [nasm.us](https://www.nasm.us/).
2. Add the NASM directory to your Windows System PATH or run it directly from the folder.

### 2. Compilation Commands

You must compile these files using the `-f bin` flag. This ensures the output is a pure stream of processor instructions.

#### For the 64-bit Thunk
```bash
nasm -f bin safe_thunk_x64.asm -o safe_thunk_x64.bin
