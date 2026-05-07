# Assembly Thunks & Test Payloads

> [!NOTE]
> This directory (`dev/asm/tests`) contains a suite of test payloads designed to verify memory allocation, parameter passing, and raw machine code execution across different architectures within VBA.

> [!IMPORTANT]
> **Technical Purpose:** Calling raw machine code from VBA requires tricking the host environment (Excel/Access/Word) using APIs like `CallWindowProcW`. These test files verify that the bridging mechanism (the "Thunk") handles pointer arithmetic, CPU registers, and stack cleanup correctly without crashing the host application.

### The VBA Assembly Execution Concept

In standard VBA, executing raw machine code natively is not supported. This limitation is typically bypassed by allocating executable memory (`VirtualAlloc` on Windows or `mmap` on macOS), injecting compiled binary opcodes, and using `CallWindowProcW` as a bridge to redirect execution flow.

However, different CPU architectures handle parameters differently. If these conventions are violated, an immediate and fatal **Host Crash** occurs.

1. **x64 (Microsoft x64 Calling Convention):** Parameters are passed via hardware registers (`RCX`, `RDX`, `R8`, `R9`). The result is returned in `RAX`.
2. **x86 (stdcall):** Parameters are pushed to the stack (`[ebp+8]`, `[ebp+12]`, etc.). The assembly payload is strictly responsible for cleaning up the stack before returning (e.g., `ret 16`). The result is returned in `EAX`.

These test cases are explicitly written to adhere to these strict calling conventions, providing a safe sandbox for low-level development.

### The Test Payloads

Beyond basic stability, this suite covers various CPU operations, memory manipulations, and logic loops to handle high-performance tasks directly in machine code:

*   **Math Addition (`add_numbers`)**: Tests standard parameter passing and integer return values.
*   **Memory Copy (`mem_copy`)**: A high-speed `memcpy` equivalent to test pointer manipulation and the `rep movsb` instruction.
*   **XOR Buffer (`xor_buffer`)**: Tests bitwise cryptography, loops, and byte-level manipulation within an allocated VBA byte array.
*   **String Length (`str_len`)**: Tests high-speed memory scanning (`repne scasb`) by finding the null-terminator of a string.
*   **Memory Set (`mem_set`)**: A `memset` equivalent to test rapid byte filling (`rep stosb`).
*   **Count Byte (`count_byte`)**: Tests conditional jumps (`cmp`, `jne`) by counting specific byte occurrences in a buffer.
*   **Memory Search (`mem_chr`)**: A `memchr` equivalent that tests searching for a specific needle (1 byte) in a haystack buffer and returning its exact memory address.
*   **Memory Reverse (`mem_reverse`)**: Tests complex two-pointer manipulation (reading from the left and right simultaneously) to reverse an array in-place.

### Files in this Directory

| File | Architecture | Description |
| :--- | :--- | :--- |
| ![](../../../resources/svg/assembly.svg) `add_x64.asm` / `add_x86.asm` | 64-bit / 32-bit | Simple math addition to test `RCX`/`RDX` and stack arguments. |
| ![](../../../resources/svg/assembly.svg) `memcpy_x64.asm` / `memcpy_x86.asm` | 64-bit / 32-bit | Fast block memory copy (`rep movsb`). |
| ![](../../../resources/svg/assembly.svg) `xor_x64.asm` / `xor_x86.asm` | 64-bit / 32-bit | In-place XOR bitwise operations for buffers. |
| ![](../../../resources/svg/assembly.svg) `strlen_x64.asm` / `strlen_x86.asm` | 64-bit / 32-bit | Fast null-terminated string length calculation. |
| ![](../../../resources/svg/assembly.svg) `memset_x64.asm` / `memset_x86.asm` | 64-bit / 32-bit | Optimized memory filling (`rep stosb`). |
| ![](../../../resources/svg/assembly.svg) `countbyte_x64.asm` / `countbyte_x86.asm`| 64-bit / 32-bit | Scans memory and counts occurrences of a specific byte. |
| ![](../../../resources/svg/assembly.svg) `memchr_x64.asm` / `memchr_x86.asm` | 64-bit / 32-bit | Locates the memory pointer of a specific byte. |
| ![](../../../resources/svg/assembly.svg) `memrev_x64.asm` / `memrev_x86.asm` | 64-bit / 32-bit | Reverses an array of bytes in-place. |

## Compilation and Verification

To test and verify these Assembly thunks, you must assemble the `.asm` source files into raw machine code (binary format).

### 1. Required Tooling: NASM

The Netwide Assembler (NASM) is the industry standard for this task. It outputs raw binary files without the overhead of OS headers (like PE or ELF).

1. Download the NASM executable from [nasm.us](https://www.nasm.us/).
2. Add the NASM directory to your Windows System PATH or run it directly from this folder.

### 2. Compilation Commands

You must compile these files using the `-f bin` flag to ensure the output is a pure stream of processor instructions.

#### Example for x64
```bash
nasm -f bin xor_x64.asm -o xor_x64.bin
```

#### Example for x86
```bash
nasm -f bin xor_x86.asm -o xor_x86.bin
```

### 3. Extracting Opcodes for VBA

Once you have the `.bin` files, you can load them directly via a binary file reader in VBA, or extract the opcodes to hardcode them into your project as Byte Arrays.

#### Method A: Using Windows CertUtil (Built-in)
```bash
certutil -dump xor_x64.bin
```

#### Method B: Using PowerShell
```powershell
[System.IO.File]::ReadAllBytes("xor_x64.bin") | ForEach-Object { "0x{0:X2}" -f $_ } | Join-String -Separator ", "
```

### 4. Integration Concept

Once compiled and loaded into executable memory, the payload can be triggered by passing pointers directly to the memory address. A conceptual execution flow in VBA looks like this:

```vba
Sub TestXorPayload()
    Dim buffer() As Byte
    Dim pBuffer As LongPtr
    
    ' 1. Setup Test Data
    buffer = StrConv("Hello World", vbFromUnicode)
    pBuffer = VarPtr(buffer(0))
    
    ' 2. Assuming `pExecMemory` is your allocated PAGE_EXECUTE_READWRITE block
    ' containing the loaded xor_x64.bin opcodes...
    
    ' 3. Run the payload via CallWindowProcW (P1 = Buffer Ptr, P2 = Length, P3 = Key)
    ' CallWindowProcW pExecMemory, pBuffer, UBound(buffer) + 1, &HAA, 0
    
    ' Result: buffer() is now encrypted in-place with XOR key 0xAA
End Sub
```

> [!CAUTION]
> Always ensure you are compiling and injecting the correct binary architecture (`x86` vs `x64`). Running a 32-bit payload on a 64-bit host installation (or vice versa) will result in an immediate application crash.
