# Obtaining stdcall zlib DLLs from Joveler.Compression.ZLib

## Prerequisites

- Internet access to download the NuGet package.
- A tool to extract ZIP files (7-Zip, WinRAR, Windows built-in).

## Steps

### 1. Download the package

Download from: https://www.nuget.org/packages/Joveler.Compression.ZLib/4.2.0

The file will be named `joveler.compression.zlib.4.2.0.nupkg`.

### 2. Extract the archive

Rename the file extension from `.nupkg` to `.zip`. Then extract it with any
ZIP utility.

### 3. Locate the DLLs

Inside the extracted folder, navigate to:

- **32-bit:** `runtimes/win-x86/native/zlibwapi.dll`
- **64-bit:** `runtimes/win-x64/native/zlibwapi.dll`

These are the stdcall builds.

### 4. Rename for Wasabi

Wasabi's `GetZlibName()` function expects architecture-specific names:

| Extracted file | Rename to |
|---|---|
| `win-x86/zlibwapi.dll` | `zlib1_x86.dll` (or `zlib1.dll` as fallback) |
| `win-x64/zlibwapi.dll` | `zlib1_x64.dll` |

The preferred names are `zlib1_x86.dll` and `zlib1_x64.dll`. Wasabi also falls
back to `zlib1.dll`, but using architecture-specific names avoids ambiguity.

### 5. Place in the correct directory

Copy the renamed DLLs to the same folder as your Excel workbook or add-in.
For redistribution, they are placed in the `libs/` folder of the Wasabi
repository.

## Verification

Run the `test_stdcall.bas` script (if available) to confirm the DLL loads and
functions correctly.

## Why not compile from source?

It is possible to compile zlib from source with `ZLIB_WINAPI` defined. However,
to avoid requiring a Visual Studio toolchain for end users of Wasabi, we
prefer to provide pre-compiled binaries from a reputable source.
