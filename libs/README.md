# Precompiled zlib DLLs for Wasabi (permessage-deflate)

This directory contains pre‑compiled builds of **zlib** that work with
the **Wasabi** WebSocket module. They are required only if you want to use
the `permessage-deflate` compression extension (RFC 7692).

## Files

| File             | Architecture | Calling convention |
|------------------|--------------|--------------------|
| `zlib1_x86.dll`  | 32‑bit (x86) | `stdcall` (WINAPI) |
| `zlib1_x64.dll`  | 64‑bit (x64) | `stdcall` (WINAPI) |

## Source

These DLLs were extracted from the **Joveler.Compression.ZLib** NuGet package
(version **4.2.0**), which distributes official zlib builds compiled with
`ZLIB_WINAPI` – the calling convention required by Wasabi’s VBA declarations.

- Package URL: [`Joveler.Compression.ZLib/4.2.0`](https://www.nuget.org/packages/Joveler.Compression.ZLib/4.2.0)

Version 4.2.0 is the last release that provides **stdcall** binaries
(`zlibwapi.dll`). Later versions (5.0.0+) switched to **cdecl** and are
**not** compatible with Wasabi.

### How they were obtained

1. Download the `.nupkg` file from the link above.
2. Rename the extension from `.nupkg` to `.zip`.
3. Extract the archive and navigate to:
   - `runtimes/win-x86/native/zlibwapi.dll` → 32‑bit stdcall
   - `runtimes/win-x64/native/zlibwapi.dll` → 64‑bit stdcall
4. Rename each DLL to match Wasabi’s expected names:
   - `zlibwapi.dll` (x86) → `zlib1_x86.dll`
   - `zlibwapi.dll` (x64) → `zlib1_x64.dll`

The renamed files are the ones you find in this folder.

## Where to place them

Copy the DLL that matches your Office architecture to the **same folder**
as your Excel workbook, Word document, or add‑in. Wasabi searches this
location first.

For a complete description of the search order, fallback names, and
compression configuration, please refer to the main documentation:

📄 [../docs/README.md – Optional zlib Dependency](../docs/README.md)

## Version info

- **zlib source version:** as shipped in Joveler 4.2.0 (based on zlib 1.2.11)
- **Compiled with:** Visual Studio, `ZLIB_WINAPI` defined
- **License:** zlib license (permissive, allows redistribution)

> [!NOTE]
> If you need a different zlib version or prefer to build from source,
> follow the compilation instructions inside the main documentation.
> Official source code is available at [https://zlib.net](https://zlib.net).
