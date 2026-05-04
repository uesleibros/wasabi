# zlib stdcall Investigation Log

## Background

Wasabi declares zlib functions without `CDecl`:

```vb
Private Declare PtrSafe Function zlib_deflateInit2 Lib "zlib1.dll" Alias ...
```

In VBA, the default calling convention is **stdcall** (WINAPI). Therefore, the DLL
must export functions with `__stdcall`. The standard Windows build of zlib
(`zlib1.dll` from zlib.net) uses `__cdecl`, causing crashes or silent corruption
when called from Wasabi.

## Attempts

### 1. Official zlib website

- **URL:** https://zlib.net
- **Result:** Provides source code and a `zlib1.dll` compiled for Windows (cdecl).
  No stdcall build available. The page links to Gilles Vollant's site for older
  builds, but that link is broken (see below).

### 2. Gilles Vollant's winimage.com

- **URL:** https://www.winimage.com/zLibDll/
- **Result:** The HTTPS URL returns a blank page. The site only works over plain
  HTTP (`http://winimage.com/zLibDll/`), which browsers often block. Even when
  accessible, the builds are ancient (zlib 1.2.3) and the page is unmaintained.
  **Dead end.**

### 3. PostgreSQL's zlib1.dll

- **Observation:** Some users reported that the `zlib1.dll` shipped with
  PostgreSQL worked. It was likely compiled with different flags.
- **Result:** Unreliable for distribution, and version may be old.
  **Not a portable solution.**

### 4. Joveler.Compression.ZLib NuGet package (first attempt)

- **Target:** Version 4.3.0 (last version assumed to have stdcall).
- **Result:** The NuGet page for 4.3.0 redirects to 6.0.1. Version 6.0.1 contains
  only `zlib1.dll` (cdecl), **not** `zlibwapi.dll`. The changelog confirms stdcall
  was dropped in version 5.0.0. **Dead end.**

### 5. Joveler.Compression.ZLib version 4.2.0 (success)

- **URL:** https://www.nuget.org/packages/Joveler.Compression.ZLib/4.2.0
- **Result:** This version still includes `zlibwapi.dll` (stdcall) for both
  `win-x86` and `win-x64` inside the `.nupkg` (ZIP archive).
  - `runtimes/win-x86/native/zlibwapi.dll`
  - `runtimes/win-x64/native/zlibwapi.dll`
- **Status:** **Adopted.** Extracted, renamed, and placed in `../libs/`.

## Key takeaways

- The calling convention **must** match between the VBA declaration and the DLL.
  Mismatches cause runtime errors, stack corruption, or crashes.
- `zlibwapi.dll` = stdcall; `zlib1.dll` (from most sources) = cdecl.
- NuGet package versions can disappear, redirect, or change ABI; always verify
  the actual content.
- The Joveler package is a reliable third-party source with a clear version
  history, making it acceptable for redistribution in binary form (zlib license).

> Point: https://www.reddit.com/r/vba/comments/1t3ed1q/comment/ojvjsue
