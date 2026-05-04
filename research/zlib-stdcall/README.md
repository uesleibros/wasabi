# zlib stdcall Build Research

Wasabi's WebSocket `permessage-deflate` extension requires a **stdcall** (WINAPI)
build of zlib. The official `zlib1.dll` from zlib.net uses **cdecl**, which is
incompatible with VBA's default calling convention for `Declare` statements
(unless `CDecl` is explicitly specified, which Wasabi does not).

This directory documents the search for a suitable build, the dead ends, and the
final solution.

## Final solution

We use `zlibwapi.dll` (stdcall) from the **Joveler.Compression.ZLib** NuGet
package, version **4.2.0**. The DLLs were extracted, renamed to match Wasabi's
expected names, and placed in `../libs/`.

- Package: [Joveler.Compression.ZLib 4.2.0](https://www.nuget.org/packages/Joveler.Compression.ZLib/4.2.0)
- Extraction and renaming instructions: see `build-sources.md`.
