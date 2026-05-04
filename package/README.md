# Package

This folder contains **Wasabi.bas**, the single, self‑contained file that
brings the entire WebSocket stack to any VBA project.

> When you import `Wasabi.bas`, the VBA runtime compiles it on the spot.

> There is no build step, no binary, and no packaging. What you see is
exactly what runs inside the Office process.

## How to use

1. In the VBA editor, click **File → Import File…**
2. Select `Wasabi.bas` from this folder.
3. No additional steps are required — no references, no tools, no setup.

After importing, you can call `WebSocketConnect` directly from any module.

> [!TIP]
> The complete API reference is available in [`docs/API_REFERENCE.md`](../docs/API_REFERENCE.md).
