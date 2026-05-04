# Zero‑Copy Receive

Wasabi offers an optional zero‑copy receive path that avoids duplicating string
data in VBA, reducing memory allocations for high‑throughput applications.

## Reference

- `WebSocketReceiveZeroCopy()` and the module‑level `m_ZeroCopyText` in `Wasabi.bas`.
