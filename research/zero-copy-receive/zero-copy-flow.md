# Zero‑Copy Receive Flow

## Motivation

VBA’s standard string handling always creates a new copy of the data when you
retrieve it from an internal buffer. For large or frequent messages, this
repeated copying can slow down the application and cause memory fragmentation.

`WebSocketReceiveZeroCopy` avoids this by returning a **pointer** directly into
Wasabi’s private buffer.

## How it works

1. `WebSocketReceive(handle)` dequeues a message and returns a VBA String. That
   string is a fresh copy.

2. `WebSocketReceiveZeroCopy` instead:
   - Checks if `ZeroCopyEnabled` is `True`.
   - Processes incoming frames and dequeues a message into the module‑level
     variable `m_ZeroCopyText`.
   - Returns the `StrPtr` of that string via the `outPtr` parameter, and its
     length via `outLen`.

   The caller can then read the string directly from memory without VBA making
   another copy.

## Lifetime and safety

- The pointer is **valid only until the next call** to
  `WebSocketReceiveZeroCopy` on the same handle (or any other receive that
  modifies `m_ZeroCopyText`). After that, the internal string may be overwritten.
- The caller **must not** modify or free the memory; it is managed by VBA.
- To use the data, the caller can either:
  - Read it immediately (e.g., pass `outPtr` to an API that expects a string
    pointer).
  - Copy it to a local buffer if it needs to persist.

## Why not always enabled?

- It’s more error‑prone: dangling pointers can cause crashes or garbled data if
  not used carefully.
- Most VBA users are more comfortable with standard string returns.
- It provides no benefit for small messages, and the internal module‑level
  variable adds a small amount of state that must be managed.

## Public API

```vb
#If VBA7 Then
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As LongPtr, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#Else
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
#End If
```

The function returns `True` if a message was available, `False` otherwise.

## Enabling

```vb
WebSocketSetZeroCopy True, handle
```

It can be enabled or disabled at any time, even while connected.
