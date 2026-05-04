# Pointer Bug Analysis: VBA7 vs Win64

## Symptoms

On 32‑bit versions of twinBASIC (and 32‑bit Office if tested), the
`WebSocketConnect` example crashed inside `ResolveAndConnect` or
`ResolveHostname` when attempting to dereference an `addrinfo` or `hostent`
pointer. The crash was a memory access violation.

## Root Cause

The original code used `#If VBA7 Then` to choose between two sets of
`CopyMemoryFromPtr` calls. However, `VBA7` is true for **both** 32‑bit and
64‑bit Office (since Office 2010+ uses VBA7). Therefore the “VBA7” branch was
used in 32‑bit environments, incorrectly reading 8‑byte pointer fields from
structures that only contain 4‑byte pointers.

### Example: `addrinfo` iteration

In `ResolveAndConnect`, the original code was:

```vb
#If VBA7 Then
    CopyMemoryFromPtr pSockaddr, pCur + 32, 8   ' expects 64-bit pointer
    CopyMemoryFromPtr pNext, pCur + 40, 8       ' expects 64-bit pointer
#Else
    CopyMemoryFromPtr pSockaddr, pCur + 24, 4   ' correct for 32-bit
    CopyMemoryFromPtr pNext, pCur + 28, 4
#End If
```

In 32-bit VBA7, the first branch runs, reading 8 bytes from an offset designed for
a 32-bit layout, picking up garbage for the upper 4 bytes and corrupting the
pointer. This leads to an invalid memory access when later dereferencing
`pSockaddr`.

### Example: `hostent` address list

Similarly, in `ResolveHostname`, the code to extract the first address from
`h_addr_list` used:

```vb
#If Win64 Then... but originally it might have been #If VBA7
```

After the report, this was confirmed to already be using `#If Win64`, but the
same mistake could have been present in an earlier version. The fix ensures
consistency.

## The Fix

Replace all such pointer-size conditionals with `#If Win64`. This constant is
defined by the compiler only on 64‑bit builds. The corrected code in both
functions now reads:

```vb
#If Win64 Then
    ' read 8-byte pointers
#Else
    ' read 4-byte pointers
#End If
```

This correctly distinguishes CPU bitness.

### Functions affected

- `ResolveAndConnect` – reading `ai_addr`, `ai_family`, `ai_addrlen`, `ai_next` from `addrinfo`.
- `ResolveHostname` – reading `h_addr_list` pointer and the first address pointer.

## Additional Notes

- The `Win64` constant is available in VB6, twinBASIC, and VBA. It’s the
  standard way to conditionally compile for pointer size.
- The `VBA7` constant is still useful for features that depend on the VBA version
  (e.g., `LongPtr` availability), but **not** for pointer size or memory layout.

## References

- [MSDN: Compiler Constants](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/compiler-constants)
- [twinBASIC testing report](https://github.com/uesleibros/wasabi/issues/1)
