# 32-bit / 64-bit Pointer Handling

This directory explains a critical bug found during testing in twinBASIC 32-bit,
which also affects 32-bit Office hosts. The bug caused memory corruption and
crashes when resolving hostnames or iterating `addrinfo` lists.

## Key takeaways

- The `#If VBA7` preprocessor constant is **not** a suitable way to differentiate
  32-bit from 64-bit environments. Use `#If Win64` instead.
- The `addrinfo` structure (used by `getaddrinfo`) has pointer fields that change
  size between 32 and 64 bits. Offsets differ accordingly.
- The `hostent` structure (used by `gethostbyname`) also contains a pointer
  to the address list. Dereferencing with the wrong pointer size leads to garbage
  data or access violations.
