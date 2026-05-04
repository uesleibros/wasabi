# MTU Discovery Flow

## Initialization

When a connection is established, `InitializeMTU` sets `CurrentMTU` to
the default (1500 bytes) and marks `ProbeEnabled = True`.

## Probing

`probeMTU(handle)` uses `getsockopt(TCP_MAXSEG)` to read the TCP
Maximum Segment Size negotiated during the three‑way handshake.

- If the call succeeds, MTU is calculated as:
  `MSS + TCP_HEADER_MIN + IP_HEADER_MIN + ETHERNET_HEADER`
  (accounting for IPv6 if `PreferIPv6` is true).

- If the call fails, MSS is assumed to be 1460 (standard Ethernet).

## Periodic probing

`AutoMTU` (enabled by default) re‑probes every
`PMTU_DISCOVERY_INTERVAL_MS` (60 seconds) to detect network changes.

## TLS overhead

When TLS is active, TLS record overhead (`Sizes.cbHeader` +
`Sizes.cbTrailer`, obtained from `QueryContextAttributes`) is subtracted
from the available space.
