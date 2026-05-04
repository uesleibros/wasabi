# Path MTU Discovery & Automatic Fragmentation

Wasabi probes the TCP connection to discover the real Path MTU and
adjusts WebSocket frame sizes to avoid IP fragmentation.

## Why it matters

Fragmented IP packets increase latency and CPU overhead. By keeping
WebSocket frames within the TCP Maximum Segment Size, Wasabi avoids
IP fragmentation transparently.
