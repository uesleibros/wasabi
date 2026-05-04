# Happy Eyeballs / Dual-Stack Connection

Wasabi implements the Happy Eyeballs algorithm (RFC 8305) to quickly
establish WebSocket connections on dual‑stack (IPv4+IPv6) networks.

## Motivation

Without Happy Eyeballs, a browser might wait for a failing IPv6
connection to time out before trying IPv4, adding seconds of delay.
Wasabi races both address families and uses whichever succeeds first.
