# HTTP CONNECT Proxy

`DoProxyHTTP(handle)` establishes a tunnel through an HTTP proxy.

1. **CONNECT request**: Sends `CONNECT host:port HTTP/1.1` with `Host`
   header.

2. **Authentication**:
   - If `proxyUser` is set and `ProxyUseNtlm` is false, adds a
     `Proxy-Authorization: Basic` header.
   - If `ProxyUseNtlm` is true, sends a hardcoded NTLM Type 1 token and
     handles the 407 challenge (see `ntlm.md`).

3. **Response parsing**: Expects `HTTP/1.x 200`. A `407` response
   triggers NTLM challenge handling or throws `ERR_PROXY_AUTH_FAILED`.

4. **Error handling**: Any other response or timeout returns
   `ERR_PROXY_TUNNEL_FAILED` or `ERR_PROXY_CONNECT_FAILED`.
