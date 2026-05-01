/*
 * wasabi-test-server.c
 * Lightweight WebSocket echo server for testing Wasabi.
 * Supports plain TCP (ws://) and TLS (wss://) via OpenSSL.
 *
 * Compilation (Linux / Termux):
 *   gcc -O2 -o wasabi-test-server wasabi-test-server.c -lssl -lcrypto
 *
 * Usage:
 *   ./wasabi-test-server [port] [cert.pem] [key.pem]
 *   Defaults: port=9001, no TLS.
 */

#define _GNU_SOURCE
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>
#include <stdint.h>
#include <inttypes.h>
#include <sys/socket.h>
#include <netinet/in.h>
#include <arpa/inet.h>
#include <openssl/sha.h>
#include <openssl/bio.h>
#include <openssl/evp.h>
#include <openssl/buffer.h>
#include <openssl/ssl.h>
#include <openssl/err.h>

#define BUFFER_SIZE 65536
#define WS_GUID "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"

/* Base64 encode using OpenSSL */
char* base64_encode(const unsigned char* buffer, size_t length) {
    BIO *bio, *b64;
    BUF_MEM *bufferPtr;
    b64 = BIO_new(BIO_f_base64());
    bio = BIO_new(BIO_s_mem());
    bio = BIO_push(b64, bio);
    BIO_set_flags(bio, BIO_FLAGS_BASE64_NO_NL);
    BIO_write(bio, buffer, length);
    BIO_flush(bio);
    BIO_get_mem_ptr(bio, &bufferPtr);
    char* res = (char*)malloc(bufferPtr->length + 1);
    memcpy(res, bufferPtr->data, bufferPtr->length);
    res[bufferPtr->length] = 0;
    BIO_free_all(bio);
    return res;
}

/* SHA1 helper */
void sha1(const unsigned char* data, size_t len, unsigned char* out) {
    SHA1(data, len, out);
}

/* Compute Sec-WebSocket-Accept */
char* compute_accept(const char* key) {
    unsigned char hash[SHA_DIGEST_LENGTH];
    char combined[256];
    snprintf(combined, sizeof(combined), "%s%s", key, WS_GUID);
    sha1((unsigned char*)combined, strlen(combined), hash);
    return base64_encode(hash, SHA_DIGEST_LENGTH);
}

/* Perform WebSocket handshake, returns 0 on success */
int ws_handshake(int fd) {
    char buf[BUFFER_SIZE];
    ssize_t n = read(fd, buf, sizeof(buf)-1);
    if (n <= 0) return -1;
    buf[n] = 0;

    /* Extract Sec-WebSocket-Key */
    char* key_start = strstr(buf, "Sec-WebSocket-Key: ");
    if (!key_start) return -1;
    key_start += 19;
    char* key_end = strstr(key_start, "\r\n");
    if (!key_end) return -1;
    char key[256];
    snprintf(key, key_end - key_start + 1, "%s", key_start);

    /* Compute accept */
    char* accept = compute_accept(key);

    /* Build response */
    char response[1024];
    snprintf(response, sizeof(response),
        "HTTP/1.1 101 Switching Protocols\r\n"
        "Upgrade: websocket\r\n"
        "Connection: Upgrade\r\n"
        "Sec-WebSocket-Accept: %s\r\n\r\n", accept);
    free(accept);

    if (write(fd, response, strlen(response)) <= 0) return -1;
    return 0;
}

/* Read a complete WebSocket frame, return payload and opcode/fin.
   Returns payload length, *opcode and *fin set, caller must free *payload. */
ssize_t ws_read_frame(int fd, uint8_t** payload, uint8_t* opcode, uint8_t* fin) {
    uint8_t header[2];
    if (read(fd, header, 2) <= 0) return -1;
    *fin = (header[0] & 0x80) >> 7;
    *opcode = header[0] & 0x0F;
    int masked = (header[1] & 0x80) >> 7;
    uint64_t plen = header[1] & 0x7F;

    if (plen == 126) {
        uint8_t ext[2];
        read(fd, ext, 2);
        plen = ((uint16_t)ext[0] << 8) | ext[1];
    } else if (plen == 127) {
        uint8_t ext[8];
        read(fd, ext, 8);
        plen = 0;
        for (int i = 0; i < 8; i++) plen = (plen << 8) | ext[i];
    }

    uint8_t mask_key[4] = {0};
    if (masked) read(fd, mask_key, 4);

    uint8_t* data = malloc(plen + 1);
    if (read(fd, data, plen) <= 0) { free(data); return -1; }

    if (masked)
        for (uint64_t i = 0; i < plen; i++)
            data[i] ^= mask_key[i % 4];

    *payload = data;
    return plen;
}

/* Send a WebSocket frame (server → client, never masked) */
void ws_send_frame(int fd, uint8_t opcode, const uint8_t* data, uint64_t len) {
    uint8_t header[14];
    int hlen = 2;
    header[0] = 0x80 | (opcode & 0x0F);
    if (len <= 125) {
        header[1] = len;
    } else if (len <= 65535) {
        header[1] = 126;
        header[2] = (len >> 8) & 0xFF;
        header[3] = len & 0xFF;
        hlen = 4;
    } else {
        header[1] = 127;
        for (int i = 0; i < 8; i++) header[2+i] = (len >> (56 - i*8)) & 0xFF;
        hlen = 10;
    }
    write(fd, header, hlen);
    if (len > 0) write(fd, data, len);
}

void handle_client(int client_fd, SSL* ssl) {
    #define READ(fd,buf,len) (ssl ? SSL_read(ssl,buf,len) : read(fd,buf,len))
    #define WRITE(fd,buf,len) (ssl ? SSL_write(ssl,buf,len) : write(fd,buf,len))

    int fd = client_fd; /* use fd for raw I/O if no TLS */

    if (ws_handshake(client_fd) != 0) {
        fprintf(stderr, "Handshake failed\n");
        goto cleanup;
    }

    for (;;) {
        uint8_t opcode, fin;
        uint8_t* payload;
        ssize_t len = ws_read_frame(fd, &payload, &opcode, &fin);
        if (len < 0) break;

        if (opcode == 0x8) { /* Close */
            if (len >= 2) {
                uint16_t close_code = (payload[0] << 8) | payload[1];
                fprintf(stderr, "Close frame received: code=%u\n", close_code);
                ws_send_frame(fd, 0x8, payload, len);
            } else {
                ws_send_frame(fd, 0x8, NULL, 0);
            }
            free(payload);
            break;
        } else if (opcode == 0x9) { /* Ping */
            ws_send_frame(fd, 0xA, payload, len); /* Pong */
        } else if (opcode == 0x1 || opcode == 0x2) { /* Text or Binary */
            /* Echo back */
            ws_send_frame(fd, opcode, payload, len);
        }
        /* Ignore other opcodes */
        free(payload);
    }

cleanup:
    if (ssl) SSL_shutdown(ssl);
    close(client_fd);
    if (ssl) SSL_free(ssl);
}

int main(int argc, char** argv) {
    int port = 9001;
    char* cert_file = NULL, *key_file = NULL;

    if (argc > 1) port = atoi(argv[1]);
    if (argc > 2) cert_file = argv[2];
    if (argc > 3) key_file = argv[3];

    int server_fd = socket(AF_INET, SOCK_STREAM, 0);
    if (server_fd < 0) { perror("socket"); return 1; }

    int opt = 1;
    setsockopt(server_fd, SOL_SOCKET, SO_REUSEADDR, &opt, sizeof(opt));

    struct sockaddr_in addr = {0};
    addr.sin_family = AF_INET;
    addr.sin_port = htons(port);
    addr.sin_addr.s_addr = INADDR_ANY;
    if (bind(server_fd, (struct sockaddr*)&addr, sizeof(addr)) < 0)
        { perror("bind"); close(server_fd); return 1; }
    if (listen(server_fd, 1) < 0)
        { perror("listen"); close(server_fd); return 1; }

    SSL_CTX* ctx = NULL;
    if (cert_file && key_file) {
        SSL_library_init();
        ctx = SSL_CTX_new(TLS_server_method());
        if (SSL_CTX_use_certificate_file(ctx, cert_file, SSL_FILETYPE_PEM) <= 0 ||
            SSL_CTX_use_PrivateKey_file(ctx, key_file, SSL_FILETYPE_PEM) <= 0) {
            fprintf(stderr, "Failed to load certificate/key\n");
            SSL_CTX_free(ctx);
            close(server_fd);
            return 1;
        }
    }

    printf("Wasabi Test Server listening on port %d (%s)\n", port, ctx ? "TLS" : "plain");
    for (;;) {
        struct sockaddr_in client_addr;
        socklen_t client_len = sizeof(client_addr);
        int client_fd = accept(server_fd, (struct sockaddr*)&client_addr, &client_len);
        if (client_fd < 0) { perror("accept"); continue; }

        SSL* ssl = NULL;
        if (ctx) {
            ssl = SSL_new(ctx);
            SSL_set_fd(ssl, client_fd);
            if (SSL_accept(ssl) <= 0) {
                ERR_print_errors_fp(stderr);
                SSL_free(ssl);
                close(client_fd);
                continue;
            }
        }
        handle_client(client_fd, ssl);
    }

    if (ctx) SSL_CTX_free(ctx);
    close(server_fd);
    return 0;
}
