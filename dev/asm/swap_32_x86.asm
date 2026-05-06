[bits 32]

start:
    mov eax, [esp + 4]
    bswap eax
    ret 4
