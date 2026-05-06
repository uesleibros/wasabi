[bits 32]

start:
    mov eax, [0x11223344]
    test eax, eax
    jz fallback
    mov eax, 0x22334455
    jmp eax
fallback:
    mov eax, 0x33445566
    jmp eax
