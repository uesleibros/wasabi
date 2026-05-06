[bits 64]

start:
    mov rax, 0x1122334455667788
    mov eax, [rax]
    test eax, eax
    jz fallback
    mov rax, 0x2233445566778899
    jmp rax
fallback:
    mov rax, 0x33445566778899AA
    jmp rax
