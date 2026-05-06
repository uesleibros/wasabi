[bits 32]

start:
    mov ecx, [esp + 8]
    mov eax, [esp + 4]
    test ecx, ecx
    jz end
zero_loop:
    mov byte [eax], 0
    inc eax
    dec ecx
    jnz zero_loop
end:
    ret 8
