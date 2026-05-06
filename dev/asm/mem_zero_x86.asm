[bits 32]

start:
    push ebp
    mov ebp, esp
    push edi
    push esi
    push ebx
    mov esi, [ebp + 8]
    mov ecx, [ebp + 12]
    mov edi, [ebp + 16]
    test ecx, ecx
    jz end
    xor edx, edx
mask_loop:
    mov al, [esi]
    mov bl, [edi + edx]
    xor al, bl
    mov [esi], al
    inc esi
    inc edx
    and edx, 3
    dec ecx
    jnz mask_loop
end:
    pop ebx
    pop esi
    pop edi
    mov esp, ebp
    pop ebp
    ret 12
