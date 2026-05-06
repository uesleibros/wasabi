[bits 64]

start:
    test rdx, rdx
    jz end
    xor r9, r9
loop:
    mov al, [rcx]
    mov r10b, [r8 + r9]
    xor al, r10b
    mov [rcx], al
    inc rcx
    inc r9
    and r9, 3
    dec rdx
    jnz loop
end:
    ret
