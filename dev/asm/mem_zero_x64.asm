[bits 64]

start:
    test rdx, rdx
    jz end
    xor al, al
loop:
    mov [rcx], al
    inc rcx
    dec rdx
    jnz loop
end:
    ret
