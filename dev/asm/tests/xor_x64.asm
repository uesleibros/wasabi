bits 64
global xor_buffer

; CallWindowProcW signature (x64):
; RCX = Buffer pointer (P1)
; RDX = Length in bytes (P2)
; R8  = XOR Key (1 byte) (P3)
; R9  = ignored

xor_buffer:
    test rdx, rdx      ; Check if length is 0
    jz .done           ; If length is 0, exit immediately
    
    mov r9, rcx        ; Save the original buffer pointer in R9 to iterate

.loop:
    mov al, byte [r9]  ; Load 1 byte from buffer into AL
    xor al, r8b        ; XOR the byte with the key (lowest byte of R8)
    mov byte [r9], al  ; Write the byte back to the buffer
    
    inc r9             ; Move to the next byte in the buffer
    dec rdx            ; Decrease the counter
    jnz .loop          ; If counter is not zero, loop again

.done:
    mov rax, rcx       ; Return the original buffer pointer
    ret
