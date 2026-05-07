bits 64
global count_byte

; CallWindowProcW signature (x64):
; RCX = Buffer pointer (P1)
; RDX = Length in bytes (P2)
; R8  = Byte to find (P3)
; R9  = ignored

count_byte:
    xor rax, rax       ; RAX = 0 (Counter for occurrences)
    test rdx, rdx      ; Check if length is 0
    jz .done           ; Exit if length is 0

.loop:
    mov r9b, byte [rcx] ; Read 1 byte from buffer into R9B
    cmp r9b, r8b       ; Compare current byte with the target byte (R8B)
    jne .skip          ; If not equal, skip increment

    inc rax            ; Increment the occurrence counter

.skip:
    inc rcx            ; Move pointer to the next byte
    dec rdx            ; Decrease length counter
    jnz .loop          ; Loop until length is 0

.done:
    ret                ; Return count in RAX
