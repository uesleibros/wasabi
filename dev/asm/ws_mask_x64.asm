bits 64
global ws_mask

; CallWindowProcW signature (x64):
; RCX = payload ptr
; RDX = payload length
; R8  = mask ptr (4 bytes)
; R9  = reserved (0)

ws_mask:
    test rdx, rdx      ; Check if length is 0
    jz .done           ; If zero, exit
    mov eax, [r8]      ; Load the 4-byte mask into EAX

.loop:
    xor [rcx], al      ; XOR current payload byte with the lowest byte of the mask (AL)
    inc rcx            ; Advance payload pointer
    ror eax, 8         ; Rotate mask (brings next mask byte into AL)
    dec rdx            ; Decrement length counter
    jnz .loop          ; If not zero, repeat loop

.done:
    ret