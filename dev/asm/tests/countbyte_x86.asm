bits 32
global count_byte

; CallWindowProcW signature (stdcall):
; [ebp+8]  = Buffer pointer (P1)
; [ebp+12] = Length in bytes (P2)
; [ebp+16] = Byte to find (P3)
; [ebp+20] = ignored

count_byte:
    push ebp
    mov ebp, esp
    push ebx           ; Save EBX

    mov edx, [ebp+8]   ; EDX = Buffer ptr
    mov ecx, [ebp+12]  ; ECX = Length
    mov ebx, [ebp+16]  ; BL  = Target byte

    xor eax, eax       ; EAX = 0 (Occurrence counter)

    test ecx, ecx      ; Check if length is zero
    jz .done

.loop:
    mov ch, byte [edx] ; Read 1 byte into CH
    cmp ch, bl         ; Compare byte with target
    jne .skip          ; If not equal, skip increment

    inc eax            ; Increment counter

.skip:
    inc edx            ; Move to next byte
    dec ecx            ; Decrease loop counter (using lowest byte of ECX technically, but let's use standard decrement)
    jnz .loop

.done:
    pop ebx            ; Restore EBX
    pop ebp
    ret 16             ; Clean up stack
