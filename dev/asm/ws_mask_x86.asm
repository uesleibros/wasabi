bits 32
global ws_mask

; CallWindowProcW signature (stdcall):
; [ebp+8]  = hwnd (used as payload ptr)
; [ebp+12] = msg  (used as length)
; [ebp+16] = wp   (used as mask ptr)
; [ebp+20] = lp   (reserved)

ws_mask:
    push ebp
    mov ebp, esp
    push ebx

    mov ecx, [ebp+12]  ; ecx = length
    test ecx, ecx
    jz .done

    mov edx, [ebp+8]   ; edx = payload ptr
    mov eax, [ebp+16]  ; eax = mask ptr
    mov ebx, [eax]     ; ebx = mask value (4 bytes)

.loop:
    mov al, bl         ; Move lowest mask byte into AL
    xor [edx], al      ; Apply XOR to payload
    inc edx            ; Advance pointer
    ror ebx, 8         ; Rotate mask
    dec ecx            ; Decrement counter
    jnz .loop          ; Repeat loop

.done:
    pop ebx
    pop ebp
    ret 16             ; Clean up the 4 arguments (16 bytes) from the stack