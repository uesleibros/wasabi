bits 64
global mem_reverse

; CallWindowProcW signature (x64):
; RCX = Buffer pointer (P1)
; RDX = Length in bytes (P2)
; R8, R9 = ignored

mem_reverse:
    cmp rdx, 1         ; If length is 0 or 1, no need to reverse
    jle .done

    mov r8, rcx        ; R8 = Start pointer (Left)
    mov r9, rcx        ; R9 = End pointer (Right)
    add r9, rdx        ; Add length
    dec r9             ; R9 points to the last valid byte

.loop:
    cmp r8, r9         ; Check if Left pointer crossed or reached Right pointer
    jge .done          ; If so, we are done

    ; Swap bytes
    mov al, byte [r8]  ; Read left byte
    mov dl, byte [r9]  ; Read right byte
    mov byte [r8], dl  ; Put right byte in left position
    mov byte [r9], al  ; Put left byte in right position

    inc r8             ; Move left pointer forward
    dec r9             ; Move right pointer backward
    jmp .loop

.done:
    mov rax, rcx       ; Return the original buffer pointer
    ret
