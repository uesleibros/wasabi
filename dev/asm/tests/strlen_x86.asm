bits 32
global str_len

; CallWindowProcW signature (stdcall):
; [ebp+8]  = String pointer (P1)
; [ebp+12], [ebp+16], [ebp+20] = ignored

str_len:
    push ebp
    mov ebp, esp
    push edi           ; Save EDI

    mov edi, [ebp+8]   ; EDI = String ptr
    xor eax, eax       ; AL = 0
    mov ecx, -1        ; ECX = -1

    repne scasb        ; Scan until 0 is found

    mov eax, -2
    sub eax, ecx       ; Calculate length

    pop edi            ; Restore EDI
    pop ebp
    ret 16             ; Clean up stack
