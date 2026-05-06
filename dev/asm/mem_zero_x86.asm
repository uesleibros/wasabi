bits 32
global mem_zero

; CallWindowProcW signature (stdcall):
; [ebp+8]  = destination ptr
; [ebp+12] = length
; [ebp+16], [ebp+20] = ignored

mem_zero:
    push ebp
    mov ebp, esp
    push edi

    mov edi, [ebp+8]   ; EDI = destination ptr
    mov ecx, [ebp+12]  ; ECX = byte counter
    xor eax, eax       ; Zero out AL
    rep stosb          ; Fill with zero

    pop edi
    pop ebp
    ret 16             ; Clean up stack