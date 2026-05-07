bits 32
global mem_copy

; CallWindowProcW signature (stdcall):
; [ebp+8]  = Destination pointer (P1)
; [ebp+12] = Source pointer (P2)
; [ebp+16] = Length in bytes (P3)
; [ebp+20] = ignored

mem_copy:
    push ebp
    mov ebp, esp
    push edi           ; Save EDI
    push esi           ; Save ESI

    mov edi, [ebp+8]   ; EDI = Destination ptr
    mov esi, [ebp+12]  ; ESI = Source ptr
    mov ecx, [ebp+16]  ; ECX = Byte counter

    rep movsb          ; Copy memory

    mov eax, [ebp+8]   ; Return the destination pointer in EAX

    pop esi            ; Restore ESI
    pop edi            ; Restore EDI
    pop ebp
    ret 16             ; Clean up stack
