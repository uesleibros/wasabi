bits 32
global mem_chr

; CallWindowProcW signature (stdcall):
; [ebp+8]  = Buffer pointer (P1)
; [ebp+12] = Byte to find (P2)
; [ebp+16] = Length to search (P3)
; [ebp+20] = ignored

mem_chr:
    push ebp
    mov ebp, esp
    push edi           ; Save EDI

    mov edi, [ebp+8]   ; EDI = Buffer ptr
    mov eax, [ebp+12]  ; AL  = Target byte
    mov ecx, [ebp+16]  ; ECX = Search length

    repne scasb        ; Scan memory
    
    je .found          ; Jump if found

    xor eax, eax       ; Return 0 if not found
    jmp .done

.found:
    mov eax, edi       ; EDI points to byte after match
    dec eax            ; Adjust back to the match

.done:
    pop edi            ; Restore EDI
    pop ebp
    ret 16             ; Clean stack
