bits 64
global mem_chr

; CallWindowProcW signature (x64):
; RCX = Buffer pointer (P1)
; RDX = Byte to find (P2)
; R8  = Length to search (P3)
; R9  = ignored

mem_chr:
    push rdi           ; Save RDI
    
    mov rdi, rcx       ; RDI = Buffer pointer (required by scasb)
    mov rax, rdx       ; AL  = Target byte (lowest byte of RAX)
    mov rcx, r8        ; RCX = Search length

    repne scasb        ; Scan string for AL, stopping if found or RCX=0
    
    je .found          ; Zero flag is set if byte was found
    
    xor rax, rax       ; Not found: Return 0
    jmp .done

.found:
    mov rax, rdi       ; Found: scasb leaves RDI pointing to the NEXT byte
    dec rax            ; Subtract 1 to get the exact address of the match

.done:
    pop rdi            ; Restore RDI
    ret
