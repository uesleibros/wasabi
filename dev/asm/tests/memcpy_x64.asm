bits 64
global mem_copy

; CallWindowProcW signature (x64):
; RCX = Destination pointer (P1)
; RDX = Source pointer (P2)
; R8  = Length in bytes (P3)
; R9  = ignored

mem_copy:
    push rdi           ; Save RDI
    push rsi           ; Save RSI

    mov rdi, rcx       ; RDI = Destination ptr (required by movsb)
    mov rsi, rdx       ; RSI = Source ptr (required by movsb)
    mov rcx, r8        ; RCX = Byte counter

    rep movsb          ; Copy RCX bytes from [RSI] to [RDI]

    mov rax, rdi       ; Return the destination pointer
    
    pop rsi            ; Restore RSI
    pop rdi            ; Restore RDI
    ret
