bits 64
global add_numbers

; CallWindowProcW signature (x64):
; RCX = Number 1 (P1)
; RDX = Number 2 (P2)
; R8, R9 = ignored

add_numbers:
    mov rax, rcx       ; Move the first parameter to RAX (return register)
    add rax, rdx       ; Add the second parameter to RAX
    ret                ; Return (Result is in RAX)
