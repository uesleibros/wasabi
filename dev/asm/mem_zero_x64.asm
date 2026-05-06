bits 64
global mem_zero

; CallWindowProcW signature (x64):
; RCX = destination ptr
; RDX = length
; R8, R9 = ignored

mem_zero:
    push rdi           ; Save register (x64 calling convention)
    mov rdi, rcx       ; RDI = destination ptr (required by stosb)
    mov rcx, rdx       ; RCX = byte counter
    xor eax, eax       ; Zero out EAX (therefore, AL = 0)
    rep stosb          ; Fill RCX bytes at [RDI] with the value of AL (0)
    pop rdi            ; Restore register
    ret