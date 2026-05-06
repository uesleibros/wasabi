bits 64
global mem_find

; CallWindowProcW signature (x64):
; RCX = haystack ptr
; RDX = haystack_len
; R8  = needle ptr
; R9  = needle_len
; Returns index in RAX (-1 if not found)

mem_find:
    push rsi
    push rdi
    push rbx

    cmp rdx, r9        ; haystack_len < needle_len?
    jb .not_found      ; Impossible to find, exit
    
    test r9, r9        ; needle_len == 0?
    jz .not_found

    sub rdx, r9        ; Loop limit: haystack_len - needle_len
    inc rdx            ; +1 (inclusive)
    
    xor rax, rax       ; Initialize return index to 0

.search_loop:
    mov rbx, rcx       ; Save current haystack pointer
    mov rcx, r9        ; RCX = needle_len (for cmpsb)
    mov rdi, r8        ; RDI = needle ptr
    mov rsi, rbx       ; RSI = current haystack ptr
    
    repe cmpsb         ; Compare memory blocks
    je .found          ; If Zero flag is set, full match found!
    
    mov rcx, rbx       ; Restore haystack pointer
    inc rcx            ; Advance 1 byte in haystack
    inc rax            ; Increment index
    dec rdx            ; Decrement loop limit
    jnz .search_loop

.not_found:
    mov rax, -1        ; Return -1
.found:
    pop rbx
    pop rdi
    pop rsi
    ret