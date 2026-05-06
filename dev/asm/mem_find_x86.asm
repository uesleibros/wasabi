bits 32
global mem_find

; CallWindowProcW signature (stdcall):
; [ebp+8]  = haystack ptr
; [ebp+12] = haystack_len
; [ebp+16] = needle ptr
; [ebp+20] = needle_len

mem_find:
    push ebp
    mov ebp, esp
    push ebx
    push esi
    push edi

    mov edx, [ebp+12]  ; EDX = haystack_len
    mov ecx, [ebp+20]  ; ECX = needle_len

    cmp edx, ecx
    jb .not_found      ; haystack_len < needle_len
    
    test ecx, ecx
    jz .not_found      ; needle_len == 0

    sub edx, ecx
    inc edx            ; Maximum possible loops
    
    xor eax, eax       ; EAX (index) = 0
    mov ebx, [ebp+8]   ; EBX = haystack ptr

.search_loop:
    push ecx           ; Save needle_len
    push ebx           ; Save current haystack position

    mov ecx, [ebp+20]  ; Amount to compare
    mov edi, [ebp+16]  ; EDI = needle ptr
    mov esi, ebx       ; ESI = current haystack ptr

    repe cmpsb         ; Compare byte by byte

    pop ebx            ; Restore position
    pop ecx            ; Restore needle_len

    je .found          ; Full match found!

    inc ebx            ; Advance 1 byte in search
    inc eax            ; Increment the index
    dec edx            ; Decrement loop count
    jnz .search_loop

.not_found:
    mov eax, -1
.found:
    pop edi
    pop esi
    pop ebx
    pop ebp
    ret 16             ; Clean up the 4 arguments