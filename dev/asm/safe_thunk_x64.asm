bits 64
global safe_thunk_x64

; CallWindowProcW/WndProc signature (x64):
; RCX = hWnd, RDX = uMsg, R8 = wParam, R9 = lParam

safe_thunk_x64:
    ; --- PROLOGUE & SAVE STATE ---
    push rbp
    mov rbp, rsp
    push rcx                        ; Save hWnd
    push rdx                        ; Save uMsg
    push r8                         ; Save wParam
    push r9                         ; Save lParam
    sub rsp, 32                     ; Shadow space (x64 ABI)

    ; --- 1. HEARTBEAT FLAG CHECK (m_AppIsAlive) ---
    mov rax, 0x1122334455667788     ; [OFFSET 16] Pointer to m_AppIsAlive
    mov eax, dword [rax]
    test eax, eax
    jz .is_dead

    ; --- 2. IDE STATE CHECK (EbMode) ---
    mov rax, 0x2233445566778899     ; [OFFSET 32] Pointer to vbe7.dll!EbMode
    test rax, rax
    jz .skip_ebmode

    call rax                        ; EbMode() — returns 1 if running normally
    cmp eax, 1
    jne .is_dead                    ; Break/edit mode → fallback

.skip_ebmode:
    ; --- 3. DISPATCH CELL CHECK (m_ptrDispatch) ---
    mov rax, 0x33445566778899AA     ; [OFFSET 54] Pointer to m_ptrDispatch cell
    mov rax, qword [rax]            ; Dereference: read actual fn ptr from cell
    test rax, rax
    jz .is_dead                     ; Cell zeroed (recompile in progress) → fallback

    ; --- 4. TAIL-CALL WasabiAsyncWndProc ---
    add rsp, 32
    pop r9
    pop r8
    pop rdx
    pop rcx
    pop rbp
    jmp rax                         ; Tail-call — callee does its own ret

.is_dead:
    ; --- 5. FALLBACK → DefWindowProcW ---
    add rsp, 32
    pop r9
    pop r8
    pop rdx
    pop rcx
    pop rbp
    mov rax, 0x445566778899AABB     ; [OFFSET 96] Pointer to user32.DefWindowProcW
    jmp rax
