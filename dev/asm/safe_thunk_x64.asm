bits 64
global safe_thunk

; CallWindowProcW/WndProc signature (x64):
; RCX = hWnd, RDX = uMsg, R8 = wParam, R9 = lParam
safe_thunk:
    ; 1. Read the Heartbeat Flag (m_AppIsAlive)
    mov rax, 0x1122334455667788   ; [OFFSET 2] Pointer to the m_AppIsAlive variable
    mov eax, dword [rax]          ; Read the 32-bit value (Long)
    test eax, eax                 ; Is it zero?
    jz .is_dead                   ; If zero, VBA was reset. Jump to fallback.

.is_alive:
    ; 2. VBA is alive. Forward to Wasabi_WndProc
    mov rax, 0x2233445566778899   ; [OFFSET 18] Pointer to Wasabi_WndProc
    jmp rax                       ; Jump (Tail Call)

.is_dead:
    ; 3. VBA is dead. Forward to default Windows handler
    mov rax, 0x33445566778899AA   ; [OFFSET 30] Pointer to user32.DefWindowProcW
    jmp rax                       ; Safe jump
