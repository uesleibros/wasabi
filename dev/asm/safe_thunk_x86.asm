bits 32
global safe_thunk

; stdcall [esp+4], [esp+8], etc...
safe_thunk:
    ; 1. Read the Heartbeat Flag (m_AppIsAlive)
    mov eax, 0x11223344           ; [OFFSET 1] Pointer to m_AppIsAlive
    mov eax, dword [eax]          ; Read the 32-bit value
    test eax, eax                 ; Is it zero?
    jz .is_dead                   ; If zero, jump to fallback.

.is_alive:
    ; 2. VBA is alive. Forward to Wasabi_WndProc
    mov eax, 0x22334455           ; [OFFSET 12] Pointer to Wasabi_WndProc
    jmp eax                       ; Jump (Tail Call)

.is_dead:
    ; 3. VBA is dead. Forward to default Windows handler
    mov eax, 0x33445566           ; [OFFSET 19] Pointer to user32.DefWindowProcW
    jmp eax                       ; Safe jump
