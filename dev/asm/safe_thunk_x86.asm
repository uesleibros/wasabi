bits 32
global safe_thunk_x86

; stdcall signature. Parameters are on the stack, not in registers.
safe_thunk_x86:
    ; --- SAVE VOLATILE REGISTERS ---
    push eax                        
    push ecx                        
    push edx                        

    ; --- 1. HEARTBEAT FLAG CHECK (m_AppIsAlive) ---
    mov eax, 0x11223344             ; [OFFSET 4] Pointer to m_AppIsAlive
    mov eax, dword [eax]            ; Read the 32-bit value
    test eax, eax                   ; Is it zero?
    jz .is_dead                     ; If zero (VBA stopped), jump to fallback handler

    ; --- 2. IDE STATE CHECK (EbMode) ---
    mov eax, 0x22334455             ; [OFFSET 15] Pointer to vba6.dll!EbMode
    test eax, eax                   ; Did we fail to find EbMode?
    jz .skip_ebmode                 ; If null, skip the IDE check and proceed normally
    
    call eax                        ; Call EbMode()
    cmp eax, 1                      ; EbMode returns 1 if running normally
    jne .is_dead                    ; If not 1 (paused or editing), jump to fallback handler

.skip_ebmode:
    ; --- 3. VBA IS SAFE & RUNNING (Forward to WasabiAsyncWndProc) ---
    pop edx                         ; Restore registers
    pop ecx                         
    pop eax                         
    
    mov eax, 0x33445566             ; [OFFSET 34] Pointer to WasabiAsyncWndProc
    jmp eax                         ; Jump (Tail Call) to VBA handler

.is_dead:
    ; --- 4. VBA IS DEAD/PAUSED (Forward to Default Windows Handler) ---
    pop edx                         ; Restore registers
    pop ecx                         
    pop eax                         
    
    mov eax, 0x44556677             ; [OFFSET 44] Pointer to user32.DefWindowProcW
    jmp eax                         ; Safe jump to Windows to discard the message
