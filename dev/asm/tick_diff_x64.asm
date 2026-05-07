bits 64
global tick_diff

; CallWindowProcW signature (x64):
; RCX = startTick (P1)
; RDX = endTick   (P2)
; R8  = unused    (P3)
; R9  = unused    (P4)
; Returns tick difference in RAX (Sign-extended from EAX)

tick_diff:
    mov eax, r8d       ; EAX = endTick (P2)
    sub eax, edx       ; EAX = EAX - startTick (P1)
    
    cdqe               ; Sign-extend EAX into RAX. Prevents VBA CLng() overflow 
                       ; if the difference is negative.
    
    ret                ; Return RAX.
