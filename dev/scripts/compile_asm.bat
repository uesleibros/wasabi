@echo off
set NASM_PATH=nasm.exe

echo Compiling WebSocket Masking...
%NASM_PATH% -f bin ..\asm\ws_mask_x64.asm -o ..\asm\ws_mask_x64.bin
%NASM_PATH% -f bin ..\asm\ws_mask_x86.asm -o ..\asm\ws_mask_x86.bin

echo Compiling Memory Utilities...
%NASM_PATH% -f bin ..\asm\mem_zero_x64.asm -o ..\asm\mem_zero_x64.bin
%NASM_PATH% -f bin ..\asm\mem_zero_x86.asm -o ..\asm\mem_zero_x86.bin

echo Compiling Find Memory Utilities...
%NASM_PATH% -f bin ..\asm\mem_find_x64.asm -o ..\asm\mem_find_x64.bin
%NASM_PATH% -f bin ..\asm\mem_find_x86.asm -o ..\asm\mem_find_x86.bin

echo Compilation complete.
pause
