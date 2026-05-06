@echo off
set NASM_PATH=nasm.exe

echo Compiling Safe Thunks...
%NASM_PATH% -f bin ..\asm\safe_thunk_x64.asm -o ..\asm\safe_thunk_x64.bin
%NASM_PATH% -f bin ..\asm\safe_thunk_x86.asm -o ..\asm\safe_thunk_x86.bin

echo Compiling Endianness Utilities...
%NASM_PATH% -f bin ..\asm\swap_32_x64.asm -o ..\asm\swap_32_x64.bin
%NASM_PATH% -f bin ..\asm\swap_32_x86.asm -o ..\asm\swap_32_x86.bin

echo Compiling WebSocket Masking...
%NASM_PATH% -f bin ..\asm\ws_mask_x64.asm -o ..\asm\ws_mask_x64.bin
%NASM_PATH% -f bin ..\asm\ws_mask_x86.asm -o ..\asm\ws_mask_x86.bin

echo Compiling Memory Utilities...
%NASM_PATH% -f bin ..\asm\mem_zero_x64.asm -o ..\asm\mem_zero_x64.bin
%NASM_PATH% -f bin ..\asm\mem_zero_x86.asm -o ..\asm\mem_zero_x86.bin

echo Compilation complete.
pause
