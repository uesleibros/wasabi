@echo off
set NASM_PATH=nasm.exe

echo Compiling x64 Thunk...
%NASM_PATH% -f bin ..\asm\safe_thunk_x64.asm -o ..\asm\safe_thunk_x64.bin

echo Compiling x86 Thunk...
%NASM_PATH% -f bin ..\asm\safe_thunk_x86.asm -o ..\asm\safe_thunk_x86.bin

echo Compilation complete.
pause
