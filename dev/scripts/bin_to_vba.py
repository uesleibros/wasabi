import os

def generate_vba_hex(filename):
    if not os.path.exists(filename):
        print(f"Error: {filename} not found.")
        return
    
    with open(filename, "rb") as f:
        bytes_data = f.read()
    
    # Format like: asm(0) = &H48: asm(1) = &HB8...
    vba_lines = []
    for i, byte in enumerate(bytes_data):
        vba_lines.append(f"asm({i}) = &H{byte:02X}")
    
    return ": ".join(vba_lines)

print("--- x64 OPCODES ---")
print(generate_vba_hex("../asm/safe_thunk_x64.bin"))
print("\n--- x86 OPCODES ---")
print(generate_vba_hex("../asm/safe_thunk_x86.bin"))
