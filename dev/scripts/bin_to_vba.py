import os

def generate_vba_hex(filepath):
    if not os.path.exists(filepath):
        return None
    
    filename = os.path.basename(filepath)
    with open(filepath, "rb") as f:
        bytes_data = f.read()
    
    vba_lines = [f"asm({i}) = &H{byte:02X}" for i, byte in enumerate(bytes_data)]
    
    return {
        "name": filename,
        "count": len(bytes_data),
        "code": ": ".join(vba_lines)
    }

binaries = [
    "safe_thunk_x64.bin", "safe_thunk_x86.bin",
    "swap_32_x64.bin", "swap_32_x86.bin",
    "ws_mask_x64.bin", "ws_mask_x86.bin",
    "mem_zero_x64.bin", "mem_zero_x86.bin"
]

asm_dir = "../asm/"

print("--- WASABI OPCODE EXTRACTOR ---")
for bin_file in binaries:
    full_path = os.path.join(asm_dir, bin_file)
    result = generate_vba_hex(full_path)
    
    if result:
        print(f"\n[ FILE: {result['name']} | SIZE: {result['count']} bytes ]")
        print(result['code'])
    else:
        print(f"\n[ SKIP: {bin_file} (not found) ]")
