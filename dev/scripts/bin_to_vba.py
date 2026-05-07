import os
import struct

ASM_DIR = "../asm/"
OUTPUT_FILE = "opcodes_output.txt"

BINARIES = {
    "x64": ["ws_mask_x64.bin", "mem_zero_x64.bin", "mem_find_x64.bin"],
    "x86": ["ws_mask_x86.bin", "mem_zero_x86.bin", "mem_find_x86.bin"],
}

FUNC_NAMES = {
    "ws_mask":  "GetWsMaskOpcodes",
    "mem_zero": "GetMemZeroOpcodes",
    "mem_find": "GetMemFindOpcodes",
}

def get_func_name(filename):
    base = filename.replace("_x64.bin", "").replace("_x86.bin", "")
    return FUNC_NAMES.get(base, base)

def bytes_to_vba_function(filepath, arch):
    if not os.path.exists(filepath):
        print(f"  [SKIP] {filepath} not found")
        return None

    with open(filepath, "rb") as f:
        raw = f.read()

    count = len(raw)
    filename = os.path.basename(filepath)
    func_name = get_func_name(filename.replace(".bin", "").replace(f"_{arch}", ""))
    full_func_name = f"{func_name}_{arch}"
    
    padded = raw + b'\x00' * ((4 - count % 4) % 4)
    longs = struct.unpack(f"<{len(padded)//4}L", padded)

    lines = []
    lines.append(f"Private Function {full_func_name}() As Byte()")
    lines.append(f"    Dim opcodes(0 To {count - 1}) As Byte")
    
    hex_vals = [f"&H{b:02X}" for b in raw]
    chunk_size = 8
    chunks = [hex_vals[i:i+chunk_size] for i in range(0, len(hex_vals), chunk_size)]
    
    lines.append(f"    Dim HexStr As Variant: HexStr = Array( _")
    for i, chunk in enumerate(chunks):
        joined = ", ".join(chunk)
        suffix = " _" if i < len(chunks) - 1 else ")"
        lines.append(f"        {joined}{suffix}")
    
    lines.append(f"    Dim i As Long")
    lines.append(f"    For i = 0 To {count - 1}: opcodes(i) = CByte(HexStr(i)): Next i")
    lines.append(f"    {full_func_name} = opcodes")
    lines.append(f"End Function")
    lines.append("")

    return {
        "name": filename,
        "arch": arch,
        "count": count,
        "longs": longs,
        "code": "\n".join(lines)
    }

def generate_summary_table(results):
    lines = ["' === OPCODE SUMMARY ==="]
    lines.append(f"' {'File':<30} {'Bytes':>6}  {'Longs':>6}")
    lines.append(f"' {'-'*30}  {'-'*6}  {'-'*6}")
    for r in results:
        if r:
            num_longs = (r['count'] + 3) // 4
            lines.append(f"' {r['name']:<30} {r['count']:>6}  {num_longs:>6}")
    return "\n".join(lines)

def main():
    results = []
    output_parts = []

    print("--- WASABI OPCODE EXTRACTOR ---\n")

    for arch, files in BINARIES.items():
        output_parts.append(f"' ===== {arch.upper()} =====\n")
        for bin_file in files:
            full_path = os.path.join(ASM_DIR, bin_file)
            result = bytes_to_vba_function(full_path, arch)
            results.append(result)

            if result:
                print(f"[OK] {result['name']:30} {result['count']:4} bytes")
                output_parts.append(result['code'])
            else:
                print(f"[--] {bin_file}")

    summary = generate_summary_table([r for r in results if r])
    output_parts.insert(0, summary + "\n\n")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(output_parts))

    print(f"\nOutput salvo em: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
