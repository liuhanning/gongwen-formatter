
def fix_spacing(file_path):
    with open(file_path, 'rb') as f:
        content = f.read().decode('utf-8')
        
    # Normalize newlines
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    lines = content.split('\n')
    
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # Check if previous line ended with continuation and this line is empty
        if i > 0 and new_lines[-1].rstrip().endswith('_') and not line.strip():
            # Skip this empty line because it breaks continuation
            pass
        else:
            new_lines.append(line)
        i += 1
        
    # Also, let's just remove ALL empty lines for a cleaner file? 
    # Or at least collapse multiple empty lines.
    # But strictly fixing the syntax error is priority.
    
    # Let's do a second pass to be cleaner:
    # Remove empty lines if the logic suggests we are inside a statement.
    # But the above check `new_lines[-1].endswith('_')` handles the critical case.
    
    # Write back
    with open(file_path, 'wb') as f:
        f.write("\r\n".join(new_lines).encode('utf-8'))

files = [
    r"d:\worksapces\wordhong\GongwenFormatter.bas",
    r"d:\worksapces\wordhong\GongwenFormatter_WPS.bas"
]

for fp in files:
    try:
        fix_spacing(fp)
        print(f"Fixed spacing in {fp}")
    except Exception as e:
        print(f"Error {fp}: {e}")
