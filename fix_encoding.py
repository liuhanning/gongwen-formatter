
import re

def to_chrw(s):
    result = []
    current_ascii = []
    
    for char in s:
        if ord(char) < 128:
            if char == '"':
                current_ascii.append('""') # Escape quote for VBA string
            else:
                current_ascii.append(char)
        else:
            if current_ascii:
                result.append(f'"{ "".join(current_ascii) }"')
                current_ascii = []
            hex_val = hex(ord(char))[2:].upper().zfill(4)
            result.append(f'ChrW(&H{hex_val})')
            
    if current_ascii:
        result.append(f'"{ "".join(current_ascii) }"')
        
    return " & ".join(result)

def process_line_content(line):
    # This handles the line content to replace string literals
    # Regex to find string literals: "..."
    # We must handle escaped quotes inside: "" 
    pattern = re.compile(r'"((?:""|[^"])*)"')

    def replace_match(match):
        original_content = match.group(1)
        
        if any(ord(c) >= 128 for c in original_content):
            # It has non-ASCII. We MUST process it.
            # To process, we first get the actual semantic string value
            semantic_value = original_content.replace('""', '"')
            return to_chrw(semantic_value)
        else:
            return match.group(0)

    return pattern.sub(replace_match, line)

def process_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    new_lines = []
    
    for line in lines:
        # separate comment
        code_part = line
        comment_part = ""
        
        # Simple comment detection
        chars = list(line)
        in_str = False
        comment_idx = -1
        idx = 0
        while idx < len(chars):
            c = chars[idx]
            if c == '"':
                # Check for escaped quote ""
                if idx + 1 < len(chars) and chars[idx+1] == '"':
                    idx += 1 # skip next
                else:
                    in_str = not in_str
            elif c == "'" and not in_str:
                comment_idx = idx
                break
            idx += 1
            
        if comment_idx != -1:
            code_part = line[:comment_idx]
            comment_part = line[comment_idx:]
        else:
            code_part = line
            comment_part = ""
            
        # Process code_part
        new_code_part = process_line_content(code_part)
        
        # Re-assemble
        new_lines.append(new_code_part + comment_part)
        
    return '\n'.join(new_lines)

input_file = r"d:\worksapces\wordhong\GongwenFormatter_WPS.bas"
new_content = process_file(input_file)

with open(input_file, 'w', encoding='utf-8') as f:
    f.write(new_content)

print(f"Processed {input_file}")
