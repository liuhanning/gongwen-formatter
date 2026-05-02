
import re

def parse_vb_msgbox(full_stmt):
    # Retrieve the content inside MsgBox
    # Stmt: MsgBox content, type, title
    
    # Remove continuations
    clean_stmt = full_stmt.replace(' _\n', '').replace(' _\r\n', '').replace(' _\r', '').replace('\r', '').replace('\n', '')
    clean_stmt = re.sub(r'\s+', ' ', clean_stmt)
    
    # Find first comma at top level
    in_quote = False
    paren_depth = 0
    comma_idx = -1
    
    # Start after "MsgBox "
    start_idx = clean_stmt.find("MsgBox ") + 7
    
    for i in range(start_idx, len(clean_stmt)):
        c = clean_stmt[i]
        if c == '"':
            # Handle escaped quotes ""
            if i+1 < len(clean_stmt) and clean_stmt[i+1] == '"':
                pass # Skip next in loop?? No, just simple toggle is dangerous.
                # If we see ", we toggle.
                # If we see "", we see " (toggle on), then " (toggle off). 
                # Wait, "" inside string is just a character.
                # In VBA "foo""bar" -> 
                # i: " (on), f, o, o, " (off?), " (on?), b, a, r, " (off?) -> OK?
                # Actually, strictly:
                # " puts us in string.
                # Next " might be end OR escape.
                # If next char is ", it's escape.
                pass 
            in_quote = not in_quote
        elif c == '(' and not in_quote:
            paren_depth += 1
        elif c == ')' and not in_quote:
            paren_depth -= 1
        elif c == ',' and not in_quote and paren_depth == 0:
            comma_idx = i
            break
            
    if comma_idx == -1:
        return None, None
        
    content = clean_stmt[start_idx:comma_idx].strip()
    rest = clean_stmt[comma_idx+1:].strip()
    return content, rest

def process_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # Check for MsgBox that is complex (has ChrW and line continuation)
        if stripped.startswith('MsgBox ') and 'ChrW(' in line and line.rstrip().endswith('_'):
            # Collect full block
            buffer = []
            indent = re.match(r'^\s*', line).group(0)
            
            while i < len(lines):
                l = lines[i].rstrip()
                buffer.append(lines[i]) # Keep original for backup
                if not l.endswith('_'):
                    i += 1
                    break
                i += 1
            
            full_text = "\n".join(buffer)
            msg_content, rest_args = parse_vb_msgbox(full_text)
            
            if msg_content:
                # Refactor
                new_block = []
                new_block.append(f"{indent}Dim msg As String")
                
                # Split content by '&'
                # We need to be careful about '&' inside strings
                # But our complex content is mostly ChrW(...) & ChrW(...) & "..."
                # Simple split by ' & ' might be mostly safe if we assume standard spacing
                # better: split by '&' and strip, checking for unbalanced quotes?
                # Actually we can use the same quote logic.
                
                parts = []
                current_part = []
                in_quote = False
                
                for char in msg_content:
                    if char == '"':
                         # simplistic quote checking
                         in_quote = not in_quote
                    
                    if char == '&' and not in_quote:
                        parts.append("".join(current_part).strip())
                        current_part = []
                    else:
                        current_part.append(char)
                parts.append("".join(current_part).strip())
                
                # Build statements
                current_line_parts = []
                current_len = 0
                is_first = True
                
                for p in parts:
                    if not p: continue
                    # approximate length
                    if current_len + len(p) > 100: # Safe line length
                        # Flush
                        expr = " & ".join(current_line_parts)
                        if is_first:
                            new_block.append(f"{indent}msg = {expr}")
                            is_first = False
                        else:
                            new_block.append(f"{indent}msg = msg & {expr}")
                        current_line_parts = [p]
                        current_len = len(p)
                    else:
                        current_line_parts.append(p)
                        current_len += len(p)
                
                if current_line_parts:
                    expr = " & ".join(current_line_parts)
                    if is_first:
                         new_block.append(f"{indent}msg = {expr}")
                    else:
                         new_block.append(f"{indent}msg = msg & {expr}")
                
                new_block.append(f"{indent}MsgBox msg, {rest_args}")
                new_lines.extend(new_block)
            else:
                # Fallback
                new_lines.extend(buffer)
        else:
            new_lines.append(line)
            i += 1
            
    return "\n".join(new_lines)

files = [
    r"d:\worksapces\wordhong\GongwenFormatter.bas",
    r"d:\worksapces\wordhong\GongwenFormatter_WPS.bas"
]

for fp in files:
    try:
        new_content = process_file(fp)
        with open(fp, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Processed {fp}")
    except Exception as e:
        print(f"Error processing {fp}: {e}")
