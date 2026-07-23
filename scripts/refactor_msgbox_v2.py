
import re

def parse_vb_msgbox(full_stmt):
    # Remove continuations
    clean_stmt = full_stmt.replace(' _\n', '').replace('\n', '')
    clean_stmt = re.sub(r'\s+', ' ', clean_stmt)
    
    # Find first comma at top level
    in_quote = False
    paren_depth = 0
    comma_idx = -1
    
    start_search = clean_stmt.find("MsgBox ")
    if start_search == -1: return None, None
    start_idx = start_search + 7
    
    for i in range(start_idx, len(clean_stmt)):
        c = clean_stmt[i]
        if c == '"':
            # toggle quote
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
    # Read binary to handle line endings explicitly
    with open(file_path, 'rb') as f:
        raw = f.read()
    
    # Detection of double returns
    # normalize to \n
    text = raw.decode('utf-8')
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # If we have double newlines everywhere, we might want to reduce them?
    # But usually code implies single spacing. 
    # Let's assume the previous corruption made it \n\n.
    # We can check if most lines are followed by an empty line.
    lines = text.split('\n')
    
    # Check emptiness density
    empty_count = sum(1 for l in lines if not l.strip())
    if len(lines) > 10 and empty_count > len(lines) * 0.4:
        # Heuristic: Remove empty lines usually
        # But valid empty lines exist.
        # Better: if we see Code \n \n Code, it's likely bad.
        # Let's just strip empty lines that are following a line ending with _ 
        # No, that's specific to the bug.
        # Let's just naive remove empty lines for now? No.
        # Let's collapse.
        pass

    # Actually, simpler: My previous script wrote `\r\n` in text mode -> `\r\r\n`.
    # `text.replace('\r\n', '\n')` handles `\r\r\n` -> `\r\n` -> `\n`?
    # `\r\r\n` in python text mode read:
    # `\r\r\n` -> `\r\n` (first \r is char, \r\n is linebreak) -> `\r` `\n`
    # So `text` has `\n\n` or `\r\n` (mixed).
    # replace `\r` with `\n` -> `\n\n`
    # replace `\n\n` with `\n`?
    
    # Let's aggressively normalise: replace multiple \n with single \n?
    # No, that removes intentional spacing.
    # How about: replace `\n\n` with `\n` ONLY IF line ends with `_`?
    # No, the previous script added empty lines EVERYWHERE.
    
    refined_lines = []
    # If we see alternate empty lines
    # Code
    # 
    # Code
    # 
    # It is safe to remove them.
    # Let's attempt to remove blank lines that appear to be duplicates.
    
    # For this specific task, let's just process the lines we have.
    # The `MsgBox ... _` followed by `\n` followed by `...` is the issue.
    # We will specifically look for `_` at end of line, and SKIP the next line if it is empty.
    
    # Re-assemble text first?
    # I'll just iterate manually.
    
    processed_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # Logic for MsgBox refactoring
        if stripped.startswith('MsgBox ') and 'ChrW(' in line and stripped.endswith('_'):
            # Collect
            buffer = []
            indent = re.match(r'^\s*', line).group(0)
            
            # Start collection
            # We need to handle the "gap" lines here too.
            original_block_lines = [] # for fallback
            
            valid_parts = []
            
            # Add first line
            buffer.append(line)
            valid_parts.append(line)
            original_block_lines.append(line)
            
            j = i + 1
            while j < len(lines):
                next_l = lines[j]
                original_block_lines.append(next_l)
                
                if not next_l.strip(): 
                    # Skip empty line inside continuation
                    j += 1
                    continue
                
                valid_parts.append(next_l)
                if not next_l.rstrip().endswith('_'):
                    # End of block
                    j += 1
                    break
                j += 1
            
            # Prepare full text from valid_parts
            full_stmt = "\n".join(valid_parts)
            msg_content, rest_args = parse_vb_msgbox(full_stmt)
            
            if msg_content:
                # Refactor!
                # Create the new block
                refactored_block = []
                refactored_block.append(f"{indent}Dim msg As String")
                
                # Split content logic (copy from before)
                chunk_parts = []
                # Split by `&` carefully
                temp_buf = ""
                in_q = False
                for char in msg_content:
                    if char == '"': in_q = not in_q
                    if char == '&' and not in_q:
                        chunk_parts.append(temp_buf.strip())
                        temp_buf = ""
                    else:
                        temp_buf += char
                chunk_parts.append(temp_buf.strip())
                
                current_line_parts = []
                current_len = 0
                is_first = True
                
                for p in chunk_parts:
                    if not p: continue
                    clean_p = p.strip()
                    if current_len + len(clean_p) > 120:
                         stmt = " & ".join(current_line_parts)
                         if is_first:
                             refactored_block.append(f"{indent}msg = {stmt}")
                             is_first = False
                         else:
                             refactored_block.append(f"{indent}msg = msg & {stmt}")
                         current_line_parts = [clean_p]
                         current_len = len(clean_p)
                    else:
                         current_line_parts.append(clean_p)
                         current_len += len(clean_p)
                         
                if current_line_parts:
                    stmt = " & ".join(current_line_parts)
                    if is_first:
                         refactored_block.append(f"{indent}msg = {stmt}")
                    else:
                         refactored_block.append(f"{indent}msg = msg & {stmt}")
                
                refactored_block.append(f"{indent}MsgBox msg, {rest_args}")
                processed_lines.extend(refactored_block)
                i = j 
            else:
                 # Clean up the original block? remove empty lines?
                 # Yes, let's append valid_parts instead of original
                 processed_lines.extend(valid_parts)
                 i = j
        else:
            # check for double spacing issue
            # If this line is empty, and previous was not empty?
            # Or if this line is empty and next is empty?
            # Let's just keep it for now unless we are sure.
            # actually if we just want to remove ALL double spacing:
            if not line.strip() and i > 0 and not lines[i-1].strip():
                 # Don't add multiple empty lines
                 pass
            else:
                 processed_lines.append(line)
            i += 1
            
    return "\r\n".join(processed_lines) # Write as CRLF

files = [
    r"d:\worksapces\wordhong\GongwenFormatter.bas",
    r"d:\worksapces\wordhong\GongwenFormatter_WPS.bas"
]

for fp in files:
    try:
        new_c = process_file(fp)
        with open(fp, 'wb') as f:
            f.write(new_c.encode('utf-8'))
        print(f"Refactored {fp}")
    except Exception as e:
        print(f"Error {fp}: {e}")
