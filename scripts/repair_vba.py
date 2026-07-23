
import re

def repair_content(content):
    lines = content.split('\n')
    new_lines = []
    
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # Detect start of a broken MsgBox block
        if stripped == "Dim msg As String":
            # Collect the block
            block_lines = []
            indent = re.match(r'^\s*', line).group(0)
            
            # This is the Dim line
            header_line = line
            
            i += 1
            # Collect msg = ... lines
            rhs_parts = []
            
            # We assume the block is contiguous
            while i < len(lines):
                l = lines[i]
                s = l.strip()
                if s.startswith("msg ="):
                    # Extract RHS
                    # format: msg = ... OR msg = msg & ...
                    # We want the content added.
                    
                    # Remove "msg = "
                    val = s[5:].strip()
                    # Remove "msg & " if present
                    if val.startswith("msg &"):
                        val = val[5:].strip()
                        
                    # Now val is the chunk added in this line.
                    rhs_parts.append(val)
                    i += 1
                elif s.startswith("MsgBox msg"):
                    # End of block
                    msgbox_line = l
                    i += 1
                    break
                else:
                    # Unexpected line? 
                    # If empty or unrelated, break block?
                    # Previous script shouldn't have put empty lines in block
                    if not s: # Empty line
                        i += 1 
                        continue
                    else:
                        # Broken block?
                        break
            
            # Now reconstruct full expression
            # We joined them with presumed "&" logic in previous broken script.
            # But here `rhs_parts` are the chunks.
            # We should join them with " & "
            full_expr = " & ".join(rhs_parts)
            
            # Fix artifacts:
            # 1. `( & H` -> `(&H`  (caused by splitting ChrW(&H...))
            # 2. `& & H` -> `&H`   (caused by splitting ChrW(&...))
            # 3. ` & H` where it should be part of &H?
            #    If we had `ChrW( & H...` -> joined with ` & ` -> `ChrW( &  & H...`?
            #    Let's look at `rhs_parts`.
            #    Part 1: `ChrW(`
            #    Part 2: `H5DF2)`
            #    Joined: `ChrW( & H5DF2)`
            #    Fix: `( & H` -> `(&H`
            
            full_expr = full_expr.replace('( & H', '(&H')
            full_expr = full_expr.replace('(& H', '(&H')
            
            # Also `ChrW` might be `ChrW ( & H...`
            full_expr = full_expr.replace('ChrW( & H', 'ChrW(&H')
            
            # Sanity check: Regex to find all functional units
            # ChrW(&H....)
            # "..."
            # Constants
            
            # We can use regex to extract all VALID tokens
            # Token: ChrW\(&H[0-9A-F]+\)
            # Token: "[^"]*"
            # Token: vb[A-Za-z0-9]+
            # Token: modeName (variable)
            # Token: Err.Description/Number
            
            # Pattern matching tokens
            token_pattern = re.compile(r'(ChrW\(&H[0-9A-F]+\)|"[^"]*"|vb[A-Za-z0-9]+|[a-zA-Z0-9_\.]+|&)')
            
            # Actually, we can just split by " & " now that we fixed `(&H`.
            # But let's be safer.
            # We want to break `full_expr` into chunks for lines.
            
            # Recalculate chunks
            # Split by ` & ` is safe now if `(&H` is fixed?
            # `ChrW(&H1234)` -> contains `&H`. No space around &.
            # `... & ...` -> space around &.
            
            # So split by ` & ` (with spaces)
            tokens = full_expr.split(' & ')
            
            # Generate new lines
            new_block = []
            new_block.append(header_line)
            
            current_line_parts = []
            current_len = 0
            is_first = True
            
            for t in tokens:
                t = t.strip()
                if not t: continue
                
                # Fix any residual damage in token?
                # e.g. `H5DF2)` -> should have been attached?
                # If we fixed `(&H`, splits should be correct?
                # What if `ChrW(&H1234)` was split nicely? `ChrW(&H1234)`.
                # What if `ChrW(&H1234)` was split as `ChrW` `(` `&H1234` `)`?
                # No, previous split was by `&` character.
                
                if current_len + len(t) > 100:
                    expr = " & ".join(current_line_parts)
                    if is_first:
                        new_block.append(f"{indent}msg = {expr}")
                        is_first = False
                    else:
                        new_block.append(f"{indent}msg = msg & {expr}")
                    current_line_parts = [t]
                    current_len = len(t)
                else:
                    current_line_parts.append(t)
                    current_len += len(t)
            
            if current_line_parts:
                expr = " & ".join(current_line_parts)
                if is_first:
                     new_block.append(f"{indent}msg = {expr}")
                else:
                     new_block.append(f"{indent}msg = msg & {expr}")
            
            # Add MsgBox line (msgbox_line contains the rest of args)
            # We need to make sure we didn't lose `MsgBox msg, ...`
            # In extraction: `elif s.startswith("MsgBox msg"):`
            if 'msgbox_line' in locals():
                new_block.append(msgbox_line)
            else:
                # Fallback if we didn't find end?
                # Should not happen if block logic is finding existing code
                pass
            
            new_lines.extend(new_block)
            
        else:
            new_lines.append(line)
            i += 1
            
    return "\r\n".join(new_lines)

files = [
    r"d:\worksapces\wordhong\GongwenFormatter.bas",
    r"d:\worksapces\wordhong\GongwenFormatter_WPS.bas"
]

for fp in files:
    try:
        with open(fp, 'rb') as f:
            raw = f.read().decode('utf-8')
            # Normalize to \n
            raw = raw.replace('\r\n', '\n').replace('\r', '\n')
            
        fixed = repair_content(raw)
        
        with open(fp, 'wb') as f:
            f.write(fixed.encode('utf-8'))
        print(f"Repaired {fp}")
    except Exception as e:
        print(f"Error {fp}: {e}")
