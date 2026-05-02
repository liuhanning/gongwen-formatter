
import re

NEW_BODY = """    Dim msg As String
    Dim result As VbMsgBoxResult
    
    msg = ChrW(&H8BF7) & ChrW(&H9009) & ChrW(&H62E9) & ChrW(&H683C) & ChrW(&H5F0F) & ChrW(&H6A21) & ChrW(&H5F0F) & ChrW(&HFF1A) & vbCrLf & vbCrLf
    msg = msg & ChrW(&H3010) & ChrW(&H662F) & ChrW(&H3011) & "= GB/T 9704" & ChrW(&H6807) & ChrW(&H51C6) & ChrW(&H6A21) & ChrW(&H5F0F) & vbCrLf
    msg = msg & "      " & ChrW(&HFF08) & ChrW(&H6807) & ChrW(&H9898) & ChrW(&H5DE6) & ChrW(&H7F29) & ChrW(&H8FDB) & "2" & ChrW(&H5B57) & ChrW(&H7B26) & ChrW(&HFF09) & vbCrLf & vbCrLf
    msg = msg & ChrW(&H3010) & ChrW(&H5426) & ChrW(&H3011) & "= " & ChrW(&H653F) & ChrW(&H5E9C) & ChrW(&H4EA4) & ChrW(&H4ED8) & ChrW(&H7248) & ChrW(&H6A21) & ChrW(&H5F0F) & vbCrLf
    msg = msg & "      " & ChrW(&HFF08) & ChrW(&H6807) & ChrW(&H9898) & ChrW(&H9996) & ChrW(&H884C) & ChrW(&H7F29) & ChrW(&H8FDB) & "2" & ChrW(&H5B57) & ChrW(&H7B26) & ChrW(&HFF09) & vbCrLf & vbCrLf
    msg = msg & ChrW(&H3010) & ChrW(&H53D6) & ChrW(&H6D88) & ChrW(&H3011) & "= " & ChrW(&H53D6) & ChrW(&H6D88) & ChrW(&H64CD) & ChrW(&H4F5C)

    result = MsgBox(msg, vbYesNoCancel + vbQuestion, ChrW(&H9009) & ChrW(&H62E9) & ChrW(&H683C) & ChrW(&H5F0F) & ChrW(&H6A21) & ChrW(&H5F0F))

    If result = vbYes Then
        SelectFormatMode = "standard"
    ElseIf result = vbNo Then
        SelectFormatMode = "government"
    Else
        SelectFormatMode = ""
    End If
End Function"""

def process_file(file_path):
    with open(file_path, 'rb') as f:
        raw = f.read().decode('utf-8')
        raw = raw.replace('\r\n', '\n').replace('\r', '\n')
        
    lines = raw.split('\n')
    new_lines = []
    
    i = 0
    in_func = False
    
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        if "Function SelectFormatMode() As String" in line:
             new_lines.append(line)
             new_lines.append(NEW_BODY)
             
             # Skip until End Function
             i += 1
             while i < len(lines):
                 if lines[i].strip() == "End Function":
                     i += 1
                     break
                 i += 1
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
        new_c = process_file(fp)
        with open(fp, 'wb') as f:
            f.write(new_c.encode('utf-8'))
        print(f"Fixed SelectFormatMode in {fp}")
    except Exception as e:
        print(f"Error {fp}: {e}")
