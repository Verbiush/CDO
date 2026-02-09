
import os

file_path = r"d:\instalar\OrganizadorArchivos\src\bot_zeus.py"

with open(file_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

new_lines = []
inside_loop = False
loop_start_index = -1
except_block_index = -1

for i, line in enumerate(lines):
    # Detect start of loop
    if "for index, row in df.iterrows():" in line and not inside_loop:
        # Check if already indented (to avoid double application)
        if line.startswith("        for"): 
            print("Already indented. Exiting.")
            exit()
            
        # Insert try block
        indent = "    "
        new_lines.append(indent + "try:\n")
        inside_loop = True
        loop_start_index = i
        
        # Indent the for loop line
        new_lines.append(indent + line)
        continue

    # Detect end of loop (the outer except block)
    if "except Exception as e_gral:" in line and inside_loop:
        inside_loop = False
        except_block_index = i
        new_lines.append(line)
        continue

    # Indent lines inside the loop
    if inside_loop:
        if line.strip() == "":
            new_lines.append(line)
        else:
            new_lines.append("    " + line)
    else:
        new_lines.append(line)

with open(file_path, "w", encoding="utf-8") as f:
    f.writelines(new_lines)

print("Indentation fixed successfully.")
