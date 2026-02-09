
import os

file_path = r"d:\instalar\OrganizadorArchivos\src\app_web.py"

with open(file_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Define ranges to indent (start_line_1_based, end_line_1_based inclusive)
ranges = [
    (6761, 6970), # RIPS
    (6976, 7085), # Visor
    (7090, 7153)  # Gemini
]

for start_line, end_line in ranges:
    start_idx = start_line - 1
    end_idx = end_line # range is exclusive at the end, so line number is correct index
    
    print(f"Indenting lines {start_line} to {end_line}...")
    for i in range(start_idx, end_idx):
        if i < len(lines) and lines[i].strip(): # Only indent non-empty lines
            lines[i] = "    " + lines[i]

with open(file_path, "w", encoding="utf-8") as f:
    f.writelines(lines)

print("Done.")
