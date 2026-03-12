
import sys
import os
import logging

# Setup logging to console
logging.basicConfig(level=logging.DEBUG)

# Add src to path to import local_agent
sys.path.append(os.path.join(os.getcwd(), "src", "local_agent"))

try:
    from main import process_search_files
except ImportError:
    # Try direct import if main.py is in src/local_agent
    sys.path.append(os.path.join(os.getcwd(), "src"))
    from local_agent.main import process_search_files

path = "G:/luisa/EMSSANAR/FEBRERO/SUBSIDIADO 1-15"
patterns = ["*"] # Search for everything
exclusion = []

print(f"Testing search in: {path}")

if not os.path.exists(path):
    print(f"ERROR: Path does not exist: {path}")
else:
    print("Path exists.")

# Test 1: Recursive, Items=both
print("\n--- Test 1: Recursive, Both ---")
result = process_search_files(
    path=path,
    patterns=patterns,
    exclusion_list=exclusion,
    search_by="name",
    item_type="both",
    recursive=True,
    search_empty_folders=False
)

print(f"Items found: {len(result['items'])}")
print(f"Errors: {result['errors']}")
if len(result['items']) > 0:
    print("First 5 items:")
    for item in result['items'][:5]:
        print(item)

# Test 2: Non-recursive
print("\n--- Test 2: Non-Recursive ---")
result = process_search_files(
    path=path,
    patterns=patterns,
    exclusion_list=exclusion,
    search_by="name",
    item_type="both",
    recursive=False,
    search_empty_folders=False
)
print(f"Items found: {len(result['items'])}")
