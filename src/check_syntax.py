
import sys
import os

try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from tabs import tab_bot_zeus
    print("Syntax: OK")
except Exception as e:
    print(f"Syntax: FAIL ({e})")
