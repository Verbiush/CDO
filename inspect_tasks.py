import sys
import os
import json
import sqlite3

# Add src to path
sys.path.append(os.path.join(os.getcwd(), "src"))

try:
    import database
except ImportError:
    from src import database

def inspect_tasks():
    print("Connecting to database...")
    try:
        conn = database.get_connection()
    except Exception as e:
        print(f"Failed to connect: {e}")
        return

    try:
        # SQLite specific query for last 5 tasks
        cursor = conn.cursor()
        cursor.execute("SELECT id, command, params, status, result, created_at FROM tasks ORDER BY id DESC LIMIT 5")
        rows = cursor.fetchall()
        
        print(f"Found {len(rows)} tasks.")
        for row in rows:
            print("-" * 50)
            print(f"ID: {row[0]}")
            print(f"Command: {row[1]}")
            print(f"Params: {row[2]}")
            print(f"Status: {row[3]}")
            
            res_str = row[4]
            if res_str:
                print(f"Result (truncated): {str(res_str)[:200]}...")
                try:
                    res_json = json.loads(res_str)
                    if isinstance(res_json, dict):
                        print(f"Result keys: {list(res_json.keys())}")
                        if "items" in res_json:
                            print(f"Items count: {len(res_json['items'])}")
                        if "errors" in res_json:
                            print(f"Errors: {res_json['errors']}")
                        if "result" in res_json and isinstance(res_json["result"], dict):
                             # Sometimes nested?
                             nested = res_json["result"]
                             if "items" in nested:
                                 print(f"Nested Items count: {len(nested['items'])}")
                except Exception as e:
                    print(f"Error parsing result JSON: {e}")
            else:
                print("Result: None")
            
            print(f"Created At: {row[5]}")

    except Exception as e:
        print(f"Error executing query: {e}")
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    inspect_tasks()
