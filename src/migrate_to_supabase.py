import sqlite3
import psycopg2
import os
import sys
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
# Source: Local SQLite or JSON
SOURCE_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.db")
SOURCE_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")

# Destination: Supabase (PostgreSQL)
SUPABASE_HOST = os.getenv("SUPABASE_HOST")
SUPABASE_USER = os.getenv("SUPABASE_USER")
SUPABASE_PASSWORD = os.getenv("SUPABASE_PASSWORD")
SUPABASE_DB = os.getenv("SUPABASE_DB")
SUPABASE_PORT = os.getenv("SUPABASE_PORT", "5432")

def get_source_users():
    """Retrieves users from SQLite or JSON fallback."""
    users = []
    
    # Try SQLite first
    if os.path.exists(SOURCE_DB):
        try:
            conn = sqlite3.connect(SOURCE_DB)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users")
            rows = cursor.fetchall()
            conn.close()
            
            if rows:
                print(f"Found {len(rows)} users in SQLite database.")
                return [dict(row) for row in rows]
        except Exception as e:
            print(f"Error reading SQLite DB: {e}")

    # Try JSON if SQLite failed or empty
    if os.path.exists(SOURCE_JSON):
        try:
            with open(SOURCE_JSON, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"Found {len(data)} users in JSON file.")
                # Convert JSON dict format to list of dicts matching DB schema
                users_list = []
                for username, user_data in data.items():
                    user_dict = {
                        'username': username,
                        'password': user_data.get('password'),
                        'role': user_data.get('role', 'user'),
                        'last_path': user_data.get('last_path'),
                        'permissions': json.dumps(user_data.get('permissions')) if isinstance(user_data.get('permissions'), (dict, list)) else user_data.get('permissions'),
                        'favorites': json.dumps(user_data.get('favorites')) if isinstance(user_data.get('favorites'), (list)) else user_data.get('favorites'),
                        'config': json.dumps(user_data.get('config')) if isinstance(user_data.get('config'), (dict)) else user_data.get('config')
                    }
                    users_list.append(user_dict)
                return users_list
        except Exception as e:
            print(f"Error reading JSON file: {e}")
            
    return []

def get_supabase_connection():
    if not all([SUPABASE_HOST, SUPABASE_USER, SUPABASE_PASSWORD, SUPABASE_DB]):
        print("Missing Supabase configuration. Please set SUPABASE_HOST, SUPABASE_USER, SUPABASE_PASSWORD, and SUPABASE_DB environment variables.")
        sys.exit(1)
    
    try:
        return psycopg2.connect(
            host=SUPABASE_HOST,
            user=SUPABASE_USER,
            password=SUPABASE_PASSWORD,
            dbname=SUPABASE_DB,
            port=SUPABASE_PORT
        )
    except Exception as e:
        print(f"Error connecting to Supabase: {e}")
        sys.exit(1)

def run_migration():
    print("--- Starting Migration to Supabase ---")
    
    # 1. Get Source Users
    print("Getting users from local source...")
    users = get_source_users()
    if not users:
        print("No users found in local SQLite or JSON. Nothing to migrate.")
        # Optional: Create default admin if really empty?
        # For now, just exit.
        return

    # 2. Connect to Destination
    print("Connecting to Supabase...")
    dest_conn = get_supabase_connection()
    dest_cursor = dest_conn.cursor()
    
    # 3. Create Tables in Supabase
    print("Creating tables in Supabase...")
    schema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema_postgres.sql")
    try:
        with open(schema_path, 'r', encoding='utf-8') as f:
            schema_sql = f.read()
            dest_cursor.execute(schema_sql)
            dest_conn.commit()
        print("Tables created successfully.")
    except Exception as e:
        print(f"Error creating tables: {e}")
        dest_conn.close()
        sys.exit(1)
    
    # 4. Migrate Users
    print(f"Migrating {len(users)} users...")
    try:
        migrated_count = 0
        for user_dict in users:
            username = user_dict.get('username')
            print(f"Migrating user: {username}")
            
            # Check if user exists
            dest_cursor.execute("SELECT username FROM users WHERE username = %s", (username,))
            if dest_cursor.fetchone():
                print(f"  User {username} already exists. Skipping.")
                continue
                
            # Insert
            dest_cursor.execute(
                """
                INSERT INTO users (username, password, role, last_path, permissions, favorites, config)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    user_dict.get('username'),
                    user_dict.get('password'),
                    user_dict.get('role', 'user'),
                    user_dict.get('last_path'),
                    # Ensure JSON fields are strings if they came from SQLite as strings, or JSON dumps if from dict
                    user_dict.get('permissions') if isinstance(user_dict.get('permissions'), str) else json.dumps(user_dict.get('permissions')),
                    user_dict.get('favorites') if isinstance(user_dict.get('favorites'), str) else json.dumps(user_dict.get('favorites')),
                    user_dict.get('config') if isinstance(user_dict.get('config'), str) else json.dumps(user_dict.get('config'))
                )
            )
            migrated_count += 1
            
        dest_conn.commit()
        print(f"Successfully migrated {migrated_count} users.")
        
    except Exception as e:
        print(f"Error migrating users: {e}")
        dest_conn.rollback()
    
    # 5. Clean up
    dest_conn.close()
    print("--- Migration Complete ---")
    print("Note: Only 'users' table data was migrated. Other tables (patients, invoices) are empty in the new database.")

if __name__ == "__main__":
    run_migration()
