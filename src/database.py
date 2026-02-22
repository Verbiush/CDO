import sqlite3
import json
import os
import threading

DB_FILE = os.path.join(os.path.dirname(__file__), "users.db")
USERS_JSON = os.path.join(os.path.dirname(__file__), "users.json")

# Lock for thread safety within the same process
_db_lock = threading.Lock()

def get_connection():
    """Returns a connection to the SQLite database."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the database table and migrates from JSON if needed."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        
        # Create users table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT DEFAULT 'user',
            last_path TEXT,
            permissions TEXT,
            favorites TEXT,
            config TEXT
        )
        ''')
        conn.commit()
        
        # Create tasks table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            command TEXT NOT NULL,
            params TEXT,
            status TEXT DEFAULT 'PENDING',
            result TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        conn.commit()

        # Create document records table (Gestión Documental)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS document_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nro_estudio TEXT,
            descripcion TEXT,
            eps TEXT,
            tipo_doc TEXT,
            no_doc TEXT,
            nombre_completo TEXT,
            nombre_tercero TEXT,
            fecha_ingreso TEXT,
            fecha_salida TEXT,
            autorizacion TEXT,
            no_factura TEXT,
            fecha_factura TEXT,
            tipo_pago TEXT,
            valor_servicio TEXT,
            copago TEXT,
            total TEXT,
            regimen TEXT DEFAULT 'SUBSIDIADO',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'PENDING'
        )
        ''')
        conn.commit()

        # Check for missing columns in document_records (Migration)
        cursor.execute("PRAGMA table_info(document_records)")
        columns = [info[1] for info in cursor.fetchall()]
        if "regimen" not in columns:
            print("Migrating document_records: Adding 'regimen' column...")
            cursor.execute("ALTER TABLE document_records ADD COLUMN regimen TEXT DEFAULT 'SUBSIDIADO'")
            conn.commit()
        
        # Check if we need to migrate from JSON
        cursor.execute("SELECT count(*) FROM users")
        count = cursor.fetchone()[0]
        
        if count == 0 and os.path.exists(USERS_JSON):
            print("Migrating users from JSON to SQLite...")
            try:
                with open(USERS_JSON, "r", encoding='utf-8') as f:
                    users_data = json.load(f)
                
                for username, data in users_data.items():
                    cursor.execute(
                        "INSERT INTO users (username, password, role, last_path, permissions, favorites, config) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (
                            username,
                            data.get("password", ""),
                            data.get("role", "user"),
                            data.get("last_path", ""),
                            json.dumps(data.get("permissions", {})),
                            json.dumps(data.get("favorites", [])),
                            json.dumps(data.get("config", {}))
                        )
                    )
                conn.commit()
                print("Migration successful.")
            except Exception as e:
                print(f"Error during migration: {e}")
        
        conn.close()

def get_user(username):
    """Retrieves a user by username."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()
    conn.close()
    
    if row:
        user_dict = dict(row)
        # Parse JSON fields
        try: user_dict["permissions"] = json.loads(user_dict["permissions"]) if user_dict["permissions"] else {}
        except: user_dict["permissions"] = {}
        
        try: user_dict["favorites"] = json.loads(user_dict["favorites"]) if user_dict["favorites"] else []
        except: user_dict["favorites"] = []
        
        try: user_dict["config"] = json.loads(user_dict["config"]) if user_dict["config"] else {}
        except: user_dict["config"] = {}
        
        return user_dict
    return None

def check_login(username, password):
    """Verifies username and password."""
    user = get_user(username)
    if user and user["password"] == password:
        return user
    return None

def update_user_last_path(username, path):
    """Updates the last_path for a user."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET last_path = ? WHERE username = ?", (path, username))
        conn.commit()
        conn.close()

def create_user(username, password, role="user"):
    """Creates a new user."""
    if get_user(username):
        return False, "El usuario ya existe"
    
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO users (username, password, role, last_path, permissions, favorites, config) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (username, password, role, "D:\\", "{}", "[]", "{}")
            )
            conn.commit()
            conn.close()
            return True, "Usuario creado exitosamente"
        except Exception as e:
            conn.close()
            return False, str(e)

def update_user_config(username, config_dict):
    """Updates the config JSON for a user."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        
        # We need to merge with existing config, so we read first
        cursor.execute("SELECT config FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        
        if row:
            current_config = json.loads(row["config"]) if row["config"] else {}
            current_config.update(config_dict)
            
            cursor.execute("UPDATE users SET config = ? WHERE username = ?", (json.dumps(current_config), username))
            conn.commit()
        conn.close()

def get_user_config(username):
    """Backwards compatibility for app_web.py usage (returns user dict, not just config)."""
    return get_user(username) or {}

def get_user_full_config(username):
    """Returns the full config dict for a user."""
    user = get_user(username)
    if user:
        return user.get("config", {})
    return {}

def add_user_favorite(username, path):
    """Adds a path to user favorites."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT favorites FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        
        if row:
            favs = json.loads(row["favorites"]) if row["favorites"] else []
            if path not in favs:
                favs.append(path)
                cursor.execute("UPDATE users SET favorites = ? WHERE username = ?", (json.dumps(favs), username))
                conn.commit()
                conn.close()
                return True
        conn.close()
    return False

def remove_user_favorite(username, path):
    """Removes a path from user favorites."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT favorites FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        
        if row:
            favs = json.loads(row["favorites"]) if row["favorites"] else []
            if path in favs:
                favs.remove(path)
                cursor.execute("UPDATE users SET favorites = ? WHERE username = ?", (json.dumps(favs), username))
                conn.commit()
                conn.close()
                return True
        conn.close()
    return False

def delete_user(username):
    """Deletes a user."""
    if username == "admin":
        return False, "No se puede eliminar al administrador principal"
        
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE username = ?", (username,))
        if cursor.rowcount > 0:
            conn.commit()
            conn.close()
            return True, "Usuario eliminado exitosamente"
        conn.close()
    return False, "El usuario no existe"

def get_all_users():
    """Returns a dictionary of all users (similar to load_users structure)."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users")
    rows = cursor.fetchall()
    conn.close()
    
    users = {}
    for row in rows:
        user_dict = dict(row)
        try: user_dict["permissions"] = json.loads(user_dict["permissions"]) if user_dict["permissions"] else {}
        except: user_dict["permissions"] = {}
        
        try: user_dict["favorites"] = json.loads(user_dict["favorites"]) if user_dict["favorites"] else []
        except: user_dict["favorites"] = []
        
        try: user_dict["config"] = json.loads(user_dict["config"]) if user_dict["config"] else {}
        except: user_dict["config"] = {}
        
        users[row["username"]] = user_dict
    return users

def update_user_role(username, role):
    """Updates the role of a user."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET role = ? WHERE username = ?", (role, username))
        conn.commit()
        conn.close()

def update_user_permissions(username, permissions_dict):
    """Updates the permissions JSON for a user."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET permissions = ? WHERE username = ?", (json.dumps(permissions_dict), username))
        conn.commit()
        conn.close()

# --- Task Queue Management ---

def create_task(username, command, params=None):
    """Creates a new task for the agent."""
    if params is None:
        params = {}
    
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO tasks (username, command, params, status) VALUES (?, ?, ?, ?)",
            (username, command, json.dumps(params), 'PENDING')
        )
        task_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return task_id

def get_pending_tasks(username, limit=1):
    """Retrieves pending tasks for a user and marks them as PROCESSING."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        
        # Select pending tasks
        cursor.execute(
            "SELECT * FROM tasks WHERE username = ? AND status = 'PENDING' ORDER BY created_at ASC LIMIT ?",
            (username, limit)
        )
        rows = cursor.fetchall()
        
        tasks = []
        for row in rows:
            task = dict(row)
            try: task["params"] = json.loads(task["params"]) if task["params"] else {}
            except: task["params"] = {}
            tasks.append(task)
            
            # Mark as processing
            cursor.execute("UPDATE tasks SET status = 'PROCESSING', updated_at = CURRENT_TIMESTAMP WHERE id = ?", (task["id"],))
        
        conn.commit()
        conn.close()
        return tasks

def update_task_result(task_id, status, result=None):
    """Updates the status and result of a task."""
    with _db_lock:
        conn = get_connection()
        cursor = conn.cursor()
        
        result_json = json.dumps(result) if result is not None else None
        
        cursor.execute(
            "UPDATE tasks SET status = ?, result = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (status, result_json, task_id)
        )
        conn.commit()
        conn.close()

def get_task_status(task_id):
    """Checks the status of a specific task."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tasks WHERE id = ?", (task_id,))
    row = cursor.fetchone()
    conn.close()
    
    if row:
        task = dict(row)
        try: task["result"] = json.loads(task["result"]) if task["result"] else None
        except: task["result"] = None
        return task
    return None

# Initialize DB on module load
init_db()
