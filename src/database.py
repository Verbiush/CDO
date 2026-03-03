import sqlite3
import json
import os
import threading
import sys

# Try to import mysql.connector, but don't fail if not present (unless configured to use MySQL)
try:
    import mysql.connector
    from mysql.connector import Error
    HAS_MYSQL_LIB = True
except ImportError:
    HAS_MYSQL_LIB = False
    # Define dummy Error class to avoid NameError if used in except blocks
    class Error(Exception): pass

# Detect if we are running as a frozen executable (PyInstaller)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Allow overriding DB path via environment variable (useful for Docker/Cloud)
DB_FILE = os.getenv("DB_PATH", os.path.join(BASE_DIR, "users.db"))
USERS_JSON = os.path.join(BASE_DIR, "users.json")

# MySQL Configuration
MYSQL_HOST = os.getenv("MYSQL_HOST")
MYSQL_USER = os.getenv("MYSQL_USER")
MYSQL_PASSWORD = os.getenv("MYSQL_PASSWORD")
MYSQL_DATABASE = os.getenv("MYSQL_DATABASE")
MYSQL_PORT = os.getenv("MYSQL_PORT", "3306")

USE_MYSQL = all([MYSQL_HOST, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DATABASE])

# Lock for thread safety within the same process
_db_lock = threading.Lock()

def get_connection():
    """Returns a connection to the database (SQLite or MySQL)."""
    if USE_MYSQL:
        if not HAS_MYSQL_LIB:
            raise ImportError("MySQL configuration found (USE_MYSQL=True) but 'mysql-connector-python' library is not installed. Please run 'pip install mysql-connector-python'")
            
        try:
            conn = mysql.connector.connect(
                host=MYSQL_HOST,
                user=MYSQL_USER,
                password=MYSQL_PASSWORD,
                database=MYSQL_DATABASE,
                port=int(MYSQL_PORT)
            )
            return conn
        except Error as e:
            print(f"Error connecting to MySQL: {e}")
            # Fallback or re-raise? For now, let's assume if env vars are set, we want MySQL or fail.
            raise e
    else:
        conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

def execute_query(conn, query, params=None):
    """Executes a query handling differences between SQLite and MySQL placeholders."""
    if params is None:
        params = ()
    
    if USE_MYSQL:
        # Convert ? to %s for MySQL
        query = query.replace('?', '%s')
        cursor = conn.cursor(dictionary=True)
        cursor.execute(query, params)
        return cursor
    else:
        cursor = conn.cursor()
        cursor.execute(query, params)
        return cursor

def init_db():
    """Initializes the database table and migrates from JSON if needed."""
    with _db_lock:
        conn = get_connection()
        try:
            # Create users table
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS users (
                    username VARCHAR(255) PRIMARY KEY,
                    password TEXT NOT NULL,
                    role VARCHAR(50) DEFAULT 'user',
                    last_path TEXT,
                    permissions TEXT,
                    favorites TEXT,
                    config TEXT
                )
                ''')
            else:
                execute_query(conn, '''
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
            if USE_MYSQL:
                 execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS tasks (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(255) NOT NULL,
                    command TEXT NOT NULL,
                    params TEXT,
                    status VARCHAR(50) DEFAULT 'PENDING',
                    result TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
                ''')
            else:
                execute_query(conn, '''
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
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS document_records (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    nro_estudio VARCHAR(255),
                    descripcion TEXT,
                    eps VARCHAR(255),
                    tipo_doc VARCHAR(50),
                    no_doc VARCHAR(255),
                    nombre_completo VARCHAR(255),
                    nombre_tercero VARCHAR(255),
                    fecha_ingreso VARCHAR(50),
                    fecha_salida VARCHAR(50),
                    autorizacion VARCHAR(255),
                    no_factura VARCHAR(255),
                    fecha_factura VARCHAR(50),
                    tipo_pago VARCHAR(50),
                    valor_servicio VARCHAR(255),
                    copago VARCHAR(255),
                    total VARCHAR(255),
                    regimen VARCHAR(50) DEFAULT 'SUBSIDIADO',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    status VARCHAR(50) DEFAULT 'PENDING'
                )
                ''')
            else:
                execute_query(conn, '''
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
            
            # --- NEW TABLES FOR NORMALIZED SCHEMA ---
            
            # Create Pacientes table
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS pacientes (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    tipo_doc VARCHAR(50),
                    no_doc VARCHAR(255) UNIQUE,
                    nombre_completo VARCHAR(255),
                    nombre_tercero VARCHAR(255),
                    eps VARCHAR(255),
                    regimen VARCHAR(50) DEFAULT 'SUBSIDIADO',
                    categoria VARCHAR(50) DEFAULT 'NIVEL 1',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')
            else:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS pacientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo_doc TEXT,
                    no_doc TEXT UNIQUE,
                    nombre_completo TEXT,
                    nombre_tercero TEXT,
                    eps TEXT,
                    regimen TEXT DEFAULT 'SUBSIDIADO',
                    categoria TEXT DEFAULT 'NIVEL 1',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')

            # Create Facturas table
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS facturas (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    no_factura VARCHAR(255) UNIQUE,
                    fecha_factura VARCHAR(50),
                    tipo_pago VARCHAR(50),
                    valor_servicio VARCHAR(255),
                    copago VARCHAR(255),
                    radicado VARCHAR(255),
                    total VARCHAR(255),
                    status VARCHAR(50) DEFAULT 'PENDING',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    fecha_radicado VARCHAR(50),
                    tipo_servicio VARCHAR(255)
                )
                ''')
            else:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS facturas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    no_factura TEXT UNIQUE,
                    fecha_factura TEXT,
                    tipo_pago TEXT,
                    valor_servicio TEXT,
                    copago TEXT,
                    radicado TEXT,
                    total TEXT,
                    status TEXT DEFAULT 'PENDING',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    fecha_radicado TEXT,
                    tipo_servicio TEXT
                )
                ''')

            # Create Atenciones table
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS atenciones (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    paciente_id INT NOT NULL,
                    factura_id INT,
                    nro_estudio VARCHAR(255),
                    descripcion_cups TEXT,
                    fecha_ingreso VARCHAR(50),
                    fecha_salida VARCHAR(50),
                    autorizacion VARCHAR(255),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(paciente_id) REFERENCES pacientes(id),
                    FOREIGN KEY(factura_id) REFERENCES facturas(id)
                )
                ''')
            else:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS atenciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    paciente_id INTEGER NOT NULL,
                    factura_id INTEGER,
                    nro_estudio TEXT,
                    descripcion_cups TEXT,
                    fecha_ingreso TEXT,
                    fecha_salida TEXT,
                    autorizacion TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(paciente_id) REFERENCES pacientes(id),
                    FOREIGN KEY(factura_id) REFERENCES facturas(id)
                )
                ''')
            
            conn.commit()

            # Check for missing columns in document_records (Migration)
            # PRAGMA table_info is SQLite specific. MySQL uses SHOW COLUMNS or DESCRIBE
            if USE_MYSQL:
                # Check document_records for regimen
                cursor = execute_query(conn, "SHOW COLUMNS FROM document_records LIKE 'regimen'")
                if not cursor.fetchone():
                    print("Migrating document_records: Adding 'regimen' column...")
                    execute_query(conn, "ALTER TABLE document_records ADD COLUMN regimen VARCHAR(50) DEFAULT 'SUBSIDIADO'")
                    conn.commit()
                
                # Check pacientes for categoria
                cursor = execute_query(conn, "SHOW COLUMNS FROM pacientes LIKE 'categoria'")
                if not cursor.fetchone():
                    print("Migrating pacientes: Adding 'categoria' column...")
                    execute_query(conn, "ALTER TABLE pacientes ADD COLUMN categoria VARCHAR(50) DEFAULT 'NIVEL 1'")
                    conn.commit()
                
                # Check facturas for tipo_servicio
                cursor = execute_query(conn, "SHOW COLUMNS FROM facturas LIKE 'tipo_servicio'")
                if not cursor.fetchone():
                    print("Migrating facturas: Adding 'tipo_servicio' column...")
                    execute_query(conn, "ALTER TABLE facturas ADD COLUMN tipo_servicio VARCHAR(255) DEFAULT 'EVENTO'")
                    conn.commit()

                # Check facturas for fecha_radicado
                cursor = execute_query(conn, "SHOW COLUMNS FROM facturas LIKE 'fecha_radicado'")
                if not cursor.fetchone():
                    print("Migrating facturas: Adding 'fecha_radicado' column...")
                    execute_query(conn, "ALTER TABLE facturas ADD COLUMN fecha_radicado VARCHAR(50)")
                    conn.commit()

            else:
                # Check document_records for regimen
                cursor = execute_query(conn, "PRAGMA table_info(document_records)")
                columns = [info[1] for info in cursor.fetchall()]
                if "regimen" not in columns:
                    print("Migrating document_records: Adding 'regimen' column...")
                    execute_query(conn, "ALTER TABLE document_records ADD COLUMN regimen TEXT DEFAULT 'SUBSIDIADO'")
                    conn.commit()
                
                # Check pacientes for categoria
                cursor = execute_query(conn, "PRAGMA table_info(pacientes)")
                columns = [info[1] for info in cursor.fetchall()]
                if "categoria" not in columns:
                    print("Migrating pacientes: Adding 'categoria' column...")
                    execute_query(conn, "ALTER TABLE pacientes ADD COLUMN categoria TEXT DEFAULT 'NIVEL 1'")
                    conn.commit()
                
                # Check facturas for tipo_servicio
                cursor = execute_query(conn, "PRAGMA table_info(facturas)")
                columns = [info[1] for info in cursor.fetchall()]
                if "tipo_servicio" not in columns:
                    print("Migrating facturas: Adding 'tipo_servicio' column...")
                    execute_query(conn, "ALTER TABLE facturas ADD COLUMN tipo_servicio TEXT DEFAULT 'EVENTO'")
                    conn.commit()

                # Check facturas for fecha_radicado
                if "fecha_radicado" not in columns:
                    print("Migrating facturas: Adding 'fecha_radicado' column...")
                    execute_query(conn, "ALTER TABLE facturas ADD COLUMN fecha_radicado TEXT")
                    conn.commit()

            
            # Check if we need to migrate from JSON
            cursor = execute_query(conn, "SELECT count(*) as count FROM users")
            res = cursor.fetchone()
            count = res['count'] if USE_MYSQL else res[0]
            
            if count == 0:
                # Try to migrate from users.json first
                if os.path.exists(USERS_JSON):
                    print("Migrating users from JSON to DB...")
                    try:
                        with open(USERS_JSON, "r", encoding='utf-8') as f:
                            users_data = json.load(f)
                        
                        for username, data in users_data.items():
                            execute_query(conn,
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
                
                # Check again if still empty
                cursor = execute_query(conn, "SELECT count(*) as count FROM users")
                res = cursor.fetchone()
                count = res['count'] if USE_MYSQL else res[0]
                
                if count == 0:
                    # Try backup file
                    json_bak = os.path.join(BASE_DIR, "users.json.bak")
                    if os.path.exists(json_bak):
                        print("Migrating users from JSON backup...")
                        try:
                            with open(json_bak, "r", encoding='utf-8') as f:
                                users_data = json.load(f)
                            
                            for username, data in users_data.items():
                                execute_query(conn,
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
                            print("Backup migration successful.")
                        except Exception as e:
                            print(f"Error during backup migration: {e}")

                # Check again, if still empty, create default admin
                cursor = execute_query(conn, "SELECT count(*) as count FROM users")
                res = cursor.fetchone()
                count = res['count'] if USE_MYSQL else res[0]
                
                if count == 0:
                    print("Creating default admin user...")
                    execute_query(conn,
                        "INSERT INTO users (username, password, role, last_path, permissions, favorites, config) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        ("admin", "admin", "admin", "", "{}", "[]", "{}")
                    )
                    conn.commit()
            
        except Exception as e:
            print(f"Error initializing database: {e}")
            raise e
        finally:
            conn.close()

def get_user(username):
    """Retrieves a user by username."""
    conn = get_connection()
    try:
        cursor = execute_query(conn, "SELECT * FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        
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
    finally:
        conn.close()

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
        try:
            execute_query(conn, "UPDATE users SET last_path = ? WHERE username = ?", (path, username))
            conn.commit()
        finally:
            conn.close()

def create_user(username, password, role="user"):
    """Creates a new user."""
    if get_user(username):
        return False, "El usuario ya existe"
    
    with _db_lock:
        conn = get_connection()
        try:
            execute_query(conn,
                "INSERT INTO users (username, password, role, last_path, permissions, favorites, config) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (username, password, role, "D:\\", "{}", "[]", "{}")
            )
            conn.commit()
            return True, "Usuario creado exitosamente"
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

def update_user_config(username, config_dict):
    """Updates the config JSON for a user."""
    with _db_lock:
        conn = get_connection()
        try:
            # We need to merge with existing config, so we read first
            cursor = execute_query(conn, "SELECT config FROM users WHERE username = ?", (username,))
            row = cursor.fetchone()
            
            if row:
                config_str = row["config"] if USE_MYSQL else row["config"]
                current_config = json.loads(config_str) if config_str else {}
                current_config.update(config_dict)
                
                execute_query(conn, "UPDATE users SET config = ? WHERE username = ?", (json.dumps(current_config), username))
                conn.commit()
        finally:
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
        try:
            cursor = execute_query(conn, "SELECT favorites FROM users WHERE username = ?", (username,))
            row = cursor.fetchone()
            
            if row:
                favs_str = row["favorites"] if USE_MYSQL else row["favorites"]
                favs = json.loads(favs_str) if favs_str else []
                if path not in favs:
                    favs.append(path)
                    execute_query(conn, "UPDATE users SET favorites = ? WHERE username = ?", (json.dumps(favs), username))
                    conn.commit()
                    return True
            return False
        finally:
            conn.close()

def remove_user_favorite(username, path):
    """Removes a path from user favorites."""
    with _db_lock:
        conn = get_connection()
        try:
            cursor = execute_query(conn, "SELECT favorites FROM users WHERE username = ?", (username,))
            row = cursor.fetchone()
            
            if row:
                favs_str = row["favorites"] if USE_MYSQL else row["favorites"]
                favs = json.loads(favs_str) if favs_str else []
                if path in favs:
                    favs.remove(path)
                    execute_query(conn, "UPDATE users SET favorites = ? WHERE username = ?", (json.dumps(favs), username))
                    conn.commit()
                    return True
            return False
        finally:
            conn.close()

def delete_user(username):
    """Deletes a user."""
    if username == "admin":
        return False, "No se puede eliminar al administrador principal"
        
    with _db_lock:
        conn = get_connection()
        try:
            cursor = execute_query(conn, "DELETE FROM users WHERE username = ?", (username,))
            if cursor.rowcount > 0:
                conn.commit()
                return True, "Usuario eliminado exitosamente"
            return False, "El usuario no existe"
        finally:
            conn.close()

def get_all_users():
    """Returns a dictionary of all users (similar to load_users structure)."""
    conn = get_connection()
    try:
        cursor = execute_query(conn, "SELECT * FROM users")
        rows = cursor.fetchall()
        
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
    finally:
        conn.close()

def update_user_role(username, role):
    """Updates the role of a user."""
    with _db_lock:
        conn = get_connection()
        try:
            execute_query(conn, "UPDATE users SET role = ? WHERE username = ?", (role, username))
            conn.commit()
        finally:
            conn.close()

def update_user_permissions(username, permissions_dict):
    """Updates the permissions JSON for a user."""
    with _db_lock:
        conn = get_connection()
        try:
            execute_query(conn, "UPDATE users SET permissions = ? WHERE username = ?", (json.dumps(permissions_dict), username))
            conn.commit()
        finally:
            conn.close()

# --- Task Queue Management ---

def create_task(username, command, params=None):
    """Creates a new task for the agent."""
    if params is None:
        params = {}
    
    with _db_lock:
        conn = get_connection()
        try:
            cursor = execute_query(conn,
                "INSERT INTO tasks (username, command, params, status) VALUES (?, ?, ?, ?)",
                (username, command, json.dumps(params), 'PENDING')
            )
            task_id = cursor.lastrowid
            conn.commit()
            return task_id
        finally:
            conn.close()

def get_pending_tasks(username, limit=1):
    """Retrieves pending tasks for a user and marks them as PROCESSING."""
    with _db_lock:
        conn = get_connection()
        try:
            # Select pending tasks
            if USE_MYSQL:
                cursor = execute_query(conn,
                    "SELECT * FROM tasks WHERE username = ? AND status = 'PENDING' ORDER BY created_at ASC LIMIT %s",
                    (username, limit)
                )
            else:
                 cursor = execute_query(conn,
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
                execute_query(conn, "UPDATE tasks SET status = 'PROCESSING', updated_at = CURRENT_TIMESTAMP WHERE id = ?", (task["id"],))
            
            conn.commit()
            return tasks
        finally:
            conn.close()

def update_task_result(task_id, status, result=None):
    """Updates the status and result of a task."""
    with _db_lock:
        conn = get_connection()
        try:
            result_json = json.dumps(result) if result is not None else None
            
            execute_query(conn,
                "UPDATE tasks SET status = ?, result = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (status, result_json, task_id)
            )
            conn.commit()
        finally:
            conn.close()

def get_task_status(task_id):
    """Checks the status of a specific task."""
    conn = get_connection()
    try:
        cursor = execute_query(conn, "SELECT * FROM tasks WHERE id = ?", (task_id,))
        row = cursor.fetchone()
        
        if row:
            task = dict(row)
            try: task["result"] = json.loads(task["result"]) if task["result"] else None
            except: task["result"] = None
            return task
        return None
    finally:
        conn.close()

# --- ADMIN REPORTS & BACKUP ---

def get_db_path():
    """Returns the absolute path to the database file."""
    return DB_FILE

def get_all_invoices():
    """Returns all invoices with patient details."""
    conn = get_connection()
    try:
        query = """
            SELECT f.*, p.eps, p.regimen, p.nombre_completo, p.no_doc
            FROM facturas f
            LEFT JOIN atenciones a ON a.factura_id = f.id
            LEFT JOIN pacientes p ON a.paciente_id = p.id
            GROUP BY f.id
        """
        cursor = execute_query(conn, query)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()

def get_pending_invoices():
    """Returns pending invoices with patient details."""
    conn = get_connection()
    try:
        query = """
            SELECT f.*, p.eps, p.regimen, p.nombre_completo, p.no_doc
            FROM facturas f
            LEFT JOIN atenciones a ON a.factura_id = f.id
            LEFT JOIN pacientes p ON a.paciente_id = p.id
            WHERE f.status = 'PENDING'
            GROUP BY f.id
        """
        cursor = execute_query(conn, query)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()

def get_radicado_invoices():
    """Returns invoices with radicado with patient details."""
    conn = get_connection()
    try:
        query = """
            SELECT f.*, p.eps, p.regimen, p.nombre_completo, p.no_doc
            FROM facturas f
            LEFT JOIN atenciones a ON a.factura_id = f.id
            LEFT JOIN pacientes p ON a.paciente_id = p.id
            WHERE f.radicado IS NOT NULL AND f.radicado != ''
            GROUP BY f.id
        """
        cursor = execute_query(conn, query)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()

# Initialize DB on module load
init_db()
