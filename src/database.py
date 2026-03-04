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

# Try to import psycopg2 for PostgreSQL (Supabase)
try:
    import psycopg2
    from psycopg2 import extras as psycopg2_extras
    HAS_POSTGRES_LIB = True
except ImportError:
    HAS_POSTGRES_LIB = False

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

# PostgreSQL / Supabase Configuration
POSTGRES_HOST = os.getenv("POSTGRES_HOST") or os.getenv("SUPABASE_HOST")
POSTGRES_USER = os.getenv("POSTGRES_USER") or os.getenv("SUPABASE_USER")
POSTGRES_PASSWORD = os.getenv("POSTGRES_PASSWORD") or os.getenv("SUPABASE_PASSWORD")
POSTGRES_DB = os.getenv("POSTGRES_DB") or os.getenv("SUPABASE_DB")
POSTGRES_PORT = os.getenv("POSTGRES_PORT", "5432")

USE_POSTGRES = all([POSTGRES_HOST, POSTGRES_USER, POSTGRES_PASSWORD, POSTGRES_DB])

# Lock for thread safety within the same process
_db_lock = threading.Lock()

def get_connection():
    """Returns a connection to the database (SQLite, MySQL, or PostgreSQL)."""
    if USE_POSTGRES:
        if not HAS_POSTGRES_LIB:
            raise ImportError("PostgreSQL/Supabase configuration found but 'psycopg2' or 'psycopg2-binary' library is not installed. Please run 'pip install psycopg2-binary'")
        try:
            conn = psycopg2.connect(
                host=POSTGRES_HOST,
                user=POSTGRES_USER,
                password=POSTGRES_PASSWORD,
                dbname=POSTGRES_DB,
                port=POSTGRES_PORT,
                cursor_factory=psycopg2_extras.RealDictCursor
            )
            # Enable autocommit for consistency with other adapters if needed, 
            # or handle commits manually. For now, we return the raw connection.
            return conn
        except Exception as e:
            print(f"Error connecting to PostgreSQL: {e}")
            raise e
    elif USE_MYSQL:
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
    """Executes a query handling differences between SQLite, MySQL and PostgreSQL placeholders."""
    if params is None:
        params = ()
    
    # SQLite uses ?, MySQL/Postgres use %s
    if USE_MYSQL or USE_POSTGRES:
        query = query.replace('?', '%s')
        
    if USE_MYSQL:
        cursor = conn.cursor(dictionary=True)
    else:
        cursor = conn.cursor()
    try:
        cursor.execute(query, params)
        return cursor
    except Exception as e:
        print(f"Query Error: {e}")
        print(f"Query: {query}")
        print(f"Params: {params}")
        raise e

def init_db():
    """Initializes the database tables."""
    with _db_lock:
        conn = get_connection()
        try:
            # Create Users table
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS users (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(255) UNIQUE NOT NULL,
                    password VARCHAR(255) NOT NULL,
                    role VARCHAR(50) DEFAULT 'user',
                    last_path VARCHAR(255),
                    permissions JSON,
                    favorites JSON,
                    config JSON,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')
            elif USE_POSTGRES:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username VARCHAR(255) UNIQUE NOT NULL,
                    password VARCHAR(255) NOT NULL,
                    role VARCHAR(50) DEFAULT 'user',
                    last_path VARCHAR(255),
                    permissions JSON,
                    favorites JSON,
                    config JSON,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')
            else:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    role TEXT DEFAULT 'user',
                    last_path TEXT,
                    permissions TEXT,
                    favorites TEXT,
                    config TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')

            # Create Document Records table
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
            elif USE_POSTGRES:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS document_records (
                    id SERIAL PRIMARY KEY,
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
            elif USE_POSTGRES:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS pacientes (
                    id SERIAL PRIMARY KEY,
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
            elif USE_POSTGRES:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS facturas (
                    id SERIAL PRIMARY KEY,
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
            elif USE_POSTGRES:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS atenciones (
                    id SERIAL PRIMARY KEY,
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
            
            # Create Tasks table (for Agent)
            if USE_MYSQL:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS tasks (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(255),
                    command VARCHAR(255),
                    params TEXT,
                    status VARCHAR(50) DEFAULT 'PENDING',
                    result TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
                ''')
            elif USE_POSTGRES:
                 execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS tasks (
                    id SERIAL PRIMARY KEY,
                    username VARCHAR(255),
                    command VARCHAR(255),
                    params TEXT,
                    status VARCHAR(50) DEFAULT 'PENDING',
                    result TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')
            else:
                execute_query(conn, '''
                CREATE TABLE IF NOT EXISTS tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT,
                    command TEXT,
                    params TEXT,
                    status TEXT DEFAULT 'PENDING',
                    result TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')

            conn.commit()

            # Check for missing columns in document_records (Migration)
            # PRAGMA table_info is SQLite specific. MySQL uses SHOW COLUMNS or DESCRIBE
            if USE_MYSQL or USE_POSTGRES:
                # Check document_records for regimen
                if USE_POSTGRES:
                    cursor = execute_query(conn, "SELECT column_name FROM information_schema.columns WHERE table_name='document_records' AND column_name='regimen'")
                else:
                    cursor = execute_query(conn, "SHOW COLUMNS FROM document_records LIKE 'regimen'")

                if not cursor.fetchone():
                    print("Migrating document_records: Adding 'regimen' column...")
                    execute_query(conn, "ALTER TABLE document_records ADD COLUMN regimen VARCHAR(50) DEFAULT 'SUBSIDIADO'")
                    conn.commit()
                
                # Check pacientes for categoria
                if USE_POSTGRES:
                    cursor = execute_query(conn, "SELECT column_name FROM information_schema.columns WHERE table_name='pacientes' AND column_name='categoria'")
                else:
                    cursor = execute_query(conn, "SHOW COLUMNS FROM pacientes LIKE 'categoria'")

                if not cursor.fetchone():
                    print("Migrating pacientes: Adding 'categoria' column...")
                    execute_query(conn, "ALTER TABLE pacientes ADD COLUMN categoria VARCHAR(50) DEFAULT 'NIVEL 1'")
                    conn.commit()
                
                # Check facturas for tipo_servicio
                if USE_POSTGRES:
                    cursor = execute_query(conn, "SELECT column_name FROM information_schema.columns WHERE table_name='facturas' AND column_name='tipo_servicio'")
                else:
                    cursor = execute_query(conn, "SHOW COLUMNS FROM facturas LIKE 'tipo_servicio'")

                if not cursor.fetchone():
                    print("Migrating facturas: Adding 'tipo_servicio' column...")
                    execute_query(conn, "ALTER TABLE facturas ADD COLUMN tipo_servicio VARCHAR(255) DEFAULT 'EVENTO'")
                    conn.commit()

                # Check facturas for fecha_radicado
                if USE_POSTGRES:
                    cursor = execute_query(conn, "SELECT column_name FROM information_schema.columns WHERE table_name='facturas' AND column_name='fecha_radicado'")
                else:
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

def get_all_users():
    """Returns all users from the database as a dictionary {username: user_data}."""
    conn = get_connection()
    try:
        if USE_MYSQL:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT * FROM users")
            rows = cursor.fetchall()
        elif USE_POSTGRES:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users")
            rows = cursor.fetchall()
        else:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users")
            rows = [dict(row) for row in cursor.fetchall()]
            
        # Convert to dictionary keyed by username
        users_dict = {}
        for row in rows:
            # Parse JSON fields if they exist and are strings
            if isinstance(row.get("permissions"), str):
                try: row["permissions"] = json.loads(row["permissions"])
                except: row["permissions"] = {}
            
            if isinstance(row.get("favorites"), str):
                try: row["favorites"] = json.loads(row["favorites"])
                except: row["favorites"] = []
                
            if isinstance(row.get("config"), str):
                try: row["config"] = json.loads(row["config"])
                except: row["config"] = {}
                
            users_dict[row["username"]] = row
            
        return users_dict
    except Exception as e:
        print(f"Error getting users: {e}")
        return {}
    finally:
        conn.close()

def get_all_invoices():
    """Returns all document records (invoices) from the database."""
    conn = get_connection()
    try:
        # We use document_records as it contains EPS and Regimen info required by tab_admin.py
        # Ordered by created_at DESC to show newest first
        query = "SELECT * FROM document_records ORDER BY created_at DESC"
        
        if USE_MYSQL:
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query)
            return cursor.fetchall()
        elif USE_POSTGRES:
            cursor = conn.cursor()
            cursor.execute(query)
            return cursor.fetchall()
        else:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query)
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error getting invoices (records): {e}")
        return []
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

def create_task(username, command, params=None):
    """Creates a new task for the agent. Returns (success, task_id)."""
    if params is None: params = {}
    with _db_lock:
        conn = get_connection()
        try:
            if USE_POSTGRES:
                cursor = execute_query(conn, 
                    "INSERT INTO tasks (username, command, params, status) VALUES (?, ?, ?, 'PENDING') RETURNING id",
                    (username, command, json.dumps(params))
                )
                task_id = cursor.fetchone()['id']
            else:
                cursor = execute_query(conn, 
                    "INSERT INTO tasks (username, command, params, status) VALUES (?, ?, ?, 'PENDING')",
                    (username, command, json.dumps(params))
                )
                task_id = cursor.lastrowid
                
            conn.commit()
            return True, task_id
        except Exception as e:
            print(f"Error creating task: {e}")
            return False, None
        finally:
            conn.close()

def get_task_result(task_id):
    """Retrieves the status and result of a specific task."""
    with _db_lock:
        conn = get_connection()
        try:
            cursor = execute_query(conn, "SELECT status, result FROM tasks WHERE id = ?", (task_id,))
            row = cursor.fetchone()
            if row:
                result_data = dict(row)
                try:
                    result_data["result"] = json.loads(result_data["result"]) if result_data["result"] else None
                except:
                    result_data["result"] = None
                return result_data
            return None
        finally:
            conn.close()

def get_pending_tasks(username):
    """Retrieves pending tasks for a user and marks them as DISPATCHED."""
    with _db_lock:
        conn = get_connection()
        try:
            cursor = execute_query(conn, "SELECT * FROM tasks WHERE username = ? AND status = 'PENDING'", (username,))
            rows = cursor.fetchall()
            tasks = []
            if rows:
                ids = []
                for row in rows:
                    task = dict(row)
                    # Handle Row object (sqlite3) vs dictionary (mysql/postgres)
                    if hasattr(task, 'keys'): # It's likely a Row or dict-like
                        task = dict(task)
                    
                    try: task["params"] = json.loads(task["params"]) if task["params"] else {}
                    except: task["params"] = {}
                    tasks.append(task)
                    ids.append(task['id'])
                
                # Mark as DISPATCHED
                if ids:
                    for tid in ids:
                        execute_query(conn, "UPDATE tasks SET status = 'DISPATCHED' WHERE id = ?", (tid,))
                    conn.commit()
            return tasks
        finally:
            conn.close()

def update_task_result(task_id, status, result=None):
    """Updates the status and result of a task."""
    with _db_lock:
        conn = get_connection()
        try:
            result_json = json.dumps(result) if result else None
            execute_query(conn, 
                "UPDATE tasks SET status = ?, result = ? WHERE id = ?",
                (status, result_json, task_id)
            )
            conn.commit()
        finally:
            conn.close()
