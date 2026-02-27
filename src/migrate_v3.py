import sqlite3
import os

# Path to the database
DB_FILE = os.path.join(os.path.dirname(__file__), "users.db")

def migrate():
    print(f"Migrating database: {DB_FILE}")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # 1. Add 'categoria' to 'pacientes'
    try:
        cursor.execute("ALTER TABLE pacientes ADD COLUMN categoria TEXT DEFAULT 'NIVEL 1'")
        print("Added column 'categoria' to 'pacientes'")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e):
            print("Column 'categoria' already exists in 'pacientes'")
        else:
            print(f"Error adding 'categoria' column: {e}")

    # 2. Add 'tipo_servicio' to 'facturas'
    try:
        cursor.execute("ALTER TABLE facturas ADD COLUMN tipo_servicio TEXT DEFAULT 'EVENTO'")
        print("Added column 'tipo_servicio' to 'facturas'")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e):
            print("Column 'tipo_servicio' already exists in 'facturas'")
        else:
            print(f"Error adding 'tipo_servicio' column: {e}")

    # 3. Update existing records with default values
    print("Updating existing records with default values...")
    
    # Update pacientes where categoria is NULL
    cursor.execute("UPDATE pacientes SET categoria = 'NIVEL 1' WHERE categoria IS NULL")
    print(f"Updated {cursor.rowcount} rows in 'pacientes'")

    # Update facturas where tipo_servicio is NULL
    cursor.execute("UPDATE facturas SET tipo_servicio = 'EVENTO' WHERE tipo_servicio IS NULL")
    print(f"Updated {cursor.rowcount} rows in 'facturas'")

    conn.commit()
    conn.close()
    print("Migration completed successfully.")

if __name__ == "__main__":
    migrate()
