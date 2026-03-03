
import sys
import os

# Add the parent directory to sys.path to allow imports if run directly
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from src import db_gestion
except ImportError:
    import db_gestion

def migrate():
    print("Starting migration (using db_gestion.ensure_schema_updates)...")
    try:
        db_gestion.ensure_schema_updates()
        print("Migration completed successfully.")
    except Exception as e:
        print(f"Migration failed: {e}")

if __name__ == "__main__":
    migrate()
