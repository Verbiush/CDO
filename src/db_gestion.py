
import sqlite3
import json
from datetime import datetime
import os

# Import database module relative to this file location
try:
    import database as db
except ImportError:
    from src import database as db

def execute_sql_script(script_path):
    """Executes a SQL script from a file."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            with open(script_path, 'r', encoding='utf-8') as f:
                script = f.read()
            cursor.executescript(script)
            conn.commit()
            conn.close()
            return True, "Script ejecutado correctamente."
        except Exception as e:
            conn.close()
            return False, str(e)

def ensure_schema_updates():
    """Checks and applies necessary schema updates."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # Check if 'fecha_radicado' column exists in 'facturas' table
            cursor.execute("PRAGMA table_info(facturas)")
            columns = [info[1] for info in cursor.fetchall()]
            if "fecha_radicado" not in columns:
                cursor.execute("ALTER TABLE facturas ADD COLUMN fecha_radicado TEXT")
                conn.commit()
                print("Added 'fecha_radicado' column to 'facturas' table.")
        except Exception as e:
            print(f"Schema update error: {e}")
        finally:
            conn.close()

def migrate_schema_v2():
    """Migrates the database schema to support One-Invoice-Many-Attentions."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # Check if migration is needed (if facturas table has 'atencion_id')
            cursor.execute("PRAGMA table_info(facturas)")
            columns = [info[1] for info in cursor.fetchall()]
            if "atencion_id" not in columns:
                conn.close()
                return # Already migrated
            
            print("Migrating schema to V2 (Invoice -> Many Attentions)...")
            
            # 1. Rename old tables
            cursor.execute("ALTER TABLE facturas RENAME TO facturas_old")
            cursor.execute("ALTER TABLE atenciones RENAME TO atenciones_old")
            
            # 2. Create new tables
            cursor.execute('''
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
            
            cursor.execute('''
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
            
            # 3. Migrate Data
            # Facturas: Copy everything except atencion_id
            cursor.execute('''
                INSERT INTO facturas (id, no_factura, fecha_factura, tipo_pago, valor_servicio, copago, radicado, total, status, created_at, fecha_radicado, tipo_servicio)
                SELECT id, no_factura, fecha_factura, tipo_pago, valor_servicio, copago, radicado, total, status, created_at, fecha_radicado, tipo_servicio
                FROM facturas_old
            ''')
            
            # Atenciones: Copy old data AND link to factura
            cursor.execute('''
                INSERT INTO atenciones (id, paciente_id, factura_id, nro_estudio, descripcion_cups, fecha_ingreso, fecha_salida, autorizacion, created_at)
                SELECT a.id, a.paciente_id, f.id, a.nro_estudio, a.descripcion_cups, a.fecha_ingreso, a.fecha_salida, a.autorizacion, a.created_at
                FROM atenciones_old a
                LEFT JOIN facturas_old f ON f.atencion_id = a.id
            ''')
            
            # 4. Drop old tables
            cursor.execute("DROP TABLE facturas_old")
            cursor.execute("DROP TABLE atenciones_old")
            
            conn.commit()
            print("Migration V2 successful.")
            
        except Exception as e:
            conn.rollback()
            print(f"Migration V2 failed: {e}")
        finally:
            conn.close()

# Run migration logic
migrate_schema_v2()
ensure_schema_updates()

def reset_database():
    """Resets the database by dropping tables and recreating them from schema."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # Drop existing tables
            cursor.execute("DROP TABLE IF EXISTS facturas")
            cursor.execute("DROP TABLE IF EXISTS atenciones")
            cursor.execute("DROP TABLE IF EXISTS pacientes")
            conn.commit()
            conn.close()
            
            # Recreate from schema.sql
            import os
            schema_path = os.path.join(os.path.dirname(__file__), "schema.sql")
            if os.path.exists(schema_path):
                return execute_sql_script(schema_path)
            else:
                return False, "No se encontró el archivo schema.sql"
                
        except Exception as e:
            conn.close()
            return False, str(e)

def insert_document_record(data):
    """Inserts a new document record into the normalized database (Pacientes -> Facturas -> Atenciones)."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # 1. Handle Patient (Pacientes)
            no_doc = data.get('no_doc', '').strip()
            
            cursor.execute("SELECT id FROM pacientes WHERE no_doc = ?", (no_doc,))
            row = cursor.fetchone()
            
            if row and no_doc:
                patient_id = row['id']
                # Update existing patient if new data is provided
                update_p_fields = []
                update_p_values = []
                p_fields = {
                    'tipo_doc': data.get('tipo_doc'),
                    'nombre_completo': data.get('nombre_completo'),
                    'nombre_tercero': data.get('nombre_tercero'),
                    'eps': data.get('eps'),
                    'regimen': data.get('regimen'),
                    'categoria': data.get('categoria')
                }
                for f, v in p_fields.items():
                    if v is not None and str(v).strip():
                        update_p_fields.append(f"{f} = ?")
                        update_p_values.append(v)
                
                if update_p_fields:
                    update_p_values.append(patient_id)
                    cursor.execute(f"UPDATE pacientes SET {', '.join(update_p_fields)} WHERE id = ?", tuple(update_p_values))

            else:
                cursor.execute('''
                    INSERT INTO pacientes (tipo_doc, no_doc, nombre_completo, nombre_tercero, eps, regimen, categoria)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    data.get('tipo_doc', ''),
                    no_doc,
                    data.get('nombre_completo', ''),
                    data.get('nombre_tercero', ''),
                    data.get('eps', ''),
                    data.get('regimen', 'SUBSIDIADO'),
                    data.get('categoria', 'NIVEL 1')
                ))
                patient_id = cursor.lastrowid

            # 2. Handle Invoice (Facturas)
            no_factura = data.get('no_factura', '').strip()
            factura_id = None
            msg = ""

            if no_factura:
                cursor.execute("SELECT id, radicado, fecha_radicado FROM facturas WHERE no_factura = ?", (no_factura,))
                existing_factura = cursor.fetchone()
                
                if existing_factura:
                    factura_id = existing_factura['id']
                    # Update Factura
                    update_fields = []
                    update_values = []
                    
                    # Auto-Status Logic
                    current_rad = existing_factura['radicado']
                    current_fecha = existing_factura['fecha_radicado']
                    
                    # Helper to determine effective value (Input > Current)
                    # We only use input if it's a non-empty string, matching the update logic below
                    def get_effective_val(key, current):
                        val = data.get(key)
                        if val is not None and str(val).strip():
                            return val
                        return current

                    new_rad = get_effective_val('radicado', current_rad)
                    new_fecha = get_effective_val('fecha_radicado', current_fecha)
                    
                    auto_status = None
                    if new_rad and str(new_rad).strip() and new_fecha and str(new_fecha).strip():
                        auto_status = 'Resolved'

                    fields_to_check = {
                        'fecha_factura': data.get('fecha_factura'),
                        'tipo_pago': data.get('tipo_pago'),
                        'valor_servicio': data.get('valor_servicio'),
                        'copago': data.get('copago'),
                        'radicado': data.get('radicado'),
                        'total': data.get('total'),
                        'fecha_radicado': data.get('fecha_radicado'),
                        'tipo_servicio': data.get('tipo_servicio'),
                        'status': data.get('status') or auto_status
                    }
                    for field, value in fields_to_check.items():
                        if value is not None and str(value).strip():
                            update_fields.append(f"{field} = ?")
                            update_values.append(value)
                    
                    if update_fields:
                        update_values.append(factura_id)
                        cursor.execute(f"UPDATE facturas SET {', '.join(update_fields)} WHERE id = ?", tuple(update_values))
                        msg = f"Factura Actualizada: {', '.join([f.split(' =')[0] for f in update_fields])}"
                    else:
                        msg = "Factura existente (sin cambios)"
                else:
                    # Create New Invoice
                    rad_val = data.get('radicado', '')
                    fecha_rad_val = data.get('fecha_radicado', '')
                    
                    initial_status = 'PENDING'
                    if rad_val and str(rad_val).strip() and fecha_rad_val and str(fecha_rad_val).strip():
                        initial_status = 'Resolved'
                    
                    if data.get('status'):
                        initial_status = data.get('status')

                    cursor.execute('''
                        INSERT INTO facturas (no_factura, fecha_factura, tipo_pago, valor_servicio, copago, radicado, total, tipo_servicio, status, fecha_radicado)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        no_factura,
                        data.get('fecha_factura', ''),
                        data.get('tipo_pago', ''),
                        data.get('valor_servicio', ''),
                        data.get('copago', ''),
                        rad_val,
                        data.get('total', ''),
                        data.get('tipo_servicio', 'EVENTO'),
                        initial_status,
                        fecha_rad_val
                    ))
                    factura_id = cursor.lastrowid
                    msg = "Factura creada"

            # 3. Handle Attention (Atenciones)
            nro_estudio = data.get('nro_estudio', '').strip()
            atencion_id = None
            
            # Check if attention exists
            cursor.execute("SELECT id FROM atenciones WHERE nro_estudio = ?", (nro_estudio,))
            row = cursor.fetchone()
            
            if row and nro_estudio:
                atencion_id = row['id']
                # Update existing attention
                update_a_fields = []
                update_a_values = []
                a_fields = {
                    'descripcion_cups': data.get('descripcion'),
                    'fecha_ingreso': data.get('fecha_ingreso'),
                    'fecha_salida': data.get('fecha_salida'),
                    'autorizacion': data.get('autorizacion'),
                    'factura_id': factura_id # Link to invoice (updates if changed)
                }
                # Always update factura_id if we have one
                if factura_id:
                     update_a_fields.append("factura_id = ?")
                     update_a_values.append(factura_id)

                for f, v in a_fields.items():
                    if f == 'factura_id': continue # Handled above
                    if v is not None and str(v).strip():
                        update_a_fields.append(f"{f} = ?")
                        update_a_values.append(v)
                
                if update_a_fields:
                    update_a_values.append(atencion_id)
                    cursor.execute(f"UPDATE atenciones SET {', '.join(update_a_fields)} WHERE id = ?", tuple(update_a_values))

            else:
                cursor.execute('''
                    INSERT INTO atenciones (paciente_id, factura_id, nro_estudio, descripcion_cups, fecha_ingreso, fecha_salida, autorizacion)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    patient_id,
                    factura_id,
                    nro_estudio,
                    data.get('descripcion', ''),
                    data.get('fecha_ingreso', ''),
                    data.get('fecha_salida', ''),
                    data.get('autorizacion', '')
                ))
                atencion_id = cursor.lastrowid

            conn.commit()
            conn.close()
            return factura_id, msg
        except Exception as e:
            conn.close()
            return None, str(e)

def get_all_document_records():
    """Retrieves all document records by JOINing Facturas, Atenciones, and Pacientes."""
    conn = db.get_connection()
    cursor = conn.cursor()
    
    # We return one row per Attention (linked to Invoice)
    # ID returned is ATENCION ID to ensure uniqueness for UI operations
    query = '''
        SELECT 
            a.id as id, 
            f.id as factura_id,
            p.tipo_doc, p.no_doc, p.nombre_completo, p.nombre_tercero, p.eps, p.regimen, p.categoria,
            a.nro_estudio, a.descripcion_cups as descripcion, a.fecha_ingreso, a.fecha_salida, a.autorizacion,
            f.no_factura, f.fecha_factura, f.tipo_pago, f.valor_servicio, f.copago, f.radicado, f.total, f.tipo_servicio, f.status, f.fecha_radicado
        FROM atenciones a
        LEFT JOIN facturas f ON a.factura_id = f.id
        LEFT JOIN pacientes p ON a.paciente_id = p.id
        ORDER BY a.id DESC
    '''
    
    try:
        cursor.execute(query)
        columns = [column[0] for column in cursor.description]
        results = []
        for row in cursor.fetchall():
            results.append(dict(zip(columns, row)))
        conn.close()
        return results
    except Exception as e:
        conn.close()
        print(f"Error fetching records: {e}")
        return []

def update_document_field(record_id, field, value):
    """
    Updates a specific field of a document record.
    record_id is 'atenciones.id' (from get_all_document_records).
    """
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # 1. Identify which table the field belongs to
            table_map = {
                'tipo_doc': 'pacientes', 'no_doc': 'pacientes', 'nombre_completo': 'pacientes', 
                'nombre_tercero': 'pacientes', 'eps': 'pacientes', 'regimen': 'pacientes', 'categoria': 'pacientes',
                
                'nro_estudio': 'atenciones', 'descripcion': 'atenciones', 'descripcion_cups': 'atenciones',
                'fecha_ingreso': 'atenciones', 'fecha_salida': 'atenciones', 'autorizacion': 'atenciones',
                
                'no_factura': 'facturas', 'fecha_factura': 'facturas', 'tipo_pago': 'facturas', 
                'valor_servicio': 'facturas', 'copago': 'facturas', 'radicado': 'facturas', 
                'total': 'facturas', 'tipo_servicio': 'facturas', 'status': 'facturas', 'fecha_radicado': 'facturas'
            }
            
            target_table = table_map.get(field)
            if not target_table:
                return False, f"Campo desconocido: {field}"
                
            # 2. Get IDs from the record ID (which is atenciones.id)
            cursor.execute('''
                SELECT a.id as a_id, a.factura_id as f_id, a.paciente_id as p_id 
                FROM atenciones a
                WHERE a.id = ?
            ''', (record_id,))
            ids = cursor.fetchone()
            
            if not ids:
                return False, "Registro no encontrado"
                
            target_id = ids['a_id'] if target_table == 'atenciones' else (ids['f_id'] if target_table == 'facturas' else ids['p_id'])
            
            if target_table == 'facturas' and not target_id:
                 return False, "Esta atención no tiene factura asociada."

            # Remap field name if necessary
            db_field = field
            if field == 'descripcion': db_field = 'descripcion_cups'
            
            # 3. Execute Update
            sql = f"UPDATE {target_table} SET {db_field} = ? WHERE id = ?"
            cursor.execute(sql, (value, target_id))
            
            # Auto-Status Logic
            if target_table == 'facturas' and field in ['radicado', 'fecha_radicado']:
                cursor.execute("SELECT radicado, fecha_radicado FROM facturas WHERE id = ?", (target_id,))
                row = cursor.fetchone()
                if row:
                    r, f = row['radicado'], row['fecha_radicado']
                    if r and str(r).strip() and f and str(f).strip():
                         cursor.execute("UPDATE facturas SET status = 'Resolved' WHERE id = ?", (target_id,))

            conn.commit()
            conn.close()
            return True, "Campo actualizado exitosamente"
            
        except Exception as e:
            conn.close()
            return False, str(e)

def delete_document_record(record_id):
    """
    Deletes a document record.
    record_id is 'atenciones.id'.
    Logic: Delete the attention. If the associated factura has no other attentions, delete the factura too.
    """
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            # Get factura_id before deleting
            cursor.execute("SELECT factura_id FROM atenciones WHERE id = ?", (record_id,))
            row = cursor.fetchone()
            factura_id = row['factura_id'] if row else None
            
            # Delete Attention
            cursor.execute("DELETE FROM atenciones WHERE id = ?", (record_id,))
            
            # Check if Factura is empty
            if factura_id:
                cursor.execute("SELECT count(*) as count FROM atenciones WHERE factura_id = ?", (factura_id,))
                count = cursor.fetchone()['count']
                if count == 0:
                    cursor.execute("DELETE FROM facturas WHERE id = ?", (factura_id,))
            
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False

def update_record_status(record_id, status):
    """
    Updates the status of a record.
    record_id is 'atenciones.id'. We need to update the associated factura.
    """
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        
        # Get factura_id
        cursor.execute("SELECT factura_id FROM atenciones WHERE id = ?", (record_id,))
        row = cursor.fetchone()
        factura_id = row['factura_id'] if row else None
        
        if factura_id:
            cursor.execute("UPDATE facturas SET status = ? WHERE id = ?", (status, factura_id))
            conn.commit()
        
        conn.close()
