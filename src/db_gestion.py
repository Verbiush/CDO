
import sqlite3
import json
from datetime import datetime

# Import database module relative to this file location
try:
    import database as db
except ImportError:
    from src import database as db

def insert_document_record(data):
    """Inserts a new document record into the database."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO document_records (
                    nro_estudio, descripcion, eps, tipo_doc, no_doc, 
                    nombre_completo, nombre_tercero, fecha_ingreso, fecha_salida, 
                    autorizacion, no_factura, fecha_factura, tipo_pago, 
                    valor_servicio, copago, total, regimen, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('nro_estudio', ''),
                data.get('descripcion', ''),
                data.get('eps', ''),
                data.get('tipo_doc', ''),
                data.get('no_doc', ''),
                data.get('nombre_completo', ''),
                data.get('nombre_tercero', ''),
                data.get('fecha_ingreso', ''),
                data.get('fecha_salida', ''),
                data.get('autorizacion', ''),
                data.get('no_factura', ''),
                data.get('fecha_factura', ''),
                data.get('tipo_pago', ''),
                data.get('valor_servicio', ''),
                data.get('copago', ''),
                data.get('total', ''),
                data.get('regimen', 'SUBSIDIADO'),
                'PENDING'
            ))
            record_id = cursor.lastrowid
            conn.commit()
            conn.close()
            return record_id, "Registro creado exitosamente"
        except Exception as e:
            conn.close()
            return None, str(e)

def get_all_document_records():
    """Retrieves all document records."""
    conn = db.get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM document_records ORDER BY created_at DESC")
    rows = cursor.fetchall()
    conn.close()
    
    records = []
    for row in rows:
        records.append(dict(row))
    return records

def delete_document_record(record_id):
    """Deletes a document record."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM document_records WHERE id = ?", (record_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False

def update_record_status(record_id, status):
    """Updates the status of a record."""
    with db._db_lock:
        conn = db.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE document_records SET status = ? WHERE id = ?", (status, record_id))
        conn.commit()
        conn.close()
