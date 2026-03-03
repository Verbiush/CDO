-- SCHEMA: OrganizadorArchivos (PostgreSQL / Supabase)
-- Version: 2.1 (Relational)

-- 1. Tabla: PACIENTES
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
);

-- 2. Tabla: FACTURAS
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
    tipo_servicio VARCHAR(255) DEFAULT 'EVENTO'
);

-- 3. Tabla: ATENCIONES (Estudios/Admisiones)
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
    FOREIGN KEY(paciente_id) REFERENCES pacientes(id) ON DELETE CASCADE,
    FOREIGN KEY(factura_id) REFERENCES facturas(id) ON DELETE SET NULL
);

-- Índices
CREATE INDEX IF NOT EXISTS idx_pacientes_doc ON pacientes(no_doc);
CREATE INDEX IF NOT EXISTS idx_atenciones_estudio ON atenciones(nro_estudio);
CREATE INDEX IF NOT EXISTS idx_facturas_numero ON facturas(no_factura);

-- 4. Tabla: USERS
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username VARCHAR(255) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role VARCHAR(50) DEFAULT 'user',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 5. Tabla: TASKS (Cola de procesamiento)
CREATE TABLE IF NOT EXISTS tasks (
    id VARCHAR(255) PRIMARY KEY,
    status VARCHAR(50),
    progress INT DEFAULT 0,
    message TEXT,
    result_data TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 6. Tabla: DOCUMENT_RECORDS (Registro de archivos procesados)
CREATE TABLE IF NOT EXISTS document_records (
    id SERIAL PRIMARY KEY,
    filename VARCHAR(255) NOT NULL,
    filepath TEXT NOT NULL,
    file_type VARCHAR(50),
    upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    processed BOOLEAN DEFAULT FALSE,
    atencion_id INT,
    FOREIGN KEY(atencion_id) REFERENCES atenciones(id) ON DELETE SET NULL
);
