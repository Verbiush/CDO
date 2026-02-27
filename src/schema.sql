
-- SCHEMA: OrganizadorArchivos (SQLite)
-- Version: 2.0 (Relational)

-- 1. Tabla: PACIENTES
-- Almacena la información demográfica única del paciente.
CREATE TABLE IF NOT EXISTS pacientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_doc TEXT,           -- CC, TI, RC, etc.
    no_doc TEXT UNIQUE,      -- Número de documento (Clave Única)
    nombre_completo TEXT,    -- Nombre completo del paciente
    nombre_tercero TEXT,     -- Nombre del tercero (responsable)
    eps TEXT,                -- EPS o Aseguradora
    regimen TEXT DEFAULT 'SUBSIDIADO', -- Régimen (Subsidiado, Contributivo, etc.)
    categoria TEXT DEFAULT 'NIVEL 1', -- Categoría del paciente (NIVEL 1, NIVEL 2, etc.)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Tabla: ATENCIONES (Estudios/Admisiones)
-- Relaciona una atención médica con un paciente.
CREATE TABLE IF NOT EXISTS atenciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paciente_id INTEGER NOT NULL, -- Clave Foránea a PACIENTES
    nro_estudio TEXT,        -- Número de Estudio o Admisión (Puede ser único por ingreso)
    descripcion_cups TEXT,   -- Descripción del procedimiento o servicio (CUPS)
    fecha_ingreso TEXT,      -- Fecha de Ingreso
    fecha_salida TEXT,       -- Fecha de Salida
    autorizacion TEXT,       -- Número de Autorización
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(paciente_id) REFERENCES pacientes(id) ON DELETE CASCADE
);

-- 3. Tabla: FACTURAS
-- Relaciona una factura con una atención específica.
CREATE TABLE IF NOT EXISTS facturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    atencion_id INTEGER NOT NULL, -- Clave Foránea a ATENCIONES
    no_factura TEXT UNIQUE,  -- Número de Factura (Clave Única)
    fecha_factura TEXT,      -- Fecha de emisión de la factura
    tipo_pago TEXT,          -- Tipo de Pago (Efectivo, Tarjeta, etc.)
    valor_servicio TEXT,     -- Valor del servicio
    copago TEXT,             -- Copago o Cuota Moderadora
    radicado TEXT,           -- Número de Radicado
    total TEXT,              -- Total de la factura
    tipo_servicio TEXT DEFAULT 'EVENTO', -- Tipo de servicio (EVENTO, CAPITA, etc.)
    status TEXT DEFAULT 'PENDING', -- Estado del proceso (PENDING, PROCESSED, ERROR)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(atencion_id) REFERENCES atenciones(id) ON DELETE CASCADE
);

-- Índices para mejorar el rendimiento de búsquedas
CREATE INDEX IF NOT EXISTS idx_pacientes_doc ON pacientes(no_doc);
CREATE INDEX IF NOT EXISTS idx_atenciones_estudio ON atenciones(nro_estudio);
CREATE INDEX IF NOT EXISTS idx_facturas_numero ON facturas(no_factura);
