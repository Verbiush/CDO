
-- SCHEMA: OrganizadorArchivos (MySQL)
-- Version: 2.1 (Relational - Corrected V2)

-- 1. Tabla: PACIENTES
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
);

-- 2. Tabla: FACTURAS
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
    tipo_servicio VARCHAR(255) DEFAULT 'EVENTO'
);

-- 3. Tabla: ATENCIONES (Estudios/Admisiones)
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
    FOREIGN KEY(paciente_id) REFERENCES pacientes(id) ON DELETE CASCADE,
    FOREIGN KEY(factura_id) REFERENCES facturas(id) ON DELETE SET NULL
);

-- Índices
CREATE INDEX idx_pacientes_doc ON pacientes(no_doc);
CREATE INDEX idx_atenciones_estudio ON atenciones(nro_estudio);
CREATE INDEX idx_facturas_numero ON facturas(no_factura);
