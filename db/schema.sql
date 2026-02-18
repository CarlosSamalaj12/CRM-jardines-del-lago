USE crm_jdl;

CREATE TABLE IF NOT EXISTS salones (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(120) NOT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_salones_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS usuarios (
  id VARCHAR(80) NOT NULL,
  nombre VARCHAR(160) NOT NULL,
  nombre_usuario VARCHAR(120) NULL,
  nombre_completo VARCHAR(200) NULL,
  correo VARCHAR(200) NULL,
  telefono VARCHAR(80) NULL,
  contrasena VARCHAR(255) NULL,
  firma_data_url LONGTEXT NULL,
  avatar_data_url LONGTEXT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  influye_meta_ventas TINYINT(1) NOT NULL DEFAULT 0,
  metas_mensuales_json LONGTEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS empresas (
  id VARCHAR(80) NOT NULL,
  nombre VARCHAR(200) NOT NULL,
  encargado_principal VARCHAR(200) NULL,
  correo VARCHAR(200) NULL,
  nit VARCHAR(64) NULL,
  razon_social VARCHAR(220) NULL,
  tipo_evento VARCHAR(120) NULL,
  direccion VARCHAR(300) NULL,
  telefono VARCHAR(80) NULL,
  notas TEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS encargados_empresa (
  id VARCHAR(80) NOT NULL,
  id_empresa VARCHAR(80) NOT NULL,
  nombre VARCHAR(200) NOT NULL,
  telefono VARCHAR(80) NULL,
  correo VARCHAR(200) NULL,
  direccion VARCHAR(300) NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_encargados_empresa (id_empresa),
  CONSTRAINT fk_encargados_empresa
    FOREIGN KEY (id_empresa) REFERENCES empresas(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS servicios (
  id VARCHAR(80) NOT NULL,
  nombre VARCHAR(220) NOT NULL,
  precio DECIMAL(12,2) NOT NULL DEFAULT 0,
  descripcion TEXT NULL,
  id_categoria BIGINT UNSIGNED NULL,
  id_subcategoria BIGINT UNSIGNED NULL,
  modo_cantidad VARCHAR(12) NOT NULL DEFAULT 'MANUAL',
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS categorias_servicio (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(140) NOT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_categorias_servicio_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS subcategorias_servicio (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  id_categoria BIGINT UNSIGNED NOT NULL,
  nombre VARCHAR(140) NOT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_subcategorias_servicio (id_categoria, nombre),
  KEY idx_subcategorias_servicio_categoria (id_categoria),
  CONSTRAINT fk_subcategorias_categoria
    FOREIGN KEY (id_categoria) REFERENCES categorias_servicio(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS eventos (
  id VARCHAR(80) NOT NULL,
  id_grupo VARCHAR(120) NULL,
  nombre VARCHAR(240) NOT NULL,
  nombre_salon VARCHAR(120) NOT NULL,
  fecha_evento DATE NOT NULL,
  hora_inicio TIME NOT NULL,
  hora_fin TIME NOT NULL,
  estado VARCHAR(80) NOT NULL,
  id_usuario VARCHAR(80) NULL,
  pax INT NULL,
  notas TEXT NULL,
  cotizacion_json LONGTEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_eventos_grupo (id_grupo),
  KEY idx_eventos_fecha_salon (fecha_evento, nombre_salon),
  KEY idx_eventos_usuario (id_usuario),
  CONSTRAINT fk_eventos_usuario
    FOREIGN KEY (id_usuario) REFERENCES usuarios(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS cotizaciones_evento (
  id_evento VARCHAR(80) NOT NULL,
  id_empresa VARCHAR(80) NULL,
  id_encargado VARCHAR(80) NULL,
  nombre_empresa VARCHAR(200) NULL,
  nombre_encargado VARCHAR(200) NULL,
  contacto VARCHAR(200) NULL,
  correo VARCHAR(200) NULL,
  facturar_a VARCHAR(220) NULL,
  direccion VARCHAR(300) NULL,
  tipo_evento VARCHAR(120) NULL,
  lugar VARCHAR(160) NULL,
  horario_texto VARCHAR(180) NULL,
  codigo VARCHAR(120) NULL,
  fecha_documento DATE NULL,
  telefono VARCHAR(80) NULL,
  nit VARCHAR(64) NULL,
  personas INT NULL,
  fecha_evento DATE NULL,
  folio VARCHAR(120) NULL,
  fecha_fin DATE NULL,
  fecha_max_pago DATE NULL,
  tipo_pago VARCHAR(120) NULL,
  notas_internas TEXT NULL,
  notas TEXT NULL,
  cotizado_en_iso VARCHAR(50) NULL,
  json_crudo LONGTEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id_evento),
  KEY idx_cotizaciones_empresa (id_empresa),
  KEY idx_cotizaciones_encargado (id_encargado),
  CONSTRAINT fk_cotizaciones_evento
    FOREIGN KEY (id_evento) REFERENCES eventos(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS items_cotizacion_evento (
  id VARCHAR(80) NOT NULL,
  id_evento VARCHAR(80) NOT NULL,
  id_servicio VARCHAR(80) NULL,
  fecha_servicio DATE NULL,
  cantidad DECIMAL(12,2) NOT NULL DEFAULT 0,
  precio DECIMAL(12,2) NOT NULL DEFAULT 0,
  nombre VARCHAR(260) NOT NULL,
  descripcion TEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_items_cotizacion_evento (id_evento),
  KEY idx_items_cotizacion_servicio (id_servicio),
  CONSTRAINT fk_items_cotizacion_evento
    FOREIGN KEY (id_evento) REFERENCES cotizaciones_evento(id_evento)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS historial_evento (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  clave_evento VARCHAR(120) NOT NULL,
  cambiado_en_iso VARCHAR(50) NULL,
  cambiado_en DATETIME NULL,
  id_usuario_actor VARCHAR(80) NULL,
  nombre_actor VARCHAR(200) NULL,
  cambio_texto TEXT NOT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_historial_clave_evento (clave_evento),
  KEY idx_historial_usuario_actor (id_usuario_actor)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS recordatorios_evento (
  id VARCHAR(80) NOT NULL,
  clave_evento VARCHAR(120) NOT NULL,
  fecha_recordatorio DATE NULL,
  hora_recordatorio TIME NULL,
  medio VARCHAR(80) NOT NULL,
  notas TEXT NULL,
  creado_en_iso VARCHAR(50) NULL,
  creado_en DATETIME NULL,
  id_usuario_creador VARCHAR(80) NULL,
  creado_en_fila TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_recordatorios_clave_evento (clave_evento),
  KEY idx_recordatorios_fecha_hora (fecha_recordatorio, hora_recordatorio)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS bitacora_migracion (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  origen VARCHAR(80) NOT NULL,
  detalle TEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
