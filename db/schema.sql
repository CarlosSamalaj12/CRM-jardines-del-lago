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

CREATE TABLE IF NOT EXISTS menu_platos_fuertes (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  tipo_plato VARCHAR(20) NOT NULL DEFAULT 'NORMAL',
  es_sin_proteina TINYINT(1) NOT NULL DEFAULT 0,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_platos_fuertes_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_preparaciones (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  id_plato_fuerte BIGINT UNSIGNED NOT NULL,
  nombre VARCHAR(180) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_preparaciones_plato_nombre (id_plato_fuerte, nombre),
  KEY idx_menu_preparaciones_plato (id_plato_fuerte),
  CONSTRAINT fk_menu_preparaciones_plato
    FOREIGN KEY (id_plato_fuerte) REFERENCES menu_platos_fuertes(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_salsas (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_salsas_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_preparacion_salsa_sugerida (
  id_preparacion BIGINT UNSIGNED NOT NULL,
  id_salsa BIGINT UNSIGNED NOT NULL,
  prioridad INT NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_preparacion, id_salsa),
  KEY idx_menu_preparacion_salsa_sugerida_salsa (id_salsa),
  CONSTRAINT fk_menu_preparacion_sugerida_preparacion
    FOREIGN KEY (id_preparacion) REFERENCES menu_preparaciones(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_preparacion_sugerida_salsa
    FOREIGN KEY (id_salsa) REFERENCES menu_salsas(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_preparacion_postre_sugerido (
  id_preparacion BIGINT UNSIGNED NOT NULL,
  id_postre BIGINT UNSIGNED NOT NULL,
  prioridad INT NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_preparacion, id_postre),
  KEY idx_menu_preparacion_postre_sugerido_postre (id_postre),
  CONSTRAINT fk_menu_preparacion_postre_preparacion
    FOREIGN KEY (id_preparacion) REFERENCES menu_preparaciones(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_preparacion_postre_postre
    FOREIGN KEY (id_postre) REFERENCES menu_postres(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_plato_guarnicion_sugerida (
  id_plato_fuerte BIGINT UNSIGNED NOT NULL,
  id_guarnicion BIGINT UNSIGNED NOT NULL,
  prioridad INT NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_plato_fuerte, id_guarnicion),
  KEY idx_menu_plato_guarnicion_sugerida_guarnicion (id_guarnicion),
  CONSTRAINT fk_menu_plato_guarnicion_plato
    FOREIGN KEY (id_plato_fuerte) REFERENCES menu_platos_fuertes(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_plato_guarnicion_guarnicion
    FOREIGN KEY (id_guarnicion) REFERENCES menu_guarniciones(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_guarniciones (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_guarniciones_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_postres (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_postres_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_comentarios_adicionales (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(240) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_comentarios_adicionales_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS montaje_tipos (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_montaje_tipos_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS montaje_adicionales (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(180) NOT NULL,
  tipo VARCHAR(120) NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_montaje_adicionales_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS montaje_tipo_adicional_sugerido (
  id_montaje_tipo BIGINT UNSIGNED NOT NULL,
  id_adicional BIGINT UNSIGNED NOT NULL,
  prioridad INT NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_montaje_tipo, id_adicional),
  KEY idx_montaje_tipo_adicional_sugerido_adicional (id_adicional),
  CONSTRAINT fk_montaje_tipo_adicional_sugerido_tipo
    FOREIGN KEY (id_montaje_tipo) REFERENCES montaje_tipos(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_montaje_tipo_adicional_sugerido_adicional
    FOREIGN KEY (id_adicional) REFERENCES montaje_adicionales(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_montaje_plantillas (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  nombre VARCHAR(200) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_menu_montaje_plantillas_nombre (nombre)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_montaje_plantilla_detalle (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  id_plantilla BIGINT UNSIGNED NOT NULL,
  id_plato_fuerte BIGINT UNSIGNED NULL,
  id_preparacion BIGINT UNSIGNED NULL,
  id_salsa BIGINT UNSIGNED NULL,
  id_postre BIGINT UNSIGNED NULL,
  cantidad INT NULL,
  notas TEXT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_menu_montaje_plantilla_detalle_plantilla (id_plantilla),
  KEY idx_menu_montaje_plantilla_detalle_plato (id_plato_fuerte),
  KEY idx_menu_montaje_plantilla_detalle_preparacion (id_preparacion),
  KEY idx_menu_montaje_plantilla_detalle_salsa (id_salsa),
  KEY idx_menu_montaje_plantilla_detalle_postre (id_postre),
  CONSTRAINT fk_menu_montaje_plantilla_detalle_plantilla
    FOREIGN KEY (id_plantilla) REFERENCES menu_montaje_plantillas(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_detalle_plato
    FOREIGN KEY (id_plato_fuerte) REFERENCES menu_platos_fuertes(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_detalle_preparacion
    FOREIGN KEY (id_preparacion) REFERENCES menu_preparaciones(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_detalle_salsa
    FOREIGN KEY (id_salsa) REFERENCES menu_salsas(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_detalle_postre
    FOREIGN KEY (id_postre) REFERENCES menu_postres(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_montaje_plantilla_guarnicion (
  id_detalle BIGINT UNSIGNED NOT NULL,
  id_guarnicion BIGINT UNSIGNED NOT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_detalle, id_guarnicion),
  KEY idx_menu_montaje_plantilla_guarnicion_guarnicion (id_guarnicion),
  CONSTRAINT fk_menu_montaje_plantilla_guarnicion_detalle
    FOREIGN KEY (id_detalle) REFERENCES menu_montaje_plantilla_detalle(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_guarnicion_guarnicion
    FOREIGN KEY (id_guarnicion) REFERENCES menu_guarniciones(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS menu_montaje_plantilla_adicional (
  id_detalle BIGINT UNSIGNED NOT NULL,
  id_montaje_tipo BIGINT UNSIGNED NULL,
  id_adicional BIGINT UNSIGNED NOT NULL,
  cantidad INT NULL,
  creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id_detalle, id_adicional),
  KEY idx_menu_montaje_plantilla_adicional_tipo (id_montaje_tipo),
  KEY idx_menu_montaje_plantilla_adicional_adicional (id_adicional),
  CONSTRAINT fk_menu_montaje_plantilla_adicional_detalle
    FOREIGN KEY (id_detalle) REFERENCES menu_montaje_plantilla_detalle(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_adicional_tipo
    FOREIGN KEY (id_montaje_tipo) REFERENCES montaje_tipos(id)
    ON DELETE SET NULL
    ON UPDATE CASCADE,
  CONSTRAINT fk_menu_montaje_plantilla_adicional_adicional
    FOREIGN KEY (id_adicional) REFERENCES montaje_adicionales(id)
    ON DELETE CASCADE
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
