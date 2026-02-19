const path = require("path");
const express = require("express");
const mariadb = require("mariadb");
const crypto = require("crypto");
require("dotenv").config();

const APP_PORT = Number(process.env.APP_PORT || 3000);
const DB_HOST = process.env.DB_HOST || "127.0.0.1";
const DB_PORT = Number(process.env.DB_PORT || 3306);
const DB_USER = process.env.DB_USER || "root";
const DB_PASSWORD = process.env.DB_PASSWORD || "";
const DB_NAME = process.env.DB_NAME || "crm_jdl";

const REQUIRED_TABLES = [
  "salones",
  "usuarios",
  "empresas",
  "encargados_empresa",
  "categorias_servicio",
  "subcategorias_servicio",
  "servicios",
  "eventos",
  "cotizaciones_evento",
  "items_cotizacion_evento",
  "cotizacion_versiones_evento",
  "items_cotizacion_version_evento",
  "historial_evento",
  "recordatorios_evento",
  "menu_platos_fuertes",
  "menu_preparaciones",
  "menu_salsas",
  "menu_preparacion_salsa_sugerida",
  "menu_preparacion_postre_sugerido",
  "menu_plato_guarnicion_sugerida",
  "menu_guarniciones",
  "menu_postres",
  "menu_bebidas",
  "menu_comentarios_adicionales",
  "montaje_tipos",
  "montaje_adicionales",
  "montaje_tipo_adicional_sugerido",
  "menu_montaje_plantillas",
  "menu_montaje_plantilla_detalle",
  "menu_montaje_plantilla_guarnicion",
  "menu_montaje_plantilla_adicional",
];

const pool = mariadb.createPool({
  host: DB_HOST,
  port: DB_PORT,
  user: DB_USER,
  password: DB_PASSWORD,
  database: DB_NAME,
  connectionLimit: 5,
});

const app = express();
app.use(express.json({ limit: "25mb" }));

function hashPasswordScrypt(password, saltHex) {
  return new Promise((resolve, reject) => {
    const salt = Buffer.from(saltHex, "hex");
    crypto.scrypt(password, salt, 64, { N: 16384, r: 8, p: 1 }, (err, derivedKey) => {
      if (err) return reject(err);
      resolve(derivedKey.toString("hex"));
    });
  });
}

async function ensurePasswordHash(rawPassword) {
  const value = String(rawPassword || "").trim();
  if (!value) return null;
  if (value.startsWith("scrypt$")) return value;
  const saltHex = crypto.randomBytes(16).toString("hex");
  const hashHex = await hashPasswordScrypt(value, saltHex);
  return `scrypt$16384$8$1$${saltHex}$${hashHex}`;
}

async function verifyPassword(rawPassword, storedPassword) {
  const raw = String(rawPassword || "");
  const stored = String(storedPassword || "");
  if (!stored) return false;
  if (!stored.startsWith("scrypt$")) {
    return stored === raw;
  }
  const parts = stored.split("$");
  if (parts.length !== 6) return false;
  const saltHex = parts[4];
  const expectedHex = parts[5];
  const gotHex = await hashPasswordScrypt(raw, saltHex);
  try {
    return crypto.timingSafeEqual(Buffer.from(expectedHex, "hex"), Buffer.from(gotHex, "hex"));
  } catch (_) {
    return false;
  }
}

function toIsoDate(value) {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return String(value).slice(0, 10);
  return d.toISOString().slice(0, 10);
}

function toHHmm(value) {
  if (!value) return "";
  const raw = String(value);
  if (raw.length >= 5) return raw.slice(0, 5);
  return raw;
}

function asDate(value) {
  const v = String(value || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : null;
}

function asTime(value) {
  const v = String(value || "").trim();
  if (/^\d{2}:\d{2}:\d{2}$/.test(v)) return v;
  if (/^\d{2}:\d{2}$/.test(v)) return `${v}:00`;
  return null;
}

function asDateTime(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return null;
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

async function ensureAppStateExtraStructure() {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(`
      CREATE TABLE IF NOT EXISTS app_state_kv (
        clave VARCHAR(120) NOT NULL,
        valor_json LONGTEXT NULL,
        actualizado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        PRIMARY KEY (clave)
      )
    `);
  } finally {
    if (conn) conn.release();
  }
}

async function ensureUsersExtendedStructure() {
  let conn;
  try {
    conn = await pool.getConnection();
    const cols = await conn.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = ? AND table_name = 'usuarios'`,
      [DB_NAME]
    );
    const colSet = new Set(cols.map((r) => String(r.column_name || "").toLowerCase()));
    if (!colSet.has("nombre_usuario")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN nombre_usuario VARCHAR(120) NULL");
    }
    if (!colSet.has("nombre_completo")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN nombre_completo VARCHAR(200) NULL");
    }
    if (!colSet.has("correo")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN correo VARCHAR(200) NULL");
    }
    if (!colSet.has("telefono")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN telefono VARCHAR(80) NULL");
    }
    if (!colSet.has("contrasena")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN contrasena VARCHAR(255) NULL");
    }
    if (!colSet.has("firma_data_url")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN firma_data_url LONGTEXT NULL");
    }
    if (!colSet.has("avatar_data_url")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN avatar_data_url LONGTEXT NULL");
    }
    if (!colSet.has("activo")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN activo TINYINT(1) NOT NULL DEFAULT 1");
    }
    if (!colSet.has("influye_meta_ventas")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN influye_meta_ventas TINYINT(1) NOT NULL DEFAULT 0");
    }
    if (!colSet.has("metas_mensuales_json")) {
      await conn.query("ALTER TABLE usuarios ADD COLUMN metas_mensuales_json LONGTEXT NULL");
    }
  } finally {
    if (conn) conn.release();
  }
}

function str(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function normalizeQuantityMode(mode) {
  const m = String(mode || "").trim().toUpperCase();
  return m === "PAX" ? "PAX" : "MANUAL";
}

function normalizeDiscountType(type) {
  return String(type || "").trim().toUpperCase() === "PERCENT" ? "PERCENT" : "AMOUNT";
}

function normalizeMenuDishType(type) {
  const value = String(type || "").trim().toUpperCase();
  if (value === "VEGETARIANO") return "VEGETARIANO";
  if (value === "VEGANO") return "VEGANO";
  return "NORMAL";
}

function parseQuoteCodeNumber(rawCode) {
  const m = String(rawCode || "").trim().toUpperCase().match(/^COT-(\d{1,})$/);
  if (!m) return 0;
  const n = Number(m[1]);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function formatQuoteCode(num) {
  const n = Math.max(1, Number(num || 1));
  return `COT-${String(Math.floor(n)).padStart(3, "0")}`;
}

async function ensureDocumentSequenceStructure() {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(`
      CREATE TABLE IF NOT EXISTS doc_sequence (
        scope VARCHAR(40) NOT NULL,
        last_value BIGINT UNSIGNED NOT NULL DEFAULT 0,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        PRIMARY KEY (scope)
      )
    `);
    await conn.query(
      `INSERT INTO doc_sequence (scope, last_value) VALUES ('COT', 0)
       ON DUPLICATE KEY UPDATE scope = VALUES(scope)`
    );
  } finally {
    if (conn) conn.release();
  }
}

async function ensureDocSequenceAtLeast(conn, scope, minValue) {
  const safeMin = Math.max(0, Number(minValue || 0));
  await conn.query(
    `UPDATE doc_sequence
     SET last_value = GREATEST(last_value, ?)
     WHERE scope = ?`,
    [safeMin, String(scope || "COT").trim() || "COT"]
  );
}

async function reserveNextDocCodeInConn(conn, scope = "COT") {
  const key = String(scope || "COT").trim() || "COT";
  await conn.query(
    `INSERT INTO doc_sequence (scope, last_value) VALUES (?, 0)
     ON DUPLICATE KEY UPDATE scope = VALUES(scope)`,
    [key]
  );
  await conn.query(
    `UPDATE doc_sequence
     SET last_value = last_value + 1
     WHERE scope = ?`,
    [key]
  );
  const rows = await conn.query(
    `SELECT last_value FROM doc_sequence WHERE scope = ? LIMIT 1`,
    [key]
  );
  const value = Number(rows?.[0]?.last_value || 1);
  return formatQuoteCode(value);
}

async function reserveNextDocCode(scope = "COT") {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.beginTransaction();
    const code = await reserveNextDocCodeInConn(conn, scope);
    await conn.commit();
    return code;
  } catch (error) {
    if (conn) await conn.rollback();
    throw error;
  } finally {
    if (conn) conn.release();
  }
}

function buildQuoteItemPrimaryKey(eventId, item, idx) {
  const e = str(eventId).trim() || "ev";
  const raw = str(item?.rowId).trim() || `row_${idx + 1}`;
  const composed = `${e}__${raw}__${idx + 1}`;
  return composed.slice(0, 100);
}

function calculateQuoteTotals(quote) {
  const q = quote && typeof quote === "object" ? quote : {};
  const items = Array.isArray(q.items) ? q.items : [];
  const subtotal = items.reduce((acc, x) => acc + Number(x?.qty || 0) * Number(x?.price || 0), 0);
  const discountType = normalizeDiscountType(q.discountType);
  const discountValueRaw = Math.max(0, Number(q.discountValue || 0));
  const discountValue = Number.isFinite(discountValueRaw) ? discountValueRaw : 0;
  const discountAmount = discountType === "PERCENT"
    ? Math.max(0, Math.min(subtotal, subtotal * Math.min(100, discountValue) / 100))
    : Math.max(0, Math.min(subtotal, discountValue));
  const total = Math.max(0, subtotal - discountAmount);
  return { subtotal, discountType, discountValue, discountAmount, total };
}

async function ensureServiceCatalogStructure() {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(`
      CREATE TABLE IF NOT EXISTS categorias_servicio (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(140) NOT NULL,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_categorias_servicio_nombre (nombre)
      )
    `);
    await conn.query(`
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
      )
    `);

    const cols = await conn.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = ? AND table_name = 'servicios'`,
      [DB_NAME]
    );
    const colSet = new Set(cols.map((r) => String(r.column_name || "").toLowerCase()));
    if (!colSet.has("id_categoria")) {
      await conn.query("ALTER TABLE servicios ADD COLUMN id_categoria BIGINT UNSIGNED NULL");
    }
    if (!colSet.has("id_subcategoria")) {
      await conn.query("ALTER TABLE servicios ADD COLUMN id_subcategoria BIGINT UNSIGNED NULL");
    }
    if (!colSet.has("modo_cantidad")) {
      await conn.query("ALTER TABLE servicios ADD COLUMN modo_cantidad VARCHAR(12) NOT NULL DEFAULT 'MANUAL'");
    }

    const stats = await conn.query(
      `SELECT DISTINCT index_name FROM information_schema.statistics WHERE table_schema = ? AND table_name = 'servicios'`,
      [DB_NAME]
    );
    const idxSet = new Set(stats.map((r) => String(r.index_name || "").toLowerCase()));
    if (!idxSet.has("idx_servicios_categoria")) {
      await conn.query("ALTER TABLE servicios ADD KEY idx_servicios_categoria (id_categoria)");
    }
    if (!idxSet.has("idx_servicios_subcategoria")) {
      await conn.query("ALTER TABLE servicios ADD KEY idx_servicios_subcategoria (id_subcategoria)");
    }

    const constraints = await conn.query(
      `SELECT constraint_name FROM information_schema.table_constraints WHERE table_schema = ? AND table_name = 'servicios' AND constraint_type = 'FOREIGN KEY'`,
      [DB_NAME]
    );
    const fkSet = new Set(constraints.map((r) => String(r.constraint_name || "").toLowerCase()));
    if (!fkSet.has("fk_servicios_categoria")) {
      await conn.query(`
        ALTER TABLE servicios
        ADD CONSTRAINT fk_servicios_categoria
          FOREIGN KEY (id_categoria) REFERENCES categorias_servicio(id)
          ON DELETE SET NULL
          ON UPDATE CASCADE
      `);
    }
    if (!fkSet.has("fk_servicios_subcategoria")) {
      await conn.query(`
        ALTER TABLE servicios
        ADD CONSTRAINT fk_servicios_subcategoria
          FOREIGN KEY (id_subcategoria) REFERENCES subcategorias_servicio(id)
          ON DELETE SET NULL
          ON UPDATE CASCADE
      `);
    }
  } finally {
    if (conn) conn.release();
  }
}

async function ensureQuoteVersionStructure() {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(`
      CREATE TABLE IF NOT EXISTS cotizacion_versiones_evento (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        id_evento VARCHAR(100) NOT NULL,
        version_num INT NOT NULL,
        subtotal DECIMAL(12,2) NOT NULL DEFAULT 0,
        descuento_tipo VARCHAR(12) NOT NULL DEFAULT 'AMOUNT',
        descuento_valor DECIMAL(12,2) NOT NULL DEFAULT 0,
        descuento_monto DECIMAL(12,2) NOT NULL DEFAULT 0,
        total_neto DECIMAL(12,2) NOT NULL DEFAULT 0,
        cotizado_en_iso VARCHAR(40) NULL,
        json_crudo LONGTEXT NULL,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_cotizacion_version_evento (id_evento, version_num),
        KEY idx_cotizacion_version_evento (id_evento)
      )
    `);
    await conn.query(`
      CREATE TABLE IF NOT EXISTS items_cotizacion_version_evento (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        id_evento VARCHAR(100) NOT NULL,
        version_num INT NOT NULL,
        fila_num INT NOT NULL,
        id_servicio VARCHAR(100) NULL,
        fecha_servicio DATE NULL,
        cantidad DECIMAL(12,2) NOT NULL DEFAULT 0,
        precio DECIMAL(12,2) NOT NULL DEFAULT 0,
        nombre VARCHAR(240) NOT NULL,
        descripcion TEXT NULL,
        PRIMARY KEY (id),
        KEY idx_items_cotizacion_version_evento (id_evento, version_num)
      )
    `);

    const itemCols = await conn.query(
      `SELECT table_name, column_name FROM information_schema.columns WHERE table_schema = ? AND table_name IN ('items_cotizacion_evento','items_cotizacion_version_evento')`,
      [DB_NAME]
    );
    const itemMap = new Map();
    for (const row of itemCols) {
      const t = String(row.table_name || "").toLowerCase();
      const c = String(row.column_name || "").toLowerCase();
      if (!itemMap.has(t)) itemMap.set(t, new Set());
      itemMap.get(t).add(c);
    }
    const itemTables = ["items_cotizacion_evento", "items_cotizacion_version_evento"];
    for (const table of itemTables) {
      const set = itemMap.get(table) || new Set();
      if (!set.has("precio_unitario")) {
        await conn.query(`ALTER TABLE ${table} ADD COLUMN precio_unitario DECIMAL(12,2) NOT NULL DEFAULT 0`);
      }
      if (!set.has("modo_cantidad")) {
        await conn.query(`ALTER TABLE ${table} ADD COLUMN modo_cantidad VARCHAR(12) NOT NULL DEFAULT 'MANUAL'`);
      }
      if (!set.has("total_linea")) {
        await conn.query(`ALTER TABLE ${table} ADD COLUMN total_linea DECIMAL(12,2) NOT NULL DEFAULT 0`);
      }
    }

    const cotizacionCols = await conn.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = ? AND table_name = 'cotizaciones_evento'`,
      [DB_NAME]
    );
    const cotizacionColSet = new Set(cotizacionCols.map((r) => String(r.column_name || "").toLowerCase()));
    if (!cotizacionColSet.has("version_actual")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN version_actual INT NOT NULL DEFAULT 1");
    }
    if (!cotizacionColSet.has("subtotal")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN subtotal DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!cotizacionColSet.has("descuento_tipo")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN descuento_tipo VARCHAR(12) NOT NULL DEFAULT 'AMOUNT'");
    }
    if (!cotizacionColSet.has("descuento_valor")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN descuento_valor DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!cotizacionColSet.has("descuento_monto")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN descuento_monto DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!cotizacionColSet.has("total_neto")) {
      await conn.query("ALTER TABLE cotizaciones_evento ADD COLUMN total_neto DECIMAL(12,2) NOT NULL DEFAULT 0");
    }

    const versionCols = await conn.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = ? AND table_name = 'cotizacion_versiones_evento'`,
      [DB_NAME]
    );
    const versionColSet = new Set(versionCols.map((r) => String(r.column_name || "").toLowerCase()));
    if (!versionColSet.has("subtotal")) {
      await conn.query("ALTER TABLE cotizacion_versiones_evento ADD COLUMN subtotal DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!versionColSet.has("descuento_tipo")) {
      await conn.query("ALTER TABLE cotizacion_versiones_evento ADD COLUMN descuento_tipo VARCHAR(12) NOT NULL DEFAULT 'AMOUNT'");
    }
    if (!versionColSet.has("descuento_valor")) {
      await conn.query("ALTER TABLE cotizacion_versiones_evento ADD COLUMN descuento_valor DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!versionColSet.has("descuento_monto")) {
      await conn.query("ALTER TABLE cotizacion_versiones_evento ADD COLUMN descuento_monto DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
    if (!versionColSet.has("total_neto")) {
      await conn.query("ALTER TABLE cotizacion_versiones_evento ADD COLUMN total_neto DECIMAL(12,2) NOT NULL DEFAULT 0");
    }
  } finally {
    if (conn) conn.release();
  }
}

async function ensureMenuMontajeCatalogStructure() {
  let conn;
  try {
    conn = await pool.getConnection();

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_platos_fuertes (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        tipo_plato VARCHAR(20) NOT NULL DEFAULT 'NORMAL',
        es_sin_proteina TINYINT(1) NOT NULL DEFAULT 0,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_platos_fuertes_nombre (nombre)
      )
    `);

    const menuPlatoCols = await conn.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = ? AND table_name = 'menu_platos_fuertes'`,
      [DB_NAME]
    );
    const menuPlatoColSet = new Set(menuPlatoCols.map((r) => String(r.column_name || "").toLowerCase()));
    if (!menuPlatoColSet.has("tipo_plato")) {
      await conn.query("ALTER TABLE menu_platos_fuertes ADD COLUMN tipo_plato VARCHAR(20) NOT NULL DEFAULT 'NORMAL'");
    }
    if (!menuPlatoColSet.has("es_sin_proteina")) {
      await conn.query("ALTER TABLE menu_platos_fuertes ADD COLUMN es_sin_proteina TINYINT(1) NOT NULL DEFAULT 0");
    }

    await conn.query(`
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
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_salsas (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_salsas_nombre (nombre)
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_guarniciones (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_guarniciones_nombre (nombre)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_postres (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_postres_nombre (nombre)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_bebidas (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_bebidas_nombre (nombre)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_comentarios_adicionales (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(240) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_comentarios_adicionales_nombre (nombre)
      )
    `);

    // Migracion legacy: si bebidas existian en app_state_kv JSON, pasarlas a tabla relacional.
    try {
      const legacyRows = await conn.query(
        "SELECT valor_json FROM app_state_kv WHERE clave = 'menuMontajeBebidas' LIMIT 1"
      );
      const raw = str(legacyRows?.[0]?.valor_json);
      if (raw) {
        let parsed = [];
        try {
          parsed = JSON.parse(raw);
        } catch (_) {
          parsed = [];
        }
        if (Array.isArray(parsed)) {
          for (const row of parsed) {
            const nombre = str(row?.nombre || row?.name || row).trim();
            if (!nombre) continue;
            const activo = row?.activo === false ? 0 : 1;
            await conn.query(
              `INSERT INTO menu_bebidas (nombre, activo) VALUES (?, ?)
               ON DUPLICATE KEY UPDATE activo = VALUES(activo)`,
              [nombre, activo]
            );
          }
        }
      }
    } catch (_) {
      // app_state_kv puede no existir todavia en instalaciones viejas durante bootstrap.
    }

    await conn.query(`
      CREATE TABLE IF NOT EXISTS montaje_tipos (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_montaje_tipos_nombre (nombre)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS montaje_adicionales (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(180) NOT NULL,
        tipo VARCHAR(120) NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_montaje_adicionales_nombre (nombre)
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS menu_montaje_plantillas (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        nombre VARCHAR(200) NOT NULL,
        activo TINYINT(1) NOT NULL DEFAULT 1,
        creado_en TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_menu_montaje_plantillas_nombre (nombre)
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
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
      )
    `);

    await conn.query(`
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
      )
    `);
  } finally {
    if (conn) conn.release();
  }
}

async function ensureRequiredTables() {
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      `SELECT table_name FROM information_schema.tables WHERE table_schema = ?`,
      [DB_NAME]
    );
    const existing = new Set(rows.map((r) => String(r.table_name || "").toLowerCase()));
    const missing = REQUIRED_TABLES.filter((t) => !existing.has(t));
    if (missing.length) {
      throw new Error(`Faltan tablas requeridas: ${missing.join(", ")}`);
    }
  } finally {
    if (conn) conn.release();
  }
}

async function readStateFromTables() {
  let conn;
  try {
    conn = await pool.getConnection();

    const [
      salones,
      usuarios,
      empresas,
      encargados,
      servicios,
      eventos,
      historial,
      recordatorios,
      bebidasCatalog,
      appStateRows,
    ] = await Promise.all([
      conn.query("SELECT id, nombre FROM salones ORDER BY id"),
      conn.query("SELECT id, nombre, nombre_usuario, nombre_completo, correo, telefono, contrasena, firma_data_url, avatar_data_url, activo, influye_meta_ventas, metas_mensuales_json FROM usuarios ORDER BY creado_en, id"),
      conn.query("SELECT id, nombre, encargado_principal, correo, nit, razon_social, tipo_evento, direccion, telefono, notas FROM empresas ORDER BY creado_en, id"),
      conn.query("SELECT id, id_empresa, nombre, telefono, correo, direccion FROM encargados_empresa ORDER BY creado_en, id"),
      conn.query(`
        SELECT
          s.id,
          s.nombre,
          s.precio,
          s.descripcion,
          s.id_categoria,
          s.id_subcategoria,
          s.modo_cantidad,
          c.nombre AS categoria_nombre,
          sc.nombre AS subcategoria_nombre
        FROM servicios s
        LEFT JOIN categorias_servicio c ON c.id = s.id_categoria
        LEFT JOIN subcategorias_servicio sc ON sc.id = s.id_subcategoria
        ORDER BY s.creado_en, s.id
      `),
      conn.query("SELECT id, id_grupo, nombre, nombre_salon, fecha_evento, hora_inicio, hora_fin, estado, id_usuario, pax, notas, cotizacion_json FROM eventos ORDER BY fecha_evento, hora_inicio, id"),
      conn.query("SELECT clave_evento, cambiado_en_iso, id_usuario_actor, nombre_actor, cambio_texto FROM historial_evento ORDER BY id DESC"),
      conn.query("SELECT id, clave_evento, fecha_recordatorio, hora_recordatorio, medio, notas, creado_en_iso, id_usuario_creador FROM recordatorios_evento ORDER BY id"),
      conn.query("SELECT id, nombre, activo FROM menu_bebidas ORDER BY nombre ASC"),
      conn.query("SELECT clave, valor_json FROM app_state_kv WHERE clave IN ('quickTemplates','disabledCompanies','disabledServices','disabledManagers','disabledSalones','globalMonthlyGoals','checklistTemplateItems','checklistTemplateSections','menuMontajeSections','menuMontajeBebidas','eventChecklists')"),
    ]);

    const hasData = salones.length || usuarios.length || empresas.length || servicios.length || eventos.length;
    if (!hasData) return null;

    const managersByCompany = new Map();
    for (const m of encargados) {
      const companyId = str(m.id_empresa);
      if (!companyId) continue;
      if (!managersByCompany.has(companyId)) managersByCompany.set(companyId, []);
      managersByCompany.get(companyId).push({
        id: str(m.id),
        name: str(m.nombre),
        phone: str(m.telefono),
        email: str(m.correo),
        address: str(m.direccion),
      });
    }

    const state = {
      salones: salones.map((r) => str(r.nombre)).filter(Boolean),
      users: usuarios.map((u) => {
        let monthlyGoals = [];
        try {
          const parsed = JSON.parse(str(u.metas_mensuales_json) || "[]");
          if (Array.isArray(parsed)) {
            monthlyGoals = parsed
              .map((g) => ({
                month: str(g?.month),
                amount: Number(g?.amount || 0),
              }))
              .filter((g) => /^\d{4}-\d{2}$/.test(g.month) && Number.isFinite(g.amount) && g.amount > 0);
          }
        } catch (_) { }
        return {
          id: str(u.id),
          name: str(u.nombre),
          username: str(u.nombre_usuario),
          fullName: str(u.nombre_completo) || str(u.nombre),
          email: str(u.correo),
          phone: str(u.telefono),
          password: str(u.contrasena),
          signatureDataUrl: str(u.firma_data_url),
          avatarDataUrl: str(u.avatar_data_url),
          active: Number(u.activo) !== 0,
          salesTargetEnabled: Number(u.influye_meta_ventas) !== 0,
          monthlyGoals,
        };
      }).filter((u) => u.id && u.name),
      companies: empresas.map((c) => ({
        id: str(c.id),
        name: str(c.nombre),
        owner: str(c.encargado_principal),
        email: str(c.correo),
        nit: str(c.nit) || "CF",
        businessName: str(c.razon_social) || str(c.nombre),
        eventType: str(c.tipo_evento) || "Social",
        address: str(c.direccion),
        phone: str(c.telefono),
        notes: str(c.notas),
        managers: managersByCompany.get(str(c.id)) || [],
      })),
      services: servicios.map((s) => ({
        id: str(s.id),
        name: str(s.nombre),
        price: Number(s.precio || 0),
        description: str(s.descripcion),
        categoryId: s.id_categoria == null ? null : Number(s.id_categoria),
        subcategoryId: s.id_subcategoria == null ? null : Number(s.id_subcategoria),
        category: str(s.categoria_nombre),
        subcategory: str(s.subcategoria_nombre),
        quantityMode: normalizeQuantityMode(s.modo_cantidad),
      })),
      quickTemplates: [],
      disabledCompanies: [],
      disabledServices: [],
      disabledManagers: [],
      disabledSalones: [],
      globalMonthlyGoals: [],
      checklistTemplateItems: [],
      checklistTemplateSections: ["General"],
      menuMontajeSections: ["General"],
      menuMontajeBebidas: (Array.isArray(bebidasCatalog) ? bebidasCatalog : []).map((b) => ({
        id: str(b.id),
        nombre: str(b.nombre),
        activo: Number(b.activo) !== 0,
      })),
      eventChecklists: {},
      changeHistory: {},
      reminders: {},
      events: eventos.map((e) => {
        let quote = null;
        if (e.cotizacion_json) {
          try {
            quote = JSON.parse(e.cotizacion_json);
          } catch (_) {
            quote = null;
          }
        }
        return {
          id: str(e.id),
          groupId: e.id_grupo ? str(e.id_grupo) : null,
          name: str(e.nombre),
          salon: str(e.nombre_salon),
          date: toIsoDate(e.fecha_evento),
          status: str(e.estado),
          startTime: toHHmm(e.hora_inicio),
          endTime: toHHmm(e.hora_fin),
          userId: str(e.id_usuario),
          pax: e.pax === null || e.pax === undefined ? null : Number(e.pax),
          notes: str(e.notas),
          quote,
        };
      }),
    };

    for (const row of historial) {
      const key = str(row.clave_evento);
      if (!key) continue;
      if (!Array.isArray(state.changeHistory[key])) state.changeHistory[key] = [];
      state.changeHistory[key].push({
        at: str(row.cambiado_en_iso) || new Date().toISOString(),
        actorUserId: str(row.id_usuario_actor),
        actorName: str(row.nombre_actor) || "Sistema",
        change: str(row.cambio_texto),
      });
    }

    const quickTemplatesRow = appStateRows.find((r) => str(r.clave) === "quickTemplates");
    if (quickTemplatesRow?.valor_json) {
      try {
        const parsed = JSON.parse(quickTemplatesRow.valor_json);
        state.quickTemplates = Array.isArray(parsed) ? parsed : [];
      } catch (_) {
        state.quickTemplates = [];
      }
    }
    const disabledCompaniesRow = appStateRows.find((r) => str(r.clave) === "disabledCompanies");
    const disabledServicesRow = appStateRows.find((r) => str(r.clave) === "disabledServices");
    const disabledManagersRow = appStateRows.find((r) => str(r.clave) === "disabledManagers");
    const disabledSalonesRow = appStateRows.find((r) => str(r.clave) === "disabledSalones");
    const globalMonthlyGoalsRow = appStateRows.find((r) => str(r.clave) === "globalMonthlyGoals");
    const checklistTemplateItemsRow = appStateRows.find((r) => str(r.clave) === "checklistTemplateItems");
    const checklistTemplateSectionsRow = appStateRows.find((r) => str(r.clave) === "checklistTemplateSections");
    const menuMontajeSectionsRow = appStateRows.find((r) => str(r.clave) === "menuMontajeSections");
    const menuMontajeBebidasRow = appStateRows.find((r) => str(r.clave) === "menuMontajeBebidas");
    const eventChecklistsRow = appStateRows.find((r) => str(r.clave) === "eventChecklists");

    const parseArray = (row) => {
      if (!row?.valor_json) return [];
      try {
        const parsed = JSON.parse(row.valor_json);
        return Array.isArray(parsed) ? parsed.map((x) => str(x)).filter(Boolean) : [];
      } catch (_) {
        return [];
      }
    };
    state.disabledCompanies = parseArray(disabledCompaniesRow);
    state.disabledServices = parseArray(disabledServicesRow);
    state.disabledManagers = parseArray(disabledManagersRow);
    state.disabledSalones = parseArray(disabledSalonesRow);
    try {
      const parsed = JSON.parse(str(globalMonthlyGoalsRow?.valor_json) || "[]");
      state.globalMonthlyGoals = Array.isArray(parsed)
        ? parsed
          .map((g) => ({ month: str(g?.month), amount: Math.max(0, Number(g?.amount || 0)) }))
          .filter((g) => /^\d{4}-\d{2}$/.test(g.month))
          .sort((a, b) => a.month.localeCompare(b.month))
        : [];
    } catch (_) {
      state.globalMonthlyGoals = [];
    }
    try {
      const parsed = JSON.parse(str(checklistTemplateItemsRow?.valor_json) || "[]");
      state.checklistTemplateItems = Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      state.checklistTemplateItems = [];
    }
    try {
      const parsed = JSON.parse(str(checklistTemplateSectionsRow?.valor_json) || '["General"]');
      state.checklistTemplateSections = Array.isArray(parsed)
        ? parsed.map((s) => str(s)).filter(Boolean)
        : ["General"];
      if (!state.checklistTemplateSections.length) state.checklistTemplateSections = ["General"];
    } catch (_) {
      state.checklistTemplateSections = ["General"];
    }
    try {
      const parsed = JSON.parse(str(menuMontajeSectionsRow?.valor_json) || '["General"]');
      state.menuMontajeSections = Array.isArray(parsed)
        ? parsed.map((s) => str(s)).filter(Boolean)
        : ["General"];
      if (!state.menuMontajeSections.length) state.menuMontajeSections = ["General"];
    } catch (_) {
      state.menuMontajeSections = ["General"];
    }
    if (!Array.isArray(state.menuMontajeBebidas) || !state.menuMontajeBebidas.length) {
      try {
        const parsed = JSON.parse(str(menuMontajeBebidasRow?.valor_json) || "[]");
        state.menuMontajeBebidas = Array.isArray(parsed)
          ? parsed
            .map((b) => ({
              id: str(b?.id),
              nombre: str(b?.nombre || b?.name),
              activo: b?.activo === false ? false : true,
            }))
            .filter((b) => b.id && b.nombre)
          : [];
      } catch (_) {
        state.menuMontajeBebidas = [];
      }
    }
    try {
      const parsed = JSON.parse(str(eventChecklistsRow?.valor_json) || "{}");
      state.eventChecklists = (parsed && typeof parsed === "object") ? parsed : {};
    } catch (_) {
      state.eventChecklists = {};
    }

    for (const row of recordatorios) {
      const key = str(row.clave_evento);
      if (!key) continue;
      if (!Array.isArray(state.reminders[key])) state.reminders[key] = [];
      state.reminders[key].push({
        id: str(row.id),
        date: toIsoDate(row.fecha_recordatorio),
        time: toHHmm(row.hora_recordatorio),
        channel: str(row.medio),
        notes: str(row.notas),
        createdAt: str(row.creado_en_iso) || new Date().toISOString(),
        createdByUserId: str(row.id_usuario_creador),
      });
    }

    return { state, updatedAt: new Date().toISOString() };
  } finally {
    if (conn) conn.release();
  }
}

async function readSalonesFromTables() {
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query("SELECT nombre FROM salones ORDER BY id");
    return rows.map((r) => str(r.nombre)).filter(Boolean);
  } finally {
    if (conn) conn.release();
  }
}

async function readCategoriasServicioFromTables() {
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      "SELECT id, nombre FROM categorias_servicio ORDER BY nombre ASC"
    );
    return rows.map((r) => ({ id: Number(r.id), nombre: str(r.nombre) }));
  } finally {
    if (conn) conn.release();
  }
}

async function readSubcategoriasServicioFromTables(categoriaId = null) {
  let conn;
  try {
    conn = await pool.getConnection();
    let rows;
    if (categoriaId !== null && categoriaId !== undefined && Number.isFinite(Number(categoriaId))) {
      rows = await conn.query(
        "SELECT id, id_categoria, nombre FROM subcategorias_servicio WHERE id_categoria = ? ORDER BY nombre ASC",
        [Number(categoriaId)]
      );
    } else {
      rows = await conn.query(
        "SELECT id, id_categoria, nombre FROM subcategorias_servicio ORDER BY id_categoria ASC, nombre ASC"
      );
    }
    return rows.map((r) => ({
      id: Number(r.id),
      id_categoria: Number(r.id_categoria),
      nombre: str(r.nombre),
    }));
  } finally {
    if (conn) conn.release();
  }
}

const MENU_CATALOG_TABLES = {
  plato_fuerte: "menu_platos_fuertes",
  salsa: "menu_salsas",
  guarnicion: "menu_guarniciones",
  postre: "menu_postres",
  bebida: "menu_bebidas",
  comentario: "menu_comentarios_adicionales",
  montaje_tipo: "montaje_tipos",
  montaje_adicional: "montaje_adicionales",
};

async function readSimpleMenuCatalog(kind) {
  const table = MENU_CATALOG_TABLES[String(kind || "").trim()];
  if (!table) return [];
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      table === "menu_platos_fuertes"
        ? `SELECT id, nombre, activo, tipo_plato, es_sin_proteina FROM ${table} ORDER BY nombre ASC`
        : (
      table === "montaje_adicionales"
        ? `SELECT id, nombre, activo, tipo FROM ${table} ORDER BY nombre ASC`
        : `SELECT id, nombre, activo FROM ${table} ORDER BY nombre ASC`
        )
    );
    return rows.map((r) => ({
      id: Number(r.id),
      nombre: str(r.nombre),
      activo: Number(r.activo) !== 0,
      tipo: str(r.tipo),
      tipo_plato: normalizeMenuDishType(r.tipo_plato),
      es_sin_proteina: Number(r.es_sin_proteina) !== 0,
    }));
  } finally {
    if (conn) conn.release();
  }
}

async function createSimpleMenuCatalog(kind, nombre, extras = {}) {
  const table = MENU_CATALOG_TABLES[String(kind || "").trim()];
  if (!table) throw new Error("catalog_kind_invalid");
  const cleanName = str(nombre).trim();
  if (!cleanName) throw new Error("catalog_name_required");

  let conn;
  try {
    conn = await pool.getConnection();
    if (table === "menu_platos_fuertes") {
      const tipoPlato = normalizeMenuDishType(extras?.tipo_plato);
      const esSinProteina = Number(extras?.es_sin_proteina) ? 1 : 0;
      await conn.query(
        `INSERT INTO menu_platos_fuertes (nombre, tipo_plato, es_sin_proteina, activo) VALUES (?, ?, ?, 1)
         ON DUPLICATE KEY UPDATE
           nombre = VALUES(nombre),
           tipo_plato = VALUES(tipo_plato),
           es_sin_proteina = VALUES(es_sin_proteina),
           activo = 1`,
        [cleanName, tipoPlato, esSinProteina]
      );
      return;
    }
    if (table === "montaje_adicionales") {
      const tipo = str(extras?.tipo).trim() || null;
      await conn.query(
        `INSERT INTO montaje_adicionales (nombre, tipo, activo) VALUES (?, ?, 1)
         ON DUPLICATE KEY UPDATE nombre = VALUES(nombre), tipo = COALESCE(VALUES(tipo), tipo), activo = 1`,
        [cleanName, tipo]
      );
      return;
    }
    await conn.query(
      `INSERT INTO ${table} (nombre, activo) VALUES (?, 1)
       ON DUPLICATE KEY UPDATE nombre = VALUES(nombre), activo = 1`,
      [cleanName]
    );
  } finally {
    if (conn) conn.release();
  }
}

async function readPreparacionesByPlato(idPlatoFuerte) {
  const id = Number(idPlatoFuerte);
  if (!Number.isFinite(id) || id <= 0) return [];
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      `SELECT id, id_plato_fuerte, nombre, activo
       FROM menu_preparaciones
       WHERE id_plato_fuerte = ?
       ORDER BY nombre ASC`,
      [id]
    );
    return rows.map((r) => ({
      id: Number(r.id),
      id_plato_fuerte: Number(r.id_plato_fuerte),
      nombre: str(r.nombre),
      activo: Number(r.activo) !== 0,
    }));
  } finally {
    if (conn) conn.release();
  }
}

async function createPreparacionForPlato(idPlatoFuerte, nombre) {
  const id = Number(idPlatoFuerte);
  const cleanName = str(nombre).trim();
  if (!Number.isFinite(id) || id <= 0) throw new Error("plato_fuerte_required");
  if (!cleanName) throw new Error("preparacion_required");
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(
      `INSERT INTO menu_preparaciones (id_plato_fuerte, nombre, activo)
       VALUES (?, ?, 1)
       ON DUPLICATE KEY UPDATE nombre = VALUES(nombre), activo = 1`,
      [id, cleanName]
    );
  } finally {
    if (conn) conn.release();
  }
}

async function updateSimpleMenuCatalog(kind, id, changes = {}) {
  const table = MENU_CATALOG_TABLES[String(kind || "").trim()];
  const itemId = Number(id);
  if (!table) throw new Error("catalog_kind_invalid");
  if (!Number.isFinite(itemId) || itemId <= 0) throw new Error("catalog_id_invalid");

  const hasName = Object.prototype.hasOwnProperty.call(changes || {}, "nombre");
  const hasActive = Object.prototype.hasOwnProperty.call(changes || {}, "activo");
  const hasTipo = table === "montaje_adicionales" && Object.prototype.hasOwnProperty.call(changes || {}, "tipo");
  const hasTipoPlato = table === "menu_platos_fuertes" && Object.prototype.hasOwnProperty.call(changes || {}, "tipo_plato");
  const hasSinProteina = table === "menu_platos_fuertes" && Object.prototype.hasOwnProperty.call(changes || {}, "es_sin_proteina");

  const setParts = [];
  const values = [];
  if (hasName) {
    const cleanName = str(changes.nombre).trim();
    if (!cleanName) throw new Error("catalog_name_required");
    setParts.push("nombre = ?");
    values.push(cleanName);
  }
  if (hasActive) {
    setParts.push("activo = ?");
    values.push(Number(changes.activo) ? 1 : 0);
  }
  if (hasTipo) {
    const tipo = str(changes.tipo).trim() || null;
    setParts.push("tipo = ?");
    values.push(tipo);
  }
  if (hasTipoPlato) {
    setParts.push("tipo_plato = ?");
    values.push(normalizeMenuDishType(changes.tipo_plato));
  }
  if (hasSinProteina) {
    setParts.push("es_sin_proteina = ?");
    values.push(Number(changes.es_sin_proteina) ? 1 : 0);
  }
  if (!setParts.length) throw new Error("catalog_changes_required");

  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(
      `UPDATE ${table} SET ${setParts.join(", ")} WHERE id = ?`,
      [...values, itemId]
    );
  } finally {
    if (conn) conn.release();
  }
}

async function updatePreparacionById(id, changes = {}) {
  const itemId = Number(id);
  if (!Number.isFinite(itemId) || itemId <= 0) throw new Error("catalog_id_invalid");

  const hasName = Object.prototype.hasOwnProperty.call(changes || {}, "nombre");
  const hasActive = Object.prototype.hasOwnProperty.call(changes || {}, "activo");
  const hasPlato = Object.prototype.hasOwnProperty.call(changes || {}, "id_plato_fuerte");
  const setParts = [];
  const values = [];

  if (hasName) {
    const cleanName = str(changes.nombre).trim();
    if (!cleanName) throw new Error("preparacion_required");
    setParts.push("nombre = ?");
    values.push(cleanName);
  }
  if (hasActive) {
    setParts.push("activo = ?");
    values.push(Number(changes.activo) ? 1 : 0);
  }
  if (hasPlato) {
    const platoId = Number(changes.id_plato_fuerte);
    if (!Number.isFinite(platoId) || platoId <= 0) throw new Error("plato_fuerte_required");
    setParts.push("id_plato_fuerte = ?");
    values.push(platoId);
  }
  if (!setParts.length) throw new Error("catalog_changes_required");

  let conn;
  try {
    conn = await pool.getConnection();
    await conn.query(
      `UPDATE menu_preparaciones SET ${setParts.join(", ")} WHERE id = ?`,
      [...values, itemId]
    );
  } finally {
    if (conn) conn.release();
  }
}

function normalizeIdList(rawList) {
  if (!Array.isArray(rawList)) return [];
  const out = [];
  for (const x of rawList) {
    const n = Number(x);
    if (!Number.isFinite(n) || n <= 0) continue;
    out.push(Math.floor(n));
  }
  return Array.from(new Set(out));
}

async function readMenuSuggestionLinks({ idPlatoFuerte, idPreparacion }) {
  const platoId = Number(idPlatoFuerte);
  const prepId = Number(idPreparacion);
  let conn;
  try {
    conn = await pool.getConnection();
    const [salsas, postres, guarniciones] = await Promise.all([
      Number.isFinite(prepId) && prepId > 0
        ? conn.query(
          "SELECT id_salsa FROM menu_preparacion_salsa_sugerida WHERE id_preparacion = ? ORDER BY prioridad, id_salsa",
          [prepId]
        )
        : Promise.resolve([]),
      Number.isFinite(prepId) && prepId > 0
        ? conn.query(
          "SELECT id_postre FROM menu_preparacion_postre_sugerido WHERE id_preparacion = ? ORDER BY prioridad, id_postre",
          [prepId]
        )
        : Promise.resolve([]),
      Number.isFinite(platoId) && platoId > 0
        ? conn.query(
          "SELECT id_guarnicion FROM menu_plato_guarnicion_sugerida WHERE id_plato_fuerte = ? ORDER BY prioridad, id_guarnicion",
          [platoId]
        )
        : Promise.resolve([]),
    ]);
    return {
      salsaIds: salsas.map((r) => Number(r.id_salsa)).filter((x) => Number.isFinite(x)),
      postreIds: postres.map((r) => Number(r.id_postre)).filter((x) => Number.isFinite(x)),
      guarnicionIds: guarniciones.map((r) => Number(r.id_guarnicion)).filter((x) => Number.isFinite(x)),
    };
  } finally {
    if (conn) conn.release();
  }
}

async function saveMenuSuggestionLinks({ idPlatoFuerte, idPreparacion, salsaIds, postreIds, guarnicionIds }) {
  const platoId = Number(idPlatoFuerte);
  const prepId = Number(idPreparacion);
  if (!Number.isFinite(platoId) || platoId <= 0) throw new Error("plato_fuerte_required");
  if (!Number.isFinite(prepId) || prepId <= 0) throw new Error("preparacion_required");

  const salsaList = normalizeIdList(salsaIds);
  const postreList = normalizeIdList(postreIds);
  const guarnicionList = normalizeIdList(guarnicionIds);

  let conn;
  try {
    conn = await pool.getConnection();
    await conn.beginTransaction();

    await conn.query("DELETE FROM menu_preparacion_salsa_sugerida WHERE id_preparacion = ?", [prepId]);
    for (let i = 0; i < salsaList.length; i++) {
      await conn.query(
        "INSERT INTO menu_preparacion_salsa_sugerida (id_preparacion, id_salsa, prioridad) VALUES (?, ?, ?)",
        [prepId, salsaList[i], i + 1]
      );
    }

    await conn.query("DELETE FROM menu_preparacion_postre_sugerido WHERE id_preparacion = ?", [prepId]);
    for (let i = 0; i < postreList.length; i++) {
      await conn.query(
        "INSERT INTO menu_preparacion_postre_sugerido (id_preparacion, id_postre, prioridad) VALUES (?, ?, ?)",
        [prepId, postreList[i], i + 1]
      );
    }

    await conn.query("DELETE FROM menu_plato_guarnicion_sugerida WHERE id_plato_fuerte = ?", [platoId]);
    for (let i = 0; i < guarnicionList.length; i++) {
      await conn.query(
        "INSERT INTO menu_plato_guarnicion_sugerida (id_plato_fuerte, id_guarnicion, prioridad) VALUES (?, ?, ?)",
        [platoId, guarnicionList[i], i + 1]
      );
    }

    await conn.commit();
  } catch (error) {
    if (conn) await conn.rollback();
    throw error;
  } finally {
    if (conn) conn.release();
  }
}

async function readLoginUsers() {
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      `
        SELECT id, nombre, nombre_usuario, nombre_completo, avatar_data_url, firma_data_url
        FROM usuarios
        WHERE activo = 1
        ORDER BY nombre_completo, nombre_usuario, nombre
      `
    );
    return rows
      .map((r) => ({
        id: str(r.id),
        name: str(r.nombre),
        fullName: str(r.nombre_completo) || str(r.nombre),
        username: str(r.nombre_usuario),
        avatarDataUrl: str(r.avatar_data_url),
        signatureDataUrl: str(r.firma_data_url),
      }))
      .filter((u) => u.id);
  } finally {
    if (conn) conn.release();
  }
}

async function authenticateUser(userId, password) {
  let conn;
  try {
    conn = await pool.getConnection();
    const rows = await conn.query(
      `
        SELECT id, nombre, nombre_usuario, nombre_completo, contrasena, avatar_data_url, firma_data_url
        FROM usuarios
        WHERE id = ? AND activo = 1
        LIMIT 1
      `,
      [str(userId).trim()]
    );
    const row = rows?.[0];
    if (!row) return null;
    const ok = await verifyPassword(password, str(row.contrasena));
    if (!ok) return null;
    return {
      id: str(row.id),
      name: str(row.nombre),
      fullName: str(row.nombre_completo) || str(row.nombre),
      username: str(row.nombre_usuario),
      avatarDataUrl: str(row.avatar_data_url),
      signatureDataUrl: str(row.firma_data_url),
    };
  } finally {
    if (conn) conn.release();
  }
}

async function ensureDefaultUserCarlos() {
  let conn;
  try {
    conn = await pool.getConnection();
    const username = "csamalaj";
    const existing = await conn.query(
      "SELECT id FROM usuarios WHERE nombre_usuario = ? LIMIT 1",
      [username]
    );
    if (existing.length) return;
    const id = `usr_${Date.now()}`;
    const passwordHash = await ensurePasswordHash("12031991");
    await conn.query(
      `
        INSERT INTO usuarios
          (id, nombre, nombre_usuario, nombre_completo, correo, telefono, contrasena, firma_data_url, avatar_data_url, activo)
        VALUES (?, ?, ?, ?, ?, ?, ?, NULL, NULL, 1)
      `,
      [
        id,
        "Carlos Roberto Samalaj",
        username,
        "Carlos Roberto Samalaj",
        "sistemas@jardinesdellago.com",
        "56325547",
        passwordHash,
      ]
    );
  } finally {
    if (conn) conn.release();
  }
}

async function writeStateToTables(state) {
  let conn;
  try {
    conn = await pool.getConnection();
    await conn.beginTransaction();

    const salones = Array.isArray(state.salones) ? state.salones : [];
    const users = Array.isArray(state.users) ? state.users : [];
    const companies = Array.isArray(state.companies) ? state.companies : [];
    const services = Array.isArray(state.services) ? state.services : [];
    const quickTemplates = Array.isArray(state.quickTemplates) ? state.quickTemplates : [];
    const disabledCompanies = Array.isArray(state.disabledCompanies) ? state.disabledCompanies : [];
    const disabledServices = Array.isArray(state.disabledServices) ? state.disabledServices : [];
    const disabledManagers = Array.isArray(state.disabledManagers) ? state.disabledManagers : [];
    const disabledSalones = Array.isArray(state.disabledSalones) ? state.disabledSalones : [];
    const globalMonthlyGoals = Array.isArray(state.globalMonthlyGoals) ? state.globalMonthlyGoals : [];
    const checklistTemplateItems = Array.isArray(state.checklistTemplateItems) ? state.checklistTemplateItems : [];
    const checklistTemplateSections = Array.isArray(state.checklistTemplateSections) ? state.checklistTemplateSections : [];
    const menuMontajeSections = Array.isArray(state.menuMontajeSections) ? state.menuMontajeSections : [];
    const menuMontajeBebidas = Array.isArray(state.menuMontajeBebidas) ? state.menuMontajeBebidas : [];
    const eventChecklists = (state.eventChecklists && typeof state.eventChecklists === "object") ? state.eventChecklists : {};
    const events = Array.isArray(state.events)
      ? state.events.map((e) => ({ ...e, salon: e?.salon ?? "" }))
      : [];
    const changeHistory = (state.changeHistory && typeof state.changeHistory === "object") ? state.changeHistory : {};
    const reminders = (state.reminders && typeof state.reminders === "object") ? state.reminders : {};
    const usedQuoteCodes = new Set();
    let maxDesiredQuoteCodeNum = 0;

    for (const e of events) {
      const q = (e?.quote && typeof e.quote === "object") ? e.quote : null;
      if (!q) continue;
      const codeNum = parseQuoteCodeNumber(q.code);
      if (codeNum > maxDesiredQuoteCodeNum) maxDesiredQuoteCodeNum = codeNum;
    }
    if (maxDesiredQuoteCodeNum > 0) {
      await ensureDocSequenceAtLeast(conn, "COT", maxDesiredQuoteCodeNum);
    }

    await conn.query("DELETE FROM items_cotizacion_evento");
    await conn.query("DELETE FROM items_cotizacion_version_evento");
    await conn.query("DELETE FROM cotizacion_versiones_evento");
    await conn.query("DELETE FROM cotizaciones_evento");
    await conn.query("DELETE FROM recordatorios_evento");
    await conn.query("DELETE FROM historial_evento");
    await conn.query("DELETE FROM eventos");
    await conn.query("DELETE FROM encargados_empresa");
    await conn.query("DELETE FROM empresas");
    await conn.query("DELETE FROM servicios");
    await conn.query("DELETE FROM subcategorias_servicio");
    await conn.query("DELETE FROM categorias_servicio");
    await conn.query("DELETE FROM usuarios");
    await conn.query("DELETE FROM salones");

    for (const roomName of salones) {
      const nombre = str(roomName).trim();
      if (!nombre) continue;
      await conn.query("INSERT INTO salones (nombre) VALUES (?)", [nombre]);
    }

    for (const u of users) {
      const id = str(u?.id).trim();
      const nombre = str(u?.name || u?.fullName).trim();
      const username = str(u?.username).trim() || null;
      const fullName = str(u?.fullName || u?.name).trim() || null;
      const email = str(u?.email).trim() || null;
      const phone = str(u?.phone).trim() || null;
      const passwordRaw = str(u?.password || u?.passwordHash).trim() || null;
      const password = passwordRaw ? await ensurePasswordHash(passwordRaw) : null;
      const signatureDataUrl = str(u?.signatureDataUrl).trim() || null;
      const avatarDataUrl = str(u?.avatarDataUrl).trim() || null;
      const active = u?.active === false ? 0 : 1;
      const salesTargetEnabled = u?.salesTargetEnabled === true ? 1 : 0;
      const monthlyGoals = Array.isArray(u?.monthlyGoals)
        ? u.monthlyGoals
          .map((g) => ({ month: str(g?.month), amount: Math.max(0, Number(g?.amount || 0)) }))
          .filter((g) => /^\d{4}-\d{2}$/.test(g.month) && Number.isFinite(g.amount) && g.amount > 0)
        : [];
      if (!id || !nombre) continue;
      await conn.query(
        `
          INSERT INTO usuarios
            (id, nombre, nombre_usuario, nombre_completo, correo, telefono, contrasena, firma_data_url, avatar_data_url, activo, influye_meta_ventas, metas_mensuales_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
        [
          id,
          nombre,
          username,
          fullName,
          email,
          phone,
          password,
          signatureDataUrl,
          avatarDataUrl,
          active,
          salesTargetEnabled,
          JSON.stringify(monthlyGoals),
        ]
      );
    }

    for (const c of companies) {
      const id = str(c?.id).trim();
      if (!id) continue;
      await conn.query(
        `
          INSERT INTO empresas
            (id, nombre, encargado_principal, correo, nit, razon_social, tipo_evento, direccion, telefono, notas)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
        [
          id,
          str(c?.name).trim() || "Empresa",
          str(c?.owner).trim() || null,
          str(c?.email).trim() || null,
          str(c?.nit).trim() || null,
          str(c?.businessName).trim() || null,
          str(c?.eventType).trim() || null,
          str(c?.address).trim() || null,
          str(c?.phone).trim() || null,
          str(c?.notes).trim() || null,
        ]
      );

      const managers = Array.isArray(c?.managers) ? c.managers : [];
      for (const m of managers) {
        const managerId = str(m?.id).trim();
        const managerName = str(m?.name).trim();
        if (!managerId || !managerName) continue;
        await conn.query(
          `
            INSERT INTO encargados_empresa
              (id, id_empresa, nombre, telefono, correo, direccion)
            VALUES (?, ?, ?, ?, ?, ?)
          `,
          [
            managerId,
            id,
            managerName,
            str(m?.phone).trim() || null,
            str(m?.email).trim() || null,
            str(m?.address).trim() || null,
          ]
        );
      }
    }

    const categoryMap = new Map();
    for (const s of services) {
      const categoryName = str(s?.category || s?.categoria).trim();
      if (!categoryName) continue;
      if (categoryMap.has(categoryName.toLowerCase())) continue;
      const res = await conn.query(
        "INSERT INTO categorias_servicio (nombre) VALUES (?)",
        [categoryName]
      );
      categoryMap.set(categoryName.toLowerCase(), Number(res.insertId));
    }

    const subcategoryMap = new Map();
    for (const s of services) {
      const categoryName = str(s?.category || s?.categoria).trim();
      const subcategoryName = str(s?.subcategory || s?.subcategoria).trim();
      if (!categoryName || !subcategoryName) continue;
      const categoryId = categoryMap.get(categoryName.toLowerCase());
      if (!categoryId) continue;
      const key = `${categoryId}|${subcategoryName.toLowerCase()}`;
      if (subcategoryMap.has(key)) continue;
      const res = await conn.query(
        "INSERT INTO subcategorias_servicio (id_categoria, nombre) VALUES (?, ?)",
        [categoryId, subcategoryName]
      );
      subcategoryMap.set(key, Number(res.insertId));
    }

    for (const s of services) {
      const id = str(s?.id).trim();
      const nombre = str(s?.name).trim();
      if (!id || !nombre) continue;
      const categoryName = str(s?.category || s?.categoria).trim();
      const subcategoryName = str(s?.subcategory || s?.subcategoria).trim();
      const categoryId = categoryName ? (categoryMap.get(categoryName.toLowerCase()) || null) : null;
      const subcategoryId = (categoryId && subcategoryName)
        ? (subcategoryMap.get(`${categoryId}|${subcategoryName.toLowerCase()}`) || null)
        : null;
      await conn.query(
        "INSERT INTO servicios (id, nombre, precio, descripcion, id_categoria, id_subcategoria, modo_cantidad) VALUES (?, ?, ?, ?, ?, ?, ?)",
        [
          id,
          nombre,
          Number(s?.price || 0),
          str(s?.description).trim() || null,
          categoryId,
          subcategoryId,
          normalizeQuantityMode(s?.quantityMode || s?.modoCantidad || s?.indicadorCantidad),
        ]
      );
    }

    for (const e of events) {
      const id = str(e?.id).trim();
      if (!id) continue;
      await conn.query(
        `
          INSERT INTO eventos
            (id, id_grupo, nombre, nombre_salon, fecha_evento, hora_inicio, hora_fin, estado, id_usuario, pax, notas, cotizacion_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
        [
          id,
          str(e?.groupId).trim() || null,
          str(e?.name).trim() || "(sin nombre)",
          str(e?.salon).trim() || "(sin salon)",
          asDate(e?.date) || "1970-01-01",
          asTime(e?.startTime) || "00:00:00",
          asTime(e?.endTime) || "00:30:00",
          str(e?.status).trim() || "Lista de Espera",
          str(e?.userId).trim() || null,
          e?.pax === null || e?.pax === undefined || e?.pax === "" ? null : Math.max(0, Number(e.pax)),
          str(e?.notes).trim() || null,
          e?.quote ? JSON.stringify(e.quote) : null,
        ]
      );

      if (e?.quote && typeof e.quote === "object") {
        const q = e.quote;
        const desiredCodeNum = parseQuoteCodeNumber(q.code);
        const desiredCode = desiredCodeNum > 0 ? formatQuoteCode(desiredCodeNum) : "";
        if (desiredCode && !usedQuoteCodes.has(desiredCode)) {
          q.code = desiredCode;
          usedQuoteCodes.add(desiredCode);
        } else {
          let nextCode = "";
          for (let guard = 0; guard < 5000; guard++) {
            const candidate = await reserveNextDocCodeInConn(conn, "COT");
            if (!usedQuoteCodes.has(candidate)) {
              nextCode = candidate;
              break;
            }
          }
          if (!nextCode) {
            throw new Error("No se pudo reservar un codigo de cotizacion unico.");
          }
          q.code = nextCode;
          usedQuoteCodes.add(nextCode);
        }
        const quoteTotals = calculateQuoteTotals(q);
        const currentVersion = Math.max(1, Number(q.version || 1));
        await conn.query(
          `
            INSERT INTO cotizaciones_evento
              (id_evento, id_empresa, id_encargado, nombre_empresa, nombre_encargado, contacto, correo, facturar_a, direccion, tipo_evento, lugar, horario_texto, codigo, fecha_documento, telefono, nit, personas, fecha_evento, folio, fecha_fin, fecha_max_pago, tipo_pago, notas_internas, notas, version_actual, subtotal, descuento_tipo, descuento_valor, descuento_monto, total_neto, cotizado_en_iso, json_crudo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            id,
            str(q.companyId).trim() || null,
            str(q.managerId).trim() || null,
            str(q.companyName).trim() || null,
            str(q.managerName).trim() || null,
            str(q.contact).trim() || null,
            str(q.email).trim() || null,
            str(q.billTo).trim() || null,
            str(q.address).trim() || null,
            str(q.eventType).trim() || null,
            str(q.venue).trim() || null,
            str(q.schedule).trim() || null,
            str(q.code).trim() || null,
            asDate(q.docDate),
            str(q.phone).trim() || null,
            str(q.nit).trim() || null,
            q.people === null || q.people === undefined || q.people === "" ? null : Math.max(0, Number(q.people)),
            asDate(q.eventDate),
            str(q.folio).trim() || null,
            asDate(q.endDate),
            asDate(q.dueDate),
            str(q.paymentType).trim() || null,
            str(q.internalNotes).trim() || null,
            str(q.notes).trim() || null,
            currentVersion,
            Number(quoteTotals.subtotal || 0),
            normalizeDiscountType(q.discountType),
            Number(quoteTotals.discountValue || 0),
            Number(quoteTotals.discountAmount || 0),
            Number(quoteTotals.total || 0),
            str(q.quotedAt).trim() || null,
            JSON.stringify(q),
          ]
        );

        const items = Array.isArray(q.items) ? q.items : [];
        for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
          const item = items[itemIdx];
          const itemId = buildQuoteItemPrimaryKey(id, item, itemIdx);
          await conn.query(
            `
              INSERT INTO items_cotizacion_evento
                (id, id_evento, id_servicio, fecha_servicio, cantidad, precio, precio_unitario, modo_cantidad, total_linea, nombre, descripcion)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            `,
            [
              itemId,
              id,
              str(item?.serviceId).trim() || null,
              asDate(item?.serviceDate),
              Number(item?.qty || 0),
              Number(item?.price || 0),
              Number(item?.unitPrice || item?.price || 0),
              normalizeQuantityMode(item?.quantityMode),
              Number(item?.qty || 0) * Number(item?.price || 0),
              str(item?.name || item?.description).trim() || "(sin descripcion)",
              str(item?.description).trim() || null,
            ]
          );
        }

        const versionRows = [];
        const rawVersions = Array.isArray(q.versions) ? q.versions : [];
        for (const v of rawVersions) {
          if (!v || typeof v !== "object") continue;
          const versionNum = Math.max(1, Number(v.version || 0));
          versionRows.push({
            version: versionNum,
            snapshot: { ...v, version: versionNum, versions: [] },
          });
        }
        const currentVersionForRows = Math.max(1, Number(q.version || (versionRows.length + 1)));
        if (!versionRows.some((x) => Number(x.version) === currentVersionForRows)) {
          versionRows.push({
            version: currentVersionForRows,
            snapshot: { ...q, version: currentVersionForRows, versions: [] },
          });
        }
        versionRows.sort((a, b) => Number(a.version) - Number(b.version));

        for (const v of versionRows) {
          const vTotals = calculateQuoteTotals(v.snapshot);
          await conn.query(
            `
              INSERT INTO cotizacion_versiones_evento
                (id_evento, version_num, subtotal, descuento_tipo, descuento_valor, descuento_monto, total_neto, cotizado_en_iso, json_crudo)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            `,
            [
              id,
              Number(v.version),
              Number(vTotals.subtotal || 0),
              normalizeDiscountType(v.snapshot?.discountType),
              Number(vTotals.discountValue || 0),
              Number(vTotals.discountAmount || 0),
              Number(vTotals.total || 0),
              str(v.snapshot?.quotedAt).trim() || null,
              JSON.stringify(v.snapshot),
            ]
          );

          const itemsVersion = Array.isArray(v.snapshot?.items) ? v.snapshot.items : [];
          for (let idx = 0; idx < itemsVersion.length; idx++) {
            const item = itemsVersion[idx];
            await conn.query(
              `
                INSERT INTO items_cotizacion_version_evento
                  (id_evento, version_num, fila_num, id_servicio, fecha_servicio, cantidad, precio, precio_unitario, modo_cantidad, total_linea, nombre, descripcion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
              `,
              [
                id,
                Number(v.version),
                idx + 1,
                str(item?.serviceId).trim() || null,
                asDate(item?.serviceDate),
                Number(item?.qty || 0),
                Number(item?.price || 0),
                Number(item?.unitPrice || item?.price || 0),
                normalizeQuantityMode(item?.quantityMode),
                Number(item?.qty || 0) * Number(item?.price || 0),
                str(item?.name || item?.description).trim() || "(sin descripcion)",
                str(item?.description).trim() || null,
              ]
            );
          }
        }
      }
    }

    for (const [eventKey, rows] of Object.entries(changeHistory)) {
      if (!Array.isArray(rows)) continue;
      for (const row of rows) {
        await conn.query(
          `
            INSERT INTO historial_evento
              (clave_evento, cambiado_en_iso, cambiado_en, id_usuario_actor, nombre_actor, cambio_texto)
            VALUES (?, ?, ?, ?, ?, ?)
          `,
          [
            str(eventKey).trim(),
            str(row?.at).trim() || null,
            asDateTime(row?.at),
            str(row?.actorUserId).trim() || null,
            str(row?.actorName).trim() || null,
            str(row?.change).trim() || "",
          ]
        );
      }
    }

    for (const [eventKey, rows] of Object.entries(reminders)) {
      if (!Array.isArray(rows)) continue;
      for (const row of rows) {
        const reminderId = str(row?.id).trim() || `rem_${Math.random().toString(16).slice(2)}`;
        await conn.query(
          `
            INSERT INTO recordatorios_evento
              (id, clave_evento, fecha_recordatorio, hora_recordatorio, medio, notas, creado_en_iso, creado_en, id_usuario_creador)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            reminderId,
            str(eventKey).trim(),
            asDate(row?.date),
            asTime(row?.time),
            str(row?.channel).trim() || "(sin canal)",
            str(row?.notes).trim() || null,
            str(row?.createdAt).trim() || null,
            asDateTime(row?.createdAt),
            str(row?.createdByUserId).trim() || null,
          ]
        );
      }
    }

    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('quickTemplates', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(quickTemplates)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('disabledCompanies', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(disabledCompanies)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('disabledServices', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(disabledServices)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('disabledManagers', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(disabledManagers)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('disabledSalones', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(disabledSalones)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('globalMonthlyGoals', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [
        JSON.stringify(
          globalMonthlyGoals
            .map((g) => ({ month: str(g?.month), amount: Math.max(0, Number(g?.amount || 0)) }))
            .filter((g) => /^\d{4}-\d{2}$/.test(g.month))
        ),
      ]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('checklistTemplateItems', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(checklistTemplateItems)]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('checklistTemplateSections', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(checklistTemplateSections.map((s) => str(s)).filter(Boolean))]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('menuMontajeSections', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(menuMontajeSections.map((s) => str(s)).filter(Boolean))]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('menuMontajeBebidas', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [
        JSON.stringify(
          menuMontajeBebidas
            .map((b) => ({ id: str(b?.id), nombre: str(b?.nombre || b?.name), activo: b?.activo === false ? false : true }))
            .filter((b) => b.id && b.nombre)
        ),
      ]
    );
    await conn.query(
      `
        INSERT INTO app_state_kv (clave, valor_json)
        VALUES ('eventChecklists', ?)
        ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
      `,
      [JSON.stringify(eventChecklists)]
    );

    await conn.commit();
  } catch (error) {
    if (conn) await conn.rollback();
    throw error;
  } finally {
    if (conn) conn.release();
  }
}

app.get("/api/health", async (_req, res) => {
  try {
    let conn;
    try {
      conn = await pool.getConnection();
      await conn.query("SELECT 1");
    } finally {
      if (conn) conn.release();
    }
    return res.json({ ok: true, db: "connected" });
  } catch (error) {
    return res.status(500).json({ ok: false, db: "error", message: error.message });
  }
});

app.get("/api/state", async (_req, res) => {
  try {
    const result = await readStateFromTables();
    if (!result) {
      return res.status(404).json({ message: "No hay estado guardado en tablas relacionales todavia." });
    }
    return res.json(result);
  } catch (error) {
    return res.status(500).json({ message: "No se pudo leer el estado desde tablas.", detail: error.message });
  }
});

app.get("/api/salones", async (_req, res) => {
  try {
    const salones = await readSalonesFromTables();
    return res.json({ salones });
  } catch (error) {
    return res.status(500).json({ message: "No se pudieron leer los salones.", detail: error.message });
  }
});

app.get("/api/login-users", async (_req, res) => {
  try {
    const users = await readLoginUsers();
    return res.json({ users });
  } catch (error) {
    return res.status(500).json({ message: "No se pudo cargar usuarios para login.", detail: error.message });
  }
});

app.post("/api/login", async (req, res) => {
  const userId = str(req.body?.userId).trim();
  const password = str(req.body?.password);
  if (!userId || !password) {
    return res.status(400).json({ ok: false, message: "Credenciales incompletas." });
  }
  try {
    const user = await authenticateUser(userId, password);
    if (!user) {
      return res.status(401).json({ ok: false, message: "Usuario o contrasena incorrecta." });
    }
    return res.json({ ok: true, user });
  } catch (error) {
    return res.status(500).json({ ok: false, message: "Error validando credenciales.", detail: error.message });
  }
});

app.post("/api/doc-code-next", async (_req, res) => {
  try {
    const code = await reserveNextDocCode("COT");
    return res.json({ ok: true, code });
  } catch (error) {
    return res.status(500).json({ ok: false, message: "No se pudo generar codigo de documento.", detail: error.message });
  }
});

app.get("/api/categorias-servicio", async (_req, res) => {
  try {
    const categorias = await readCategoriasServicioFromTables();
    return res.json({ categorias });
  } catch (error) {
    return res.status(500).json({ message: "No se pudieron leer las categorias.", detail: error.message });
  }
});

app.get("/api/subcategorias-servicio", async (req, res) => {
  try {
    const categoriaIdRaw = req.query?.categoria_id;
    const categoriaId = categoriaIdRaw === undefined ? null : Number(categoriaIdRaw);
    const subcategorias = await readSubcategoriasServicioFromTables(
      Number.isFinite(categoriaId) ? categoriaId : null
    );
    return res.json({ subcategorias });
  } catch (error) {
    return res.status(500).json({ message: "No se pudieron leer las subcategorias.", detail: error.message });
  }
});

app.get("/api/menu-catalog/:kind", async (req, res) => {
  try {
    const kind = str(req.params?.kind).trim();
    if (kind === "preparacion") {
      const platoId = Number(req.query?.plato_id);
      const rows = await readPreparacionesByPlato(platoId);
      return res.json({ items: rows });
    }
    const rows = await readSimpleMenuCatalog(kind);
    return res.json({ items: rows });
  } catch (error) {
    return res.status(500).json({ message: "No se pudo leer el catalogo de Menu & Montaje.", detail: error.message });
  }
});

app.post("/api/menu-catalog/:kind", async (req, res) => {
  try {
    const kind = str(req.params?.kind).trim();
    const nombre = str(req.body?.nombre).trim();
    if (kind === "preparacion") {
      const idPlatoFuerte = Number(req.body?.id_plato_fuerte);
      await createPreparacionForPlato(idPlatoFuerte, nombre);
      return res.json({ ok: true });
    }
    await createSimpleMenuCatalog(kind, nombre, {
      tipo: req.body?.tipo,
      tipo_plato: req.body?.tipo_plato,
      es_sin_proteina: req.body?.es_sin_proteina,
    });
    return res.json({ ok: true });
  } catch (error) {
    return res.status(400).json({ message: "No se pudo guardar el catalogo de Menu & Montaje.", detail: error.message });
  }
});

app.put("/api/menu-catalog/:kind/:id", async (req, res) => {
  try {
    const kind = str(req.params?.kind).trim();
    const id = Number(req.params?.id);
    if (kind === "preparacion") {
      await updatePreparacionById(id, {
        nombre: req.body?.nombre,
        activo: req.body?.activo,
        id_plato_fuerte: req.body?.id_plato_fuerte,
      });
      return res.json({ ok: true });
    }
    await updateSimpleMenuCatalog(kind, id, {
      nombre: req.body?.nombre,
      activo: req.body?.activo,
      tipo: req.body?.tipo,
      tipo_plato: req.body?.tipo_plato,
      es_sin_proteina: req.body?.es_sin_proteina,
    });
    return res.json({ ok: true });
  } catch (error) {
    return res.status(400).json({ message: "No se pudo actualizar el catalogo de Menu & Montaje.", detail: error.message });
  }
});

app.get("/api/menu-suggestions", async (req, res) => {
  try {
    const idPlatoFuerte = Number(req.query?.plato_id);
    const idPreparacion = Number(req.query?.preparacion_id);
    const payload = await readMenuSuggestionLinks({ idPlatoFuerte, idPreparacion });
    return res.json(payload);
  } catch (error) {
    return res.status(400).json({ message: "No se pudieron leer las sugerencias de Menu.", detail: error.message });
  }
});

app.put("/api/menu-suggestions", async (req, res) => {
  try {
    await saveMenuSuggestionLinks({
      idPlatoFuerte: req.body?.id_plato_fuerte,
      idPreparacion: req.body?.id_preparacion,
      salsaIds: req.body?.salsaIds,
      postreIds: req.body?.postreIds,
      guarnicionIds: req.body?.guarnicionIds,
    });
    return res.json({ ok: true });
  } catch (error) {
    return res.status(400).json({ message: "No se pudieron guardar las sugerencias de Menu.", detail: error.message });
  }
});

app.put("/api/state", async (req, res) => {
  const nextState = req.body && req.body.state;
  if (!nextState || typeof nextState !== "object" || Array.isArray(nextState)) {
    return res.status(400).json({ message: "Body invalido. Use { state: {...} }." });
  }

  try {
    await writeStateToTables(nextState);
    return res.json({ ok: true });
  } catch (error) {
    console.error("Error PUT /api/state:", error);
    return res.status(500).json({ message: "No se pudo guardar en tablas reales.", detail: error.message });
  }
});

app.use(express.static(path.join(__dirname)));

app.get("*", (req, res) => {
  if (req.path.startsWith("/api/")) {
    return res.status(404).json({ message: "Ruta API no encontrada." });
  }
  return res.sendFile(path.join(__dirname, "index.html"));
});

async function start() {
  try {
    await ensureAppStateExtraStructure();
    await ensureServiceCatalogStructure();
    await ensureQuoteVersionStructure();
    await ensureMenuMontajeCatalogStructure();
    await ensureDocumentSequenceStructure();
    await ensureUsersExtendedStructure();
    await ensureRequiredTables();
    await ensureDefaultUserCarlos();
    app.listen(APP_PORT, () => {
      console.log(`CRM listo en http://localhost:${APP_PORT}`);
      console.log(`MariaDB -> ${DB_USER}@${DB_HOST}:${DB_PORT}/${DB_NAME}`);
      console.log("Persistencia activa en tablas relacionales.");
    });
  } catch (error) {
    console.error("Error al iniciar el servidor:", error.message);
    process.exit(1);
  }
}

start();


