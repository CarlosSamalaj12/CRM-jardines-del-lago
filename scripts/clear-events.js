require("dotenv").config();
const mariadb = require("mariadb");

const DB_HOST = process.env.DB_HOST || "127.0.0.1";
const DB_PORT = Number(process.env.DB_PORT || 3306);
const DB_USER = process.env.DB_USER || "root";
const DB_PASSWORD = process.env.DB_PASSWORD || "";
const DB_NAME = process.env.DB_NAME || "crm_jdl";

function parseJsonSafe(value, fallback) {
  try {
    const parsed = JSON.parse(String(value || ""));
    return parsed && typeof parsed === "object" ? parsed : fallback;
  } catch {
    return fallback;
  }
}

async function tableExists(conn, tableName) {
  const rows = await conn.query("SHOW TABLES LIKE ?", [String(tableName || "").trim()]);
  return Array.isArray(rows) && rows.length > 0;
}

async function main() {
  const force = process.argv.includes("--yes");
  if (!force) {
    console.error("Este script eliminara todos los eventos actuales y sus cotizaciones relacionadas.");
    console.error("Ejecute: node scripts/clear-events.js --yes");
    process.exit(1);
  }

  const conn = await mariadb.createConnection({
    host: DB_HOST,
    port: DB_PORT,
    user: DB_USER,
    password: DB_PASSWORD,
    database: DB_NAME,
    multipleStatements: false,
  });

  try {
    const hasEventos = await tableExists(conn, "eventos");
    const hasItemsVersion = await tableExists(conn, "items_cotizacion_version_evento");
    const hasItems = await tableExists(conn, "items_cotizacion_evento");
    const hasQuoteVersions = await tableExists(conn, "cotizacion_versiones_evento");
    const hasQuotes = await tableExists(conn, "cotizaciones_evento");
    const hasReminders = await tableExists(conn, "recordatorios_evento");
    const hasHistory = await tableExists(conn, "historial_evento");
    const hasAppStateKv = await tableExists(conn, "app_state_kv");
    const hasCrmAppState = await tableExists(conn, "crm_app_state");

    const eventRows = hasEventos ? await conn.query("SELECT id FROM eventos") : [];
    const eventCount = Array.isArray(eventRows) ? eventRows.length : 0;

    const stateRows = hasCrmAppState
      ? await conn.query("SELECT state_json FROM crm_app_state WHERE id = 1 LIMIT 1")
      : [];
    const state = Array.isArray(stateRows) && stateRows[0]
      ? parseJsonSafe(stateRows[0].state_json, {})
      : {};

    const nextState = {
      ...state,
      events: [],
      changeHistory: {},
      reminders: {},
      eventChecklists: {},
    };

    await conn.beginTransaction();

    if (hasItemsVersion) await conn.query("DELETE FROM items_cotizacion_version_evento");
    if (hasItems) await conn.query("DELETE FROM items_cotizacion_evento");
    if (hasQuoteVersions) await conn.query("DELETE FROM cotizacion_versiones_evento");
    if (hasQuotes) await conn.query("DELETE FROM cotizaciones_evento");
    if (hasReminders) await conn.query("DELETE FROM recordatorios_evento");
    if (hasHistory) await conn.query("DELETE FROM historial_evento");
    if (hasEventos) await conn.query("DELETE FROM eventos");

    if (hasAppStateKv) {
      await conn.query(
        `
          INSERT INTO app_state_kv (clave, valor_json)
          VALUES ('eventChecklists', ?)
          ON DUPLICATE KEY UPDATE valor_json = VALUES(valor_json)
        `,
        [JSON.stringify({})]
      );
    }

    if (hasCrmAppState) {
      await conn.query(
        `
          INSERT INTO crm_app_state (id, state_json)
          VALUES (1, ?)
          ON DUPLICATE KEY UPDATE state_json = VALUES(state_json), updated_at = CURRENT_TIMESTAMP
        `,
        [JSON.stringify(nextState)]
      );
    }

    await conn.commit();

    console.log(`Eventos eliminados: ${eventCount}`);
    console.log("Tambien se limpiaron cotizaciones, historial, recordatorios y checklists asociados.");
  } catch (error) {
    try {
      await conn.rollback();
    } catch {}
    console.error("No se pudieron limpiar los eventos:", error.message || error);
    process.exitCode = 1;
  } finally {
    await conn.end();
  }
}

main();

