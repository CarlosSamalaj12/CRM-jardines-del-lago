require("dotenv").config();
const fs = require("fs");
const path = require("path");
const mariadb = require("mariadb");

const DB_HOST = process.env.DB_HOST || "127.0.0.1";
const DB_PORT = Number(process.env.DB_PORT || 3306);
const DB_USER = process.env.DB_USER || "root";
const DB_PASSWORD = process.env.DB_PASSWORD || "";
const DB_NAME = process.env.DB_NAME || "crm_jdl";

function uid(prefix = "ev") {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function todayIso() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function buildDefaultState() {
  const users = [
    { id: uid("usr"), name: "Carlos" },
    { id: uid("usr"), name: "Ana" },
    { id: uid("usr"), name: "Recepcion" },
    { id: uid("usr"), name: "Ventas" },
  ];
  return {
    salones: [
      "Arenal",
      "Santa Isabel",
      "Rancho Grande",
      "Jardincito",
      "Casa Flores",
      "Hotel Principal",
    ],
    users,
    companies: [
      {
        id: uid("cmp"),
        name: "Cliente General",
        owner: "Encargado General",
        email: "cliente@general.com",
        nit: "CF",
        businessName: "Cliente General",
        eventType: "Social",
        address: "",
        phone: "",
        notes: "",
        managers: [
          {
            id: uid("mgr"),
            name: "Encargado General",
            phone: "",
            email: "cliente@general.com",
            address: "",
          },
        ],
      },
    ],
    services: [
      { id: uid("svc"), name: "Alquiler de salon", price: 2500, description: "Uso del salon por jornada" },
      { id: uid("svc"), name: "Decoracion basica", price: 1200, description: "Montaje y decoracion estandar" },
      { id: uid("svc"), name: "Catering por persona", price: 85, description: "Menu base por invitado" },
      { id: uid("svc"), name: "Sonido e iluminacion", price: 1800, description: "Equipo profesional y operador" },
      { id: uid("svc"), name: "Fotografia", price: 2200, description: "Cobertura del evento" },
    ],
    changeHistory: {},
    reminders: {},
    events: [
      {
        id: uid("evt"),
        name: "Evento demo - Boda",
        salon: "Arenal",
        date: todayIso(),
        status: "Confirmado",
        startTime: "16:00",
        endTime: "18:00",
        userId: users[0].id,
        notes: "Ejemplo",
      },
    ],
  };
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

function asNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function str(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

async function runSchema(conn) {
  const schemaPath = path.join(__dirname, "..", "db", "schema.sql");
  const sql = fs.readFileSync(schemaPath, "utf8");
  const statements = sql
    .split(";")
    .map((x) => x.trim())
    .filter(Boolean);

  for (const stmt of statements) {
    await conn.query(stmt);
  }
}

async function clearRelationalTables(conn) {
  await conn.query("DELETE FROM items_cotizacion_evento");
  await conn.query("DELETE FROM cotizaciones_evento");
  await conn.query("DELETE FROM recordatorios_evento");
  await conn.query("DELETE FROM historial_evento");
  await conn.query("DELETE FROM eventos");
  await conn.query("DELETE FROM encargados_empresa");
  await conn.query("DELETE FROM empresas");
  await conn.query("DELETE FROM servicios");
  await conn.query("DELETE FROM usuarios");
  await conn.query("DELETE FROM salones");
}

async function readAppState(conn) {
  const rows = await conn.query(
    "SELECT state_json FROM crm_app_state WHERE id = 1 LIMIT 1"
  );
  if (!rows.length) return null;
  const parsed = JSON.parse(rows[0].state_json);
  if (!parsed || typeof parsed !== "object") return null;
  return parsed;
}

async function ensureStateRow(conn) {
  const existing = await readAppState(conn);
  if (existing) return existing;

  await conn.query(
    `
      CREATE TABLE IF NOT EXISTS crm_app_state (
        id TINYINT UNSIGNED PRIMARY KEY,
        state_json LONGTEXT NOT NULL,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      )
    `
  );
  const initialState = buildDefaultState();
  await conn.query(
    `
      INSERT INTO crm_app_state (id, state_json)
      VALUES (1, ?)
      ON DUPLICATE KEY UPDATE state_json = VALUES(state_json), updated_at = CURRENT_TIMESTAMP
    `,
    [JSON.stringify(initialState)]
  );
  return initialState;
}

async function migrateState(conn, state) {
  const salones = Array.isArray(state.salones) ? state.salones : [];
  const users = Array.isArray(state.users) ? state.users : [];
  const companies = Array.isArray(state.companies) ? state.companies : [];
  const services = Array.isArray(state.services) ? state.services : [];
  const events = Array.isArray(state.events) ? state.events : [];
  const changeHistory =
    state.changeHistory && typeof state.changeHistory === "object"
      ? state.changeHistory
      : {};
  const reminders =
    state.reminders && typeof state.reminders === "object" ? state.reminders : {};

  for (const salonNombre of salones) {
    const name = str(salonNombre).trim();
    if (!name) continue;
    await conn.query("INSERT INTO salones (nombre) VALUES (?)", [name]);
  }

  for (const user of users) {
    const id = str(user?.id).trim();
    const name = str(user?.name).trim();
    if (!id || !name) continue;
    await conn.query("INSERT INTO usuarios (id, nombre) VALUES (?, ?)", [id, name]);
  }

  for (const company of companies) {
    const companyId = str(company?.id).trim();
    if (!companyId) continue;
    await conn.query(
      `
        INSERT INTO empresas
          (id, nombre, encargado_principal, correo, nit, razon_social, tipo_evento, direccion, telefono, notas)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
      [
        companyId,
        str(company?.name || "Empresa").trim(),
        str(company?.owner).trim() || null,
        str(company?.email).trim() || null,
        str(company?.nit).trim() || null,
        str(company?.businessName).trim() || null,
        str(company?.eventType).trim() || null,
        str(company?.address).trim() || null,
        str(company?.phone).trim() || null,
        str(company?.notes).trim() || null,
      ]
    );

    const managers = Array.isArray(company?.managers) ? company.managers : [];
    for (const manager of managers) {
      const managerId = str(manager?.id).trim();
      const name = str(manager?.name).trim();
      if (!managerId || !name) continue;
      await conn.query(
        `
          INSERT INTO encargados_empresa
            (id, id_empresa, nombre, telefono, correo, direccion)
          VALUES (?, ?, ?, ?, ?, ?)
        `,
        [
          managerId,
          companyId,
          name,
          str(manager?.phone).trim() || null,
          str(manager?.email).trim() || null,
          str(manager?.address).trim() || null,
        ]
      );
    }
  }

  for (const service of services) {
    const serviceId = str(service?.id).trim();
    const name = str(service?.name).trim();
    if (!serviceId || !name) continue;
    await conn.query(
      "INSERT INTO servicios (id, nombre, precio, descripcion) VALUES (?, ?, ?, ?)",
      [serviceId, name, asNumber(service?.price), str(service?.description).trim() || null]
    );
  }

  for (const ev of events) {
    const id = str(ev?.id).trim();
    if (!id) continue;
    const eventDate = asDate(ev?.date) || "1970-01-01";
    const startTime = asTime(ev?.startTime) || "00:00:00";
    const endTime = asTime(ev?.endTime) || "00:30:00";
    const userId = str(ev?.userId).trim();

    await conn.query(
      `
        INSERT INTO eventos
          (id, id_grupo, nombre, nombre_salon, fecha_evento, hora_inicio, hora_fin, estado, id_usuario, pax, notas, cotizacion_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
      [
        id,
        str(ev?.groupId).trim() || null,
        str(ev?.name).trim() || "(sin nombre)",
        str(ev?.salon).trim() || "(sin salon)",
        eventDate,
        startTime,
        endTime,
        str(ev?.status).trim() || "Lista de Espera",
        userId || null,
        ev?.pax === null || ev?.pax === undefined || ev?.pax === "" ? null : Math.max(0, Number(ev.pax)),
        str(ev?.notes).trim() || null,
        ev?.quote ? JSON.stringify(ev.quote) : null,
      ]
    );

    if (ev?.quote && typeof ev.quote === "object") {
      const q = ev.quote;
      await conn.query(
        `
          INSERT INTO cotizaciones_evento
            (id_evento, id_empresa, id_encargado, nombre_empresa, nombre_encargado, contacto, correo, facturar_a, direccion, tipo_evento, lugar, horario_texto, codigo, fecha_documento, telefono, nit, personas, fecha_evento, folio, fecha_fin, fecha_max_pago, tipo_pago, notas_internas, notas, cotizado_en_iso, json_crudo)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
          str(q.quotedAt).trim() || null,
          JSON.stringify(q),
        ]
      );

      const items = Array.isArray(q.items) ? q.items : [];
      for (const item of items) {
        const itemId = str(item?.rowId).trim() || `item_${id}_${Math.random().toString(16).slice(2)}`;
        await conn.query(
          `
            INSERT INTO items_cotizacion_evento
              (id, id_evento, id_servicio, fecha_servicio, cantidad, precio, nombre, descripcion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            itemId,
            id,
            str(item?.serviceId).trim() || null,
            asDate(item?.serviceDate),
            asNumber(item?.qty),
            asNumber(item?.price),
            str(item?.name || item?.description).trim() || "(sin descripcion)",
            str(item?.description).trim() || null,
          ]
        );
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
    "INSERT INTO bitacora_migracion (origen, detalle) VALUES (?, ?)",
    [
      "crm_app_state",
      JSON.stringify({
        salones: salones.length,
        users: users.length,
        companies: companies.length,
        services: services.length,
        events: events.length,
        historyKeys: Object.keys(changeHistory).length,
        reminderKeys: Object.keys(reminders).length,
      }),
    ]
  );
}

async function main() {
  const conn = await mariadb.createConnection({
    host: DB_HOST,
    port: DB_PORT,
    user: DB_USER,
    password: DB_PASSWORD,
    database: DB_NAME,
    multipleStatements: false,
  });

  try {
    await conn.beginTransaction();
    await runSchema(conn);
    const state = await ensureStateRow(conn);
    await clearRelationalTables(conn);
    await migrateState(conn, state);
    await conn.commit();

    const counts = {};
    for (const tableName of [
      "salones",
      "usuarios",
      "empresas",
      "encargados_empresa",
      "servicios",
      "eventos",
      "cotizaciones_evento",
      "items_cotizacion_evento",
      "historial_evento",
      "recordatorios_evento",
    ]) {
      const rows = await conn.query(`SELECT COUNT(*) AS total FROM ${tableName}`);
      counts[tableName] = rows[0].total;
    }
    console.log("Migracion completada. Conteos:", counts);
  } catch (error) {
    await conn.rollback();
    console.error("Error en migracion:", error.message);
    process.exitCode = 1;
  } finally {
    await conn.end();
  }
}

main();

