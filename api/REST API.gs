
/***************************************************************
 * GOOGLE SHEETS SUPER REST API
 * ULTIMATE SINGLE-FILE EDITION
 *
 * Authentication: NONE
 * Required: Deployment ID only
 *
 * Read methods: GET
 * Write methods: POST
 *
 * Database: Google Sheets
 ***************************************************************/

// =============================================================
// 01. GLOBAL CONFIGURATION
// =============================================================

const SUPER_CONFIG = {
  SPREADSHEET_ID: "",

  DEFAULT_SHEET: "Users",

  ID_COLUMN: "ID",

  CREATED_COLUMN: "created_at",

  UPDATED_COLUMN: "updated_at",

  AUDIT_SHEET: "__API_AUDIT",

  BACKUP_PREFIX: "__BACKUP_",

  MAX_LIMIT: 1000,

  DEFAULT_LIMIT: 100,

  MAX_BULK_SIZE: 500,

  ENABLE_AUDIT: true,

  ENABLE_AUTO_ID: true,

  ENABLE_TIMESTAMPS: true,

  TIMEZONE: "Asia/Kolkata"
};


// =============================================================
// 02. HTTP ENTRY POINTS
// =============================================================

function doGet(e) {
  return superRouter("GET", e);
}

function doPost(e) {
  return superRouter("POST", e);
}


// =============================================================
// 03. MAIN API ROUTER
// =============================================================

function superRouter(method, e) {

  try {

    const params = e && e.parameter ? e.parameter : {};

    let body = {};

    if (e && e.postData && e.postData.contents) {

      body = JSON.parse(e.postData.contents);

      if (!body || typeof body !== "object" || Array.isArray(body)) {
        throw new Error("JSON body must be an object.");
      }

    }

    const req = Object.assign({}, params, body);

    const action = String(req.action || "info").toLowerCase();

    const sheet = req.sheet || SUPER_CONFIG.DEFAULT_SHEET;

    const readActions = [
      "info",
      "sheets",
      "schema",
      "list",
      "get",
      "search",
      "count",
      "stats",
      "distinct",
      "aggregate",
      "duplicates",
      "export",
      "history"
    ];

    const writeActions = [
      "create",
      "update",
      "upsert",
      "delete",
      "bulk_create",
      "bulk_update",
      "bulk_delete",
      "clear",
      "create_sheet",
      "rename_sheet",
      "delete_sheet",
      "add_column",
      "rename_column",
      "delete_column",
      "import",
      "backup",
      "restore_backup"
    ];

    if (readActions.includes(action) && method !== "GET") {
      throw new Error("Use GET for this operation.");
    }

    if (writeActions.includes(action) && method !== "POST") {
      throw new Error("Use POST for this operation.");
    }

    switch (action) {

      case "info":
        return apiInfo();

      case "sheets":
        return getSheets();

      case "schema":
        return getSchema(sheet);

      case "list":
        return listData(sheet, req);

      case "get":
        return getById(sheet, req.id);

      case "search":
        return searchData(sheet, req);

      case "count":
        return countData(sheet, req);

      case "stats":
        return statistics(sheet);

      case "distinct":
        return distinctData(sheet, req.column);

      case "aggregate":
        return aggregateData(sheet, req);

      case "duplicates":
        return findDuplicates(sheet, req.column);

      case "export":
        return exportData(sheet, req);

      case "history":
        return getHistory(req);

      case "create":
        return createData(sheet, req.data);

      case "update":
        return updateData(sheet, req.id, req.data);

      case "upsert":
        return upsertData(sheet, req.data);

      case "delete":
        return deleteData(sheet, req.id);

      case "bulk_create":
        return bulkCreate(sheet, req.data);

      case "bulk_update":
        return bulkUpdate(sheet, req.data);

      case "bulk_delete":
        return bulkDelete(sheet, req.ids);

      case "clear":
        return clearData(sheet);

      case "create_sheet":
        return createSheet(req.sheet_name, req.headers);

      case "rename_sheet":
        return renameSheet(sheet, req.new_name);

      case "delete_sheet":
        return deleteSheet(sheet);

      case "add_column":
        return addColumn(sheet, req.column, req.default_value);

      case "rename_column":
        return renameColumn(sheet, req.column, req.new_name);

      case "delete_column":
        return deleteColumn(sheet, req.column);

      case "import":
        return importData(sheet, req.data);

      case "backup":
        return backupSheet(sheet);

      case "restore_backup":
        return restoreBackup(sheet, req.backup_name);

      default:
        throw new Error("Unknown action: " + action);

    }

  } catch (err) {

    return response({
      success: false,
      error: err.message
    });

  }

}


// =============================================================
// 04. DATABASE CONNECTION
// =============================================================

function database() {

  if (SUPER_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SUPER_CONFIG.SPREADSHEET_ID);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Spreadsheet connection failed.");
  }

  return ss;

}


function getTable(name) {

  const sheet = database().getSheetByName(name);

  if (!sheet) {
    throw new Error("Table not found: " + name);
  }

  if (sheet.getLastColumn() < 1) {
    throw new Error("Table has no columns.");
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);

  if (headers.some(h => !h.trim())) {
    throw new Error("Empty column heading detected.");
  }

  if (new Set(headers).size !== headers.length) {
    throw new Error("Duplicate column headings detected.");
  }

  return {
    sheet: sheet,
    headers: headers
  };

}


function getRows(table) {

  const sheet = table.sheet;

  const count = sheet.getLastRow() - 1;

  if (count <= 0) return [];

  return sheet
    .getRange(2, 1, count, table.headers.length)
    .getValues()
    .map((row, index) => ({
      rowNumber: index + 2,
      values: row,
      data: rowObject(table.headers, row)
    }));

}


function rowObject(headers, row) {

  const result = {};

  headers.forEach((h, i) => {
    result[h] = row[i];
  });

  return result;

}


function idIndex(headers) {

  const index = headers.indexOf(SUPER_CONFIG.ID_COLUMN);

  if (index < 0) {
    throw new Error("Required ID column is missing.");
  }

  return index;

}


// =============================================================
// 05. LOCK AND AUDIT ENGINE
// =============================================================

function withLock(callback) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }

}


function audit(action, sheet, id, details) {

  if (!SUPER_CONFIG.ENABLE_AUDIT) return;

  const ss = database();

  let log = ss.getSheetByName(SUPER_CONFIG.AUDIT_SHEET);

  if (!log) {

    log = ss.insertSheet(SUPER_CONFIG.AUDIT_SHEET);

    log.appendRow([
      "Timestamp",
      "Action",
      "Sheet",
      "Record_ID",
      "Details"
    ]);

  }

  log.appendRow([
    new Date(),
    action,
    sheet,
    id || "",
    JSON.stringify(details || {})
  ]);

}


// =============================================================
// 06. API INFORMATION
// =============================================================

function apiInfo() {

  return response({
    success: true,
    name: "Google Sheets Super REST API",
    version: "3.0",
    authentication: "None",
    database: "Google Sheets",
    architecture: "Single File",
    methods: {
      GET: [
        "info",
        "sheets",
        "schema",
        "list",
        "get",
        "search",
        "count",
        "stats",
        "distinct",
        "aggregate",
        "duplicates",
        "export",
        "history"
      ],
      POST: [
        "create",
        "update",
        "upsert",
        "delete",
        "bulk_create",
        "bulk_update",
        "bulk_delete",
        "clear",
        "create_sheet",
        "rename_sheet",
        "delete_sheet",
        "add_column",
        "rename_column",
        "delete_column",
        "import",
        "backup",
        "restore_backup"
      ]
    }
  });

}


// =============================================================
// 07. SHEET MANAGEMENT
// =============================================================

function getSheets() {

  const sheets = database().getSheets().map(s => ({
    name: s.getName(),
    rows: Math.max(0, s.getLastRow() - 1),
    columns: s.getLastColumn()
  }));

  return response({
    success: true,
    total: sheets.length,
    sheets: sheets
  });

}


function getSchema(name) {

  const table = getTable(name);

  return response({
    success: true,
    sheet: name,
    columns: table.headers.map((name, index) => ({
      name: name,
      position: index + 1,
      is_id: name === SUPER_CONFIG.ID_COLUMN,
      is_timestamp: [
        SUPER_CONFIG.CREATED_COLUMN,
        SUPER_CONFIG.UPDATED_COLUMN
      ].includes(name)
    })),
    total_records: Math.max(0, table.sheet.getLastRow() - 1)
  });

}


function createSheet(name, headers) {

  if (!name || !Array.isArray(headers) || !headers.length) {
    throw new Error("Sheet name and headers are required.");
  }

  if (
    headers.some(h => typeof h !== "string" || !h.trim()) ||
    new Set(headers).size !== headers.length
  ) {
    throw new Error("Headers must be non-empty and unique.");
  }

  if (!headers.includes(SUPER_CONFIG.ID_COLUMN)) {
    throw new Error("An ID column is required.");
  }

  return withLock(() => {

    const ss = database();

    if (ss.getSheetByName(name)) {
      throw new Error("Sheet already exists.");
    }

    const sheet = ss.insertSheet(name);

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    audit("CREATE_SHEET", name, "", { headers: headers });

    return response({
      success: true,
      message: "Sheet created.",
      sheet: name,
      headers: headers
    });

  });

}


function renameSheet(oldName, newName) {

  if (!newName || database().getSheetByName(newName)) {
    throw new Error("Invalid or existing sheet name.");
  }

  return withLock(() => {

    getTable(oldName).sheet.setName(newName);

    audit("RENAME_SHEET", oldName, "", { new_name: newName });

    return response({
      success: true,
      old_name: oldName,
      new_name: newName
    });

  });

}


function deleteSheet(name) {

  return withLock(() => {

    const ss = database();

    const sheet = ss.getSheetByName(name);

    if (!sheet) throw new Error("Sheet not found.");

    if (ss.getSheets().length <= 1) {
      throw new Error("Cannot delete the last spreadsheet tab.");
    }

    if (name === SUPER_CONFIG.AUDIT_SHEET) {
      throw new Error("Cannot delete the audit sheet through this action.");
    }

    ss.deleteSheet(sheet);

    return response({
      success: true,
      message: "Sheet deleted.",
      sheet: name
    });

  });

}


// =============================================================
// 08. COLUMN MANAGEMENT
// =============================================================

function addColumn(name, column, defaultValue) {

  if (!column) throw new Error("Column name required.");

  return withLock(() => {

    const table = getTable(name);

    if (table.headers.includes(column)) {
      throw new Error("Column already exists.");
    }

    const sheet = table.sheet;

    const newIndex = sheet.getLastColumn() + 1;

    sheet.getRange(1, newIndex).setValue(column);

    const rows = sheet.getLastRow() - 1;

    if (rows > 0 && defaultValue !== undefined) {
      sheet.getRange(2, newIndex, rows, 1)
        .setValues(Array.from({ length: rows }, () => [defaultValue]));
    }

    audit("ADD_COLUMN", name, "", { column: column });

    return response({
      success: true,
      message: "Column added.",
      column: column
    });

  });

}


function renameColumn(name, oldName, newName) {

  if (!newName) throw new Error("New column name required.");

  return withLock(() => {

    const table = getTable(name);

    const index = table.headers.indexOf(oldName);

    if (index < 0) throw new Error("Column not found.");

    if (table.headers.includes(newName)) {
      throw new Error("New column already exists.");
    }

    table.sheet.getRange(1, index + 1).setValue(newName);

    audit("RENAME_COLUMN", name, "", {
      old_name: oldName,
      new_name: newName
    });

    return response({
      success: true,
      old_name: oldName,
      new_name: newName
    });

  });

}


function deleteColumn(name, column) {

  return withLock(() => {

    const table = getTable(name);

    if (column === SUPER_CONFIG.ID_COLUMN) {
      throw new Error("Cannot delete the ID column.");
    }

    const index = table.headers.indexOf(column);

    if (index < 0) throw new Error("Column not found.");

    table.sheet.deleteColumn(index + 1);

    audit("DELETE_COLUMN", name, "", { column: column });

    return response({
      success: true,
      deleted_column: column
    });

  });

}


// =============================================================
// 09. ADVANCED FILTER ENGINE
// =============================================================

function matchesCondition(actual, condition) {

  if (
    condition === null ||
    typeof condition !== "object" ||
    Array.isArray(condition) ||
    condition instanceof Date
  ) {
    return String(actual) === String(condition);
  }

  const op = String(condition.op || "eq").toLowerCase();

  const expected = condition.value;

  const a = actual;
  const b = expected;

  switch (op) {

    case "eq":
      return String(a) === String(b);

    case "ne":
      return String(a) !== String(b);

    case "contains":
      return String(a).toLowerCase().includes(String(b).toLowerCase());

    case "not_contains":
      return !String(a).toLowerCase().includes(String(b).toLowerCase());

    case "starts":
      return String(a).toLowerCase().startsWith(String(b).toLowerCase());

    case "ends":
      return String(a).toLowerCase().endsWith(String(b).toLowerCase());

    case "gt":
      return compareValues(a, b) > 0;

    case "gte":
      return compareValues(a, b) >= 0;

    case "lt":
      return compareValues(a, b) < 0;

    case "lte":
      return compareValues(a, b) <= 0;

    case "in":
      return Array.isArray(b) && b.map(String).includes(String(a));

    case "not_in":
      return Array.isArray(b) && !b.map(String).includes(String(a));

    case "between":
      return Array.isArray(b) &&
        compareValues(a, b[0]) >= 0 &&
        compareValues(a, b[1]) <= 0;

    case "is_empty":
      return a === "" || a === null;

    case "not_empty":
      return a !== "" && a !== null;

    case "regex":
      return new RegExp(String(b), "i").test(String(a));

    default:
      throw new Error("Unsupported filter operator: " + op);

  }

}


function compareValues(a, b) {

  const x = Number(a);
  const y = Number(b);

  if (a !== "" && b !== "" && Number.isFinite(x) && Number.isFinite(y)) {
    return x - y;
  }

  return String(a).localeCompare(String(b), undefined, {
    numeric: true,
    sensitivity: "base"
  });

}


function filterRecords(records, headers, req) {

  let result = records;

  const filters = req.filters || req.where || {};

  if (Array.isArray(filters)) {

    filters.forEach(filter => {

      if (!headers.includes(filter.column)) {
        throw new Error("Unknown filter column: " + filter.column);
      }

      result = result.filter(row =>
        matchesCondition(row.data[filter.column], filter)
      );

    });

  } else {

    Object.keys(filters).forEach(column => {

      if (!headers.includes(column)) {
        throw new Error("Unknown filter column: " + column);
      }

      result = result.filter(row =>
        matchesCondition(row.data[column], filters[column])
      );

    });

  }

  // Global search
  if (req.q !== undefined && req.q !== "") {

    const q = String(req.q).toLowerCase();

    result = result.filter(row =>
      headers.some(h =>
        String(row.data[h]).toLowerCase().includes(q)
      )
    );

  }

  // Multiple sorting rules
  if (req.sort) {

    const sorts = Array.isArray(req.sort)
      ? req.sort
      : [{ column: req.sort, order: req.order || "asc" }];

    sorts.forEach(sort => {

      if (!headers.includes(sort.column)) {
        throw new Error("Unknown sort column: " + sort.column);
      }

    });

    result.sort((a, b) => {

      for (const sort of sorts) {

        const comparison = compareValues(
          a.data[sort.column],
          b.data[sort.column]
        );

        if (comparison !== 0) {
          return String(sort.order).toLowerCase() === "desc"
            ? -comparison
            : comparison;
        }

      }

      return 0;

    });

  }

  return result;

}


// =============================================================
// 10. LIST / SEARCH / COUNT
// =============================================================

function listData(name, req) {

  const table = getTable(name);

  let records = filterRecords(
    getRows(table),
    table.headers,
    req
  );

  const total = records.length;

  const offset = Math.max(0, Number(req.offset) || 0);

  const limit = Math.min(
    SUPER_CONFIG.MAX_LIMIT,
    Math.max(1, Number(req.limit) || SUPER_CONFIG.DEFAULT_LIMIT)
  );

  if (req.fields) {

    const fields = String(req.fields).split(",");

    fields.forEach(f => {
      if (!table.headers.includes(f)) {
        throw new Error("Unknown selected field: " + f);
      }
    });

    records = records.slice(offset, offset + limit).map(row => {

      const selected = {};

      fields.forEach(f => selected[f] = row.data[f]);

      return selected;

    });

  } else {

    records = records
      .slice(offset, offset + limit)
      .map(row => row.data);

  }

  return response({
    success: true,
    sheet: name,
    total: total,
    offset: offset,
    limit: limit,
    has_more: offset + limit < total,
    data: records
  });

}


function searchData(name, req) {

  return listData(name, req);

}


function countData(name, req) {

  const table = getTable(name);

  const rows = filterRecords(
    getRows(table),
    table.headers,
    req
  );

  return response({
    success: true,
    sheet: name,
    count: rows.length
  });

}


// =============================================================
// 11. GET BY ID
// =============================================================

function getById(name, id) {

  if (id === undefined || id === "") {
    throw new Error("Record ID required.");
  }

  const table = getTable(name);

  const index = idIndex(table.headers);

  const rows = getRows(table);

  const record = rows.find(row =>
    String(row.values[index]) === String(id)
  );

  if (!record) throw new Error("Record not found: " + id);

  return response({
    success: true,
    sheet: name,
    data: record.data
  });

}


// =============================================================
// 12. RECORD DATA PREPARATION
// =============================================================

function prepareRecord(headers, data, oldRecord) {

  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("A valid record object is required.");
  }

  const result = {};

  const now = new Date();

  headers.forEach(header => {

    if (header === SUPER_CONFIG.ID_COLUMN) {

      result[header] = oldRecord
        ? oldRecord[header]
        : (
            data[header] ||
            (SUPER_CONFIG.ENABLE_AUTO_ID ? Utilities.getUuid() : "")
          );

    } else if (header === SUPER_CONFIG.CREATED_COLUMN) {

      result[header] = oldRecord
        ? oldRecord[header]
        : (data[header] || now);

    } else if (header === SUPER_CONFIG.UPDATED_COLUMN) {

      result[header] = SUPER_CONFIG.ENABLE_TIMESTAMPS
        ? now
        : (data[header] || "");

    } else {

      result[header] = Object.prototype.hasOwnProperty.call(data, header)
        ? data[header]
        : (oldRecord ? oldRecord[header] : "");

    }

  });

  return result;

}


// =============================================================
// 13. CREATE RECORD
// =============================================================

function createData(name, data) {

  return withLock(() => {

    const table = getTable(name);

    const headers = table.headers;

    const index = idIndex(headers);

    const record = prepareRecord(headers, data);

    const newId = String(record[SUPER_CONFIG.ID_COLUMN]);

    if (!newId) throw new Error("ID generation failed.");

    const rows = getRows(table);

    if (rows.some(row => String(row.values[index]) === newId)) {
      throw new Error("Duplicate ID: " + newId);
    }

    table.sheet.appendRow(headers.map(h => record[h]));

    audit("CREATE", name, newId, record);

    return response({
      success: true,
      message: "Record created.",
      data: record
    });

  });

}


// =============================================================
// 14. UPDATE RECORD
// =============================================================

function updateData(name, id, data) {

  if (id === undefined || id === "") {
    throw new Error("Record ID required.");
  }

  return withLock(() => {

    const table = getTable(name);

    const headers = table.headers;

    const index = idIndex(headers);

    const rows = getRows(table);

    const found = rows.find(row =>
      String(row.values[index]) === String(id)
    );

    if (!found) throw new Error("Record not found.");

    const updated = prepareRecord(headers, data, found.data);

    table.sheet
      .getRange(found.rowNumber, 1, 1, headers.length)
      .setValues([headers.map(h => updated[h])]);

    audit("UPDATE", name, id, data);

    return response({
      success: true,
      message: "Record updated.",
      data: updated
    });

  });

}


// =============================================================
// 15. UPSERT RECORD
// =============================================================

function upsertData(name, data) {

  if (!data || data[SUPER_CONFIG.ID_COLUMN] === undefined) {
    throw new Error("Upsert requires an ID.");
  }

  const id = data[SUPER_CONFIG.ID_COLUMN];

  try {

    const table = getTable(name);

    const index = idIndex(table.headers);

    const found = getRows(table).some(row =>
      String(row.values[index]) === String(id)
    );

    if (found) {
      return updateData(name, id, data);
    }

  } catch (err) {
    throw err;
  }

  return createData(name, data);

}


// =============================================================
// 16. DELETE RECORD
// =============================================================

function deleteData(name, id) {

  if (id === undefined || id === "") {
    throw new Error("Record ID required.");
  }

  return withLock(() => {

    const table = getTable(name);

    const index = idIndex(table.headers);

    const found = getRows(table).find(row =>
      String(row.values[index]) === String(id)
    );

    if (!found) throw new Error("Record not found.");

    table.sheet.deleteRow(found.rowNumber);

    audit("DELETE", name, id, {});

    return response({
      success: true,
      message: "Record deleted.",
      deleted_id: id
    });

  });

}


// =============================================================
// 17. BULK CREATE
// =============================================================

function bulkCreate(name, data) {

  if (!Array.isArray(data) || !data.length) {
    throw new Error("Provide an array of records.");
  }

  if (data.length > SUPER_CONFIG.MAX_BULK_SIZE) {
    throw new Error("Bulk size limit exceeded.");
  }

  return withLock(() => {

    const table = getTable(name);

    const headers = table.headers;

    const index = idIndex(headers);

    const existing = new Set(
      getRows(table).map(row => String(row.values[index]))
    );

    const newIds = new Set();

    const rows = data.map(item => {

      const record = prepareRecord(headers, item);

      const id = String(record[SUPER_CONFIG.ID_COLUMN]);

      if (!id || existing.has(id) || newIds.has(id)) {
        throw new Error("Duplicate or invalid ID: " + id);
      }

      newIds.add(id);

      return headers.map(h => record[h]);

    });

    table.sheet.getRange(
      table.sheet.getLastRow() + 1,
      1,
      rows.length,
      headers.length
    ).setValues(rows);

    audit("BULK_CREATE", name, "", { count: rows.length });

    return response({
      success: true,
      message: "Bulk creation completed.",
      total: rows.length,
      ids: Array.from(newIds)
    });

  });

}


// =============================================================
// 18. BULK UPDATE
// =============================================================

function bulkUpdate(name, data) {

  if (!Array.isArray(data) || !data.length) {
    throw new Error("Provide update objects.");
  }

  if (data.length > SUPER_CONFIG.MAX_BULK_SIZE) {
    throw new Error("Bulk size limit exceeded.");
  }

  return withLock(() => {

    const table = getTable(name);

    const headers = table.headers;

    const index = idIndex(headers);

    const rows = getRows(table);

    const updates = new Map();

    data.forEach(item => {

      if (!item || item.id === undefined || !item.data) {
        throw new Error("Each update requires id and data.");
      }

      const found = rows.find(row =>
        String(row.values[index]) === String(item.id)
      );

      if (!found) throw new Error("Record not found: " + item.id);

      if (updates.has(found.rowNumber)) {
        throw new Error("Duplicate update ID: " + item.id);
      }

      updates.set(
        found.rowNumber,
        prepareRecord(headers, item.data, found.data)
      );

    });

    updates.forEach((record, rowNumber) => {

      table.sheet.getRange(
        rowNumber,
        1,
        1,
        headers.length
      ).setValues([headers.map(h => record[h])]);

    });

    audit("BULK_UPDATE", name, "", { count: updates.size });

    return response({
      success: true,
      updated_count: updates.size
    });

  });

}


// =============================================================
// 19. BULK DELETE
// =============================================================

function bulkDelete(name, ids) {

  if (!Array.isArray(ids) || !ids.length) {
    throw new Error("Provide an array of IDs.");
  }

  if (ids.length > SUPER_CONFIG.MAX_BULK_SIZE) {
    throw new Error("Bulk size limit exceeded.");
  }

  return withLock(() => {

    const table = getTable(name);

    const index = idIndex(table.headers);

    const wanted = new Set(ids.map(String));

    const rows = getRows(table);

    const found = rows
      .filter(row => wanted.has(String(row.values[index])))
      .map(row => row.rowNumber)
      .sort((a, b) => b - a);

    found.forEach(rowNumber => {
      table.sheet.deleteRow(rowNumber);
    });

    audit("BULK_DELETE", name, "", {
      deleted_count: found.length
    });

    return response({
      success: true,
      deleted_count: found.length,
      requested_count: ids.length
    });

  });

}


// =============================================================
// 20. CLEAR DATA
// =============================================================

function clearData(name) {

  return withLock(() => {

    const table = getTable(name);

    const sheet = table.sheet;

    const rows = sheet.getLastRow() - 1;

    if (rows > 0) {
      sheet.getRange(
        2,
        1,
        rows,
        table.headers.length
      ).clearContent();
    }

    audit("CLEAR", name, "", { rows: rows });

    return response({
      success: true,
      message: "Data cleared. Headers preserved.",
      cleared_rows: rows
    });

  });

}


// =============================================================
// 21. DISTINCT VALUES
// =============================================================

function distinctData(name, column) {

  if (!column) throw new Error("Column required.");

  const table = getTable(name);

  const index = table.headers.indexOf(column);

  if (index < 0) throw new Error("Column not found.");

  const values = getRows(table).map(row => row.values[index]);

  const unique = Array.from(
    new Map(values.map(v => [JSON.stringify(v), v])).values()
  );

  return response({
    success: true,
    column: column,
    total: unique.length,
    data: unique
  });

}


// =============================================================
// 22. STATISTICS
// =============================================================

function statistics(name) {

  const table = getTable(name);

  const rows = getRows(table);

  const result = {};

  table.headers.forEach((header, index) => {

    const values = rows.map(row => row.values[index]);

    const numbers = values
      .filter(v => v !== "" && v !== null)
      .map(Number)
      .filter(Number.isFinite);

    result[header] = {
      non_empty: values.filter(v => v !== "" && v !== null).length,
      empty: values.filter(v => v === "" || v === null).length,
      unique: new Set(values.map(String)).size,
      numeric_count: numbers.length,
      sum: numbers.reduce((a, b) => a + b, 0),
      average: numbers.length
        ? numbers.reduce((a, b) => a + b, 0) / numbers.length
        : null,
      min: numbers.length ? Math.min(...numbers) : null,
      max: numbers.length ? Math.max(...numbers) : null
    };

  });

  return response({
    success: true,
    sheet: name,
    records: rows.length,
    columns: table.headers.length,
    statistics: result
  });

}


// =============================================================
// 23. AGGREGATION ENGINE
// =============================================================

function aggregateData(name, req) {

  const table = getTable(name);

  const column = req.column;

  if (!table.headers.includes(column)) {
    throw new Error("Aggregation column not found.");
  }

  const values = filterRecords(
    getRows(table),
    table.headers,
    req
  ).map(row => Number(row.data[column]))
   .filter(Number.isFinite);

  const operation = String(req.operation || "sum").toLowerCase();

  let result;

  switch (operation) {

    case "sum":
      result = values.reduce((a, b) => a + b, 0);
      break;

    case "average":
    case "avg":
      result = values.length
        ? values.reduce((a, b) => a + b, 0) / values.length
        : null;
      break;

    case "min":
      result = values.length ? Math.min(...values) : null;
      break;

    case "max":
      result = values.length ? Math.max(...values) : null;
      break;

    case "count":
      result = values.length;
      break;

    case "median": {
      const sorted = values.slice().sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);

      result = !sorted.length
        ? null
        : sorted.length % 2
          ? sorted[mid]
          : (sorted[mid - 1] + sorted[mid]) / 2;

      break;
    }

    default:
      throw new Error("Unsupported aggregation: " + operation);

  }

  return response({
    success: true,
    column: column,
    operation: operation,
    count: values.length,
    result: result
  });

}


// =============================================================
// 24. DUPLICATE DETECTION
// =============================================================

function findDuplicates(name, column) {

  const table = getTable(name);

  if (!table.headers.includes(column)) {
    throw new Error("Column not found.");
  }

  const groups = {};

  getRows(table).forEach(row => {

    const key = String(row.data[column]);

    if (!groups[key]) groups[key] = [];

    groups[key].push(row.data);

  });

  const duplicates = Object.keys(groups)
    .filter(key => groups[key].length > 1)
    .map(key => ({
      value: key,
      count: groups[key].length,
      records: groups[key]
    }));

  return response({
    success: true,
    column: column,
    duplicate_groups: duplicates.length,
    data: duplicates
  });

}


// =============================================================
// 25. IMPORT DATA
// =============================================================

function importData(name, data) {

  if (!Array.isArray(data) || !data.length) {
    throw new Error("Import requires a non-empty data array.");
  }

  return bulkCreate(name, data);

}


// =============================================================
// 26. EXPORT DATA
// =============================================================

function exportData(name, req) {

  const table = getTable(name);

  const records = filterRecords(
    getRows(table),
    table.headers,
    req
  );

  const format = String(req.format || "json").toLowerCase();

  if (format === "json") {

    return response({
      success: true,
      format: "json",
      total: records.length,
      data: records.map(row => row.data)
    });

  }

  if (format === "csv") {

    const escapeCSV = value => {

      let text = value instanceof Date
        ? value.toISOString()
        : String(value === null || value === undefined ? "" : value);

      return '"' + text.replace(/"/g, '""') + '"';

    };

    const csv = [
      table.headers.map(escapeCSV).join(",")
    ].concat(
      records.map(row =>
        row.values.map(escapeCSV).join(",")
      )
    ).join("\r\n");

    return response({
      success: true,
      format: "csv",
      total: records.length,
      csv: csv
    });

  }

  throw new Error("Supported formats: json, csv.");

}


// =============================================================
// 27. BACKUP ENGINE
// =============================================================

function backupSheet(name) {

  return withLock(() => {

    const ss = database();

    const source = getTable(name).sheet;

    const timestamp = Utilities.formatDate(
      new Date(),
      SUPER_CONFIG.TIMEZONE,
      "yyyyMMdd_HHmmss"
    );

    const backupName = (
      SUPER_CONFIG.BACKUP_PREFIX + name + "_" + timestamp
    ).substring(0, 99);

    if (ss.getSheetByName(backupName)) {
      throw new Error("Backup name already exists.");
    }

    const backup = source.copyTo(ss);

    backup.setName(backupName);

    audit("BACKUP", name, "", {
      backup_name: backupName
    });

    return response({
      success: true,
      message: "Backup created.",
      original_sheet: name,
      backup_name: backupName
    });

  });

}


// =============================================================
// 28. RESTORE BACKUP
// =============================================================

function restoreBackup(name, backupName) {

  if (!backupName ||
      !backupName.startsWith(SUPER_CONFIG.BACKUP_PREFIX)) {
    throw new Error("A valid backup sheet name is required.");
  }

  return withLock(() => {

    const ss = database();

    const backup = ss.getSheetByName(backupName);

    const target = ss.getSheetByName(name);

    if (!backup || !target) {
      throw new Error("Target or backup sheet not found.");
    }

    const temporary = backup.copyTo(ss);

    temporary.setName(
      ("__RESTORE_" + Date.now()).substring(0, 99)
    );

    ss.deleteSheet(target);

    temporary.setName(name);

    audit("RESTORE_BACKUP", name, "", {
      backup_name: backupName
    });

    return response({
      success: true,
      message: "Backup restored.",
      sheet: name,
      backup_name: backupName
    });

  });

}


// =============================================================
// 29. AUDIT HISTORY
// =============================================================

function getHistory(req) {

  const log = database().getSheetByName(
    SUPER_CONFIG.AUDIT_SHEET
  );

  if (!log || log.getLastRow() < 2) {
    return response({
      success: true,
      total: 0,
      data: []
    });
  }

  const values = log.getRange(
    2,
    1,
    log.getLastRow() - 1,
    5
  ).getValues();

  let records = values.map(row => ({
    timestamp: row[0],
    action: row[1],
    sheet: row[2],
    id: row[3],
    details: row[4]
  }));

  if (req.sheet) {
    records = records.filter(r => r.sheet === req.sheet);
  }

  if (req.action_filter) {
    records = records.filter(r =>
      String(r.action).toLowerCase() ===
      String(req.action_filter).toLowerCase()
    );
  }

  records.reverse();

  const total = records.length;

  const offset = Math.max(0, Number(req.offset) || 0);

  const limit = Math.min(
    SUPER_CONFIG.MAX_LIMIT,
    Math.max(1, Number(req.limit) || 100)
  );

  return response({
    success: true,
    total: total,
    offset: offset,
    limit: limit,
    data: records.slice(offset, offset + limit)
  });

}


// =============================================================
// 30. JSON RESPONSE ENGINE
// =============================================================

function response(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

}

