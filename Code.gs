/**
 * Creating a Google Sheets Data Entry Form for CRUD Operations
 * By: bpwebs.com
 * Post URL: https://www.bpwebs.com/crud-operations-on-google-sheets-with-online-forms
 */

//CONSTANTS
const SPREADSHEETID = "18TZDPPICngT-It2DW_F4brvvmiIVkpjwK7gO7L01_lU";
const DATARANGE = "Data!A2:P";
const DATASHEET = "Data";
const DATASHEETID = "0";
const LASTCOL = "P";
const IDRANGE = "Data!A2:A";
const DROPDOWNRANGE1 = "Helpers!A1:A9"; //GRADOS OFIC LIST
//const DROPDOWNRANGE3 = "Helpers!A10:A17"; //GRADOS VOL LIST
const DROPDOWNRANGE2 = "Helpers!C1:C335"; //UNIDADES LIST
const USERS_SHEET = "Users";


//Display HTML page
function doGet(e) {
  const userEmail = e.parameter.email; // o cualquier otro parámetro que uses para la sesión
  if (!userEmail) { // Si no hay sesión, redirige al login
    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('Login');
  }
  // Si hay sesión, muestra la página principal
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const file = (page === 'index') ? 'Index' : 'Login';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(page === 'index' ? 'Sistema PMP' : 'Login');
}




//PROCESS SUBMITTED FORM DATA
function processForm(formObject) {
  if (formObject.recId && checkId(formObject.recId)) {
    const values = [[
      formObject.recId,
      formObject.proceso,
      formObject.tipoPMP,
      formObject.gradoMilitar,
      formObject.funcion,
      formObject.unidad,
      formObject.documento,
      formObject.sexo,
      formObject.fecha_doc,
      formObject.autoridad,
      formObject.nombre_req,
      formObject.clasificacion,
      formObject.cantidadPMP,
      formObject.estado,
      formObject.observ_justif,
      new Date().toLocaleString()
    ]];
    const updateRange = getRangeById(formObject.recId);
    //Update the record
    updateRecord(values, updateRange);
  } else {
    //Prepare new row of data
    let values = [[
      generateUniqueId(),
      formObject.proceso,
      formObject.tipoPMP,
      formObject.gradoMilitar,
      formObject.funcion,
      formObject.unidad,
      formObject.documento,
      formObject.sexo,
      formObject.fecha_doc,
      formObject.autoridad,
      formObject.nombre_req,
      formObject.clasificacion,
      formObject.cantidadPMP,
      formObject.estado,
      formObject.observ_justif,
      new Date().toLocaleString()
    ]];

    //Create new record
    createRecord(values);
  }

  //Return the last 10 records
  return getLastTenRecords();
}

//Login
// === HELPERS ===
function _usersSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEETID);
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.appendRow(["email","name","password_hash","salt","role","enabled","created_at","last_login"]);
  }
  return sh;
}
function _findUserRow(email) {
  email = (email || "").trim().toLowerCase();
  if (!email) return null;

  const sh = _usersSheet();
  const last = sh.getLastRow();
  if (last < 2) return null;

  const vals = sh.getRange(2, 1, last - 1, 8).getValues(); // A:H
  for (let i = 0; i < vals.length; i++) {
    const rowEmail = (vals[i][0] || "").toString().toLowerCase();
    if (rowEmail === email) {
      return { row: i + 2, data: vals[i] };
    }
  }
  return null;
}

function loginUser(email, password) {
  email = (email || "").trim();
  if (!email || !password) return { ok:false, msg:"Credenciales incompletas" };

  const hit = _findUserRow(email);
  if (!hit) return { ok:false, msg:"Usuario no encontrado" };

  const [uEmail, uName, uHash, uSalt, uRole, uEnabled] = hit.data;
  if (!uEnabled) return { ok:false, msg:"Usuario deshabilitado" };
  if (_hash(password, uSalt) !== uHash) return { ok:false, msg:"Contraseña incorrecta" };

  _usersSheet().getRange(hit.row, 8).setValue(_now()); // last_login (col H)
  return { ok:true, msg:"Login OK", user:{ email:uEmail, name:uName, role:uRole, enabled:uEnabled } };
}

function listUsers() {
  const sh = _usersSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2,1,last-1,8).getValues();
  return vals.map(r => ({ email:r[0], name:r[1], role:r[4], enabled:r[5], created_at:r[6], last_login:r[7] }));
}

function setUserEnabled(email, enabled) {
  const hit = _findUserRow(email || "");
  if (!hit) return { ok:false, msg:"Usuario no encontrado" };
  _usersSheet().getRange(hit.row, 6).setValue(!!enabled); // col F = enabled
  return { ok:true, msg:`Usuario ${enabled ? "habilitado" : "deshabilitado"}` };
}

function _hash(password, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password);
  return Utilities.base64EncodeWebSafe(raw);
}

function _now() {
  return new Date();
}

function getIndexUrl() {
  var base = ScriptApp.getService().getUrl(); // URL del deployment (exec)
  return base + (base.indexOf('?') === -1 ? '?page=index' : '&page=index');
}

// --- REGISTRO ---
function registerUser(name, email, password, role) {
  try {
    email = (email||"").trim(); name = (name||"").trim(); role=(role||"user").trim();
    if (!email || !password) return {ok:false, msg:"Email y contraseña son obligatorios"};
    const hit = _findUserRow(email);
    if (hit) return {ok:false, msg:"El correo ya está registrado"};

    const sh   = _usersSheet();                 // <- aquí fallaba por permisos
    const salt = Utilities.getUuid();
    const hash = _hash(password, salt);
    sh.appendRow([email, name, hash, salt, role, true, new Date(), ""]);
    return {ok:true, msg:"Usuario registrado correctamente"};
  } catch (e) {
    return {ok:false, msg:"No se pudo registrar: " + (e && e.message ? e.message : e)};
  }
}



/**
 * CREATE RECORD
 * REF:
 * https://developers.google.com/sheets/api/guides/values#append_values
 */
function createRecord(values) {
  try {
    Sheets.Spreadsheets.Values.append(
      { values: values },
      SPREADSHEETID,
      DATARANGE,
      { valueInputOption: "RAW" }
    );
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}

/**
 * READ RECORD
 * REF:
 * https://developers.google.com/sheets/api/guides/values#read
 */
function readRecord(range) {
  try {
    let result = Sheets.Spreadsheets.Values.get(SPREADSHEETID, range);
    return result.values;
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}

/**
 * UPDATE RECORD
 * REF:
 * https://developers.google.com/sheets/api/guides/values#write_to_a_single_range
 */
function updateRecord(values, updateRange) {
  try {
    let valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.update(valueRange, SPREADSHEETID, updateRange, { valueInputOption: "RAW" });
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}


/**
 * DELETE RECORD
 * Ref:
 * https://developers.google.com/sheets/api/guides/batchupdate
 * https://developers.google.com/sheets/api/samples/rowcolumn#delete_rows_or_columns
*/

function _dataSheetId() {
  const sh = SpreadsheetApp.openById(SPREADSHEETID).getSheetByName(DATASHEET);
  return sh.getSheetId(); // número correcto, no asumir "0"
}


function deleteRecord(id) {
  const rowToDelete = getRowIndexById(id); // índice base 0 OK
  const deleteRequest = {
    deleteDimension: {
      range: {
        sheetId: _dataSheetId(),
        dimension: "ROWS",
        startIndex: rowToDelete,
        endIndex: rowToDelete + 1
      }
    }
  };
  Sheets.Spreadsheets.batchUpdate({ requests: [deleteRequest] }, SPREADSHEETID);
  return getLastTenRecords();
}


/**
 * RETURN LAST 10 RECORDS IN THE SHEET
 */
function getLastTenRecords() {
  let lastRow = readRecord(DATARANGE).length + 1;
  let startRow = lastRow - 9;
  if (startRow < 2) { //If less than 10 records, eleminate the header row and start from second row
    startRow = 2;
  }
  let range = DATASHEET + "!A" + startRow + ":" + LASTCOL + lastRow;
  let lastTenRecords = readRecord(range);
  Logger.log(lastTenRecords);
  return lastTenRecords;
}


//GET ALL RECORDS
function getAllRecords() {
  const allRecords = readRecord(DATARANGE);
  return allRecords;
}

//GET RECORD FOR THE GIVEN ID
function getRecordById(id) {
  if (!id || !checkId(id)) {
    return null;
  }
  const range = getRangeById(id);
  if (!range) {
    return null;
  }
  const result = readRecord(range);
  return result;
}

function getRowIndexById(id) {
  if (!id) {
    throw new Error('Invalid ID');
  }

  const idList = readRecord(IDRANGE);
  for (var i = 0; i < idList.length; i++) {
    if (id == idList[i][0]) {
      var rowIndex = parseInt(i + 1);
      return rowIndex;
    }
  }
}


//VALIDATE ID
function checkId(id) {
  const idList = readRecord(IDRANGE).flat();
  return idList.includes(id);
}


//GET DATA RANGE IN A1 NOTATION FOR GIVEN ID
function getRangeById(id) {
  if (!id) {
    return null;
  }
  const idList = readRecord(IDRANGE);
  const rowIndex = idList.findIndex(item => item[0] === id);
  if (rowIndex === -1) {
    return null;
  }
  const range = `Data!A${rowIndex + 2}:${LASTCOL}${rowIndex + 2}`;
  return range;
}


//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

//GENERATE UNIQUE ID
function generateUniqueId() {
  let id = Utilities.getUuid();
  return id;
}

function getGradeList() {
  gradeList = readRecord(DROPDOWNRANGE1);
  return gradeList;
}

function getUnidadeList() {
  unidadList = readRecord(DROPDOWNRANGE2);
  return unidadList;
}


//SEARCH RECORDS
function searchRecords(formObject) {
  let result = [];
  try {
    if (formObject.searchText) {//Execute if form passes search text
      const data = readRecord(DATARANGE);
      const q = String(formObject.searchText).toLowerCase();

      // Loop through each row and column to search for matches
      for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
          const cell = (data[i][j] == null) ? "" : String(data[i][j]);
          if (cell.toLowerCase().includes(q)) {
            result.push(data[i]);
            break; // Stop searching for other matches in this row
          }
        }
      }
    }
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
  return result;
}
