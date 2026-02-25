// ============================================
// CONFIGURACIÓN
// ============================================
const MASTER_SHEET_ID = '1NgwrfhTjw8TifTEflk66YQM3M5lwDRbzNusJlK1NfgU';

// Nombres de las hojas (en spreadsheets per-user)
const SHEET_MOVIMIENTOS = 'Movimientos';
const SHEET_CONFIG = 'Configuración';
const SHEET_PRESUPUESTOS = 'Presupuestos';
const SHEET_CATEGORIAS = 'Categorías';

// Nombre de la hoja de usuarios (en master sheet)
const SHEET_USERS = 'Usuarios';

// ============================================
// HELPERS
// ============================================

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function hashPassword(password, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password);
  return raw.map(function(b) {
    return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2);
  }).join('');
}

function generateToken() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

function generateSalt() {
  return Utilities.getUuid();
}

// Busca un usuario por username en la tabla Usuarios
function findUser(username) {
  const sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === username.toLowerCase()) {
      return {
        rowNum: i + 1,
        username: data[i][0],
        password_hash: data[i][1],
        salt: data[i][2],
        spreadsheet_id: data[i][3],
        created_at: data[i][4],
        session_token: data[i][5],
        email: data[i][6] || '',
        reset_token: data[i][7] || '',
        reset_expiry: data[i][8] || ''
      };
    }
  }
  return null;
}

// Busca un usuario por email en la tabla Usuarios
function findUserByEmail(email) {
  if (!email) return null;
  const sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  var emailLower = email.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().toLowerCase().trim() === emailLower) {
      return {
        rowNum: i + 1,
        username: data[i][0],
        password_hash: data[i][1],
        salt: data[i][2],
        spreadsheet_id: data[i][3],
        created_at: data[i][4],
        session_token: data[i][5],
        email: data[i][6] || '',
        reset_token: data[i][7] || '',
        reset_expiry: data[i][8] || ''
      };
    }
  }
  return null;
}

// Resuelve un token de sesión a {username, spreadsheet_id}
function resolveSession(token) {
  if (!token) return { error: 'Sesión inválida' };
  const sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  if (!sheet) return { error: 'Sesión inválida' };
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] && data[i][5] === token) {
      return {
        username: data[i][0],
        spreadsheet_id: data[i][3]
      };
    }
  }
  return { error: 'Sesión inválida' };
}

// ============================================
// ENDPOINTS - GET y POST
// ============================================

function doGet(e) {
  var action = e.parameter.action;
  var result;

  try {
    // --- Endpoints públicos (sin auth) ---
    if (action === 'ping') {
      return jsonResponse({ ok: true, timestamp: new Date().toISOString() });
    }

    if (action === 'validateSession') {
      var session = resolveSession(e.parameter.token);
      if (session.error) return jsonResponse({ error: session.error });
      return jsonResponse({ ok: true, username: session.username });
    }

    // --- Endpoints autenticados ---
    var sess = resolveSession(e.parameter.token);
    if (sess.error) return jsonResponse(sess);
    var userSheet = SpreadsheetApp.openById(sess.spreadsheet_id);

    switch (action) {
      case 'getMovimientos':
        result = getMovimientos(userSheet, e.parameter.mes);
        break;
      case 'getConfig':
        result = getConfig(userSheet);
        break;
      case 'getCategorias':
        result = getCategorias(userSheet);
        break;
      case 'getPresupuestos':
        result = getPresupuestos(userSheet);
        break;
      default:
        result = { error: 'Acción no reconocida' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return jsonResponse(result);
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'JSON inválido' });
  }

  var action = body.action;
  var result;

  try {
    // --- Endpoints públicos (sin auth) ---
    if (action === 'register') {
      return jsonResponse(register(body.data));
    }
    if (action === 'login') {
      return jsonResponse(login(body.data));
    }
    if (action === 'setPassword') {
      return jsonResponse(setPassword(body.data));
    }
    if (action === 'forgotPassword') {
      return jsonResponse(forgotPassword(body.data));
    }
    if (action === 'resetPassword') {
      return jsonResponse(resetPassword(body.data));
    }

    // --- Endpoints autenticados ---
    var sess = resolveSession(body.token);
    if (sess.error) return jsonResponse(sess);
    var userSheet = SpreadsheetApp.openById(sess.spreadsheet_id);

    switch (action) {
      case 'addMovimiento':
        result = addMovimiento(userSheet, body.data);
        break;
      case 'deleteMovimiento':
        result = deleteMovimiento(userSheet, body.data.id);
        break;
      case 'updateMovimiento':
        result = updateMovimiento(userSheet, body.data);
        break;
      case 'updateConfig':
        result = updateConfig(userSheet, body.data);
        break;
      case 'updatePresupuesto':
        result = updatePresupuesto(userSheet, body.data);
        break;
      case 'addCategoria':
        result = addCategoria(userSheet, body.data);
        break;
      case 'syncBatch':
        result = syncBatch(userSheet, body.data);
        break;
      default:
        result = { error: 'Acción no reconocida' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return jsonResponse(result);
}

// ============================================
// AUTENTICACIÓN
// ============================================

function register(data) {
  var username = (data.username || '').trim();
  var password = data.password || '';
  var email = (data.email || '').trim().toLowerCase();

  if (!username || username.length < 2) {
    return { error: 'El nombre debe tener al menos 2 caracteres' };
  }
  if (!password || password.length < 4) {
    return { error: 'La contraseña debe tener al menos 4 caracteres' };
  }
  if (!email || email.indexOf('@') === -1 || email.indexOf('.') === -1) {
    return { error: 'Ingresá un email válido' };
  }

  // Check if username already exists
  var existing = findUser(username);
  if (existing) {
    return { error: 'Ese nombre de usuario ya existe' };
  }

  // Check if email already registered
  var existingEmail = findUserByEmail(email);
  if (existingEmail) {
    return { error: 'Ese email ya está registrado' };
  }

  // Create user's personal spreadsheet
  var newSS = SpreadsheetApp.create('Gastos - ' + username);
  var newSheetId = newSS.getId();

  // Setup standard tabs in the new spreadsheet
  setupUserSheet(newSS);

  // Generate credentials
  var salt = generateSalt();
  var hash = hashPassword(password, salt);
  var token = generateToken();

  // Add user to master sheet (9 columns: username, hash, salt, sheet_id, created_at, token, email, reset_token, reset_expiry)
  var usersSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  usersSheet.appendRow([username, hash, salt, newSheetId, new Date(), token, email, '', '']);

  return { ok: true, token: token, username: username };
}

function login(data) {
  var username = (data.username || '').trim();
  var password = data.password || '';

  if (!username) return { error: 'Ingresá tu usuario' };
  if (!password) return { error: 'Ingresá tu contraseña' };

  var user = findUser(username);
  if (!user) {
    return { error: 'Usuario no encontrado' };
  }

  // Check if migrated user needs to set password
  if (user.password_hash === 'PENDING') {
    return { needsPassword: true, username: user.username };
  }

  // Verify password
  var hash = hashPassword(password, user.salt);
  if (hash !== user.password_hash) {
    return { error: 'Contraseña incorrecta' };
  }

  // Generate new session token
  var token = generateToken();
  var sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowNum, 6).setValue(token); // column 6 = session_token

  return { ok: true, token: token, username: user.username };
}

function setPassword(data) {
  var username = (data.username || '').trim();
  var password = data.password || '';

  if (!username) return { error: 'Usuario requerido' };
  if (!password || password.length < 4) {
    return { error: 'La contraseña debe tener al menos 4 caracteres' };
  }

  var user = findUser(username);
  if (!user) return { error: 'Usuario no encontrado' };
  if (user.password_hash !== 'PENDING') {
    return { error: 'Este usuario ya tiene contraseña. Usá login.' };
  }

  // Set password
  var salt = generateSalt();
  var hash = hashPassword(password, salt);
  var token = generateToken();

  var sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowNum, 2).setValue(hash);    // password_hash
  sheet.getRange(user.rowNum, 3).setValue(salt);     // salt
  sheet.getRange(user.rowNum, 6).setValue(token);    // session_token

  return { ok: true, token: token, username: user.username };
}

function forgotPassword(data) {
  var email = (data.email || '').trim().toLowerCase();

  if (!email || email.indexOf('@') === -1) {
    return { error: 'Ingresá un email válido' };
  }

  var user = findUserByEmail(email);
  if (!user) {
    return { error: 'No hay cuenta con ese email' };
  }

  if (user.password_hash === 'PENDING') {
    return { error: 'Tu cuenta aún no tiene contraseña. Usá "Iniciar sesión" con tu usuario.' };
  }

  // Generate 6-digit code
  var code = (Math.floor(Math.random() * 900000) + 100000).toString();
  var expiry = new Date(new Date().getTime() + 15 * 60 * 1000).toISOString(); // 15 minutes

  // Save reset token and expiry
  var sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowNum, 8).setValue(code);    // reset_token (col H)
  sheet.getRange(user.rowNum, 9).setValue(expiry);   // reset_expiry (col I)

  // Send email
  var subject = 'Código de recuperación - Mis Gastos';
  var body = 'Hola ' + user.username + ',\n\n'
    + 'Tu código de recuperación es: ' + code + '\n\n'
    + 'Expira en 15 minutos.\n\n'
    + 'Si no solicitaste esto, ignorá este email.';

  MailApp.sendEmail(email, subject, body);

  return { ok: true, message: 'Código enviado a tu email' };
}

function resetPassword(data) {
  var email = (data.email || '').trim().toLowerCase();
  var code = (data.code || '').trim();
  var password = data.password || '';

  if (!email || !code || !password) {
    return { error: 'Datos incompletos' };
  }
  if (password.length < 4) {
    return { error: 'La contraseña debe tener al menos 4 caracteres' };
  }

  var user = findUserByEmail(email);
  if (!user) {
    return { error: 'Usuario no encontrado' };
  }

  // Verify code
  if (!user.reset_token || user.reset_token.toString() !== code) {
    return { error: 'Código incorrecto' };
  }

  // Verify not expired
  if (!user.reset_expiry || new Date() > new Date(user.reset_expiry)) {
    return { error: 'El código expiró. Pedí uno nuevo.' };
  }

  // Set new password
  var salt = generateSalt();
  var hash = hashPassword(password, salt);
  var token = generateToken();

  var sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowNum, 2).setValue(hash);     // password_hash (col B)
  sheet.getRange(user.rowNum, 3).setValue(salt);      // salt (col C)
  sheet.getRange(user.rowNum, 6).setValue(token);     // session_token (col F)
  sheet.getRange(user.rowNum, 8).setValue('');         // clear reset_token (col H)
  sheet.getRange(user.rowNum, 9).setValue('');         // clear reset_expiry (col I)

  return { ok: true, token: token, username: user.username };
}

// ============================================
// FUNCIONES CRUD - MOVIMIENTOS
// ============================================

function getMovimientos(userSheet, mes) {
  var sheet = userSheet.getSheetByName(SHEET_MOVIMIENTOS);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  var headers = data[0];
  var movimientos = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var mov = {};
    headers.forEach(function(h, idx) {
      mov[h] = row[idx];
    });

    // Formatear fecha
    if (mov.fecha instanceof Date) {
      mov.fecha = Utilities.formatDate(mov.fecha, 'America/Montevideo', 'yyyy-MM-dd');
    }

    // Filtrar por mes si se especifica (formato: "2026-02")
    if (mes) {
      if (typeof mov.fecha === 'string' && !mov.fecha.startsWith(mes)) {
        continue;
      }
    }

    movimientos.push(mov);
  }

  return movimientos;
}

function addMovimiento(userSheet, data) {
  var sheet = userSheet.getSheetByName(SHEET_MOVIMIENTOS);

  var id = Utilities.getUuid();

  var row = [
    id,
    data.fecha || Utilities.formatDate(new Date(), 'America/Montevideo', 'yyyy-MM-dd'),
    parseFloat(data.monto),
    data.moneda || 'UYU',
    data.tipo || 'gasto',
    data.categoria || 'Otros',
    data.descripcion || ''
  ];

  sheet.appendRow(row);

  return { ok: true, id: id };
}

function deleteMovimiento(userSheet, id) {
  var sheet = userSheet.getSheetByName(SHEET_MOVIMIENTOS);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }

  return { error: 'Movimiento no encontrado' };
}

function updateMovimiento(userSheet, data) {
  var sheet = userSheet.getSheetByName(SHEET_MOVIMIENTOS);
  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      var rowNum = i + 1;
      headers.forEach(function(h, idx) {
        if (h !== 'id' && data[h] !== undefined) {
          sheet.getRange(rowNum, idx + 1).setValue(data[h]);
        }
      });
      return { ok: true };
    }
  }

  return { error: 'Movimiento no encontrado' };
}

// Sync batch para offline queue
function syncBatch(userSheet, items) {
  var results = [];
  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    try {
      var res;
      switch (item.action) {
        case 'addMovimiento':
          res = addMovimiento(userSheet, item.data);
          break;
        case 'deleteMovimiento':
          res = deleteMovimiento(userSheet, item.data.id);
          break;
        case 'updateMovimiento':
          res = updateMovimiento(userSheet, item.data);
          break;
        default:
          res = { error: 'Acción no reconocida en batch' };
      }
      results.push({ tempId: item.tempId, result: res });
    } catch (err) {
      results.push({ tempId: item.tempId, error: err.message });
    }
  }
  return { ok: true, results: results };
}

// ============================================
// FUNCIONES - CONFIGURACIÓN
// ============================================

function getConfig(userSheet) {
  var sheet = userSheet.getSheetByName(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();
  var config = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      config[data[i][0]] = data[i][1];
    }
  }

  return config;
}

function updateConfig(userSheet, data) {
  var sheet = userSheet.getSheetByName(SHEET_CONFIG);
  var allData = sheet.getDataRange().getValues();

  for (var key in data) {
    var found = false;
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, data[key]]);
    }
  }

  return { ok: true };
}

// ============================================
// FUNCIONES - PRESUPUESTOS
// ============================================

function getPresupuestos(userSheet) {
  var sheet = userSheet.getSheetByName(SHEET_PRESUPUESTOS);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  var headers = data[0];
  var presupuestos = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var p = {};
    headers.forEach(function(h, idx) {
      p[h] = data[i][idx];
    });
    presupuestos.push(p);
  }

  return presupuestos;
}

function updatePresupuesto(userSheet, data) {
  var sheet = userSheet.getSheetByName(SHEET_PRESUPUESTOS);
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.categoria) {
      sheet.getRange(i + 1, 2).setValue(parseFloat(data.presupuesto_mensual));
      sheet.getRange(i + 1, 3).setValue(data.moneda || 'UYU');
      return { ok: true };
    }
  }

  // Si no existe, agregar
  sheet.appendRow([data.categoria, parseFloat(data.presupuesto_mensual), data.moneda || 'UYU']);
  return { ok: true };
}

// ============================================
// FUNCIONES - CATEGORÍAS
// ============================================

function getCategorias(userSheet) {
  var sheet = userSheet.getSheetByName(SHEET_CATEGORIAS);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  var headers = data[0];
  var categorias = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var c = {};
    headers.forEach(function(h, idx) {
      c[h] = data[i][idx];
    });
    categorias.push(c);
  }

  return categorias;
}

function addCategoria(userSheet, data) {
  var sheet = userSheet.getSheetByName(SHEET_CATEGORIAS);
  sheet.appendRow([data.nombre, data.emoji || '📋', data.color || '#9E9E9E', data.tipo || 'gasto']);
  return { ok: true };
}

// ============================================
// SETUP - Crea las hojas en un spreadsheet de usuario
// ============================================

function setupUserSheet(ss) {
  // Crear hoja Movimientos (sin registrado_por - es per-user)
  var sheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (sheet) {
    sheet.setName(SHEET_MOVIMIENTOS);
  } else {
    sheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    if (!sheet) sheet = ss.insertSheet(SHEET_MOVIMIENTOS);
  }
  sheet.getRange(1, 1, 1, 7).setValues([['id', 'fecha', 'monto', 'moneda', 'tipo', 'categoria', 'descripcion']]);

  // Crear hoja Configuración
  sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG);
  sheet.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]);
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['tipo_cambio_usd', 42.5]);
    sheet.appendRow(['presupuesto_total_mensual_uyu', 50000]);
  }

  // Crear hoja Presupuestos
  sheet = ss.getSheetByName(SHEET_PRESUPUESTOS);
  if (!sheet) sheet = ss.insertSheet(SHEET_PRESUPUESTOS);
  sheet.getRange(1, 1, 1, 3).setValues([['categoria', 'presupuesto_mensual', 'moneda']]);

  // Crear hoja Categorías con defaults
  sheet = ss.getSheetByName(SHEET_CATEGORIAS);
  if (!sheet) sheet = ss.insertSheet(SHEET_CATEGORIAS);
  sheet.getRange(1, 1, 1, 4).setValues([['nombre', 'emoji', 'color', 'tipo']]);
  if (sheet.getLastRow() < 2) {
    var cats = [
      ['Supermercado', '🛒', '#4CAF50', 'gasto'],
      ['Farmacia', '💊', '#E91E63', 'gasto'],
      ['Médico', '🏥', '#F44336', 'gasto'],
      ['Transporte', '🚌', '#FF9800', 'gasto'],
      ['Casa y Hogar', '🏠', '#795548', 'gasto'],
      ['Servicios', '⚡', '#FFC107', 'gasto'],
      ['Ropa', '👗', '#9C27B0', 'gasto'],
      ['Comida afuera', '🍽️', '#FF5722', 'gasto'],
      ['Entretenimiento', '🎬', '#3F51B5', 'gasto'],
      ['Regalos', '🎁', '#E91E63', 'gasto'],
      ['Otros', '📋', '#9E9E9E', 'gasto'],
      ['Jubilación', '💰', '#4CAF50', 'ingreso'],
      ['Banco/Intereses', '🏦', '#2196F3', 'ingreso'],
      ['Regalos recibidos', '🎁', '#E91E63', 'ingreso'],
      ['Otros ingresos', '📋', '#9E9E9E', 'ingreso']
    ];
    sheet.getRange(2, 1, cats.length, 4).setValues(cats);
  }
}

// ============================================
// SETUP MASTER - Crea la hoja Usuarios en el master sheet
// ============================================

function setupMasterUsersTab() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 1, 9).setValues([['username', 'password_hash', 'salt', 'spreadsheet_id', 'created_at', 'session_token', 'email', 'reset_token', 'reset_expiry']]);
  }
  return 'Usuarios tab created.';
}

// Agrega columnas email/reset_token/reset_expiry a tabla Usuarios existente (correr UNA VEZ)
function addEmailColumnToExistingUsers() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return 'No Users tab found.';

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('email') === -1) {
    var nextCol = headers.length + 1;
    sheet.getRange(1, nextCol).setValue('email');
    sheet.getRange(1, nextCol + 1).setValue('reset_token');
    sheet.getRange(1, nextCol + 2).setValue('reset_expiry');
    return 'Columns email, reset_token, reset_expiry added at column ' + nextCol;
  }
  return 'Email column already exists.';
}

// ============================================
// MIGRACIÓN - Corre una sola vez para migrar datos existentes
// ============================================

function migrateExistingData() {
  // 1. Ensure Users tab exists
  setupMasterUsersTab();

  var masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var movSheet = masterSS.getSheetByName('Movimientos');

  if (!movSheet) return 'No Movimientos tab found in master sheet.';

  var data = movSheet.getDataRange().getValues();
  if (data.length <= 1) return 'No data to migrate.';

  var headers = data[0];
  var regPorIdx = headers.indexOf('registrado_por');

  // 2. Group movements by registrado_por
  var userMovs = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var user = regPorIdx >= 0 ? (row[regPorIdx] || 'Anónimo') : 'Anónimo';
    if (!userMovs[user]) userMovs[user] = [];
    userMovs[user].push(row);
  }

  // 3. Get existing config, presupuestos, categorias from master
  var configSheet = masterSS.getSheetByName(SHEET_CONFIG);
  var configData = configSheet ? configSheet.getDataRange().getValues() : [];
  var presupSheet = masterSS.getSheetByName(SHEET_PRESUPUESTOS);
  var presupData = presupSheet ? presupSheet.getDataRange().getValues() : [];
  var catSheet = masterSS.getSheetByName(SHEET_CATEGORIAS);
  var catData = catSheet ? catSheet.getDataRange().getValues() : [];

  var usersSheet = masterSS.getSheetByName(SHEET_USERS);
  var results = [];

  // 4. For each user, create their spreadsheet
  for (var username in userMovs) {
    // Check if user already migrated
    var existing = findUser(username);
    if (existing) {
      results.push(username + ': SKIPPED (already exists)');
      continue;
    }

    // Create spreadsheet
    var newSS = SpreadsheetApp.create('Gastos - ' + username);
    var newSheetId = newSS.getId();

    // Setup standard tabs
    setupUserSheet(newSS);

    // Copy movimientos (without registrado_por column)
    var userMovSheet = newSS.getSheetByName(SHEET_MOVIMIENTOS);
    var movRows = userMovs[username].map(function(row) {
      // Take columns 0-6 (id, fecha, monto, moneda, tipo, categoria, descripcion), skip index 7 (registrado_por)
      return [row[0], row[1], row[2], row[3], row[4], row[5], row[6]];
    });
    if (movRows.length > 0) {
      userMovSheet.getRange(2, 1, movRows.length, 7).setValues(movRows);
    }

    // Copy config
    if (configData.length > 1) {
      var userConfigSheet = newSS.getSheetByName(SHEET_CONFIG);
      userConfigSheet.clear();
      userConfigSheet.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
    }

    // Copy presupuestos
    if (presupData.length > 1) {
      var userPresSheet = newSS.getSheetByName(SHEET_PRESUPUESTOS);
      userPresSheet.clear();
      userPresSheet.getRange(1, 1, presupData.length, presupData[0].length).setValues(presupData);
    }

    // Copy categorias
    if (catData.length > 1) {
      var userCatSheet = newSS.getSheetByName(SHEET_CATEGORIAS);
      userCatSheet.clear();
      userCatSheet.getRange(1, 1, catData.length, catData[0].length).setValues(catData);
    }

    // Add user to Users tab with PENDING password (9 cols)
    usersSheet.appendRow([username, 'PENDING', '', newSheetId, new Date(), '', '', '', '']);

    results.push(username + ': ' + movRows.length + ' movimientos migrated -> ' + newSheetId);
  }

  return 'Migration complete:\n' + results.join('\n');
}
