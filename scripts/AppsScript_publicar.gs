/**
 * AUDITORIA DE LOTES - LUNA GI
 * Publica el dashboard vivo directamente desde Google Sheets.
 *
 * Lee la hoja LOTES, arma el mismo datos.json que produce auditar.py,
 * lo inyecta en la plantilla y sube ambos archivos a GitHub. GitHub Pages
 * publica solo. No hace falta Python ni esta computadora.
 *
 * INSTALACION (una sola vez):
 *   1. En el Sheet:  Extensiones > Apps Script
 *   2. Borra lo que haya y pega TODO este archivo. Guarda.
 *   3. Recarga el Sheet. Aparece el menu "Dashboard".
 *   4. Dashboard > Configurar token de GitHub  (pegas tu token)
 *   5. Dashboard > Publicar ahora
 */

const REPO = 'afernandezfalconi/DASHBOARD_PROYECTOS_LUNAGI';
const RAMA = 'main';
const HOJA = 'LOTES';
const URL_DASHBOARD = 'https://afernandezfalconi.github.io/DASHBOARD_PROYECTOS_LUNAGI/';

const ESTATUS = {
  'DISPONIBLE': 'disponible',
  'APARTADO': 'apartado',
  'VENDIDO': 'vendido',
  'VENDIDO SIN DATO': 'vendido_sin_dato'
};

/**
 * Socios de la empresa. Que su nombre aparezca en CLIENTE no es una venta:
 * significa que el lote esta apartado para el socio y fuera del sistema.
 *
 * Se compara por nombre COMPLETO y exacto, a proposito. En GUAYACÁN existe
 * un comprador real llamado CRISTHIAN HERNANDEZ MATTERN; una regla que
 * buscara solo "CRISTHIAN" lo arrastraria por error.
 */
const SOCIOS_QUE_APARTAN = [
  'CRISTHIAN ENRIQUE LUNA ORDOÑEZ'
];

function esSocio(nombre) {
  const n = norm(nombre);
  return SOCIOS_QUE_APARTAN.some(function (s) { return norm(s) === n; });
}

// --------------------------------------------------------------- menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard')
    .addItem('Publicar ahora', 'publicar')
    .addItem('Revisar sin publicar', 'revisar')
    .addSeparator()
    .addItem('Marcar apartados de socios', 'normalizarApartadosDeSocios')
    .addItem('Traer m2 desde las bases', 'traerMetrosCuadrados')
    .addItem('Traer tipo de venta desde las bases', 'traerTipoDeVenta')
    .addSeparator()
    .addItem('Abrir el dashboard', 'abrirDashboard')
    .addItem('Configurar token de GitHub', 'configurarToken')
    .addToUi();
}

function configurarToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'Token de GitHub',
    'Pega un token con permiso de escritura SOLO sobre el repositorio\n' +
    REPO + '\n\n' +
    'Se guarda en las propiedades de este script, no en la hoja.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (!t) { ui.alert('No guardé nada: el token venía vacío.'); return; }
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', t);
  ui.alert('Token guardado. Ya puedes usar "Publicar ahora".');
}

function abrirDashboard() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p style="font:14px system-ui">El dashboard vivo:</p>' +
      '<p><a href="' + URL_DASHBOARD + '" target="_blank" style="font:14px system-ui">' +
      URL_DASHBOARD + '</a></p>').setWidth(420).setHeight(120),
    'Dashboard vivo');
}

// ----------------------------------------------------------- utilidades
function norm(s) {
  return String(s === null || s === undefined ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function texto(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return '';
  return String(v).replace(/\s+/g, ' ').trim();
}

function num(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (v === null || v === undefined || v instanceof Date) return 0;
  const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
  return isFinite(n) ? n : 0;
}

function r2(x) { return Math.round((x + Number.EPSILON) * 100) / 100; }

// clave que ordena "10" despues de "9" y no alfabeticamente
function claveNum(x) {
  const s = String(x).trim();
  return /^\d+$/.test(s) ? [0, parseInt(s, 10), ''] : [1, 0, s];
}
function cmpClave(a, b) {
  for (let i = 0; i < a.length; i++) {
    if (a[i] < b[i]) return -1;
    if (a[i] > b[i]) return 1;
  }
  return 0;
}

/**
 * Precio promedio por m2, calculado SOLO con lotes vendidos que tengan a la vez
 * monto y superficie. Sirve para traducir el inventario disponible a dinero.
 */
function precioPorM2(proyectos) {
  let dinero = 0, metros = 0;
  proyectos.forEach(function (p) {
    p.lotes.forEach(function (l) {
      if (l.m2 && l.monto && l.estatus === 'vendido') { dinero += l.monto; metros += l.m2; }
    });
  });
  return metros ? r2(dinero / metros) : 0;
}

// ------------------------------------------------------- armar los datos
function construirDatos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  if (!hoja) throw new Error('No encuentro la hoja "' + HOJA + '"');

  const filas = hoja.getDataRange().getValues();
  if (!filas.length) throw new Error('La hoja ' + HOJA + ' está vacía');

  let iCab = -1;
  for (let i = 0; i < Math.min(filas.length, 10); i++) {
    if (norm(filas[i][0]) === 'PROYECTO') { iCab = i; break; }
  }
  if (iCab < 0) throw new Error('No encuentro el encabezado (fila con PROYECTO)');

  const cab = filas[iCab].map(norm);
  const need = ['PROYECTO', 'MANZANA', 'LOTE', 'CLIENTE', 'ESTATUS',
                'MONTO', 'ENGANCHE', 'INGRESO', 'CATEGORIA', 'NOTA'];
  const col = {};
  need.forEach(function (c) {
    const j = cab.indexOf(c);
    if (j < 0) throw new Error('Falta la columna ' + c + ' en la hoja ' + HOJA);
    col[c] = j;
  });
  col.M2 = cab.indexOf('M2');                  // opcional: -1 si aun no se ha traido
  col.TIPO = cab.indexOf('TIPO VENTA');        // opcional: contado / credito

  const porProyecto = {}, orden = [], problemas = [];

  for (let i = iCab + 1; i < filas.length; i++) {
    const f = filas[i];
    const proyecto = texto(f[col.PROYECTO]);
    const lote = texto(f[col.LOTE]);
    if (!proyecto && !lote) continue;
    if (!proyecto || !lote) {
      problemas.push('Fila ' + (i + 1) + ': falta PROYECTO o LOTE');
      continue;
    }

    const est = ESTATUS[norm(f[col.ESTATUS])];
    if (!est) {
      problemas.push('Fila ' + (i + 1) + ': estatus no válido "' + texto(f[col.ESTATUS]) + '"');
      continue;
    }

    const monto = num(f[col.MONTO]);
    const enganche = num(f[col.ENGANCHE]);
    const reg = {
      mza: texto(f[col.MANZANA]),
      lote: lote,
      cliente: texto(f[col.CLIENTE]),
      estatus: est,
      monto: monto ? r2(monto) : null,
      enganche: enganche ? r2(enganche) : null,
      ingreso: r2(num(f[col.INGRESO]))
    };
    const sup = col.M2 >= 0 ? num(f[col.M2]) : 0;
    if (sup > 0) reg.m2 = r2(sup);
    const tipo = col.TIPO >= 0 ? texto(f[col.TIPO]) : '';
    if (tipo) reg.tipo = tipo;
    if (est === 'vendido_sin_dato') {
      reg.categoria = texto(f[col.CATEGORIA]) || 'Cliente sin dato capturado';
    }
    if (texto(f[col.NOTA])) reg.nota = texto(f[col.NOTA]);

    if (!porProyecto[proyecto]) { porProyecto[proyecto] = []; orden.push(proyecto); }
    porProyecto[proyecto].push(reg);
  }

  const proyectos = [], alertas = [];
  orden.forEach(function (nombre) {
    const lotes = porProyecto[nombre];
    lotes.forEach(function (l) {
      if (l.estatus === 'vendido_sin_dato') {
        alertas.push({ proyecto: nombre, mza: l.mza, lote: l.lote,
                       cliente: l.cliente, categoria: l.categoria || '' });
      }
    });
    const cuenta = function (e) { return lotes.filter(function (l) { return l.estatus === e; }).length; };
    const sup = function (filtro) {
      return r2(lotes.reduce(function (s, l) {
        return s + ((l.m2 && filtro(l)) ? l.m2 : 0); }, 0));
    };
    const vendido = function (l) { return l.estatus === 'vendido' || l.estatus === 'vendido_sin_dato'; };
    proyectos.push({
      nombre: nombre,
      total: lotes.length,
      disponibles: cuenta('disponible'),
      apartados: cuenta('apartado'),
      vendidos: cuenta('vendido') + cuenta('vendido_sin_dato'),
      vendidos_con_ingreso: cuenta('vendido'),
      vendidos_sin_ingreso: cuenta('vendido_sin_dato'),
      ingresos: r2(lotes.reduce(function (s, l) { return s + l.ingreso; }, 0)),
      m2_total: sup(function () { return true; }),
      m2_disponible: sup(function (l) { return l.estatus === 'disponible'; }),
      m2_vendido: sup(vendido),
      lotes_con_m2: lotes.filter(function (l) { return !!l.m2; }).length,
      contado: lotes.filter(function (l) { return vendido(l) && norm(l.tipo) === 'CONTADO'; }).length,
      credito: lotes.filter(function (l) { return vendido(l) && norm(l.tipo) === 'CREDITO'; }).length,
      ingresos_contado: r2(lotes.reduce(function (s, l) {
        return s + (vendido(l) && norm(l.tipo) === 'CONTADO' ? l.ingreso : 0); }, 0)),
      ingresos_credito: r2(lotes.reduce(function (s, l) {
        return s + (vendido(l) && norm(l.tipo) === 'CREDITO' ? l.ingreso : 0); }, 0)),
      lotes: lotes
    });
  });

  proyectos.sort(function (a, b) { return b.total - a.total; });
  alertas.sort(function (a, b) {
    if (a.proyecto !== b.proyecto) return a.proyecto < b.proyecto ? -1 : 1;
    const m = cmpClave(claveNum(a.mza), claveNum(b.mza));
    return m !== 0 ? m : cmpClave(claveNum(a.lote), claveNum(b.lote));
  });

  const suma = function (k) {
    return proyectos.reduce(function (s, p) { return s + p[k]; }, 0);
  };

  return {
    datos: {
      generated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      grand: {
        total: suma('total'),
        available: suma('disponibles'),
        reserved: suma('apartados'),
        sold: suma('vendidos'),
        sold_with_data: suma('vendidos_con_ingreso'),
        sold_no_data: suma('vendidos_sin_ingreso'),
        income: r2(suma('ingresos')),
        m2_total: r2(suma('m2_total')),
        m2_available: r2(suma('m2_disponible')),
        m2_sold: r2(suma('m2_vendido')),
        lots_with_m2: suma('lotes_con_m2'),
        price_per_m2: precioPorM2(proyectos),
        cash_lots: suma('contado'),
        credit_lots: suma('credito'),
        cash_income: r2(suma('ingresos_contado')),
        credit_income: r2(suma('ingresos_credito'))
      },
      projects: proyectos,
      alerts: alertas
    },
    problemas: problemas
  };
}

// ------------------------------------------------- avisos de coherencia
/**
 * Contradicciones que la estructura no detecta pero que ensucian el tablero.
 * No impiden publicar: avisan.
 */
function avisosDeCoherencia(proyectos) {
  const avisos = [];

  proyectos.forEach(function (p) {
    const vistos = {};
    p.lotes.forEach(function (l) {
      const donde = p.nombre + ' mza ' + l.mza + ' lote ' + l.lote;

      // el mismo lote fisico dos veces: fue el caso de GUIDXILAYÚ
      const clave = norm(l.mza) + '|' + norm(l.lote);
      if (vistos[clave]) avisos.push('Repetido: ' + donde);
      vistos[clave] = true;

      if (l.estatus === 'disponible' && l.cliente && !l.nota) {
        avisos.push('Disponible pero con cliente: ' + donde + ' — ' + l.cliente);
      }
      if (l.estatus === 'disponible' && l.ingreso) {
        avisos.push('Disponible pero con ingreso: ' + donde +
                    ' — $' + l.ingreso.toLocaleString('es-MX'));
      }
      if (esSocio(l.cliente) && l.estatus !== 'apartado') {
        avisos.push('Socio sin marcar APARTADO: ' + donde + ' — ' + l.cliente);
      }
    });
  });

  return avisos;
}

function textoAvisos(avisos) {
  if (!avisos.length) return '';
  let t = '\n\nAVISOS DE COHERENCIA (' + avisos.length + '):\n' +
          avisos.slice(0, 15).join('\n');
  if (avisos.length > 15) t += '\n  ...y ' + (avisos.length - 15) + ' más';
  return t;
}

// -------------------------------------------- normalizar lotes de socios
/**
 * Pone en APARTADO los lotes cuyo CLIENTE es exactamente un socio.
 * Pide confirmacion y dice cuantos va a tocar antes de escribir.
 */
function normalizarApartadosDeSocios() {
  const ui = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  if (!hoja) { ui.alert('No encuentro la hoja ' + HOJA); return; }

  const filas = hoja.getDataRange().getValues();
  let iCab = -1;
  for (let i = 0; i < Math.min(filas.length, 10); i++) {
    if (norm(filas[i][0]) === 'PROYECTO') { iCab = i; break; }
  }
  if (iCab < 0) { ui.alert('No encuentro el encabezado'); return; }

  const cab = filas[iCab].map(norm);
  const jCliente = cab.indexOf('CLIENTE');
  const jEstatus = cab.indexOf('ESTATUS');
  if (jCliente < 0 || jEstatus < 0) { ui.alert('Faltan columnas CLIENTE o ESTATUS'); return; }

  const cambios = [];
  for (let i = iCab + 1; i < filas.length; i++) {
    if (!esSocio(filas[i][jCliente])) continue;
    if (norm(filas[i][jEstatus]) === 'APARTADO') continue;
    cambios.push({ fila: i + 1, de: texto(filas[i][jEstatus]),
                   proyecto: texto(filas[i][0]), mza: texto(filas[i][1]), lote: texto(filas[i][2]) });
  }

  if (!cambios.length) {
    ui.alert('Nada que cambiar', 'Todos los lotes de socios ya están en APARTADO.',
             ui.ButtonSet.OK);
    return;
  }

  const detalle = cambios.slice(0, 12).map(function (c) {
    return '  ' + c.proyecto + ' mza ' + c.mza + ' lote ' + c.lote + '  (' + c.de + ')';
  }).join('\n');

  const ok = ui.alert('Marcar como APARTADO',
    'Se van a cambiar ' + cambios.length + ' lote(s) a APARTADO:\n\n' + detalle +
    (cambios.length > 12 ? '\n  ...y ' + (cambios.length - 12) + ' más' : '') +
    '\n\nSolo se toca la columna ESTATUS. ¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  cambios.forEach(function (c) {
    hoja.getRange(c.fila, jEstatus + 1).setValue('APARTADO');
  });

  ui.alert('Listo',
    cambios.length + ' lote(s) marcados como APARTADO.\n\n' +
    'Revisa el RESUMEN y publica cuando estés conforme.', ui.ButtonSet.OK);
}

// ------------------------------------------------------- traer los m2
/**
 * Rellena la columna M2 leyendo m2.json del repositorio, que se genera desde
 * las bases originales. Empareja por PROYECTO + MANZANA + LOTE, asi que no
 * importa el orden de las filas ni las ediciones que se hayan hecho.
 *
 * Solo escribe donde la celda esta vacia: nunca pisa un m2 capturado a mano.
 */
function traerTipoDeVenta() {
  traerColumna('tipo_venta.json', 'TIPO VENTA', 'Tipo de venta', 130, 'CLIENTE');
}

function traerMetrosCuadrados() {
  traerColumna('m2.json', 'M2', 'Superficies', 90, 'LOTE');
}

/**
 * Deja la columna en su sitio y devuelve su indice (0-based).
 *
 * Se usa insertColumnBefore / moveColumns y no un simple setValue al final
 * porque la hoja RESUMEN apunta a LOTES por letra de columna: al insertar de
 * esta forma, Google Sheets reajusta esas formulas solo. Escribir el
 * encabezado a mano las dejaria apuntando a la columna equivocada.
 */
function ubicarColumna(hoja, filaCab, cab, encabezado, despuesDe, ancho) {
  const destino = cab.indexOf(despuesDe) + 1;      // justo despues de esa columna
  let j = cab.indexOf(encabezado);

  if (j < 0) {
    hoja.insertColumnBefore(destino + 1);
    j = destino;
  } else if (j !== destino) {
    const col = hoja.getRange(1, j + 1, hoja.getMaxRows(), 1);
    hoja.moveColumns(col, j + 1 > destino + 1 ? destino + 1 : destino + 2);
    j = destino;
  } else {
    return j;                                      // ya estaba en su lugar
  }

  hoja.getRange(filaCab, j + 1).setValue(encabezado)
      .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
  hoja.setColumnWidth(j + 1, ancho);
  return j;
}

/**
 * Rellena una columna leyendo un json del repositorio, generado desde las
 * bases originales. Empareja por PROYECTO + MANZANA + LOTE, asi que no importa
 * el orden de las filas ni las ediciones hechas a mano.
 *
 * Solo escribe donde la celda esta vacia: nunca pisa un dato ya capturado.
 */
function traerColumna(archivo, encabezado, titulo, ancho, despuesDe) {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { ui.alert('Falta el token', 'Usa: Dashboard > Configurar token', ui.ButtonSet.OK); return; }

  let mapa;
  try { mapa = JSON.parse(traerDeGitHub(archivo, token)); }
  catch (e) { ui.alert('No pude leer ' + archivo, String(e.message), ui.ButtonSet.OK); return; }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  let filas = hoja.getDataRange().getValues();

  let iCab = -1;
  for (let i = 0; i < Math.min(filas.length, 10); i++) {
    if (norm(filas[i][0]) === 'PROYECTO') { iCab = i; break; }
  }
  if (iCab < 0) { ui.alert('No encuentro el encabezado'); return; }

  const jCol = ubicarColumna(hoja, iCab + 1, filas[iCab].map(norm),
                             encabezado, despuesDe, ancho);

  filas = hoja.getDataRange().getValues();   // releer: las columnas se movieron
  const cab = filas[iCab].map(norm);
  const jProy = cab.indexOf('PROYECTO'), jMza = cab.indexOf('MANZANA'), jLote = cab.indexOf('LOTE');

  let puestos = 0, yaTenian = 0, sinDato = 0;
  const faltantes = {}, escribir = [];

  for (let i = iCab + 1; i < filas.length; i++) {
    const proyecto = texto(filas[i][jProy]);
    const lote = texto(filas[i][jLote]);
    if (!proyecto || !lote) continue;

    if (jCol < filas[i].length && texto(filas[i][jCol])) { yaTenian++; continue; }

    const v = (mapa[proyecto] || {})[texto(filas[i][jMza]) + '|' + lote];
    if (v) { escribir.push([i + 1, v]); puestos++; }
    else { sinDato++; faltantes[proyecto] = (faltantes[proyecto] || 0) + 1; }
  }

  escribir.forEach(function (e) { hoja.getRange(e[0], jCol + 1).setValue(e[1]); });

  const lineas = [
    titulo + ' desde las bases:', '',
    '  Rellenados      ' + puestos,
    '  Ya tenían dato  ' + yaTenian,
    '  Sin dato        ' + sinDato
  ];
  const pend = Object.keys(faltantes).sort(function (a, b) { return faltantes[b] - faltantes[a]; });
  if (pend.length) {
    lineas.push('', 'Proyectos con lotes sin dato (hay que capturarlos a mano):');
    pend.slice(0, 10).forEach(function (p) { lineas.push('  ' + p + ': ' + faltantes[p]); });
  }
  lineas.push('', 'No se pisó nada de lo que ya estaba capturado.');
  const msg = lineas.join(String.fromCharCode(10));
  ui.alert('Listo', msg, ui.ButtonSet.OK);
}


// ------------------------------------------------------------- revisar
function revisar() {
  const ui = SpreadsheetApp.getUi();
  let r;
  try { r = construirDatos(); }
  catch (e) { ui.alert('No pude leer la hoja', String(e.message), ui.ButtonSet.OK); return; }

  const g = r.datos.grand;
  let msg = 'Así quedaría el dashboard:\n\n' +
    '  Lotes totales   ' + g.total + '\n' +
    '  Disponibles     ' + g.available + '\n' +
    '  Apartados       ' + g.reserved + '\n' +
    '  Vendidos        ' + g.sold + '\n' +
    '  Sin ingreso     ' + g.sold_no_data + '\n' +
    '  Ingresos        $' + g.income.toLocaleString('es-MX') + '\n\n' +
    '  Proyectos       ' + r.datos.projects.length;

  if (r.problemas.length) {
    msg += '\n\nFILAS CON PROBLEMA (' + r.problemas.length + '), no se incluyen:\n' +
           r.problemas.slice(0, 12).join('\n');
    if (r.problemas.length > 12) msg += '\n  ...y ' + (r.problemas.length - 12) + ' más';
  }
  msg += textoAvisos(avisosDeCoherencia(r.datos.projects));
  ui.alert('Revisión', msg, ui.ButtonSet.OK);
}

// ------------------------------------------------------------ publicar
function publicar() {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    ui.alert('Falta el token', 'Usa primero: Dashboard > Configurar token de GitHub',
             ui.ButtonSet.OK);
    return;
  }

  let r;
  try { r = construirDatos(); }
  catch (e) { ui.alert('No pude leer la hoja', String(e.message), ui.ButtonSet.OK); return; }

  if (r.problemas.length) {
    const seguir = ui.alert(
      'Hay ' + r.problemas.length + ' fila(s) con problema',
      r.problemas.slice(0, 10).join('\n') +
      '\n\nEsas filas NO se van a incluir. ¿Publico de todos modos?',
      ui.ButtonSet.YES_NO);
    if (seguir !== ui.Button.YES) return;
  }

  const g = r.datos.grand;
  const ok = ui.alert('Confirmar publicación',
    'Se va a publicar en el dashboard vivo:\n\n' +
    '  ' + g.total + ' lotes · ' + g.available + ' disponibles · ' +
    g.reserved + ' apartados · ' + g.sold + ' vendidos\n' +
    '  Ingresos: $' + g.income.toLocaleString('es-MX') +
    textoAvisos(avisosDeCoherencia(r.datos.projects)) + '\n\n¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  try {
    const json = JSON.stringify(r.datos);
    const plantilla = traerDeGitHub('scripts/plantilla.html', token);
    const ini = '/*__DATOS_AQUI__*/', fin = '/*__FIN__*/';
    const a = plantilla.indexOf(ini), b = plantilla.indexOf(fin);
    if (a < 0 || b < 0) throw new Error('La plantilla no trae los marcadores de datos');
    const html = plantilla.slice(0, a + ini.length) +
                 json.replace(/<\//g, '<\\/') +
                 plantilla.slice(b);

    const sello = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    subirAGitHub('datos.json', json, 'Actualiza datos desde Sheets (' + sello + ')', token);
    subirAGitHub('index.html', html, 'Actualiza dashboard desde Sheets (' + sello + ')', token);

    ui.alert('Publicado',
      'Listo. El dashboard se actualiza en aproximadamente un minuto:\n\n' +
      URL_DASHBOARD + '\n\nSi no ves el cambio, recarga con Ctrl+F5.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('No se pudo publicar', String(e.message), ui.ButtonSet.OK);
  }
}

// --------------------------------------------------------- GitHub API
function encabezados(token) {
  return {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

function traerDeGitHub(ruta, token) {
  const url = 'https://api.github.com/repos/' + REPO + '/contents/' +
              encodeURI(ruta) + '?ref=' + RAMA;
  const res = UrlFetchApp.fetch(url, { headers: encabezados(token), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('No pude leer ' + ruta + ' de GitHub (código ' +
                    res.getResponseCode() + '). ¿El token tiene acceso al repositorio?');
  }
  const j = JSON.parse(res.getContentText());
  return Utilities.newBlob(Utilities.base64Decode(j.content.replace(/\n/g, ''))).getDataAsString('UTF-8');
}

function subirAGitHub(ruta, contenido, mensaje, token) {
  const url = 'https://api.github.com/repos/' + REPO + '/contents/' + encodeURI(ruta);

  let sha = null;
  const previo = UrlFetchApp.fetch(url + '?ref=' + RAMA,
    { headers: encabezados(token), muteHttpExceptions: true });
  if (previo.getResponseCode() === 200) sha = JSON.parse(previo.getContentText()).sha;

  const cuerpo = {
    message: mensaje,
    content: Utilities.base64Encode(contenido, Utilities.Charset.UTF_8),
    branch: RAMA
  };
  if (sha) cuerpo.sha = sha;

  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: encabezados(token),
    contentType: 'application/json',
    payload: JSON.stringify(cuerpo),
    muteHttpExceptions: true
  });
  const codigo = res.getResponseCode();
  if (codigo !== 200 && codigo !== 201) {
    throw new Error('GitHub rechazó ' + ruta + ' (código ' + codigo + '): ' +
                    res.getContentText().slice(0, 300));
  }
}
