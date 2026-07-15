/**
 * ============================================================
 *  CALZADO MARTINI — Asistente Virtual con Groq API + RAG Semántico
 *  Google Apps Script  ·  v3.0  ·  2026
 *  EVOLUCIÓN TÉCNICA — no validada en la prueba piloto
 * ============================================================
 * *   Completar COHERE_API_KEY, GROQ_API_KEY y SPREADSHEET_ID
 * ============================================================
 */

const GROQ_API_KEY   = '_CLAVE_GROQ_';    // https://console.groq.com/keys
const GROQ_MODEL     = 'llama-3.3-70b-versatile';

const COHERE_API_KEY = '_CLAVE_COHERE_';   // https://dashboard.cohere.com/api-keys  (gratis)
const COHERE_MODEL   = 'embed-multilingual-v3.0'; // soporta español nativo, 1024 dimensiones

// ID de tu Google Spreadsheet — del URL:
// https://docs.google.com/spreadsheets/d/[ESTE_ID]/edit

const SPREADSHEET_ID = '_SPREADSHEET_ID_';

const SHEET_CATALOG     = 'Productos';
const SHEET_ORDERS      = 'Pedidos';
const DRIVE_FOLDER_NAME = 'Comprobantes Martini';
const OWNER_EMAIL       = '_OWNER_EMAIL_'; // Correo del propietario para notificaciones

// Clave de caché para embeddings del catálogo (se regenera cada 6 horas)
const EMBED_CACHE_KEY = 'cm_catalog_embeddings_v1';
const EMBED_CACHE_TTL = 21600; // segundos (6 horas)

const BANK_INFO = {
  banco:    'Banco Pichincha',
  cuenta:   '2200123456',
  tipo:     'Cuenta de Ahorros',
  titular:  'Calzado Martini',
  ruc:      '1803456789001',
  concepto: 'Pedido CM - [tu nombre]'
};

// ============================================================
//  PUNTO DE ENTRADA — POST / GET
// ============================================================

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || 'chat';
    if (action === 'chat')       return handleChat(data);
    if (action === 'save_order') return handleSaveOrder(data);
    if (action === 'feedback')   return handleFeedback(data);
    return jsonResp({ error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResp({ error: 'Error interno: ' + err.message });
  }
}

function doGet() {
  return jsonResp({ status: 'Asistente Martini activo', version: '2.2.0' });
}

// ============================================================
//  CHAT CON RAG SEMÁNTICO
// ============================================================

function handleChat(data) {
  const message = (data.message || '').trim();
  const history = data.history || [];

  if (!message) return jsonResp({ error: 'Mensaje vacío' });

  const allProducts = getCatalog();

  // Enriquecer query si el usuario usa referencia anafórica ("ese", "este", etc.)
  const contextualPattern = /\b(ese|este|esa|el mismo|ese modelo|ese zapato|ese par|quiero ese|quiero esa|me llevo|me interesa ese|cuánto cuesta ese|no entiendo|explica|explícame|qué quieres decir|repite|o sea|osea|cómo así|y ese|y esa|y el otro|y la otra)\b/i;
  let searchQuery = message;
  if (contextualPattern.test(message) || message.trim().length < 15) {
    const recentBot = history.filter(m => m.role === 'assistant').slice(-2).map(m => m.content).join(' ');
    allProducts.forEach(p => {
      const name = (p.nombre || p.name || '').toLowerCase();
      if (name && recentBot.toLowerCase().includes(name)) {
        searchQuery += ' ' + name;
      }
    });
  }

  // RAG semántico (Cohere embeddings) con fallback a keywords
  const relevant = searchCatalogSemantic(allProducts, searchQuery, history);

  const systemPrompt = buildSystemPrompt(relevant, history);
  const reply        = callGroq(systemPrompt, history, message);

  return jsonResp({ response: reply, ok: true });
}

// ============================================================
//  GUARDAR PEDIDO — Google Sheets + Google Drive
// ============================================================

function handleSaveOrder(data) {
  const order    = data.order    || {};
  const customer = data.customer || {};
  const imageB64 = data.image    || '';
  const summary  = data.summary  || '';

  // ── 1. Guardar comprobante de pago en Drive ──
  let driveUrl = '';
  if (imageB64 && imageB64.length > 100) {
    try {
      const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
      const rawB64 = imageB64.includes(',') ? imageB64.split(',')[1] : imageB64;
      const blob   = Utilities.newBlob(
        Utilities.base64Decode(rawB64), 'image/jpeg',
        'comprobante_' + (customer.cedula || 'sin-cedula') + '_' + Date.now() + '.jpg'
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveUrl = file.getUrl();
    } catch (imgErr) {
      Logger.log('Error al guardar imagen: ' + imgErr.message);
      driveUrl = 'Error al guardar: ' + imgErr.message;
    }
  }

  // ── 2. Guardar conversación completa como .txt en Drive ──
  let convUrl = '';
  if (summary && summary.length > 10) {
    try {
      const folder  = getOrCreateFolder(DRIVE_FOLDER_NAME);
      const txtBlob = Utilities.newBlob(
        summary, 'text/plain',
        'conversacion_' + (customer.cedula || 'sin-cedula') + '_' + Date.now() + '.txt'
      );
      const txtFile = folder.createFile(txtBlob);
      txtFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      convUrl = txtFile.getUrl();
    } catch (txtErr) {
      Logger.log('Error al guardar conversación: ' + txtErr.message);
    }
  }

  // ── 3. Consolidar productos del pedido (soporta múltiples) ──
  const items    = (order.items && order.items.length > 0)
    ? order.items
    : [{ product: order.product, size: order.size, color: order.color, price: order.price }];
  const prodStr  = items.map(i => i.product || '').filter(Boolean).join(' + ') || '';
  const tallaStr = items.map(i => i.size    || '').filter(Boolean).join(' / ') || '';
  const colorStr = items.map(i => i.color   || '').filter(Boolean).join(' / ') || '';
  const total    = items.reduce((s, i) => s + parseFloat(i.price || 0), 0).toFixed(2);

  // ── 4. Registrar en Google Sheets ──
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_ORDERS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ORDERS);
    sheet.appendRow([
      'Fecha y hora', 'Estado',
      'Nombre cliente', 'Cédula/CI', 'Correo', 'Teléfono',
      'Producto', 'Talla', 'Color', 'Precio (USD)',
      'Dirección de envío', 'Comprobante (Drive)', 'Conversación (Drive)'
    ]);
    const hr = sheet.getRange(1, 1, 1, 13);
    hr.setBackground('#1a2d5a').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toLocaleString('es-EC'), 'Pendiente de envío',
    customer.name || '', customer.cedula || '', customer.email || '', customer.phone || '',
    prodStr, tallaStr, colorStr, total,
    customer.direccion || '', driveUrl, convUrl
  ]);

  // ── 5. Notificación por correo al propietario ──
  try {
    const emailLines = [
      '🛍️ NUEVO PEDIDO — Calzado Martini',
      '═══════════════════════════════════',
      '',
      '👤 CLIENTE',
      '  Nombre:    ' + (customer.name    || 'N/D'),
      '  Cédula:    ' + (customer.cedula  || 'N/D'),
      '  Correo:    ' + (customer.email   || 'N/D'),
      '  Celular:   ' + (customer.phone   || 'N/D'),
      '  Dirección: ' + (customer.direccion || 'N/D'),
      '',
      '📦 PEDIDO',
      '  Producto: ' + (prodStr  || 'N/D'),
      '  Talla:    ' + (tallaStr || 'N/D'),
      '  Color:    ' + (colorStr || 'N/D'),
      '  Total:    $' + total + ' USD',
      '',
      '🔗 ARCHIVOS',
      '  Comprobante:   ' + (driveUrl || 'No disponible'),
      '  Conversación:  ' + (convUrl  || 'No disponible'),
      '',
      'Fecha: ' + new Date().toLocaleString('es-EC')
    ].join('\n');

    MailApp.sendEmail({
      to:      OWNER_EMAIL,
      subject: '🛍️ Nuevo pedido — ' + (customer.name || 'Cliente') + ' | $' + total + ' USD',
      body:    emailLines
    });
  } catch (mailErr) {
    Logger.log('Error al enviar email: ' + mailErr.message);
  }

  // ── 6. Inicializar hoja de métricas (primera vez) ──
  setupMetricsSheet(ss);

  return jsonResp({ ok: true, driveUrl: driveUrl });
}

// ============================================================
//  FEEDBACK 👍/👎
// ============================================================

function handleFeedback(data) {
  const message  = data.message  || '';
  const positive = data.positive !== false;

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName('Feedback');
  if (!sheet) {
    sheet = ss.insertSheet('Feedback');
    sheet.appendRow(['Fecha', '¿Positivo?', 'Mensaje del bot']);
    const hr = sheet.getRange(1, 1, 1, 3);
    hr.setBackground('#1a2d5a').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 500);
  }
  sheet.appendRow([new Date().toLocaleString('es-EC'), positive ? 'Sí' : 'No', message]);
  return jsonResp({ ok: true });
}

// ============================================================
//  HOJA DE MÉTRICAS
// ============================================================

function setupMetricsSheet(ss) {
  if (ss.getSheetByName('Métricas')) return; // Ya existe
  const ms = ss.insertSheet('Métricas');

  // Título
  ms.getRange('A1:C1').merge();
  ms.getRange('A1').setValue('PANEL DE MÉTRICAS — ASISTENTE MARTÍN')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#1a2d5a').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Encabezados de tabla
  ms.getRange('A3:C3').setValues([['Métrica', 'Valor', 'Fuente']]);
  ms.getRange('A3:C3').setBackground('#c9a96e').setFontWeight('bold');

  // Filas con fórmulas (se auto-actualizan con cada nuevo pedido)
  const rows = [
    ['Total de pedidos',              '=IFERROR(COUNTA(Pedidos!A:A)-1,0)',                   'Hoja Pedidos'],
    ['Ingresos totales (USD)',         '=IFERROR(SUM(Pedidos!J2:J),0)',                       'Hoja Pedidos'],
    ['Pedidos pendientes de envío',   '=IFERROR(COUNTIF(Pedidos!B:B,"Pendiente de envío"),0)','Hoja Pedidos'],
    ['Fecha último pedido',            '=IFERROR(INDEX(Pedidos!A:A,COUNTA(Pedidos!A:A)),"")', 'Hoja Pedidos'],
    ['', '', ''],
    ['Valoraciones positivas (👍)',    '=IFERROR(COUNTIF(Feedback!B:B,"Sí"),0)',              'Hoja Feedback'],
    ['Valoraciones negativas (👎)',    '=IFERROR(COUNTIF(Feedback!B:B,"No"),0)',              'Hoja Feedback'],
    ['Total valoraciones',             '=IFERROR(COUNTA(Feedback!A:A)-1,0)',                  'Hoja Feedback'],
  ];

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0]) ms.getRange(4 + i, 1, 1, 3).setValues([rows[i]]);
  }

  ms.setColumnWidth(1, 270);
  ms.setColumnWidth(2, 160);
  ms.setColumnWidth(3, 160);
  ms.setFrozenRows(3);
}

// ============================================================
//  RAG — CARGA DEL CATÁLOGO
// ============================================================

function getCatalog() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CATALOG);
    if (!sheet) return [];

    const data    = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => h.toString().toLowerCase().replace(/\s+/g, '_').trim());
    return data.slice(1)
      .filter(row => row.some(cell => cell !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = (row[i] !== null && row[i] !== undefined) ? row[i].toString().trim() : '';
        });
        return obj;
      });
  } catch (err) {
    Logger.log('Error al cargar catálogo: ' + err.message);
    return [];
  }
}

// ============================================================
//  RAG SEMÁNTICO — EMBEDDINGS COHERE + SIMILITUD COSENO
// ============================================================

/**
 * Convierte un producto del catálogo en texto enriquecido para embedding.
 * Incluye todos los atributos relevantes en lenguaje natural.
 */
function buildProductText(p) {
  return [
    p.nombre   || '',
    p.categoria || '',
    p.genero   || '',
    'material: ' + (p.material || ''),
    'colores: '  + (p.colores  || ''),
    'tallas: '   + (p.tallas   || ''),
    'precio: $'  + (p.precio   || '') + ' USD'
  ].filter(Boolean).join('. ');
}

/**
 * Similitud coseno entre dos vectores de igual dimensión.
 * Devuelve un valor entre -1 y 1 (1 = idénticos, 0 = sin relación).
 */
function cosineSimilarity(a, b) {
  let dot = 0, na = 0, nb = 0;
  for (let i = 0; i < a.length; i++) {
    dot += a[i] * b[i];
    na  += a[i] * a[i];
    nb  += b[i] * b[i];
  }
  if (na === 0 || nb === 0) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

/**
 * Llama a la API de embeddings de Cohere.
 * inputType: 'search_document' (para ítems del catálogo)
 *            'search_query'    (para la consulta del usuario)
 * Devuelve array de vectores float[].
 */
function callCohereEmbed(texts, inputType) {
  const payload = {
    model:            COHERE_MODEL,
    texts:            texts,
    input_type:       inputType,
    embedding_types:  ['float']
  };

  const options = {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'Authorization': 'Bearer ' + COHERE_API_KEY,
      'Accept':        'application/json'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp   = UrlFetchApp.fetch('https://api.cohere.com/v2/embed', options);
  const result = JSON.parse(resp.getContentText());

  if (result.message) throw new Error('Cohere error: ' + result.message);
  return result.embeddings.float;
}

/**
 * Obtiene embeddings del catálogo desde caché (CacheService) o los genera.
 * La caché se invalida si el número de productos cambia o tras EMBED_CACHE_TTL segundos.
 */
function getCachedCatalogEmbeddings(products) {
  const cache = CacheService.getScriptCache();
  const raw   = cache.get(EMBED_CACHE_KEY);

  if (raw) {
    try {
      const cached = JSON.parse(raw);
      if (cached.count === products.length) {
        return cached.embeddings;
      }
    } catch (e) {
      // Caché corrupta — regenerar
    }
  }

  const texts      = products.map(buildProductText);
  const embeddings = callCohereEmbed(texts, 'search_document');

  try {
    cache.put(EMBED_CACHE_KEY, JSON.stringify({ count: products.length, embeddings }), EMBED_CACHE_TTL);
  } catch (e) {
    Logger.log('Caché de embeddings no almacenada (tamaño): ' + e.message);
  }

  return embeddings;
}

/**
 * Búsqueda semántica principal.
 * Intenta RAG con embeddings; si Cohere falla, cae a búsqueda por keywords.
 */
function searchCatalogSemantic(products, query, history) {
  if (!products.length) return [];

  // Detectar restricciones estructuradas en la consulta
  const normalized = (query + ' ' + (history || []).filter(m => m.role === 'user').slice(-3).map(m => m.content).join(' '))
    .toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

  const priceMatch = normalized.match(/(?:menos de|hasta|maximo|por)\s*\$?\s*(\d+)/);
  const maxPrice   = priceMatch ? parseFloat(priceMatch[1]) : null;

  const sizeMatch  = normalized.match(/talla\s*(\d{1,2})/);
  const targetSize = sizeMatch ? sizeMatch[1] : null;

  try {
    // ── Obtener embeddings del catálogo (cacheados) ───────────────
    const catalogEmbeddings = getCachedCatalogEmbeddings(products);

    // ── Embedding de la consulta del usuario ─────────────────────
    const [queryEmbedding] = callCohereEmbed([query], 'search_query');

    // ── Puntuar por similitud coseno ──────────────────────────────
    const scored = products.map((p, i) => {
      let score = cosineSimilarity(queryEmbedding, catalogEmbeddings[i]);

      // Penalizar productos fuera del presupuesto indicado
      if (maxPrice !== null) {
        const price = parseFloat(p.precio || 9999);
        if (price > maxPrice) return { ...p, _score: -1 };
        score += 0.08; // ligero boost a productos dentro del rango
      }

      // Boost por talla disponible
      if (targetSize) {
        const pText = Object.values(p).join(' ');
        if (pText.includes(targetSize)) score += 0.12;
      }

      // Boost máximo por coincidencia exacta de ID
      const pid = (p.id || '').toString().toLowerCase();
      if (pid && normalized.includes(pid)) score += 0.5;

      return { ...p, _score: score };
    });

    scored.sort((a, b) => b._score - a._score);
    let results = scored.filter(p => p._score > 0).slice(0, 5);

    // Guard de género: si query pide dama y ningún resultado es dama, forzar inclusión
    const queryNorm = query.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');
    const wantsFemale = /\b(dama|mujer|femenin|señora|chica|ella|damas|mujeres)\b/.test(queryNorm);
    const wantsMale   = /\b(caballero|hombre|masculin|señor|chico|el|caballeros|hombres)\b/.test(queryNorm);

    if (wantsFemale && !results.some(p => (p.genero||'').toLowerCase().includes('dama'))) {
      const female = scored.filter(p => (p.genero||'').toLowerCase().includes('dama')).slice(0, 3);
      results = [...female, ...results.filter(p => !(p.genero||'').toLowerCase().includes('dama'))].slice(0, 5);
    }
    if (wantsMale && !results.some(p => (p.genero||'').toLowerCase().includes('caballero'))) {
      const male = scored.filter(p => (p.genero||'').toLowerCase().includes('caballero')).slice(0, 3);
      results = [...male, ...results.filter(p => !(p.genero||'').toLowerCase().includes('caballero'))].slice(0, 5);
    }

    return results;

  } catch (err) {
    Logger.log('Embedding search falló, usando keywords: ' + err.message);
    return searchCatalogKeyword(products, query, history);
  }
}

// ============================================================
//  RAG FALLBACK — BÚSQUEDA POR PALABRAS CLAVE
//  Se activa automáticamente si Cohere no está disponible.
// ============================================================

function searchCatalogKeyword(products, query, history) {
  const histText  = (history || []).filter(m => m.role === 'user').slice(-3).map(m => m.content || '').join(' ');
  const fullQuery = (query + ' ' + histText).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

  const priceMatch = fullQuery.match(/(?:menos de|hasta|maximo|por)\s*\$?\s*(\d+)/);
  const maxPrice   = priceMatch ? parseFloat(priceMatch[1]) : null;

  const sizeMatch  = fullQuery.match(/talla\s*(\d{1,2})/);
  const targetSize = sizeMatch ? sizeMatch[1] : null;

  const synonyms = {
    dama:      ['dama','mujer','senora','femenino','ella','chica','nina','damas','mujeres','femenina'],
    caballero: ['caballero','hombre','senor','masculino','chico','nino','caballeros','hombres'],
    formal:    ['formal','oficina','trabajo','ejecutivo','traje','evento','corporativo','profesional','elegante'],
    casual:    ['casual','diario','comodo','comoda','diaria','cotidiano','sport','relajado'],
    fiesta:    ['fiesta','boda','gala','noche','evento','celebracion','graduacion','quincea','prom'],
    deporte:   ['deporte','deportivo','correr','gym','gimnasio','atletismo','running'],
    cuero:     ['cuero','piel','leather','genuino','natural','bovino'],
    charol:    ['charol','barniz','brillante','lustrado','lacado'],
    gamuza:    ['gamuza','ante','nubuck','suede'],
    sintetico: ['sintetico','pu','poliuretano','vegano','ecologico','artificial'],
    negro:     ['negro','black','oscuro','negra'],
    cafe:      ['cafe','marron','castano','brown','chocolate','camel','tabaco'],
    blanco:    ['blanco','white','crema','beige','marfil','ivory'],
    azul:      ['azul','blue','marino','navy','celeste'],
    rojo:      ['rojo','red','vino','bordo','guinda','carmesi'],
    gris:      ['gris','grey','gray','plomo','plateado'],
    verde:     ['verde','green','militar','oliva'],
  };

  const scored = products.map(p => {
    let score = 0;
    const pText = Object.values(p).join(' ').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

    if (maxPrice !== null) {
      const price = parseFloat(p.precio || 9999);
      if (price > maxPrice) return { ...p, _score: -1 };
      score += 3;
    }
    if (targetSize && pText.includes(targetSize)) score += 5;

    const pid = (p.id || '').toString().toLowerCase();
    if (pid && fullQuery.includes(pid)) score += 10;

    fullQuery.split(/\s+/).filter(w => w.length >= 3).forEach(kw => {
      if (pText.includes(kw)) score += 2;
    });

    Object.entries(synonyms).forEach(([key, variants]) => {
      if (variants.some(v => fullQuery.includes(v)) && pText.includes(key)) score += 4;
    });

    return { ...p, _score: score };
  });

  scored.sort((a, b) => b._score - a._score);
  const withScore = scored.filter(p => p._score > 0);
  return withScore.length >= 2 ? withScore.slice(0, 5) : scored.filter(p => p._score >= 0).slice(0, 4);
}

// ============================================================
//  MÁQUINA DE ESTADOS — FLUJO DE COMPRA
// ============================================================

function detectPurchaseState(history) {
  if (!history || history.length === 0) return 'VENDIENDO';
  const botAll  = history.filter(m => m.role === 'assistant').map(m => m.content || '').join(' ');
  const userRec = history.filter(m => m.role === 'user').slice(-4).map(m => (m.content || '').toLowerCase()).join(' ');
  if (botAll.includes('[PEDIDO_REGISTRADO]'))      return 'COMPLETADO';
  if (botAll.includes('[SOLICITAR_DATOS]'))        return 'ESPERANDO_DATOS';
  if (botAll.includes('[SOLICITAR_COMPROBANTE]'))  return 'ESPERANDO_COMPROBANTE';
  if (botAll.includes('[SOLICITAR_PAGO]')) {
    const pagado = /\b(pagu[eé]|transfer[ií]|ya pagu|ya transfer|hice el pago|realic[eé]|listo el pago|deposit[eé]|ya deposit|listo|hecho)\b/;
    if (pagado.test(userRec)) return 'PAGO_REALIZADO';
    return 'ESPERANDO_PAGO';
  }
  return 'VENDIENDO';
}

// ============================================================
//  SYSTEM PROMPT CON CONTEXTO RAG
// ============================================================

function buildSystemPrompt(relevant, history) {
  const productsCtx = relevant.length > 0
    ? relevant.map(p =>
        `• ${p.nombre || ''}` +
        ` (${p.categoria || ''} · ${p.genero || ''})` +
        ` | ${p.material || ''}` +
        ` | $${p.precio || '?'} USD` +
        ` | Tallas: ${p.tallas || ''}` +
        ` | Colores: ${p.colores || ''}` +
        ` | Imagen: imagenes/${p.id || ''}.jpg` +
        `\n  Descripción: ${p.descripcion || ''}`
      ).join('\n\n')
    : '— Sin productos específicos. Orienta al cliente con preguntas. —';

  const bankText =
    `- Banco: ${BANK_INFO.banco}\n` +
    `- Nro. Cuenta: ${BANK_INFO.cuenta} (${BANK_INFO.tipo})\n` +
    `- Titular: ${BANK_INFO.titular}\n` +
    `- RUC/CI: ${BANK_INFO.ruc}\n` +
    `- Concepto: "${BANK_INFO.concepto}"`;

  // Estado actual del flujo de compra — inyectado explícitamente en el prompt
  const purchaseState = detectPurchaseState(history || []);
  const stateNotes = {
    VENDIENDO:             '🟢 Asesorando — no hay compra en curso. Aplica el flujo natural de ventas.',
    ESPERANDO_PAGO:        '🟡 CLIENTE DEBE PAGAR — Ya recibió los datos bancarios. Si el cliente dice que ya pagó o transfirió: tu ÚNICA respuesta es pedir el comprobante e incluir [SOLICITAR_COMPROBANTE]. NO preguntes sobre nueva compra ni nueva ocasión.',
    PAGO_REALIZADO:        '🔴 CLIENTE INDICÓ QUE YA PAGÓ — Tu ÚNICO siguiente mensaje: pide el comprobante de pago e incluye [SOLICITAR_COMPROBANTE]. NO digas nada más.',
    ESPERANDO_COMPROBANTE: '🟡 ESPERANDO COMPROBANTE — El cliente lo acaba de enviar. Solicita los datos personales e incluye [SOLICITAR_DATOS].',
    ESPERANDO_DATOS:       '🟡 ESPERANDO DATOS DE CONTACTO — Al recibir nombre/cédula/correo/celular/dirección, confirma el pedido e incluye [PEDIDO_REGISTRADO].',
    COMPLETADO:            '✅ PEDIDO YA REGISTRADO — Solo agradece y despídete. No inicies nueva venta.'
  };
  const stateNote = stateNotes[purchaseState] || stateNotes.VENDIENDO;

  return `Eres Martín, asesor personal de Calzado Martini — empresa artesanal ecuatoriana con 20 años de trayectoria. Vendes calzado para DAMA y CABALLERO. Eres experto en ventas consultivas. Tu estilo: cálido, natural, directo — como asesor de boutique.

═══════════════════════════════════
CATÁLOGO — SIEMPRE TIENES:
═══════════════════════════════════
- Modelos para DAMA (sandalias, pumps, slingbacks, loafers, botines)
- Modelos para CABALLERO (derbys, oxfords, mocasines, chelsea boots, loafers)
- Categorías: formal, casual, fiesta
- NUNCA digas que no tienes modelos para dama o caballero — siempre existen.

═══════════════════════════════════
CÓMO VENDER — FLUJO NATURAL
═══════════════════════════════════
1. Si el cliente no especifica, haz UNA pregunta: "¿Para qué ocasión buscas el calzado?"
2. Con dos datos (género + ocasión o estilo), presenta UN producto: nombre + material + frase corta de imagen.
   → Incluye [IMG:imagenes/ID.jpg] ANTES del texto usando el ID exacto del producto del catálogo.
   → Si no tienes el ID, NO incluyas [IMG:...].
3. NO menciones tallas ni colores hasta que el cliente lo pida. Cuando pida talla, pregunta "¿Qué talla usas?" si no la mencionó. Cuando pida color, pregunta "¿Tienes preferencia de color?" — no listes todas las opciones.
4. Cuando el cliente confirme interés: pregunta SOLO la talla (si no la sabe), luego el color.
5. Una vez confirmado producto + talla + color: proporciona datos bancarios DIRECTAMENTE + [SOLICITAR_PAGO]. No pidas permiso.

═══════════════════════════════════
REGLAS DE RESPUESTA
═══════════════════════════════════
- Máximo 2-3 oraciones por mensaje. Una sola idea.
- Una acción por turno: O preguntas O informas — nunca ambas en el mismo mensaje.
- NUNCA incluyas códigos [DAM-001] en tus respuestas. Son solo internos.
- NUNCA niegues un producto que mencionaste antes en este chat.
- Si el cliente dice "no entiendo": reformula en términos más simples, no cambies de producto.
- NUNCA inventes datos. Solo la información del catálogo proporcionado.

═══════════════════════════════════
CATÁLOGO RELEVANTE PARA ESTA CONSULTA
═══════════════════════════════════
${productsCtx}

ESTADO ACTUAL DE LA COMPRA (lee esto primero):
${stateNote}

═══════════════════════════════════
PROCESO DE COMPRA — SIGUE ESTE FLUJO EXACTO
═══════════════════════════════════
PASO 1 — Confirma producto, talla y color.
PASO 2 — Al confirmar producto + talla + color, escribe EXACTAMENTE en este orden:
  a) Marcador al inicio (línea sola, sin espacios extra):
     [PEDIDO:NombreProducto|Talla|Color|PrecioUSD]
     Varios: [PEDIDO:Prod1|Talla1|Color1|Precio1§Prod2|Talla2|Color2|Precio2]
  b) Anuncia el precio: "El total es $XX USD."
  c) Los datos bancarios completos.
  d) Al final del mensaje: [SOLICITAR_PAGO]
  Ejemplo real: "[PEDIDO:Oxford Brogue Elegante|42|Negro|80] El total es $80 USD. Para proceder, realiza la transferencia a: ${bankText} [SOLICITAR_PAGO]"

PASO 3 — Si el cliente indica que pagó o transfirió → responde SOLO: pide el comprobante + [SOLICITAR_COMPROBANTE]
PASO 4 — Al recibir el comprobante → solicita datos personales e incluye: [SOLICITAR_DATOS]
  Datos: nombre completo, cédula/CI, correo, celular, dirección de envío (o URL de Google Maps).
PASO 5 — Al recibir los datos → confirma el pedido e incluye: [PEDIDO_REGISTRADO]

IMPORTANTE: Usa los marcadores ÚNICAMENTE en el paso exacto. No los repitas.`;
}

// ============================================================
//  LLAMADA A LA API DE GROQ
// ============================================================

function callGroq(systemPrompt, history, message) {
  const cleanHistory = (history || [])
    .slice(-16)
    .filter(m => m.role && m.content)
    .map(m => ({
      role:    m.role === 'user' ? 'user' : 'assistant',
      content: m.content
        .replace(/\[SOLICITAR_PAGO\]/g, '').replace(/\[SOLICITAR_COMPROBANTE\]/g, '')
        .replace(/\[SOLICITAR_DATOS\]/g, '').replace(/\[PEDIDO_REGISTRADO\]/g, '')
        .replace(/\[PEDIDO:[^\]]+\]/g, '').trim()
    }))
    .filter(m => m.content.length > 0);

  const messages = [
    { role: 'system', content: systemPrompt },
    ...cleanHistory,
    { role: 'user', content: message }
  ];

  const resp   = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'Authorization': 'Bearer ' + GROQ_API_KEY },
    payload:            JSON.stringify({ model: GROQ_MODEL, max_tokens: 1024, messages }),
    muteHttpExceptions: true
  });
  const result = JSON.parse(resp.getContentText());

  if (result.error)                      throw new Error('Groq API error: ' + result.error.message);
  if (!result.choices || !result.choices[0]) throw new Error('Respuesta vacía de Groq');

  return result.choices[0].message.content;
}

// ============================================================
//  UTILIDADES
// ============================================================

function getOrCreateFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
