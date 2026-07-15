/**
 * ============================================================
 * CALZADO MARTINI — Asistente Virtual con Groq API + RAG léxico ligero
 *  Google Apps Script  ·  v2.2.0  ·  2026
 *  PROTOTIPO EVALUADO en la prueba piloto (n = 5)
 * ============================================================
 */

// ============================================================
//  CONFIGURACIÓN — Completa estos valores antes de implementar
// ============================================================

const GROQ_API_KEY = 'TU_CLAVE_GROQ_AQUI'; // https://console.groq.com/keys
const GROQ_MODEL   = 'llama-3.3-70b-versatile';

// ⬇ ID de tu Google Spreadsheet — cópialo de la URL:
// https://docs.google.com/spreadsheets/d/[ESTE_ID]/edit
const SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI';

// Hojas del Google Sheet
const SHEET_CATALOG = 'Productos'; // nombre de tu hoja de catálogo
const SHEET_ORDERS  = 'Pedidos';   // se crea si no existe

// Carpeta de Google Drive para guardar comprobantes
const DRIVE_FOLDER_NAME = 'Comprobantes Martini';

// ⬇ Datos bancarios para transferencias — actualiza con tu cuenta real
const BANK_INFO = {
  banco:    'Banco Pichincha',
  cuenta:   '2200123456',
  tipo:     'Cuenta de Ahorros',
  titular:  'Calzado Martini',
  ruc:      '1803456789001',
  concepto: 'Pedido CM - [tu nombre]'
};

// ============================================================
//  PUNTO DE ENTRADA — POST
// ============================================================

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || 'chat';

    if (action === 'chat') {
      return handleChat(data);
    }
    if (action === 'save_order') {
      return handleSaveOrder(data);
    }
    return jsonResp({ error: 'Acción no reconocida: ' + action });

  } catch (err) {
    return jsonResp({ error: 'Error interno: ' + err.message });
  }
}

// Permite verificar que el script está activo
function doGet() {
  return jsonResp({ status: 'Asistente Martini activo', version: '2.0' });
}

// ============================================================
//  CHAT CON RAG — obtiene catálogo, busca relevantes, llama Groq
// ============================================================

function handleChat(data) {
  const message = (data.message || '').trim();
  const history = data.history || [];

  if (!message) {
    return jsonResp({ error: 'Mensaje vacío' });
  }

  // 1. RAG: cargar catálogo y buscar productos relevantes
  const allProducts = getCatalog();

  // Enriquecer query si el usuario usa referencia anafórica o pide clarificación
  const contextualPattern = /\b(ese|este|esa|el mismo|ese modelo|ese zapato|ese par|quiero ese|quiero esa|me llevo|me interesa ese|cuánto cuesta ese|no entiendo|no entiendo eso|explica|explícame|qué quieres decir|repite|o sea|osea|cómo así|y ese|y esa|y el otro|y la otra)\b/i;
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

  const relevant = searchCatalog(allProducts, searchQuery, history);

  // 2. Construir system prompt con contexto del catálogo
  const systemPrompt = buildSystemPrompt(relevant, allProducts);

  // 3. Llamar Groq API
  const reply = callGroq(systemPrompt, history, message);

  return jsonResp({ response: reply, ok: true });
}

// ============================================================
//  GUARDAR PEDIDO — Google Sheets + imagen en Google Drive
// ============================================================

function handleSaveOrder(data) {
  const order    = data.order    || {};  // { product, size, color, price }
  const customer = data.customer || {};  // { name, cedula, email, phone }
  const imageB64 = data.image    || '';  // base64 del comprobante (puede ser vacío)
  const summary  = data.summary  || '';

  let driveUrl = '';

  // --- Guardar comprobante en Drive ---
  if (imageB64 && imageB64.length > 100) {
    try {
      const folder   = getOrCreateFolder(DRIVE_FOLDER_NAME);
      const rawB64   = imageB64.includes(',') ? imageB64.split(',')[1] : imageB64;
      const blob     = Utilities.newBlob(
        Utilities.base64Decode(rawB64),
        'image/jpeg',
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

  // --- Guardar en Google Sheets ---
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_ORDERS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ORDERS);
    sheet.appendRow([
      'Fecha y hora', 'Estado',
      'Nombre cliente', 'Cédula/CI', 'Correo', 'Teléfono',
      'Producto', 'Talla', 'Color', 'Precio (USD)',
      'Comprobante (Drive)', 'Resumen conversación'
    ]);
    // Formatear encabezados
    const headerRange = sheet.getRange(1, 1, 1, 12);
    headerRange.setBackground('#1a2d5a');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toLocaleString('es-EC'),
    'Pendiente de envío',
    customer.name   || '',
    customer.cedula || '',
    customer.email  || '',
    customer.phone  || '',
    order.product   || '',
    order.size      || '',
    order.color     || '',
    order.price     || '',
    driveUrl,
    summary
  ]);

  return jsonResp({ ok: true, driveUrl: driveUrl });
}

// ============================================================
//  RAG — CARGA DEL CATÁLOGO DESDE GOOGLE SHEETS
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
//  RAG — BÚSQUEDA SEMÁNTICA LIGERA (keyword + dominio)
// ============================================================

function searchCatalog(products, query, history) {
  // Solo mensajes del USUARIO en historial — evita que las respuestas del asistente contaminen la búsqueda
  const histText = (history || [])
    .filter(m => m.role === 'user')
    .slice(-3)
    .map(h => h.content || '')
    .join(' ');
  const fullQuery = (query + ' ' + histText).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');

  // ── Detección de rango de precio ──────────────────────────────
  const priceMatch = fullQuery.match(/(?:menos de|hasta|maximo|por)\s*\$?\s*(\d+)/);
  const maxPrice = priceMatch ? parseFloat(priceMatch[1]) : null;

  // ── Detección de talla ────────────────────────────────────────
  const sizeMatch = fullQuery.match(/talla\s*(\d{1,2})/);
  const targetSize = sizeMatch ? sizeMatch[1] : null;

  // ── Mapa de sinónimos extendido del dominio calzado ───────────
  const synonyms = {
    dama:      ['dama','mujer','senora','femenino','ella','chica','nina','damas','mujeres','femenina'],
    caballero: ['caballero','hombre','senor','masculino','chico','nino','caballeros','hombres','masculina'],
    formal:    ['formal','oficina','trabajo','ejecutivo','traje','evento','corporativo','profesional','elegante'],
    casual:    ['casual','diario','comodo','comoda','diaria','cotidiano','sport','relajado','everyday'],
    fiesta:    ['fiesta','boda','gala','noche','evento','celebracion','graduacion','quincea','prom'],
    deporte:   ['deporte','deportivo','correr','gym','gimnasio','atletismo','running','tenis','activo'],
    cuero:     ['cuero','piel','leather','genuino','natural','bovino','vacuno'],
    charol:    ['charol','barniz','brillante','lustrado','lacado'],
    gamuza:    ['gamuza','ante','nubuck','suede','aterciopelado'],
    sintetico: ['sintetico','pu','poliuretano','vegano','ecologico','sintetica','artificial'],
    negro:     ['negro','black','oscuro','negra','negros'],
    cafe:      ['cafe','marron','castano','brown','chocolate','camel','tabaco','havana'],
    blanco:    ['blanco','white','crema','beige','marfil','ivory','blanca'],
    azul:      ['azul','blue','marino','navy','celeste','cobalto'],
    rojo:      ['rojo','red','vino','bordo','guinda','carmesi','roja','rojos'],
    gris:      ['gris','grey','gray','plomo','plateado','grafito'],
    verde:     ['verde','green','militar','oliva','bosque','menta'],
  };

  const scored = products.map(p => {
    let score = 0;
    const pText = Object.values(p).join(' ').toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');

    // Filtro de precio: excluir productos sobre el máximo solicitado
    if (maxPrice !== null) {
      const productPrice = parseFloat(p.precio || p.price || 9999);
      if (productPrice > maxPrice) return { ...p, _score: -1 };
      score += 3;
    }

    // Boost por talla exacta disponible
    if (targetSize && pText.includes(targetSize)) {
      score += 5;
    }

    // Coincidencia exacta por ID de producto
    const productId = (p.id || '').toString().toLowerCase();
    if (productId && fullQuery.includes(productId)) {
      score += 10;
    }

    // Coincidencia directa con palabras clave (>= 3 chars, incluye siglas cortas)
    fullQuery.split(/\s+/).filter(w => w.length >= 3).forEach(kw => {
      if (pText.includes(kw)) score += 2;
    });

    // Coincidencia semántica por sinónimos de dominio
    Object.entries(synonyms).forEach(([key, variants]) => {
      const queryHasVariant = variants.some(v => fullQuery.includes(v));
      if (queryHasVariant && pText.includes(key)) score += 4;
    });

    return { ...p, _score: score };
  });

  scored.sort((a, b) => b._score - a._score);

  const withScore = scored.filter(p => p._score > 0);
  return withScore.length >= 2
    ? withScore.slice(0, 5)
    : scored.filter(p => p._score >= 0).slice(0, 4);
}

// ============================================================
//  SYSTEM PROMPT CON CONTEXTO RAG
// ============================================================

function buildSystemPrompt(relevant, allProducts) {
  const productsCtx = relevant.length > 0
    ? relevant.map(p =>
        `• ${p.nombre || p.name || ''}` +
        ` (${p.categoria || p.cat || ''} · ${p.genero || p.gen || ''})` +
        ` | ${p.material || p.mat || ''}` +
        ` | $${p.precio || p.price || '?'} USD` +
        ` | Tallas: ${p.tallas || p.sizes || ''}` +
        ` | Colores: ${p.colores || p.colors || ''}` +
        `\n  Descripción: ${p.descripcion || p.description || ''}`
      ).join('\n\n')
    : '— Sin productos específicos para esta consulta. Orienta al cliente con preguntas. —';

  const bankText =
    `- Banco: ${BANK_INFO.banco}\n` +
    `- Nro. Cuenta: ${BANK_INFO.cuenta} (${BANK_INFO.tipo})\n` +
    `- Titular: ${BANK_INFO.titular}\n` +
    `- RUC/CI: ${BANK_INFO.ruc}\n` +
    `- Concepto: "${BANK_INFO.concepto}"`;

  return `Eres Martín, asesor personal de Calzado Martini — empresa artesanal ecuatoriana con 20 años de trayectoria. Eres experto en calzado y en ventas consultivas. Tu estilo es cálido, natural y profesional: como un asesor de boutique, no un catálogo.

═══════════════════════════════════
FASE 1 — ENTENDER LA NECESIDAD
═══════════════════════════════════
Si el cliente no especificó qué necesita, haz UNA sola pregunta abierta antes de sugerir nada:
- "¿Para qué ocasión buscas el calzado?" o "¿Es para ti o de regalo?"
- Luego afina con UNA pregunta más si hace falta (género, estilo, presupuesto).
- NO presentes productos hasta tener al menos DOS datos de necesidad del cliente.

═══════════════════════════════════
FASE 2 — PRESENTAR CON EMOCIÓN
═══════════════════════════════════
- Presenta UN producto a la vez. Máximo 2-3 oraciones.
- Primera presentación: nombre + atributo clave + frase sensorial o emocional breve.
  Ejemplo: "Te presento la Sandalia Plataforma Plateada, en cuerina premium con brillo elegante. Perfecta para destacar en cualquier evento."
- Tallas, colores y precio: solo cuando el cliente los pida o confirme interés.
- Usa el campo "Descripción" del producto para crear imagen mental: "imagina llegar con...", "está hecho para...", "es el modelo ideal si..."
- Puedes usar validación natural: "es uno de los favoritos para eventos" o "tiene muy buena acogida".

═══════════════════════════════════
FASE 3 — CIERRE NATURAL
═══════════════════════════════════
- Cuando el cliente muestre interés, pregunta: "¿Te gustaría conocer tallas y colores disponibles?"
- Cuando confirme, pregunta: "¿Quieres que te ayude a reservarlo?" — entonces inicia el flujo de compra.

═══════════════════════════════════
REGLAS CRÍTICAS
═══════════════════════════════════
1. NUNCA incluyas códigos de producto [DAM-001] en tus respuestas. Son solo internos.
2. NUNCA niegues tener un producto que mencionaste en este chat. Si el cliente pregunta por él, sigue con ese producto.
3. Si el cliente dice "no entiendo", "explica" o "¿qué?", simplifica o reformula lo último que dijiste — NO cambies de producto.
4. NUNCA inventes datos. Usa solo la información del catálogo proporcionado.
5. Si el cliente menciona un presupuesto, muestra solo productos en ese rango.
6. Si el cliente menciona una talla, prioriza productos que la incluyan.

═══════════════════════════════════
CATÁLOGO RELEVANTE PARA ESTA CONSULTA
═══════════════════════════════════
${productsCtx}

═══════════════════════════════════
PROCESO DE COMPRA — SIGUE ESTE FLUJO EXACTO
═══════════════════════════════════
Cuando el cliente indique que desea comprar un producto:

PASO 1 — Confirma con el cliente: producto, talla y color elegidos.
PASO 2 — Cuando el cliente confirme, proporciona los datos de transferencia bancaria e incluye al final de tu mensaje exactamente: [SOLICITAR_PAGO]
Datos bancarios:
${bankText}

PASO 3 — Cuando el cliente informe que realizó la transferencia, pide el comprobante e incluye al final: [SOLICITAR_COMPROBANTE]
PASO 4 — Una vez recibido el comprobante, solicita los datos personales del cliente e incluye al final: [SOLICITAR_DATOS]
  Datos requeridos: nombre completo, número de cédula/CI, correo electrónico, número de celular.
PASO 5 — Al recibir los datos del cliente, confirma el pedido e incluye al final: [PEDIDO_REGISTRADO]
  Indica que recibirá una notificación cuando su pedido sea despachado.

IMPORTANTE: Usa los marcadores [SOLICITAR_PAGO], [SOLICITAR_COMPROBANTE], [SOLICITAR_DATOS] y [PEDIDO_REGISTRADO] ÚNICAMENTE cuando corresponda al paso exacto del flujo. No los repitas ni los uses fuera de contexto.`;
}

// ============================================================
//  LLAMADA A LA API DE GROQ (compatible con OpenAI)
// ============================================================

function callGroq(systemPrompt, history, message) {
  // Limpiar historial: solo roles user/assistant, sin tokens internos
  const cleanHistory = (history || [])
    .slice(-8)
    .filter(m => m.role && m.content)
    .map(m => ({
      role:    m.role === 'user' ? 'user' : 'assistant',
      content: m.content
        .replace(/\[SOLICITAR_PAGO\]/g, '')
        .replace(/\[SOLICITAR_COMPROBANTE\]/g, '')
        .replace(/\[SOLICITAR_DATOS\]/g, '')
        .replace(/\[PEDIDO_REGISTRADO\]/g, '')
        .trim()
    }))
    .filter(m => m.content.length > 0);

  // Groq usa el formato OpenAI: system como primer mensaje del array
  const messages = [
    { role: 'system', content: systemPrompt },
    ...cleanHistory,
    { role: 'user', content: message }
  ];

  const payload = {
    model:      GROQ_MODEL,
    max_tokens: 1024,
    messages:   messages
  };

  const options = {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'Authorization': 'Bearer ' + GROQ_API_KEY
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp   = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', options);
  const result = JSON.parse(resp.getContentText());

  if (result.error) {
    throw new Error('Groq API error: ' + result.error.message);
  }
  if (!result.choices || !result.choices[0]) {
    throw new Error('Respuesta vacía de Groq API');
  }

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
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
