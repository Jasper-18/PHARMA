// api/procesar-sustento-transitos.js
//
// Lectura automática de las fotos de sustento de Tránsitos.
// Mismo patrón que api/procesar-ocr.js de PharmaSPOT (ver
// Documentacion_Tecnica_OCR_PharmaSPOT.md), adaptado a registros_transito:
//
//   1. El frontend ya subió las fotos y guardó la fila (intentos_ia +1,
//      estado_procesamiento_ia = null).
//   2. Llama a este endpoint con la lista de LPN que acaba de sustentar.
//   3. Aquí se descargan las fotos, se lee su texto (OCR) y se busca el N° de
//      LPN de CADA bulto en el texto leído.
//   4. Si lo encuentra (>= 95% de coincidencia): el bulto queda APROBADO con
//      validado_por = 'IA'. Si no: queda PENDIENTE (por validar) para el admin.
//      La IA nunca desaprueba.
//
// Variables de entorno (en Vercel, NUNCA en App.jsx):
//   SUPABASE_SERVICE_KEY        service_role key del Supabase de Tránsitos
//   TRANSITOS_OCRSPACE_API_KEY  API key gratuita de OCR.space (motor actual)
//   TRANSITOS_VISION_API_KEY    API key de Google Cloud Vision (opcional)
//
// Motor de OCR: si existe TRANSITOS_VISION_API_KEY se usa Google Cloud Vision
// (mejor con letra manuscrita); si no, OCR.space (gratis, sin tarjeta). Para
// pasar a Google basta con agregar esa variable en Vercel y redesplegar: no
// hay que tocar código. La comparación del LPN es la misma con ambos motores.

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://zffuccirauheklpxagga.supabase.co";
const UMBRAL_COINCIDENCIA = 95;   // mismo umbral que PharmaSPOT
const MAX_LPNS_POR_LLAMADA = 300;
const ESTATUS_CON_FOTO = ["INGRESO_CONFORME", "ENTREGA_MANUAL", "DERIVADO"];

// Vision + descargas + updates pueden pasar los 10s por defecto en cargas grandes.
export const config = { maxDuration: 30 };

// ---------------------------------------------------------------------------
// Comparación LPN vs. texto leído
// ---------------------------------------------------------------------------
// Igual que PharmaSPOT se compara solo con dígitos (el OCR confunde letras
// como O/0 o I/1 en el prefijo, los dígitos son lo que identifica al bulto).
// Si un LPN tuviera menos de 6 dígitos, se compara con letras y números.
//
// Diferencia con PharmaSPOT: la búsqueda se hace LÍNEA POR LÍNEA. Un voucher
// de carga puede listar varios LPN seguidos; si se juntara todo el texto en
// una sola cadena de dígitos, el final de un LPN y el inicio del siguiente
// podrían formar por casualidad otro número. Como respaldo, si ninguna línea
// alcanza el umbral, se acepta el texto completo solo con coincidencia exacta.

function claveLpn(lpn) {
  const digitos = String(lpn || "").replace(/[^0-9]/g, "");
  if (digitos.length >= 6) return { tipo: "digitos", valor: digitos };
  return { tipo: "alfa", valor: String(lpn || "").toUpperCase().replace(/[^0-9A-Z]/g, "") };
}

// Letras que el OCR suele leer en lugar de un dígito, sobre todo en letra
// manuscrita (ej. "O235" en vez de "0235"). Solo se aplica al comparar por
// dígitos. Como se exige coincidencia de al menos 95% (en la práctica, exacta
// para LPN de 10-12 dígitos), esto no abre la puerta a falsos positivos.
const LETRA_A_DIGITO = { O: "0", Q: "0", D: "0", I: "1", L: "1", Z: "2", S: "5", G: "6", B: "8" };

function normalizar(texto, tipo) {
  const t = String(texto || "").toUpperCase();
  if (tipo !== "digitos") return t.replace(/[^0-9A-Z]/g, "");
  return t.replace(/[ODQILZSGB]/g, c => LETRA_A_DIGITO[c]).replace(/[^0-9]/g, "");
}

// Prueba el valor esperado contra cada posición del segmento (distancia de
// Hamming con ventana deslizante) y devuelve la ventana más parecida.
function mejorVentana(valor, segmento) {
  const L = valor.length;
  if (L === 0 || segmento.length < L) return { ventana: null, score: -1 };
  let mejorD = Infinity, mejor = null;
  for (let i = 0; i <= segmento.length - L; i++) {
    const ventana = segmento.slice(i, i + L);
    let d = 0;
    for (let j = 0; j < L; j++) if (valor[j] !== ventana[j]) d++;
    if (d < mejorD) { mejorD = d; mejor = ventana; }
  }
  return { ventana: mejor, score: Math.round(((L - mejorD) / L) * 100) };
}

export function compararLpn(lpn, textoOcr) {
  const { tipo, valor } = claveLpn(lpn);
  if (!valor) return { coincide: false, score: 0, ocrLeyo: null };

  let mejor = { ventana: null, score: -1 };
  for (const linea of String(textoOcr || "").split(/\r?\n/)) {
    const r = mejorVentana(valor, normalizar(linea, tipo));
    if (r.score > mejor.score) mejor = r;
  }
  if (mejor.score >= UMBRAL_COINCIDENCIA) {
    return { coincide: true, score: mejor.score, ocrLeyo: mejor.ventana };
  }

  // Respaldo: el LPN partido en 2 líneas (frecuente en anotaciones a mano).
  const completo = mejorVentana(valor, normalizar(textoOcr, tipo));
  if (completo.score === 100) return { coincide: true, score: 100, ocrLeyo: completo.ventana };

  return { coincide: false, score: Math.max(mejor.score, 0), ocrLeyo: mejor.ventana };
}

// ---------------------------------------------------------------------------
// Lectura de texto (OCR)
// ---------------------------------------------------------------------------
function motorOcr() {
  if (process.env.TRANSITOS_VISION_API_KEY) return "google";
  if (process.env.TRANSITOS_OCRSPACE_API_KEY) return "ocrspace";
  return null;
}

// Google Cloud Vision. DOCUMENT_TEXT_DETECTION también lee texto manuscrito.
async function ocrGoogle(buffer) {
  const resp = await fetch(
    `https://vision.googleapis.com/v1/images:annotate?key=${process.env.TRANSITOS_VISION_API_KEY}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: [{
          image: { content: buffer.toString("base64") },
          features: [{ type: "DOCUMENT_TEXT_DETECTION" }],
          imageContext: { languageHints: ["es"] },
        }],
      }),
    }
  );
  const data = await resp.json();
  const r = data.responses?.[0];
  if (data.error) throw new Error(`Vision API: ${data.error.message}`);
  if (r?.error) throw new Error(`Vision API: ${r.error.message}`);
  return r?.fullTextAnnotation?.text || "";
}

// OCR.space (plan gratuito: archivos de hasta 1 MB; la app ya comprime las
// fotos a 800 KB como máximo). Se envía como archivo y no como base64, porque
// el base64 pesa un tercio más y podría pasar el límite de 1 MB.
// OCREngine 2 es el que mejor lee números y códigos.
async function ocrOcrSpace(buffer) {
  const form = new FormData();
  form.append("file", new Blob([buffer], { type: "image/jpeg" }), "sustento.jpg");
  form.append("language", "spa");
  form.append("OCREngine", "2");
  form.append("scale", "true");
  form.append("detectOrientation", "true");
  const resp = await fetch("https://api.ocr.space/parse/image", {
    method: "POST",
    headers: { apikey: process.env.TRANSITOS_OCRSPACE_API_KEY },
    body: form,
  });
  if (!resp.ok) throw new Error(`OCR.space respondió ${resp.status}`);
  const data = await resp.json();
  if (data.IsErroredOnProcessing) {
    const msg = Array.isArray(data.ErrorMessage) ? data.ErrorMessage.join(" ") : (data.ErrorMessage || "error desconocido");
    throw new Error(`OCR.space: ${msg}`);
  }
  return (data.ParsedResults || []).map(r => r.ParsedText || "").join("\n");
}

async function leerTextoFoto(supabase, path) {
  const { data: blob, error } = await supabase.storage.from("documentos").download(path);
  if (error) throw new Error(`Error descargando ${path}: ${error.message}`);
  const buffer = Buffer.from(await blob.arrayBuffer());
  return motorOcr() === "google" ? ocrGoogle(buffer) : ocrOcrSpace(buffer);
}

// Ejecuta tareas de a pocas a la vez (para no saturar Supabase).
async function enLotes(items, tam, fn) {
  const out = [];
  for (let i = 0; i < items.length; i += tam) {
    out.push(...(await Promise.all(items.slice(i, i + tam).map(fn))));
  }
  return out;
}

// ---------------------------------------------------------------------------
// Handler
// ---------------------------------------------------------------------------
export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  if (!process.env.SUPABASE_SERVICE_KEY || !motorOcr()) {
    return res.status(500).json({ error: "Faltan variables de entorno: SUPABASE_SERVICE_KEY y TRANSITOS_OCRSPACE_API_KEY (o TRANSITOS_VISION_API_KEY)" });
  }

  const lpns = Array.isArray(req.body?.lpns)
    ? [...new Set(req.body.lpns.map(String).filter(Boolean))]
    : [];
  if (lpns.length === 0) return res.status(400).json({ error: "Falta lpns" });
  if (lpns.length > MAX_LPNS_POR_LLAMADA) return res.status(400).json({ error: `Máximo ${MAX_LPNS_POR_LLAMADA} LPN por llamada` });

  // service_role: se salta RLS, por eso SOLO vive en el servidor.
  const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY, {
    auth: { persistSession: false },
  });

  // A diferencia de procesar-ocr.js, aquí se exige la sesión del usuario: el
  // endpoint puede aprobar bultos y cada llamada a Vision tiene costo, así que
  // solo un usuario logueado puede usarlo, y un transportista solo sobre sus
  // propios bultos.
  const token = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  if (!token) return res.status(401).json({ error: "Falta sesión" });
  const { data: authData, error: authErr } = await supabase.auth.getUser(token);
  if (authErr || !authData?.user) return res.status(401).json({ error: "Sesión inválida" });
  const meta = authData.user.user_metadata || {};
  const esAdmin = meta.role === "admin";

  let procesables = [];
  try {
    // 1. Lee solo las columnas necesarias, de a 100 LPN por consulta.
    const filas = [];
    for (let i = 0; i < lpns.length; i += 100) {
      const { data, error } = await supabase
        .from("registros_transito")
        .select("nro_lpn_final, proveedor, foto_url_1, foto_url_2, estatus_transportista, estado_validacion_admin")
        .in("nro_lpn_final", lpns.slice(i, i + 100));
      if (error) throw new Error(`Error leyendo registros: ${error.message}`);
      filas.push(...(data || []));
    }

    procesables = filas.filter(f =>
      (esAdmin || f.proveedor === meta.empresa_id) &&
      ESTATUS_CON_FOTO.includes(f.estatus_transportista) &&
      f.estado_validacion_admin === "PENDIENTE" &&
      (f.foto_url_1 || f.foto_url_2)
    );

    // 2. Lee cada foto UNA sola vez: en un sustento por carga, todos los
    //    bultos comparten las mismas fotos.
    const paths = [...new Set(procesables.flatMap(f => [f.foto_url_1, f.foto_url_2].filter(Boolean)))];
    const textoPorPath = {};
    for (const path of paths) {
      try {
        textoPorPath[path] = await leerTextoFoto(supabase, path);
      } catch (e) {
        console.error("Lectura fallida:", path, e.message);
        textoPorPath[path] = null; // los bultos que dependan solo de esta foto quedan en ERROR
      }
    }

    // 3. Compara y guarda el resultado de cada bulto.
    const ahora = new Date().toISOString();
    const resultadosProcesados = await enLotes(procesables, 10, async (f) => {
      const fotos = [f.foto_url_1, f.foto_url_2].filter(Boolean);
      const textos = fotos.map(p => textoPorPath[p]).filter(t => t != null);

      let cambios;
      let aprobado = false;
      if (textos.length === 0) {
        cambios = { estado_procesamiento_ia: "ERROR", estado_validacion_ia: null, texto_detectado_ia: null, match_ia: null };
      } else {
        // El mejor resultado entre las fotos del bulto.
        const match = textos
          .map(t => compararLpn(f.nro_lpn_final, t))
          .reduce((a, b) => (b.coincide && !a.coincide) || b.score > a.score ? b : a);
        aprobado = match.coincide;
        cambios = {
          estado_procesamiento_ia: "PROCESADO",
          estado_validacion_ia: match.coincide ? "COINCIDE" : "NO_COINCIDE",
          texto_detectado_ia: match.ocrLeyo,
          match_ia: match.score,
          ...(match.coincide ? { estado_validacion_admin: "APROBADO", validado_por: "IA", fecha_validacion: ahora } : {}),
        };
      }

      // Solo escribe si el bulto sigue igual que cuando se leyó: aún pendiente
      // y con la misma foto. Si en el intertanto el admin lo validó o el
      // transportista subió otra foto, este resultado ya no corresponde.
      let q = supabase.from("registros_transito").update(cambios)
        .eq("nro_lpn_final", f.nro_lpn_final)
        .eq("estado_validacion_admin", "PENDIENTE");
      q = f.foto_url_1 ? q.eq("foto_url_1", f.foto_url_1) : q.is("foto_url_1", null);
      const { data: act, error } = await q.select("nro_lpn_final");
      if (error) {
        console.error("Error guardando resultado:", f.nro_lpn_final, error.message);
        return { nro_lpn_final: f.nro_lpn_final, aprobado: false, cambios: null };
      }
      if (!act || act.length === 0) return { nro_lpn_final: f.nro_lpn_final, aprobado: false, cambios: null };
      return { nro_lpn_final: f.nro_lpn_final, aprobado, cambios };
    });

    // Los LPN pedidos que no se procesaron (ajenos, sin foto, ya validados)
    // vuelven sin cambios para que el frontend no los toque.
    const procesadosSet = new Set(resultadosProcesados.map(r => r.nro_lpn_final));
    const omitidos = lpns.filter(l => !procesadosSet.has(l)).map(l => ({ nro_lpn_final: l, aprobado: false, cambios: null }));

    return res.status(200).json({ resultados: [...resultadosProcesados, ...omitidos] });
  } catch (e) {
    // Igual que PharmaSPOT: nunca dejar un bulto colgado en "procesando".
    console.error("procesar-sustento-transitos:", e);
    for (const f of procesables) {
      await supabase.from("registros_transito")
        .update({ estado_procesamiento_ia: "ERROR", estado_validacion_ia: null, texto_detectado_ia: null, match_ia: null })
        .eq("nro_lpn_final", f.nro_lpn_final)
        .eq("estado_validacion_admin", "PENDIENTE")
        .is("estado_procesamiento_ia", null);
    }
    return res.status(500).json({ error: e.message });
  }
}
