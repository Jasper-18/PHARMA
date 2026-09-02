import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = "https://zffuccirauheklpxagga.supabase.co";
const UMBRAL_COINCIDENCIA = 95;

// ============================================================
// GR (guía de remisión) -- SIN CAMBIOS respecto a la versión anterior.
// Match difuso (tolera errores de OCR), pensado para números de guía.
// ============================================================

function normalizar(texto) {
  return (texto || "").replace(/[^0-9]/g, "");
}

function parsearRutas(rutasStr) {
  return (rutasStr || "").split("|").map(s => s.trim()).filter(Boolean);
}

function mejorVentana(itemNorm, textoNorm) {
  const L = itemNorm.length;
  if (L === 0) return { ventana: null, score: 0 };

  let mejorD = Infinity, mejor = null;

  for (let i = 0; i <= textoNorm.length - L; i++) {
    const ventana = textoNorm.slice(i, i + L);
    let d = 0;

    // Distancia de Hamming
    for (let j = 0; j < L; j++) {
      if (itemNorm[j] !== ventana[j]) d++;
    }

    if (d < mejorD) {
      mejorD = d;
      mejor = ventana;
    }
  }

  // Calculamos el porcentaje real de acierto matemático
  const score = Math.round(((L - mejorD) / L) * 100);

  return { ventana: mejor, score };
}

function compararConRutas(rutasStr, textoOcr) {
  const items = parsearRutas(rutasStr);
  const textoNorm = normalizar(textoOcr);

  let mejorItem = null, mejorScore = -1, ocrLeyo = null;

  for (const item of items) {
    const itemNorm = normalizar(item);
    if (!itemNorm || textoNorm.length < itemNorm.length) continue;

    // Obtenemos la ventana y su score real
    const resultado = mejorVentana(itemNorm, textoNorm);

    if (resultado.score > mejorScore) {
      mejorScore = resultado.score;
      mejorItem = item;
      ocrLeyo = resultado.ventana;
    }
  }

  return {
    mejorItem,
    ocrLeyo,
    mejorScore: mejorScore >= 0 ? mejorScore : 0,
    coincide: mejorScore >= UMBRAL_COINCIDENCIA,
  };
}

// Sufijo de columna según el documento: 1 -> "", 2 -> "_2", 3 -> "_3"
function sufijo(documento) {
  return documento === 1 ? "" : `_${documento}`;
}

// ============================================================
// COTIZACIÓN DE FLETE -- NUEVO. Match EXACTO (no difuso), pensado para
// el Importe. Requerimiento "P/OLT DESPACHO PROVINCIA" únicamente.
// ============================================================

// Extrae todos los números con pinta de monto (con o sin comas de miles,
// con o sin decimales) del texto de Vision, y los devuelve como floats.
function extraerCandidatosImporte(texto) {
  const matches = (texto || "").match(/\d[\d,]*\.?\d*/g) || [];
  return matches.map(m => parseFloat(m.replace(/,/g, ""))).filter(n => !isNaN(n));
}

// Compara el Importe del ticket (tal como esté guardado: "8200", "8200.00",
// "8,200.00") contra cualquier número que Vision haya detectado en la
// imagen -- por VALOR numérico real, no por texto, para que el formato no
// importe. Tolerancia de 0.005 solo para redondeos de punto flotante, no
// es una tolerancia de negocio.
function compararImporteExacto(importeCol, textoOcr) {
  const importeNum = parseFloat(String(importeCol ?? "").replace(/,/g, ""));
  if (isNaN(importeNum)) return { coincide: false, importeNum: null };
  const candidatos = extraerCandidatosImporte(textoOcr);
  const coincide = candidatos.some(c => Math.abs(c - importeNum) < 0.005);
  return { coincide, importeNum };
}

async function procesarCotizacion(supabase, nro_spot) {
  try {
    const { data: viaje, error: errViaje } = await supabase
      .from("viajes")
      .select("nro_spot, importe, foto_cotizacion")
      .eq("nro_spot", nro_spot)
      .single();

    if (errViaje || !viaje) return { status: 404, body: { error: "Viaje no encontrado" } };
    if (!viaje.foto_cotizacion) return { status: 400, body: { error: "La cotización no tiene foto" } };

    const { data: imgBlob, error: errDownload } = await supabase.storage
      .from("documentos")
      .download(viaje.foto_cotizacion);
    if (errDownload) throw new Error(`Error descargando imagen: ${errDownload.message}`);

    const arrayBuffer = await imgBlob.arrayBuffer();
    const base64 = Buffer.from(arrayBuffer).toString("base64");

    const visionRes = await fetch(
      `https://vision.googleapis.com/v1/images:annotate?key=${process.env.GOOGLE_VISION_API_KEY}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          requests: [{
            image: { content: base64 },
            features: [{ type: "DOCUMENT_TEXT_DETECTION" }],
          }],
        }),
      }
    );

    const visionData = await visionRes.json();
    const respuesta = visionData.responses?.[0];
    if (respuesta?.error) throw new Error(`Vision API: ${respuesta.error.message}`);

    const textoOcr = respuesta?.fullTextAnnotation?.text || "";
    const match = compararImporteExacto(viaje.importe, textoOcr);

    const payload = {
      estado_procesamiento_cotizacion: "PROCESADO",
      estado_validacion_cotizacion: match.coincide ? "COINCIDE" : "NO_COINCIDE",
      // Se guarda el texto completo (acotado a 3000 caracteres) -- acá no hay
      // una "ventana" puntual como en el GR, porque no buscamos un número
      // conocido de antemano dentro del texto, sino que extraemos candidatos.
      texto_detectado_cotizacion: textoOcr.slice(0, 3000),
    };

    const { error: errUpdate } = await supabase.from("viajes").update(payload).eq("nro_spot", nro_spot);
    if (errUpdate) throw new Error(`Error actualizando Supabase: ${errUpdate.message}`);

    return { status: 200, body: payload };

  } catch (e) {
    await supabase.from("viajes").update({
      estado_procesamiento_cotizacion: "ERROR",
      estado_validacion_cotizacion: null,
      texto_detectado_cotizacion: null,
    }).eq("nro_spot", nro_spot);

    return { status: 500, body: { error: e.message } };
  }
}

// ============================================================
// Handler principal
// ============================================================

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { nro_spot, documento } = req.body;
  if (!nro_spot) {
    return res.status(400).json({ error: "Falta nro_spot" });
  }

  const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);

  // --- Rama nueva: cotización de flete -- completamente aparte del flujo
  // de GR de abajo, para no arriesgar nada de lo que ya funciona. ---
  if (documento === "cotizacion") {
    const resultado = await procesarCotizacion(supabase, nro_spot);
    return res.status(resultado.status).json(resultado.body);
  }

  // --- Flujo de GR (placa/rutas) -- IDÉNTICO a la versión anterior ---
  const doc = [1, 2, 3].includes(documento) ? documento : 1;
  const suf = sufijo(doc);
  const colRutas = `rutas${suf}`;
  const colFoto = `foto_url${suf}`;
  const colProcesamiento = `estado_procesamiento_ia${suf}`;
  const colValidacion = `estado_validacion_ia${suf}`;
  const colTexto = `texto_detectado_ia${suf}`;
  const colMatch = `match_ia${suf}`;

  try {
    const { data: viaje, error: errViaje } = await supabase
      .from("viajes")
      .select(`nro_spot, ${colRutas}, ${colFoto}`)
      .eq("nro_spot", nro_spot)
      .single();

    if (errViaje || !viaje) return res.status(404).json({ error: "Viaje no encontrado" });
    if (!viaje[colFoto]) return res.status(400).json({ error: `El documento ${doc} no tiene foto` });

    const { data: imgBlob, error: errDownload } = await supabase.storage
      .from("documentos")
      .download(viaje[colFoto]);
    if (errDownload) throw new Error(`Error descargando imagen: ${errDownload.message}`);

    const arrayBuffer = await imgBlob.arrayBuffer();
    const base64 = Buffer.from(arrayBuffer).toString("base64");

    const visionRes = await fetch(
      `https://vision.googleapis.com/v1/images:annotate?key=${process.env.GOOGLE_VISION_API_KEY}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          requests: [{
            image: { content: base64 },
            features: [{ type: "DOCUMENT_TEXT_DETECTION" }],
          }],
        }),
      }
    );

    const visionData = await visionRes.json();
    const respuesta = visionData.responses?.[0];
    if (respuesta?.error) throw new Error(`Vision API: ${respuesta.error.message}`);

    const textoOcr = respuesta?.fullTextAnnotation?.text || "";
    const match = compararConRutas(viaje[colRutas], textoOcr);

    const payload = {
      [colProcesamiento]: "PROCESADO",
      [colValidacion]: match.coincide ? "COINCIDE" : "NO_COINCIDE",
      [colTexto]: match.ocrLeyo,
      [colMatch]: match.mejorScore,
    };

    const { error: errUpdate } = await supabase.from("viajes").update(payload).eq("nro_spot", nro_spot);
    if (errUpdate) throw new Error(`Error actualizando Supabase: ${errUpdate.message}`);

    // Se devuelven siempre con nombre SIN sufijo, para que App.jsx no tenga que
    // saber de sufijos -- el frontend ya sabe a qué documento pertenece por el
    // docIndex con el que hizo la llamada.
    return res.status(200).json({
      estado_procesamiento_ia: payload[colProcesamiento],
      estado_validacion_ia: payload[colValidacion],
      texto_detectado_ia: payload[colTexto],
      match_ia: payload[colMatch],
    });

  } catch (e) {
    await supabase.from("viajes").update({
      [colProcesamiento]: "ERROR",
      [colValidacion]: null,
      [colTexto]: null,
      [colMatch]: null,
    }).eq("nro_spot", nro_spot);

    return res.status(500).json({ error: e.message });
  }
}
