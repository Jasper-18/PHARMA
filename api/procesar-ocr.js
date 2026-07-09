import { createClient } from '@supabase/supabase-js';
import { partial_ratio } from 'fuzzball';

const SUPABASE_URL = "https://zffuccirauheklpxagga.supabase.co";
const UMBRAL_COINCIDENCIA = 95;

function normalizar(texto) {
  return (texto || "").replace(/[^0-9]/g, "");
}

function parsearRutas(rutasStr) {
  return (rutasStr || "").split("|").map(s => s.trim()).filter(Boolean);
}

function mejorVentana(itemNorm, textoNorm) {
  const L = itemNorm.length;
  let mejorD = Infinity, mejor = null;
  for (let i = 0; i <= textoNorm.length - L; i++) {
    const ventana = textoNorm.slice(i, i + L);
    let d = 0;
    for (let j = 0; j < L; j++) if (itemNorm[j] !== ventana[j]) d++;
    if (d < mejorD) { mejorD = d; mejor = ventana; }
  }
  return mejor;
}

function compararConRutas(rutasStr, textoOcr) {
  const items = parsearRutas(rutasStr);
  const textoNorm = normalizar(textoOcr);
  let mejorItem = null, mejorScore = -1, ocrLeyo = null;

  for (const item of items) {
    const itemNorm = normalizar(item);
    if (!itemNorm) continue;
    const score = partial_ratio(itemNorm, textoNorm);
    if (score > mejorScore) {
      mejorScore = score;
      mejorItem = item;
      ocrLeyo = mejorVentana(itemNorm, textoNorm);
    }
  }

  return {
    mejorItem,
    ocrLeyo,
    mejorScore: mejorScore >= 0 ? Math.round(mejorScore) : 0,
    coincide: mejorScore >= UMBRAL_COINCIDENCIA,
  };
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { nro_spot } = req.body;
  if (!nro_spot) {
    return res.status(400).json({ error: "Falta nro_spot" });
  }

  const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);

  try {
    const { data: viaje, error: errViaje } = await supabase
      .from("viajes")
      .select("nro_spot, rutas, foto_url")
      .eq("nro_spot", nro_spot)
      .single();

    if (errViaje || !viaje) return res.status(404).json({ error: "Viaje no encontrado" });
    if (!viaje.foto_url) return res.status(400).json({ error: "El viaje no tiene foto" });

    const { data: imgBlob, error: errDownload } = await supabase.storage
      .from("documentos")
      .download(viaje.foto_url);
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
    const match = compararConRutas(viaje.rutas, textoOcr);

    const payload = {
      estado_procesamiento_ia: "PROCESADO",
      estado_validacion_ia: match.coincide ? "COINCIDE" : "NO_COINCIDE",
      texto_detectado_ia: match.ocrLeyo,
      match_ia: match.mejorScore,
    };

    const { error: errUpdate } = await supabase.from("viajes").update(payload).eq("nro_spot", nro_spot);
    if (errUpdate) throw new Error(`Error actualizando Supabase: ${errUpdate.message}`);

    return res.status(200).json(payload);

  } catch (e) {
    await supabase.from("viajes").update({
      estado_procesamiento_ia: "ERROR",
      estado_validacion_ia: null,
      texto_detectado_ia: null,
      match_ia: null,
    }).eq("nro_spot", nro_spot);

    return res.status(500).json({ error: e.message });
  }
}
