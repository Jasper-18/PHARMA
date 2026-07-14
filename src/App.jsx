import { useState, useEffect, useRef, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://zffuccirauheklpxagga.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpmZnVjY2lyYXVoZWtscHhhZ2dhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI1Mjc0MzIsImV4cCI6MjA5ODEwMzQzMn0.MH8hJSktS_G3_omz9y48Vsp6PIlPcWdg6s6zdQtUNBo";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);

const RED = "#C8102E";
const RED_DARK = "#A00D24";
const RED_LIGHT = "#FCEBEB";
const GREEN = "#0f6e56";
const GREEN_LIGHT = "#e1f5ee";
const AMBER = "#854f0b";
const AMBER_LIGHT = "#faeeda";
const BLUE = "#185fa5";
const BLUE_LIGHT = "#e6f1fb";
const GRAY_50 = "#F8F8F8";
const GRAY_100 = "#F0EEEA";
const GRAY_200 = "#D3D1C7";
const GRAY_500 = "#888780";
const GRAY_900 = "#1a1a18";
const BORDER = "#e5e2db";
const MAX_MB = 10;

// Orden de columnas según diseño acordado (posición 6 = Proveedor, solo visible para admin)
// headerGroup: 'ia' = encabezado con fondo azul tenue, 'transportista' = fondo amarillo tenue
const COLS = [
  { key: "nro_spot",            label: "N° SPOT",                  width: 145, mono: true },
  { key: "fecha_carga",         label: "Fecha Servicio",           width: 118 },
  { key: "estado_final",        label: "Estado Final",             width: 150 },
  { key: "placa",               label: "N° Placa",                 width: 90,  mono: true,  headerGroup: "transportista" },
  { key: "rutas",               label: "N° GR",                    width: 140, trunc: true, headerGroup: "transportista" },
  { key: "texto_detectado_ia",  label: "Texto Detectado IA",       width: 200, trunc: true, headerGroup: "ia" },
  { key: "proveedor",           label: "Proveedor",                width: 110, adminOnly: true },
  { key: "hora_cita",           label: "Hora Cita",                width: 70 },
  { key: "cd_origen",           label: "Origen",                   width: 130 },
  { key: "cd_destino",          label: "Destino",                  width: 130 },
  { key: "tipo_traslado",       label: "Tipo de Envío",            width: 90 },
  { key: "cantidad",            label: "Cantidad",                 width: 80,  right: true },
  { key: "area",                label: "Área",                     width: 120 },
  { key: "requerimiento",       label: "Requerimiento",            width: 160, trunc: true },
  { key: "importe",             label: "Importe",                  width: 90,  right: true },
  { key: "centro_costo",        label: "CECO",                     width: 110, muted: true },
  { key: "detalle_servicio",    label: "Detalle del Servicio",     width: 180, trunc: true },
  { key: "realizado",    label: "Realizado",                width: 80 },
  { key: "estado_procesamiento_ia", label: "Procesam. IA",         width: 100, headerGroup: "ia", adminOnly: true },
  { key: "estado_validacion_ia",label: "Validación IA",            width: 110, headerGroup: "ia" },
  { key: "match_ia",            label: "Match IA",                 width: 70,  right: true, headerGroup: "ia" },
  { key: "usuario_modif",       label: "Usuario Modif.",           width: 140, muted: true, headerGroup: "ia" },
  { key: "fecha_modif",         label: "Fecha Modif.",             width: 130, muted: true, headerGroup: "ia" },
  { key: "estado_doc",          label: "Estado Doc.",              width: 100 },
  { key: "fecha_entrega_doc",   label: "Fec. Entrega Doc.",        width: 130 },
];

// Vista de transportista: panel principal resumido (orden exacto solicitado)
const COLS_TRANSPORTISTA_PRINCIPAL = [
  "nro_spot", "fecha_carga", "estado_final", "cd_origen", "cd_destino",
  "placa", "rutas", "texto_detectado_ia", "estado_validacion_ia",
];

// Vista de transportista: campos adicionales que se muestran solo en el modal de detalle
// (excluye lo que ya está en el panel principal, y explícitamente match_ia/usuario_modif/fecha_modif/hora_cita/area)
const COLS_TRANSPORTISTA_DETALLE = [
  "importe", "tipo_traslado", "cantidad", "requerimiento",
  "detalle_servicio", "realizado", "estado_doc", "fecha_entrega_doc",
];

// Extrae el segmento "Nro..." de un nro_spot para nombrar archivos
// "Serv Adcional Nro000004006268" → "Nro000004006268"
function extraerNroSpotCorto(nroSpot) {
  if (!nroSpot) return "doc";
  const m = String(nroSpot).match(/(Nro\d+)/i);
  return m ? m[1] : nroSpot.replace(/\s+/g, "_").slice(0, 30);
}

// Calcula el rango de fechas por defecto anclado en la fecha_carga MÁS RECIENTE
// que exista en la data (no en "hoy"), porque los servicios suelen programarse
// para el día siguiente -- si se ancla en "hoy", un registro con fecha de mañana
// queda fuera del rango justo el día que se crea.
async function calcularRangoAncla_(isAdminUser, empresaId) {
  let ancla = new Date(); // fallback si no hay data o falla la consulta
  try {
    let q = supabase
      .from("viajes")
      .select("fecha_carga")
      .eq("realizado", "SI")
      .not("fecha_carga", "is", null)
      .order("fecha_carga", { ascending: false })
      .limit(1);
    if (!isAdminUser && empresaId) q = q.eq("proveedor", empresaId);
    const { data } = await q;
    if (data && data[0] && data[0].fecha_carga) {
      // Parseo local (sin "Z"), para no correr el día por interpretación UTC
      ancla = new Date(data[0].fecha_carga + "T00:00:00");
    }
  } catch { /* si falla, se queda con "hoy" como ancla */ }

  const desde = new Date(ancla);
  desde.setDate(ancla.getDate() - (isAdminUser ? 6 : 30));
  const fmt = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return { desde: fmt(desde), hasta: fmt(ancla) };
}

async function comprimirImagen(file) {
  if (file.size > MAX_MB * 1024 * 1024) {
    throw new Error(`El archivo supera los ${MAX_MB} MB.`);
  }
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        let { width, height } = img;
        const maxDim = 1800;
        if (width > maxDim || height > maxDim) {
          if (width > height) { height = Math.round(height * maxDim / width); width = maxDim; }
          else { width = Math.round(width * maxDim / height); height = maxDim; }
        }
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        let quality = 0.85;
        const tryCompress = () => {
          canvas.toBlob((blob) => {
            if (blob.size > 800 * 1024 && quality > 0.3) { quality -= 0.1; tryCompress(); }
            else resolve(new File([blob], file.name.replace(/\.[^.]+$/, ".jpg"), { type: "image/jpeg" }));
          }, "image/jpeg", quality);
        };
        tryCompress();
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

const MESES_ES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const DIAS_ES  = ["D","L","M","X","J","V","S"];

function RangePicker({ desde, hasta, maxDias = 31, onChange }) {
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const parseFecha = s => s ? new Date(s + "T00:00:00") : null;
  const fmt = d => d ? d.toISOString().slice(0,10) : "";
  const inicioMes = (() => { const ref = parseFecha(desde) || hoy; return new Date(ref.getFullYear(), ref.getMonth(), 1); })();
  const [mesVista, setMesVista] = useState(inicioMes);
  const [selStart, setSelStart] = useState(parseFecha(desde));
  const [selEnd,   setSelEnd]   = useState(parseFecha(hasta));
  const [hover,    setHover]    = useState(null);

  const prevMes = () => setMesVista(m => new Date(m.getFullYear(), m.getMonth()-1, 1));
  const nextMes = () => setMesVista(m => new Date(m.getFullYear(), m.getMonth()+1, 1));

  const diasDelMes = () => {
    const dias = [];
    const inicio = new Date(mesVista.getFullYear(), mesVista.getMonth(), 1).getDay();
    for (let i = 0; i < inicio; i++) dias.push(null);
    const total = new Date(mesVista.getFullYear(), mesVista.getMonth()+1, 0).getDate();
    for (let d = 1; d <= total; d++) dias.push(new Date(mesVista.getFullYear(), mesVista.getMonth(), d));
    return dias;
  };

  const handleDia = (dia) => {
    if (!dia) return;
    if (!selStart || (selStart && selEnd)) { setSelStart(dia); setSelEnd(null); setHover(null); }
    else {
      if (dia < selStart) { setSelStart(dia); setSelEnd(null); }
      else {
        if ((dia - selStart) / 86400000 > maxDias) return;
        setSelEnd(dia);
        onChange({ desde: fmt(selStart), hasta: fmt(dia) });
      }
    }
  };

  const enRango = (d) => { if (!d) return false; const end = selEnd || hover; return selStart && end && d >= selStart && d <= end; };
  const esStart = d => d && selStart && fmt(d) === fmt(selStart);
  const esEnd   = d => d && selEnd   && fmt(d) === fmt(selEnd);
  const esHoy   = d => d && fmt(d) === fmt(hoy);
  const rangoValido = selStart && selEnd && ((selEnd - selStart) / 86400000) <= 7;

  return (
    <div style={{ userSelect: "none" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <button onClick={prevMes} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: GRAY_500, padding: "2px 6px" }}>‹</button>
        <div style={{ fontSize: 13, fontWeight: 600, color: GRAY_900, textTransform: "uppercase", letterSpacing: ".04em" }}>
          {MESES_ES[mesVista.getMonth()]} {mesVista.getFullYear()}
        </div>
        <button onClick={nextMes} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: GRAY_500, padding: "2px 6px" }}>›</button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", marginBottom: 4 }}>
        {DIAS_ES.map(d => <div key={d} style={{ textAlign: "center", fontSize: 10, color: GRAY_500, fontWeight: 500, padding: "2px 0" }}>{d}</div>)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: "2px 0" }}>
        {diasDelMes().map((dia, i) => {
          const isS = esStart(dia), isE = esEnd(dia), isH = esHoy(dia), enR = enRango(dia);
          const sobreLimite = selStart && !selEnd && hover && dia && (Math.abs(dia - selStart) / 86400000) > maxDias;
          return (
            <div key={i} onClick={() => handleDia(dia)}
              onMouseEnter={() => { if (selStart && !selEnd && dia) setHover(dia); }}
              onMouseLeave={() => setHover(null)}
              style={{ height: 32, display: "flex", alignItems: "center", justifyContent: "center",
                cursor: dia ? "pointer" : "default",
                background: (isS || isE) ? RED : enR ? "#FDDDE3" : "transparent",
                borderRadius: isS ? "50% 0 0 50%" : isE ? "0 50% 50% 0" : 0,
                color: (isS || isE) ? "white" : sobreLimite ? "#ccc" : isH ? RED : dia ? GRAY_900 : "transparent",
                fontWeight: (isS || isE || isH) ? 600 : 400, fontSize: 12,
                outline: isH && !isS && !isE ? `1.5px solid ${RED}` : "none", outlineOffset: -2 }}>
              {dia ? dia.getDate() : ""}
            </div>
          );
        })}
      </div>
      <div style={{ marginTop: 10, padding: "6px 10px", background: GRAY_100, borderRadius: 7, fontSize: 11, color: GRAY_500, textAlign: "center" }}>
        {rangoValido ? `${fmt(selStart)} – ${fmt(selEnd)}` : "Selecciona el rango de fechas"}
      </div>
    </div>
  );
}

export default function App() {
  const [session,      setSession]      = useState(null);
  const [loading,      setLoading]      = useState(true);
  const [email,        setEmail]        = useState("");
  const [password,     setPassword]     = useState("");
  const [loginErr,     setLoginErr]     = useState("");
  const [viajes,       setViajes]       = useState([]);
  const [dataLoading,  setDataLoading]  = useState(false);

  const _saved = (() => { try { return JSON.parse(sessionStorage.getItem("ps_filters") || "{}"); } catch { return {}; } })();
  const [fDesde,    setFDesde]    = useState(_saved.fDesde    ?? "");
  const [fHasta,    setFHasta]    = useState(_saved.fHasta    ?? "");
  const [fEstadoDoc,setFEstadoDoc]= useState(_saved.fEstadoDoc ?? "");
  const [fNroSpot,  setFNroSpot]  = useState(_saved.fNroSpot  ?? "");
  const [fRutas,    setFRutas]    = useState(_saved.fRutas    ?? "");
  const [fPlaca,    setFPlaca]    = useState(_saved.fPlaca    ?? "");
  const [fEstFinal, setFEstFinal] = useState(_saved.fEstFinal ?? "");
  const [fProveedor,setFProveedor]= useState(_saved.fProveedor ?? "");

  const [filterOpen, setFilterOpen] = useState(false);
  const [dNroSpot,   setDNroSpot]   = useState("");
  const [dRutas,     setDRutas]     = useState("");
  const [dPlaca,     setDPlaca]     = useState("");
  const [dEstadoDoc, setDEstadoDoc] = useState("");
  const [dEstFinal,  setDEstFinal]  = useState("");
  const [dProveedor, setDProveedor] = useState("");
  const [dDesde,     setDDesde]     = useState("");
  const [dHasta,     setDHasta]     = useState("");
  const [fechaErr,   setFechaErr]   = useState("");

  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [modal,        setModal]        = useState(null);
  const [editModal,    setEditModal]    = useState(null);
  const [vista, setVista] = useState("tabla");
  const [detalleModal, setDetalleModal] = useState(null);
  const [dashProveedor, setDashProveedor] = useState("");
  const [dashDesde, setDashDesde] = useState("");
  const [dashHasta, setDashHasta] = useState("");
  const [kpis, setKpis] = useState(null);
  const [ranking, setRanking] = useState([]);
  const [detalle, setDetalle] = useState([]);
  const [detalleTotal, setDetalleTotal] = useState(0);
  const [detallePagina, setDetallePagina] = useState(0);
  const [dashLoading, setDashLoading] = useState(false);
  const DETALLE_POR_PAGINA = 10;
  const [pwModalOpen,  setPwModalOpen]  = useState(false);
  const [pwNueva,      setPwNueva]      = useState("");
  const [pwConfirma,   setPwConfirma]   = useState("");
  const [pwErr,        setPwErr]        = useState("");
  const [pwSaving,     setPwSaving]     = useState(false);
  const [pwOk,         setPwOk]         = useState(false);
  const [editPlaca,    setEditPlaca]    = useState("");
  const [editRutaInput,setEditRutaInput]= useState("");
  const [editSaving,   setEditSaving]   = useState(false);
  const [editErr,      setEditErr]      = useState("");
  const [uploading,    setUploading]    = useState(false);
  const [uploadFile,   setUploadFile]   = useState(null);
  const [uploadErr,    setUploadErr]    = useState("");
  const [uploadSuccess,setUploadSuccess]= useState(false);
  const [validandoIA,  setValidandoIA]  = useState(false);
  const [resultadoIA,  setResultadoIA]  = useState(null);

  // Estado modal validación admin
  const [validModal,   setValidModal]   = useState(null);
  const [validSaving,  setValidSaving]  = useState(false);
  const [validErr,     setValidErr]     = useState("");

  const fileRef    = useRef();
  const userMenuRef= useRef();
  const fetchViajesRef = useRef(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setLoading(false); });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  // Al recibir la sesión, si no hay fechas guardadas en sessionStorage PARA ESTE MISMO ROL,
  // aplica el rango por defecto según el rol: 7 días para admin, 31 para transportista —
  // anclado en la fecha_carga más reciente de la data, no en "hoy" (ver calcularRangoAncla_)
  useEffect(() => {
    if (!session) return;
    const isAdminUser = session.user.user_metadata?.role === "admin";
    const rolActual = isAdminUser ? "admin" : "transportista";
    if (_saved.fDesde && _saved.fHasta && _saved._rol === rolActual) return; // mismo rol, no pisar
    (async () => {
      const { desde, hasta } = await calcularRangoAncla_(isAdminUser, session.user.user_metadata?.empresa_id);
      setFDesde(desde);
      setFHasta(hasta);
    })();
  }, [session]);

  useEffect(() => {
    const h = (e) => { if (userMenuRef.current && !userMenuRef.current.contains(e.target)) setUserMenuOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);

  const fetchViajes = useCallback(async () => {
    if (!session) return;
    setDataLoading(true);
    const meta = session.user.user_metadata;
    const isAdmin = meta?.role === "admin";
    const filters = fetchViajesRef.current || {};
    let q = supabase.from("viajes").select("*").order("fecha_carga", { ascending: false });
    q = q.eq("realizado", "SI"); // solo servicios efectivamente realizados, para ambos roles
    if (!isAdmin && meta?.empresa_id) q = q.eq("proveedor", meta.empresa_id);
    if (isAdmin && filters.fProveedor) q = q.eq("proveedor", filters.fProveedor);
    if (filters.fDesde)    q = q.gte("fecha_carga", filters.fDesde);
    if (filters.fHasta)    q = q.lte("fecha_carga", filters.fHasta);
    if (filters.fNroSpot)  q = q.ilike("nro_spot",  `%${filters.fNroSpot}%`);
    if (filters.fRutas)    q = q.ilike("rutas",      `%${filters.fRutas}%`);
    if (filters.fPlaca)    q = q.ilike("placa",      `%${filters.fPlaca}%`);
    if (filters.fEstFinal) q = q.eq("estado_final",   filters.fEstFinal);
    const { data } = await q;
    let rows = data || [];
    if (filters.fEstadoDoc === "Completo") rows = rows.filter(r => (r.foto_versiones?.length || 0) > 0);
    if (filters.fEstadoDoc === "Pendiente") rows = rows.filter(r => !(r.foto_versiones?.length > 0));
    setViajes(rows);
    setDataLoading(false);
  }, [session]);

  useEffect(() => { if (session) fetchViajes(); }, [session]);

  useEffect(() => {
    if (!session) return;
    const rolActual = session.user.user_metadata?.role === "admin" ? "admin" : "transportista";
    sessionStorage.setItem("ps_filters", JSON.stringify({ fDesde, fHasta, fEstadoDoc, fNroSpot, fRutas, fPlaca, fEstFinal, fProveedor, _rol: rolActual }));
  }, [fDesde, fHasta, fEstadoDoc, fNroSpot, fRutas, fPlaca, fEstFinal, fProveedor, session]);

  useEffect(() => {
    fetchViajesRef.current = { fDesde, fHasta, fEstadoDoc, fNroSpot, fRutas, fPlaca, fEstFinal, fProveedor };
  }, [fDesde, fHasta, fEstadoDoc, fNroSpot, fRutas, fPlaca, fEstFinal, fProveedor]);

  async function doLogin(e) {
    e.preventDefault();
    setLoginErr("");
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) setLoginErr("Usuario o contraseña incorrectos");
  }

  async function cambiarPassword() {
    setPwErr("");
    if (pwNueva.length < 6) { setPwErr("La contraseña debe tener al menos 6 caracteres."); return; }
    if (pwNueva !== pwConfirma) { setPwErr("Las contraseñas no coinciden."); return; }
    setPwSaving(true);
    try {
      const { error } = await supabase.auth.updateUser({ password: pwNueva });
      if (error) throw error;
      setPwOk(true);
      setPwNueva(""); setPwConfirma("");
      setTimeout(() => { setPwModalOpen(false); setPwOk(false); }, 1800);
    } catch (err) {
      setPwErr(err.message || "Error al cambiar la contraseña.");
    } finally {
      setPwSaving(false);
    }
  }

  function abrirModalPassword() {
    setUserMenuOpen(false);
    setPwNueva(""); setPwConfirma(""); setPwErr(""); setPwOk(false);
    setPwModalOpen(true);
  }

  const fetchDashboard = useCallback(async (pagina = 0) => {
    if (!isAdmin) return;
    setDashLoading(true);
    try {
      const [{ data: kpisData, error: kpisErr }, { data: rankingData, error: rankingErr }] = await Promise.all([
        supabase.rpc("dashboard_kpis", {
          p_proveedor: dashProveedor || null,
          p_fecha_desde: dashDesde || null,
          p_fecha_hasta: dashHasta || null,
        }),
        supabase.rpc("dashboard_ranking", {
          p_fecha_desde: dashDesde || null,
          p_fecha_hasta: dashHasta || null,
        }),
      ]);
      if (kpisErr) throw kpisErr;
      if (rankingErr) throw rankingErr;
      setKpis(kpisData?.[0] || null);
      setRanking(rankingData || []);

      let q = supabase.from("viajes")
        .select("nro_spot, proveedor, fecha_carga, realizado, rutas, estado_validacion_ia, estado_final", { count: "exact" })
        .order("fecha_registro", { ascending: false })
        .range(pagina * DETALLE_POR_PAGINA, pagina * DETALLE_POR_PAGINA + DETALLE_POR_PAGINA - 1);
      if (dashProveedor) q = q.eq("proveedor", dashProveedor);
      if (dashDesde) q = q.gte("fecha_registro", dashDesde);
      if (dashHasta) q = q.lte("fecha_registro", dashHasta);
      const { data: detalleData, count, error: detalleErr } = await q;
      if (detalleErr) throw detalleErr;
      setDetalle(detalleData || []);
      setDetalleTotal(count || 0);
      setDetallePagina(pagina);
    } catch (err) {
      console.error("Error cargando dashboard:", err);
    } finally {
      setDashLoading(false);
    }
  }, [isAdmin, dashProveedor, dashDesde, dashHasta]);

  useEffect(() => {
    if (vista !== "dashboard" || !isAdmin) return;
    fetchDashboard(0);
    const intervalo = setInterval(() => fetchDashboard(detallePagina), 30000);
    return () => clearInterval(intervalo);
  }, [vista, isAdmin, dashProveedor, dashDesde, dashHasta]);

  function limpiarFiltrosDashboard() {
    setDashProveedor(""); setDashDesde(""); setDashHasta("");
  }

  function openEditModal(viaje) {
    setEditModal(viaje);
    setEditPlaca(viaje.placa || "");
    setEditRutaInput(viaje.rutas || "");
    setEditErr("");
  }

  const PLACA_RE = /^[A-Za-z0-9]{3}-[A-Za-z0-9]{3}$/;

  async function saveEdit() {
    if (editPlaca && !PLACA_RE.test(editPlaca)) { setEditErr("Formato inválido. Usa ABC-123."); return; }
    setEditSaving(true); setEditErr("");
    try {
      const payload = { placa: editPlaca || null, rutas: editRutaInput.trim() || null };
      const { data, error } = await supabase
        .from("viajes")
        .update(payload)
        .eq("nro_spot", editModal.nro_spot)
        .select();
      if (error) throw error;
      setViajes(prev => prev.map(v =>
        v.nro_spot === editModal.nro_spot
          ? { ...v, placa: payload.placa, rutas: payload.rutas }
          : v
      ));
      setEditModal(null);
    } catch (err) {
      setEditErr(err.message || "Error al guardar.");
    }
    finally { setEditSaving(false); }
  }

  function applyFilters() {
    if (dDesde && dHasta) {
      const diff = (new Date(dHasta) - new Date(dDesde)) / (1000 * 60 * 60 * 24);
      const maxDias = isAdmin ? 7 : 31;
      if (diff < 0) { setFechaErr("La fecha 'Hasta' debe ser mayor o igual a 'Desde'."); return; }
      if (diff > maxDias) { setFechaErr(`El rango máximo permitido es de ${maxDias} días.`); return; }
    }
    setFechaErr("");
    setFNroSpot(dNroSpot); setFRutas(dRutas); setFPlaca(dPlaca);
    setFEstadoDoc(dEstadoDoc); setFEstFinal(dEstFinal); setFProveedor(dProveedor);
    setFDesde(dDesde); setFHasta(dHasta);
    // Actualizar el ref directamente para que fetchViajes use los valores nuevos
    // (los useEffect de React aún no habrán corrido cuando se llama fetchViajes)
    fetchViajesRef.current = {
      fDesde: dDesde, fHasta: dHasta, fEstadoDoc: dEstadoDoc,
      fNroSpot: dNroSpot, fRutas: dRutas, fPlaca: dPlaca,
      fEstFinal: dEstFinal, fProveedor: dProveedor
    };
    fetchViajes();
    setFilterOpen(false);
  }

  async function clearFilters() {
    const { desde, hasta } = await calcularRangoAncla_(isAdmin, meta?.empresa_id);
    const d = { desde, hasta };
    setDNroSpot(""); setDRutas(""); setDPlaca(""); setDEstadoDoc(""); setDEstFinal(""); setDProveedor("");
    setDDesde(d.desde); setDHasta(d.hasta); setFechaErr("");
    setFNroSpot(""); setFRutas(""); setFPlaca(""); setFEstadoDoc(""); setFEstFinal(""); setFProveedor("");
    setFDesde(d.desde); setFHasta(d.hasta);
    fetchViajesRef.current = {
      fDesde: d.desde, fHasta: d.hasta, fEstadoDoc: "",
      fNroSpot: "", fRutas: "", fPlaca: "", fEstFinal: "", fProveedor: ""
    };
    fetchViajes();
  }

  function openFilter() {
    setDNroSpot(fNroSpot); setDRutas(fRutas); setDPlaca(fPlaca);
    setDEstadoDoc(fEstadoDoc); setDEstFinal(fEstFinal); setDProveedor(fProveedor);
    setDDesde(fDesde); setDHasta(fHasta); setFechaErr("");
    setFilterOpen(true);
  }

  // Para columnas tipo DATE puro (solo calendario, ej. fecha_carga = "2026-07-12").
  // Se parsea directo de los dígitos, SIN pasar por new Date(): new Date() interpreta
  // ese texto como medianoche UTC, y en Perú (UTC-5) eso corre el día hacia atrás al
  // convertir a hora local (12 → 11). No hay ninguna zona horaria que aplicar aquí,
  // porque una fecha tipo DATE no es un instante, es un día del calendario.
  const fmtFechaSolo = (val) => {
    if (!val) return "";
    const m = String(val).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[3]}/${m[2]}/${m[1]}`;
    return String(val);
  };

  // Para TIMESTAMPTZ reales (fecha_modif, fecha_entrega_doc, subido_en) — estos sí
  // son un instante con hora y zona horaria, así que SÍ corresponde convertir a
  // hora local del navegador.
  const fmtFecha = (val) => {
    if (!val) return "";
    const d = new Date(val);
    if (isNaN(d)) return String(val);
    const dd   = String(d.getDate()).padStart(2, "0");
    const mm   = String(d.getMonth() + 1).padStart(2, "0");
    const aaaa = d.getFullYear();
    return `${dd}/${mm}/${aaaa}`;
  };

  const fmtFechaHora = (val) => {
    if (!val) return "";
    const d = new Date(val);
    if (isNaN(d)) return String(val);
    const dd   = String(d.getDate()).padStart(2, "0");
    const mm   = String(d.getMonth() + 1).padStart(2, "0");
    const aaaa = d.getFullYear();
    let h = d.getHours();
    const min  = String(d.getMinutes()).padStart(2, "0");
    const ampm = h >= 12 ? "pm." : "am.";
    h = h % 12 || 12; // sin cero inicial: 4 en vez de 04
    return `${dd}/${mm}/${aaaa} ${h}:${min} ${ampm}`;
  };

  function exportarExcel() {
    if (!viajes.length) return;
    // Exporta siempre el set completo de columnas (según rol), no el panel resumido en pantalla
    const colsExp = COLS.filter(c => !c.adminOnly || isAdmin);
    const headers = colsExp.map(c => c.label);
    const rows = viajes.map(v => colsExp.map(c => {
      const val = v[c.key];
      if (val === null || val === undefined) return "";
      if (c.key === "fecha_carga")
        return fmtFechaSolo(val);
      if (c.key === "fecha_modif" || c.key === "fecha_entrega_doc")
        return fmtFecha(val);
      if (c.key === "realizado" || c.key === "estado_final" || c.key === "estado_procesamiento_ia" || c.key === "estado_validacion_ia")
        return val ? String(val).toUpperCase() : "";
      if (c.key === "foto_versiones") return Array.isArray(val) ? val.length : 0;
      return String(val);
    }));

    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Auto-ancho de columnas basado en el contenido más largo
    const colWidths = headers.map((h, i) => {
      const maxLen = Math.max(
        h.length,
        ...rows.map(r => String(r[i] || "").length)
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });
    ws["!cols"] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Viajes");
    XLSX.writeFile(wb, "viajes_adicionales.xlsx");
  }

  async function abrirFoto(pathOrUrl) {
    // Si viene una URL completa (datos viejos), extraer el path relativo
    let path = pathOrUrl;
    if (pathOrUrl && pathOrUrl.startsWith("http")) {
      const match = pathOrUrl.match(/\/object\/(?:public|sign)\/documentos\/(.+?)(?:\?|$)/);
      path = match ? decodeURIComponent(match[1]) : pathOrUrl;
    }
    const { data, error } = await supabase.storage.from("documentos").createSignedUrl(path, 120);
    if (error || !data?.signedUrl) {
      console.error("Error signed URL:", error, "path:", path);
      alert("No se pudo obtener la URL de la foto.");
      return;
    }
    window.open(data.signedUrl, "_blank");
  }

  async function descargarFoto(pathOrUrl, nombre) {
    let path = pathOrUrl;
    if (pathOrUrl && pathOrUrl.startsWith("http")) {
      const match = pathOrUrl.match(/\/object\/(?:public|sign)\/documentos\/(.+?)(?:\?|$)/);
      path = match ? decodeURIComponent(match[1]) : pathOrUrl;
    }
    const { data, error } = await supabase.storage.from("documentos").createSignedUrl(path, 120);
    if (error || !data?.signedUrl) {
      console.error("Error signed URL:", error, "path:", path);
      alert("No se pudo obtener la URL de la foto.");
      return;
    }
    // Forzar descarga via blob para evitar que el navegador abra la imagen en pestaña
    try {
      const resp = await fetch(data.signedUrl);
      const blob = await resp.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = nombre || "documento.jpg";
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch {
      // Fallback: abrir en pestaña si fetch falla por CORS
      window.open(data.signedUrl, "_blank");
    }
  }

  function openModal(viaje) {
    // Verificar que el transportista haya completado placa y rutas antes de subir
    const faltaPlaca = !viaje.placa || viaje.placa.trim() === "";
    const faltaRutas = !viaje.rutas || viaje.rutas.trim() === "";
    if (faltaPlaca || faltaRutas) {
      const campos = [faltaPlaca && "N° Placa", faltaRutas && "N° GR"].filter(Boolean).join(" y ");
      setModal({ ...viaje, _bloqueado: true, _mensajeBloqueo: campos });
    } else {
      setModal(viaje);
    }
    setUploadFile(null); setUploadErr(""); setUploadSuccess(false); setValidandoIA(false); setResultadoIA(null);
  }

  function handleFileSelect(file) {
    if (!file) return;
    if (file.size > MAX_MB * 1024 * 1024) { setUploadErr(`El archivo supera los ${MAX_MB} MB.`); setUploadFile(null); return; }
    setUploadErr(""); setUploadFile(file);
  }

  async function handleUpload() {
    if (!uploadFile || !modal) return;
    setUploading(true); setUploadErr("");
    try {
      const compressed = await comprimirImagen(uploadFile);
      const versiones  = modal.foto_versiones || [];
      const nv  = versiones.length + 1;
      const ext = compressed.name.split(".").pop();
      const nroCorto = extraerNroSpotCorto(modal.nro_spot);
      const fileName = `${nroCorto}_v${nv}.${ext}`;
      const path = `${modal.proveedor}/${modal.nro_spot}/${fileName}`;
      const { error: upErr } = await supabase.storage.from("documentos").upload(path, compressed, { upsert: true });
      if (upErr) throw upErr;
      // Bucket privado — guardar el path, la URL firmada se genera al momento de ver/descargar
      const nuevasVersiones = [...versiones, { v: nv, path, nombre: fileName, subido_por: session.user.email, subido_en: new Date().toISOString() }];
      const ahora = new Date().toISOString();
      const { error: dbErr } = await supabase.from("viajes").update({
        foto_versiones:    nuevasVersiones,
        foto_url:          path,           // guardamos el path, no URL pública
        foto_nombre:       fileName,
        subido_por:        session.user.email,
        subido_en:         ahora,
        fecha_entrega_doc: versiones.length === 0 ? ahora : modal.fecha_entrega_doc,
        estado_doc:        "Completo",
        estado_procesamiento_ia: null, // resetea para forzar re-validación de esta nueva foto
      }).eq("nro_spot", modal.nro_spot);
      if (dbErr) throw dbErr;

      // La foto ya está guardada de forma segura en este punto — lo que sigue
      // (validación IA en vivo) es una mejora de UX, no puede hacer fallar la subida.
      setUploading(false);
      setValidandoIA(true);

      let resultado = null;
      try {
        const resp = await fetch("/api/procesar-ocr", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ nro_spot: modal.nro_spot }),
        });
        if (resp.ok) resultado = await resp.json();
      } catch { /* sin internet momentáneo, timeout, etc. — el batch la recoge después */ }

      setValidandoIA(false);
      setResultadoIA(resultado); // null si falló — se muestra mensaje genérico
      setViajes(prev => prev.map(v => v.nro_spot === modal.nro_spot
        ? { ...v, foto_versiones: nuevasVersiones, foto_url: path, foto_nombre: fileName,
            subido_por: session.user.email, subido_en: ahora, estado_doc: "Completo",
            ...(resultado || {}) }
        : v));

      setUploadSuccess(true);
      fetchViajes();
      const cierraSolo = !resultado || resultado.estado_validacion_ia !== "NO_COINCIDE";
      if (cierraSolo) setTimeout(() => { setModal(null); setResultadoIA(null); }, 2200);
    } catch (err) {
      setUploadErr(err.message || "Error al subir el archivo.");
      setUploading(false);
    }
  }

  // Validación manual por admin
  async function handleValidarManual() {
    if (!validModal) return;
    setValidSaving(true); setValidErr("");
    try {
      const ahora = new Date().toISOString();
      const { error } = await supabase.from("viajes").update({
        estado_validacion_ia: "MANUAL",
        estado_final:         "FINALIZADO",
        usuario_modif:        session.user.email,
        fecha_modif:          ahora,
      }).eq("nro_spot", validModal.nro_spot);
      if (error) throw error;
      setViajes(prev => prev.map(v => v.nro_spot === validModal.nro_spot
        ? { ...v, estado_validacion_ia: "MANUAL", estado_final: "FINALIZADO", usuario_modif: session.user.email, fecha_modif: ahora }
        : v));
      setValidModal(null);
    } catch (err) { setValidErr(err.message || "Error al validar."); }
    finally { setValidSaving(false); }
  }

  const meta    = session?.user?.user_metadata;
  const isAdmin = meta?.role === "admin";
  const empresa = meta?.empresa_id || "";
  const initials= (session?.user?.email || "U").substring(0, 2).toUpperCase();
  const inp     = { padding: "6px 10px", fontSize: 12, border: `0.5px solid ${BORDER}`, borderRadius: 8, background: "white", color: GRAY_900, outline: "none" };

  const filtrosActivos = fNroSpot || fRutas || fPlaca || fEstadoDoc || fEstFinal || fProveedor;

  if (loading) return <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}><p style={{ color: GRAY_500, fontSize: 13 }}>Cargando...</p></div>;

  if (!session) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: GRAY_100 }}>
      <div style={{ background: "white", border: `0.5px solid ${BORDER}`, borderRadius: 16, padding: "36px 36px", width: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <img src="/logo_fape.png" alt="FP" style={{ width: 48, height: 48, borderRadius: 10, objectFit: "contain", marginBottom: 12 }} />
          <div style={{ fontSize: 22, fontWeight: 500, color: RED, letterSpacing: "-.3px" }}>Pharma<span style={{ color: GRAY_900 }}>SPOT</span></div>
          <div style={{ fontSize: 13, color: GRAY_500, marginTop: 4 }}>Seguimiento Adicionales</div>
        </div>
        <form onSubmit={doLogin}>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Correo</div>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box", padding: "8px 10px" }} />
          </div>
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Contraseña</div>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box", padding: "8px 10px" }} />
          </div>
          {loginErr && <div style={{ fontSize: 11, color: RED, marginBottom: 10 }}>{loginErr}</div>}
          <button type="submit" style={{ width: "100%", padding: 9, background: RED, color: "white", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 500, cursor: "pointer" }}>Ingresar</button>
        </form>
      </div>
    </div>
  );

  // Columnas visibles según rol
  const colsVisibles = isAdmin
    ? COLS.filter(c => !c.adminOnly || isAdmin)
    : COLS_TRANSPORTISTA_PRINCIPAL.map(key => COLS.find(c => c.key === key)).filter(Boolean);
  const colsDetalleExtra = COLS_TRANSPORTISTA_DETALLE.map(key => COLS.find(c => c.key === key)).filter(Boolean);

  return (
    <div style={{ minHeight: "100vh", background: GRAY_100, display: "flex", flexDirection: "column" }}>
      <style>{`
        .viaje-row td { background: white; transition: background 0.08s; }
        .viaje-row:hover td { background: #FFF5F5; }
      `}</style>

      {/* TOPBAR */}
      <div style={{ height: 54, background: RED, display: "flex", alignItems: "center", padding: "0 18px", gap: 12, flexShrink: 0 }}>
        <img src="/logo_fape.png" alt="FP" style={{ width: 36, height: 36, borderRadius: 7, background: "white", objectFit: "contain", padding: 2, flexShrink: 0 }} />
        <span style={{ fontSize: 19, fontWeight: 700, color: "white", letterSpacing: "-.3px" }}>Pharma<span style={{ opacity: .5, fontWeight: 400 }}>SPOT</span></span>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          <div ref={userMenuRef} style={{ position: "relative" }}>
            <button onClick={() => setUserMenuOpen(o => !o)}
              style={{ width: 32, height: 32, borderRadius: "50%", background: "rgba(255,255,255,.2)", border: "1.5px solid rgba(255,255,255,.35)", color: "white", fontSize: 11, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
              {initials}
            </button>
            {userMenuOpen && (
              <div style={{ position: "absolute", top: 38, right: 0, background: "white", borderRadius: 10, border: `0.5px solid ${BORDER}`, minWidth: 200, padding: "8px 0", zIndex: 30, boxShadow: "0 4px 20px rgba(0,0,0,.1)" }}>
                <div style={{ padding: "6px 14px 10px", borderBottom: `0.5px solid ${BORDER}` }}>
                  <div style={{ fontSize: 12, fontWeight: 500, color: GRAY_900 }}>{session.user.email}</div>
                  <div style={{ fontSize: 10, color: GRAY_500, marginTop: 2 }}>{isAdmin ? "Administrador" : empresa}</div>
                </div>
                <button onClick={abrirModalPassword}
                  style={{ width: "100%", padding: "8px 14px", background: "none", border: "none", textAlign: "left", fontSize: 12, color: GRAY_900, cursor: "pointer" }}>
                  Cambiar contraseña
                </button>
                <button onClick={() => supabase.auth.signOut()}
                  style={{ width: "100%", padding: "8px 14px", background: "none", border: "none", textAlign: "left", fontSize: 12, color: RED, cursor: "pointer" }}>
                  Cerrar sesión
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* TOOLBAR */}
      <div style={{ background: "white", borderBottom: `0.5px solid ${BORDER}`, padding: "7px 18px", display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
        <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 1 }}>
          <span style={{ fontSize: 13, fontWeight: 600, color: GRAY_900, lineHeight: 1 }}>
            {isAdmin ? "Seguimiento General de Adicionales" : `${empresa} — Seguimiento de Adicionales`}
          </span>
          {filtrosActivos && (
            <span style={{ fontSize: 10, fontWeight: 500, color: RED }}>Filtros activos</span>
          )}
        </div>
        {isAdmin && (
          <button onClick={() => setVista(vista === "tabla" ? "dashboard" : "tabla")}
            style={{ height: 34, padding: "0 16px", borderRadius: 999, border: `0.5px solid ${vista === "dashboard" ? RED : BORDER}`, background: vista === "dashboard" ? RED : "white", color: vista === "dashboard" ? "white" : GRAY_900, cursor: "pointer", fontSize: 12, fontWeight: 500 }}>
            {vista === "dashboard" ? "Ver tabla" : "Ver dashboard"}
          </button>
        )}
        {vista === "tabla" && (
          <>
            {/* Exportar Excel */}
            <button onClick={exportarExcel} title="Exportar a Excel"
              style={{ height: 34, padding: "0 14px", borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_900, cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Exportar
            </button>

            <button onClick={clearFilters} title="Limpiar filtros"
              style={{ width: 34, height: 34, borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 15, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            <button onClick={fetchViajes} title="Actualizar"
              style={{ width: 34, height: 34, borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>↻</button>
            <button onClick={openFilter}
              style={{ height: 34, padding: "0 16px", borderRadius: 999, border: `0.5px solid ${filterOpen ? RED : BORDER}`, background: filterOpen ? RED : "white", color: filterOpen ? "white" : GRAY_900, cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}>
              ⚙ Filtrar
            </button>
          </>
        )}
      </div>

      {/* DRAWER LATERAL */}
      {vista === "tabla" && filterOpen && (
        <div style={{ position: "fixed", inset: 0, zIndex: 40 }}>
          <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,.18)" }} onClick={() => setFilterOpen(false)} />
          <div style={{ position: "absolute", top: 0, right: 0, bottom: 0, width: 320, background: "white", boxShadow: "-4px 0 24px rgba(0,0,0,.12)", display: "flex", flexDirection: "column", zIndex: 41 }}
            onClick={e => e.stopPropagation()}>
            <div style={{ padding: "18px 20px 14px", borderBottom: `0.5px solid ${BORDER}`, display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: GRAY_900 }}>Opciones de Filtros</div>
              <button onClick={() => setFilterOpen(false)} style={{ width: 28, height: 28, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 14, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "18px 20px", display: "flex", flexDirection: "column", gap: 18 }}
              onKeyDown={e => e.key === "Enter" && applyFilters()}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>N° SPOT</div>
                <input value={dNroSpot} onChange={e => setDNroSpot(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }} />
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>N° GR</div>
                <input value={dRutas} onChange={e => setDRutas(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }} />
              </div>
              {isAdmin && (
                <div>
                  <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>Transportista</div>
                  <select value={dProveedor} onChange={e => setDProveedor(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }}>
                    <option value="">Todos</option>
                    {["MLT","ANDI","TRANSA","JEDA","MUNDO","INDUAMERICA","LELY","RICPAL","HUAYRAZ","A&S","MAKOOL","BSC","E&S","RANSA"].map(p => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                </div>
              )}
              <div>
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>N° Placa</div>
                <input value={dPlaca} onChange={e => setDPlaca(e.target.value.toUpperCase())} maxLength={7} style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace" }} />
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>Estado doc</div>
                <select value={dEstadoDoc} onChange={e => setDEstadoDoc(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }}>
                  <option value="">Todos</option>
                  <option value="Completo">Completo</option>
                  <option value="Pendiente">Pendiente</option>
                </select>
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>Estado de Viaje Final</div>
                <select value={dEstFinal} onChange={e => setDEstFinal(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }}>
                  <option value="">Todos</option>
                  <option value="FINALIZADO">FINALIZADO</option>
                  <option value="PENDIENTE">PENDIENTE</option>
                </select>
              </div>
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
                  <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900 }}>Rango de fecha</div>
                  <div style={{ fontSize: 10, color: GRAY_500 }}>Máx. {isAdmin ? 7 : 31} días</div>
                </div>
                <RangePicker desde={dDesde} hasta={dHasta} maxDias={isAdmin ? 7 : 31} onChange={({ desde, hasta }) => { setDDesde(desde); setDHasta(hasta); setFechaErr(""); }} />
                {fechaErr && (
                  <div style={{ marginTop: 8, padding: "7px 10px", background: RED_LIGHT, border: `0.5px solid #f7c1c1`, borderRadius: 7, fontSize: 11, color: RED_DARK }}>
                    ⚠ {fechaErr}
                  </div>
                )}
              </div>
            </div>
            <div style={{ padding: "14px 20px", borderTop: `0.5px solid ${BORDER}`, display: "flex", gap: 8, flexShrink: 0 }}>
              <button onClick={clearFilters} style={{ flex: 1, padding: "9px 0", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500, fontWeight: 500 }}>Limpiar</button>
              <button onClick={applyFilters} style={{ flex: 2, padding: "9px 0", background: RED, color: "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Buscar</button>
            </div>
          </div>
        </div>
      )}

      {/* TABLA */}
      {vista === "tabla" && (
      <div style={{ flex: 1, padding: "14px 18px", overflow: "hidden", display: "flex", flexDirection: "column" }}>
        <div style={{ overflowX: "auto", overflowY: "auto", flex: 1, background: "white", borderRadius: 10, border: `0.5px solid ${BORDER}` }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, tableLayout: "fixed" }}>
            <colgroup>
              {colsVisibles.map(c => <col key={c.key} style={{ width: c.width || 100 }} />)}
              {!isAdmin && <col style={{ width: 74 }} />}
              <col style={{ width: 36 }} />
              <col style={{ width: 80 }} />
              {!isAdmin && <col style={{ width: 84 }} />}
            </colgroup>
            <thead>
              <tr>
                {colsVisibles.map(c => (
                  <th key={c.key} style={{ padding: "8px 11px", textAlign: c.right ? "right" : "left", fontSize: 10, fontWeight: 500,
                    color: c.headerGroup === "ia" ? BLUE : c.headerGroup === "transportista" ? AMBER : GRAY_500,
                    background: c.headerGroup === "ia" ? "#EEF4FC" : c.headerGroup === "transportista" ? "#FFF8EC" : GRAY_50,
                    borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", position: "sticky", top: 0, zIndex: 3, textTransform: "uppercase", letterSpacing: ".03em" }}>
                    {c.label}
                  </th>
                ))}
                {!isAdmin && <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: 200, zIndex: 4, borderLeft: `0.5px solid ${BORDER}`, textTransform: "uppercase", letterSpacing: ".03em" }}>Detalle</th>}
                <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: isAdmin ? 80 : 164, zIndex: 4, borderLeft: `0.5px solid ${BORDER}`, textTransform: "uppercase", letterSpacing: ".03em" }}>Edit</th>
                <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: isAdmin ? 0 : 84, zIndex: 4, borderLeft: `0.5px solid ${BORDER}`, textTransform: "uppercase", letterSpacing: ".03em" }}>Foto</th>
                {!isAdmin && <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: 0, zIndex: 4, textTransform: "uppercase", letterSpacing: ".03em", borderLeft: `0.5px solid ${BORDER}` }}>Doc. Adjunto</th>}
              </tr>
            </thead>
            <tbody>
              {viajes.length === 0 ? (
                <tr><td colSpan={colsVisibles.length + (isAdmin ? 2 : 4)} style={{ padding: 40, textAlign: "center", color: GRAY_500 }}>No hay registros para el rango seleccionado</td></tr>
              ) : viajes.map(v => {
                const completo = (v.foto_versiones?.length || 0) > 0;

                return (
                  <tr key={v.nro_spot} className="viaje-row">
                    {colsVisibles.map(c => {
                      const editable = !isAdmin && (c.key === "placa" || c.key === "rutas");
                      let content = v[c.key];
                      if (c.key === "estado_final") {
                        const ef = v.estado_final;
                        const bg = ef === "FINALIZADO" ? GREEN_LIGHT : ef === "PENDIENTE" ? AMBER_LIGHT : GRAY_100;
                        const fg = ef === "FINALIZADO" ? GREEN : ef === "PENDIENTE" ? AMBER : GRAY_500;
                        content = ef
                          ? <span style={{ display: "inline-flex", padding: "2px 8px", borderRadius: 999, fontSize: 10, fontWeight: 600, background: bg, color: fg, whiteSpace: "nowrap" }}>{ef}</span>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (["realizado","estado_doc","estado_procesamiento_ia","estado_validacion_ia"].includes(c.key)) {
                        content = v[c.key]
                          ? <span style={{ fontSize: 11, color: GRAY_900 }}>{String(v[c.key]).toUpperCase()}</span>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "fecha_modif" || c.key === "fecha_entrega_doc") {
                        content = v[c.key] ? fmtFechaHora(v[c.key]) : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "fecha_carga") {
                        content = v[c.key] ? fmtFechaSolo(v[c.key]) : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "hora_cita") {
                        // Google Sheets guarda columnas "solo hora" con una fecha base (epoch 1899),
                        // así que el valor puede llegar como "Sat Dec 30 1899 09:00:00..." — se extrae
                        // solo el HH:MM sin importar el resto del texto.
                        const m = v[c.key] ? String(v[c.key]).match(/(\d{1,2}:\d{2})/) : null;
                        content = m ? m[1] : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (content === null || content === undefined || content === "") {
                        content = <span style={{ color: GRAY_200 }}>—</span>;
                      }
                      return (
                        <td key={c.key} title={(v[c.key] !== null && v[c.key] !== undefined && v[c.key] !== "") ? String(v[c.key]) : undefined}
                          data-editable={editable ? "1" : undefined}
                          style={{ padding: "8px 11px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontFamily: c.mono ? "monospace" : "inherit", fontSize: c.mono ? 10 : 11, textAlign: c.right ? "right" : "left", color: c.muted ? GRAY_500 : GRAY_900 }}>
                          {content}
                        </td>
                      );
                    })}

                    {/* Botón Detalle — solo transportista, abre modal con campos extendidos */}
                    {!isAdmin && (
                      <td style={{ padding: "4px 6px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: 200, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center" }}>
                        <button onClick={() => setDetalleModal(v)} title="Ver detalle completo"
                          style={{ padding: "5px 10px", borderRadius: 7, border: `1px solid ${BORDER}`, background: GRAY_50, cursor: "pointer", fontSize: 10, fontWeight: 500, color: GRAY_900 }}>
                          Detalle
                        </button>
                      </td>
                    )}

                    <td style={{ padding: "4px 6px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: isAdmin ? 80 : 164, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center" }}>
                      {!isAdmin && (
                        <button onClick={() => openEditModal(v)} title="Editar Placa y Rutas"
                          style={{ width: 28, height: 28, borderRadius: 7, border: `1.5px solid ${GRAY_900}`, background: "white", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto", padding: 0 }}>
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={GRAY_900} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                          </svg>
                        </button>
                      )}
                      {isAdmin && v.estado_final !== "FINALIZADO" && (
                        <button onClick={() => { setValidModal(v); setValidErr(""); }} title="Aprobar adicional"
                          style={{ width: 28, height: 28, borderRadius: 7, border: `1.5px solid ${BLUE}`, background: "white", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={BLUE} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                            <polyline points="20 6 9 17 4 12"/>
                          </svg>
                        </button>
                      )}
                    </td>

                    {/* Botones foto — visibles para todos si hay foto */}
                    <td style={{ padding: "4px 6px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: isAdmin ? 0 : 84, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center" }}>
                      {completo && v.foto_url && (
                        <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                          <button onClick={() => abrirFoto(v.foto_url)} title="Ver foto"
                            style={{ width: 26, height: 26, borderRadius: 6, border: `1px solid ${BORDER}`, background: "white", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: GRAY_500 }}>
                            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                              <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                              <circle cx="12" cy="12" r="3"/>
                            </svg>
                          </button>
                          <button onClick={() => descargarFoto(v.foto_url, v.foto_nombre)} title="Descargar foto"
                            style={{ width: 26, height: 26, borderRadius: 6, border: `1px solid ${BORDER}`, background: "white", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: GRAY_500 }}>
                            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                              <polyline points="7 10 12 15 17 10"/>
                              <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                          </button>
                        </div>
                      )}
                    </td>

                    {/* Doc. adjuntos — solo transportista */}
                    {!isAdmin && (
                      <td style={{ padding: "8px 6px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: 0, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center" }}>
                        <button onClick={() => openModal(v)}
                          style={{ width: 28, height: 28, borderRadius: "50%", border: `1.5px solid ${completo ? GREEN : GRAY_200}`, background: "white", color: completo ? GREEN : GRAY_500, cursor: "pointer", fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto" }}>
                          {completo ? "✓" : "↑"}
                        </button>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      )}

      {/* DASHBOARD */}
      {vista === "dashboard" && isAdmin && (
        <div style={{ flex: 1, padding: "14px 18px", overflow: "auto" }}>

          {/* Filtros del dashboard */}
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16, paddingBottom: 14, borderBottom: `0.5px solid ${BORDER}` }}>
            <div>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Transportista</div>
              <select value={dashProveedor} onChange={e => setDashProveedor(e.target.value)} style={{ ...inp, width: 170 }}>
                <option value="">Todos</option>
                {["MLT","ANDI","TRANSA","JEDA","MUNDO","INDUAMERICA","LELY","RICPAL","HUAYRAZ","A&S","MAKOOL","BSC","E&S","RANSA"].map(p => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </div>
            <div>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Fecha registro desde</div>
              <input type="date" value={dashDesde} onChange={e => setDashDesde(e.target.value)} style={{ ...inp, width: 150 }} />
            </div>
            <div>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Fecha registro hasta</div>
              <input type="date" value={dashHasta} onChange={e => setDashHasta(e.target.value)} style={{ ...inp, width: 150 }} />
            </div>
            <button onClick={limpiarFiltrosDashboard} style={{ height: 34, padding: "0 14px", borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 12 }}>
              Limpiar filtros
            </button>
            <button onClick={() => fetchDashboard(detallePagina)} title="Actualizar"
              style={{ width: 34, height: 34, borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 16 }}>↻</button>
            {dashLoading && <span style={{ fontSize: 11, color: GRAY_500 }}>Actualizando…</span>}
          </div>

          {kpis && (
            <>
              {/* Tarjetas de KPI */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 24 }}>
                <div style={{ background: GRAY_50, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: GRAY_500, marginBottom: 6 }}>Tickets generados</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: GRAY_900 }}>{kpis.total_tickets}</div>
                </div>
                <div style={{ background: GRAY_50, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: GRAY_500, marginBottom: 6 }}>Validados (ejecución)</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: GRAY_900 }}>{kpis.tickets_validados_ejecucion} <span style={{ fontSize: 13, color: GRAY_500, fontWeight: 400 }}>/ {kpis.total_tickets}</span></div>
                </div>
                <div style={{ background: GRAY_50, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: GRAY_500, marginBottom: 6 }}>Escaneo transportista</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: GRAY_900 }}>{kpis.tickets_con_foto} <span style={{ fontSize: 13, color: GRAY_500, fontWeight: 400 }}>/ {kpis.tickets_realizados}</span></div>
                </div>
                <div style={{ background: RED_LIGHT, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: RED_DARK, marginBottom: 6 }}>Rechazados por IA</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: RED_DARK }}>{kpis.tickets_rechazados_ia}</div>
                </div>
                <div style={{ background: GREEN_LIGHT, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: GREEN, marginBottom: 6 }}>Listos para migrar</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: GREEN }}>{kpis.tickets_listos_migrar}</div>
                </div>
              </div>

              {/* Cascada — barra horizontal segmentada */}
              <div style={{ marginBottom: 28 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: GRAY_900, marginBottom: 10 }}>Cascada de tickets</div>
                {(() => {
                  const total = kpis.total_tickets || 1;
                  const segs = [
                    { label: "No realizados", val: kpis.cascada_no_realizados, color: GRAY_500 },
                    { label: "Pendiente subir", val: kpis.cascada_pendiente_subir, color: AMBER },
                    { label: "Pendiente validar", val: kpis.cascada_pendiente_validar, color: RED_DARK },
                    { label: "Listos para migrar", val: kpis.tickets_listos_migrar, color: GREEN },
                  ];
                  return (
                    <>
                      <div style={{ display: "flex", width: "100%", height: 28, borderRadius: 6, overflow: "hidden" }}>
                        {segs.map(s => (
                          <div key={s.label} title={`${s.label}: ${s.val}`} style={{ width: `${(s.val / total) * 100}%`, background: s.color, minWidth: s.val > 0 ? 3 : 0 }} />
                        ))}
                      </div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 16, marginTop: 10, fontSize: 11, color: GRAY_500 }}>
                        {segs.map(s => (
                          <span key={s.label} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            <span style={{ width: 9, height: 9, borderRadius: 2, background: s.color, display: "inline-block" }} />
                            {s.label}: {s.val}
                          </span>
                        ))}
                      </div>
                    </>
                  );
                })()}
              </div>
            </>
          )}

          {/* Ranking por transportista */}
          <div style={{ marginBottom: 28 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: GRAY_900, marginBottom: 10 }}>Ranking por transportista</div>
            <div style={{ overflowX: "auto", background: "white", borderRadius: 10, border: `0.5px solid ${BORDER}` }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead>
                  <tr style={{ background: GRAY_50 }}>
                    {["Proveedor","Total","Realizados","No realiz.","Pend. subir","Pend. validar","Listos migrar","% Avance"].map(h => (
                      <th key={h} style={{ padding: "8px 10px", textAlign: h === "Proveedor" ? "left" : "right", fontSize: 10, fontWeight: 500, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".03em", borderBottom: `0.5px solid ${BORDER}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {ranking.map(r => (
                    <tr key={r.proveedor}>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}` }}>{r.proveedor}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}` }}>{r.total}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}` }}>{r.realizados}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}` }}>{r.no_realizados}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}` }}>{r.pendiente_subir}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}`, color: r.pendiente_validar > 0 ? RED_DARK : GRAY_900 }}>{r.pendiente_validar}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}`, color: GREEN }}>{r.listos_migrar}</td>
                      <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: `0.5px solid ${BORDER}`, fontWeight: 500 }}>{r.pct_avance ?? 0}%</td>
                    </tr>
                  ))}
                  {ranking.length === 0 && (
                    <tr><td colSpan={8} style={{ padding: 24, textAlign: "center", color: GRAY_500 }}>Sin datos para el rango seleccionado</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Detalle de tickets filtrados */}
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: GRAY_900 }}>Detalle de tickets filtrados</div>
              <div style={{ fontSize: 11, color: GRAY_500 }}>Mostrando {detalle.length ? detallePagina * DETALLE_POR_PAGINA + 1 : 0}-{detallePagina * DETALLE_POR_PAGINA + detalle.length} de {detalleTotal}</div>
            </div>
            <div style={{ overflowX: "auto", background: "white", borderRadius: 10, border: `0.5px solid ${BORDER}` }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead>
                  <tr style={{ background: GRAY_50 }}>
                    {["N° SPOT","Proveedor","Fecha servicio","Realizado","N° GR","Validación IA","Estado final"].map(h => (
                      <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 500, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".03em", borderBottom: `0.5px solid ${BORDER}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {detalle.map(t => (
                    <tr key={t.nro_spot}>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}`, fontFamily: "monospace", fontSize: 10 }}>{t.nro_spot}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}` }}>{t.proveedor || "—"}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}` }}>{t.fecha_carga ? fmtFechaSolo(t.fecha_carga) : "—"}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}` }}>{t.realizado || "—"}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}` }}>{t.rutas || "—"}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}`, color: t.estado_validacion_ia === "COINCIDE" ? GREEN : t.estado_validacion_ia === "NO_COINCIDE" ? RED_DARK : GRAY_500 }}>{t.estado_validacion_ia || "—"}</td>
                      <td style={{ padding: "8px 10px", borderBottom: `0.5px solid ${BORDER}`, fontWeight: 500 }}>{t.estado_final || "—"}</td>
                    </tr>
                  ))}
                  {detalle.length === 0 && (
                    <tr><td colSpan={7} style={{ padding: 24, textAlign: "center", color: GRAY_500 }}>Sin registros</td></tr>
                  )}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
              <button onClick={() => fetchDashboard(Math.max(0, detallePagina - 1))} disabled={detallePagina === 0}
                style={{ padding: "6px 14px", borderRadius: 8, border: `0.5px solid ${BORDER}`, background: "white", fontSize: 12, cursor: detallePagina === 0 ? "default" : "pointer", color: detallePagina === 0 ? GRAY_200 : GRAY_900 }}>Anterior</button>
              <span style={{ fontSize: 11, color: GRAY_500 }}>Página {detallePagina + 1} de {Math.max(1, Math.ceil(detalleTotal / DETALLE_POR_PAGINA))}</span>
              <button onClick={() => fetchDashboard(detallePagina + 1)} disabled={(detallePagina + 1) * DETALLE_POR_PAGINA >= detalleTotal}
                style={{ padding: "6px 14px", borderRadius: 8, border: `0.5px solid ${BORDER}`, background: "white", fontSize: 12, cursor: (detallePagina + 1) * DETALLE_POR_PAGINA >= detalleTotal ? "default" : "pointer", color: (detallePagina + 1) * DETALLE_POR_PAGINA >= detalleTotal ? GRAY_200 : GRAY_900 }}>Siguiente</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL EDICIÓN PLACA / RUTAS */}
      {/* MODAL CAMBIAR CONTRASEÑA */}
      {pwModalOpen && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && !pwSaving && setPwModalOpen(false)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 360, maxWidth: "94vw" }}>
            {pwOk ? (
              <div style={{ textAlign: "center", padding: "20px 0" }}>
                <div style={{ fontSize: 32, marginBottom: 8 }}>✅</div>
                <div style={{ fontSize: 13, color: GREEN, fontWeight: 500 }}>Contraseña actualizada</div>
              </div>
            ) : (
              <>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18 }}>
                  <div style={{ fontSize: 14, fontWeight: 500 }}>Cambiar contraseña</div>
                  <button onClick={() => !pwSaving && setPwModalOpen(false)} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                </div>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Nueva contraseña</div>
                  <input type="password" value={pwNueva} onChange={e => { setPwNueva(e.target.value); setPwErr(""); }}
                    onKeyDown={e => e.key === "Enter" && !pwSaving && cambiarPassword()}
                    style={{ ...inp, width: "100%", boxSizing: "border-box" }} />
                </div>
                <div style={{ marginBottom: 6 }}>
                  <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Confirmar contraseña</div>
                  <input type="password" value={pwConfirma} onChange={e => { setPwConfirma(e.target.value); setPwErr(""); }}
                    onKeyDown={e => e.key === "Enter" && !pwSaving && cambiarPassword()}
                    style={{ ...inp, width: "100%", boxSizing: "border-box" }} />
                </div>
                <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 12 }}>Mínimo 6 caracteres.</div>
                {pwErr && <div style={{ fontSize: 11, color: RED, marginBottom: 10 }}>{pwErr}</div>}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                  <button onClick={() => setPwModalOpen(false)} disabled={pwSaving}
                    style={{ padding: "7px 16px", background: "none", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: pwSaving ? "default" : "pointer", color: GRAY_500 }}>
                    Cancelar
                  </button>
                  <button onClick={cambiarPassword} disabled={pwSaving || !pwNueva || !pwConfirma}
                    style={{ padding: "7px 16px", background: (pwSaving || !pwNueva || !pwConfirma) ? GRAY_200 : RED, color: (pwSaving || !pwNueva || !pwConfirma) ? GRAY_500 : "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: (pwSaving || !pwNueva || !pwConfirma) ? "default" : "pointer" }}>
                    {pwSaving ? "Guardando..." : "Guardar"}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* MODAL DETALLE — vista transportista, campos extendidos fuera del panel principal */}
      {detalleModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && setDetalleModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 420, maxWidth: "94vw", maxHeight: "90vh", overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
              <div style={{ fontSize: 14, fontWeight: 500 }}>Detalle del viaje</div>
              <button onClick={() => setDetalleModal(null)} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 18, fontFamily: "monospace" }}>{detalleModal.nro_spot}</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {colsDetalleExtra.map(c => {
                let val = detalleModal[c.key];
                if (c.key === "fecha_entrega_doc") {
                  val = val ? fmtFechaHora(val) : "—";
                } else if (c.key === "realizado" || c.key === "estado_doc") {
                  val = val ? String(val).toUpperCase() : "—";
                } else if (c.right) {
                  val = (val === null || val === undefined || val === "") ? "—" : val;
                } else if (val === null || val === undefined || val === "") {
                  val = "—";
                }
                return (
                  <div key={c.key}>
                    <div style={{ fontSize: 10, color: GRAY_500, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 3 }}>{c.label}</div>
                    <div style={{ fontSize: 13, color: GRAY_900 }}>{val}</div>
                  </div>
                );
              })}
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 20 }}>
              <button onClick={() => setDetalleModal(null)} style={{ padding: "7px 20px", background: GRAY_100, border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", color: GRAY_900 }}>Cerrar</button>
            </div>
          </div>
        </div>
      )}

      {editModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && !editSaving && setEditModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 420, maxWidth: "94vw", maxHeight: "90vh", overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
              <div style={{ fontSize: 14, fontWeight: 500 }}>Editar datos del viaje</div>
              <button onClick={() => setEditModal(null)} disabled={editSaving} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 18, fontFamily: "monospace" }}>{editModal.nro_spot}</div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° Placa</div>
              <input value={editPlaca} onChange={e => { setEditPlaca(e.target.value.toUpperCase()); setEditErr(""); }}
                maxLength={7}
                style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 13, letterSpacing: ".08em", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
            </div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° GR</div>
              <input value={editRutaInput} onChange={e => { setEditRutaInput(e.target.value); setEditErr(""); }}
                onKeyDown={e => e.key === "Enter" && (e.preventDefault(), !editSaving && saveEdit())}
                style={{ ...inp, width: "100%", boxSizing: "border-box", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>Presiona Enter para guardar.</div>
            </div>
            {editErr && <div style={{ padding: "8px 12px", background: RED_LIGHT, borderRadius: 8, fontSize: 11, color: RED_DARK, marginBottom: 12, border: `0.5px solid #f7c1c1` }}>⚠ {editErr}</div>}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setEditModal(null)} disabled={editSaving} style={{ padding: "7px 14px", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500 }}>Cancelar</button>
              <button onClick={saveEdit} disabled={editSaving} style={{ padding: "7px 16px", background: editSaving ? GRAY_200 : RED, color: editSaving ? GRAY_500 : "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: editSaving ? "default" : "pointer" }}>
                {editSaving ? "Guardando..." : "Guardar"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL VALIDACIÓN MANUAL ADMIN */}
      {validModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && !validSaving && setValidModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 380, maxWidth: "94vw" }}>
            {/* Header con ícono */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 12, marginBottom: 16 }}>
              <div style={{ width: 38, height: 38, borderRadius: "50%", background: BLUE_LIGHT, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={BLUE} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                  <line x1="12" y1="9" x2="12" y2="13"/>
                  <line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>
              </div>
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: GRAY_900, textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 4 }}>Aprobar Adicional</div>
                <div style={{ fontSize: 12, color: GRAY_500 }}>
                  ¿Estás seguro que desea confirmar la validación del viaje{" "}
                  <span style={{ fontFamily: "monospace", color: GRAY_900, fontWeight: 500 }}>{validModal.nro_spot}</span>?
                </div>
              </div>
            </div>
            {validErr && <div style={{ padding: "8px 12px", background: RED_LIGHT, borderRadius: 8, fontSize: 11, color: RED_DARK, marginBottom: 12 }}>⚠ {validErr}</div>}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setValidModal(null)} disabled={validSaving}
                style={{ padding: "7px 20px", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500 }}>No</button>
              <button onClick={handleValidarManual} disabled={validSaving}
                style={{ padding: "7px 20px", background: validSaving ? GRAY_200 : BLUE, color: validSaving ? GRAY_500 : "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: validSaving ? "default" : "pointer" }}>
                {validSaving ? "Validando..." : "Sí"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL SUBIDA DOC */}
      {modal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && setModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 390, maxWidth: "92vw", maxHeight: "88vh", overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
              <div style={{ fontSize: 14, fontWeight: 500 }}>{modal._bloqueado ? "Campos incompletos" : "Subir documento"}</div>
              <button onClick={() => setModal(null)} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 14, fontFamily: "monospace" }}>{modal.nro_spot}</div>

            {/* Vista bloqueada — faltan campos */}
            {modal._bloqueado ? (
              <div>
                <div style={{ padding: "14px 16px", background: AMBER_LIGHT, borderRadius: 10, marginBottom: 18, border: `0.5px solid #e8c87a` }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: AMBER, marginBottom: 6 }}>⚠ Completa los datos del viaje</div>
                  <div style={{ fontSize: 12, color: AMBER }}>
                    Para subir el documento primero debes registrar: <strong>{modal._mensajeBloqueo}</strong>.
                  </div>
                  <div style={{ fontSize: 11, color: AMBER, marginTop: 8 }}>
                    Usa el botón ✏️ de la fila para completar esos campos.
                  </div>
                </div>
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <button onClick={() => setModal(null)}
                    style={{ padding: "7px 20px", background: GRAY_100, border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", color: GRAY_900 }}>
                    Entendido
                  </button>
                </div>
              </div>
            ) : uploadSuccess ? (
              <div style={{ textAlign: "center", padding: "24px 0" }}>
                {!resultadoIA ? (
                  <>
                    <div style={{ fontSize: 32, marginBottom: 8 }}>✅</div>
                    <div style={{ fontSize: 13, color: GREEN, fontWeight: 500 }}>Documento guardado correctamente</div>
                    <div style={{ fontSize: 11, color: GRAY_500, marginTop: 6 }}>La validación automática se completará en breve.</div>
                  </>
                ) : resultadoIA.estado_validacion_ia === "COINCIDE" ? (
                  <>
                    <div style={{ fontSize: 32, marginBottom: 8 }}>✅</div>
                    <div style={{ fontSize: 13, color: GREEN, fontWeight: 500 }}>Documento validado correctamente</div>
                    <div style={{ fontSize: 11, color: GRAY_500, marginTop: 6 }}>Coincidencia: {resultadoIA.match_ia}%</div>
                  </>
                ) : (
                  <>
                    <div style={{ fontSize: 32, marginBottom: 8 }}>⚠️</div>
                    <div style={{ fontSize: 13, color: AMBER, fontWeight: 500 }}>El documento no coincide con lo registrado</div>
                    <div style={{ fontSize: 11, color: GRAY_500, marginTop: 6 }}>Detectado: {resultadoIA.texto_detectado_ia || "—"} ({resultadoIA.match_ia ?? 0}%)</div>
                    <div style={{ fontSize: 11, color: GRAY_500, marginTop: 4 }}>Un administrador revisará tu documento.</div>
                    <button onClick={() => { setModal(null); setResultadoIA(null); }}
                      style={{ marginTop: 14, padding: "7px 20px", background: GRAY_100, border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", color: GRAY_900 }}>
                      Entendido
                    </button>
                  </>
                )}
              </div>
            ) : validandoIA ? (
              <div style={{ textAlign: "center", padding: "24px 0" }}>
                <div style={{ fontSize: 24, marginBottom: 8 }}>⏳</div>
                <div style={{ fontSize: 13, color: GRAY_500 }}>Validando documento...</div>
              </div>
            ) : (
              <>
                {(modal.foto_versiones?.length || 0) > 0 && (
                  <div style={{ marginBottom: 14, padding: "10px 12px", background: GREEN_LIGHT, borderRadius: 8, fontSize: 12, color: GREEN }}>
                    ✓ Ya tienes {modal.foto_versiones.length} documento(s) subido(s). Puedes agregar otro si necesitas corregir.
                  </div>
                )}
                <div onClick={() => fileRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleFileSelect(e.dataTransfer.files[0]); }}
                  style={{ border: `1.5px dashed ${GRAY_200}`, borderRadius: 10, padding: "22px 16px", textAlign: "center", cursor: "pointer", marginBottom: 12 }}>
                  <div style={{ fontSize: 22, color: GRAY_200, marginBottom: 6 }}>📷</div>
                  <div style={{ fontSize: 12, color: GRAY_500 }}>{uploadFile ? uploadFile.name : "Clic o arrastra tu foto aquí"}</div>
                </div>
                {uploadFile && (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", background: GRAY_100, borderRadius: 8, fontSize: 11, color: GRAY_900, marginBottom: 12, border: `0.5px solid ${BORDER}` }}>
                    <span>📎</span><span style={{ flex: 1 }}>Archivo: {uploadFile.name}</span>
                  </div>
                )}
                {uploadErr && <div style={{ padding: "8px 12px", background: RED_LIGHT, borderRadius: 8, fontSize: 11, color: RED_DARK, marginBottom: 12, border: `0.5px solid #f7c1c1` }}>⚠ {uploadErr}</div>}
                <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => handleFileSelect(e.target.files[0])} />
                <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                  <button onClick={() => setModal(null)} style={{ padding: "7px 14px", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500 }}>Cancelar</button>
                  <button onClick={handleUpload} disabled={!uploadFile || uploading}
                    style={{ padding: "7px 16px", background: uploadFile && !uploading ? RED : GRAY_200, color: uploadFile && !uploading ? "white" : GRAY_500, border: "none", borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: uploadFile && !uploading ? "pointer" : "default" }}>
                    {uploading ? "Subiendo..." : "Guardar documento"}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
