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

// Orden de columnas según diseño acordado
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
  { key: "realizado",           label: "Realizado",                width: 80 },
  { key: "estado_procesamiento_ia", label: "Procesam. IA",         width: 100, headerGroup: "ia", adminOnly: true },
  { key: "estado_validacion_ia",label: "Validación IA",            width: 110, headerGroup: "ia" },
  { key: "match_ia",            label: "Match IA",                 width: 70,  right: true, headerGroup: "ia" },
  { key: "usuario_modif",       label: "Usuario Modif.",           width: 140, muted: true, headerGroup: "ia" },
  { key: "fecha_modif",         label: "Fecha Modif.",             width: 130, muted: true, headerGroup: "ia" },
  { key: "estado_doc",          label: "Estado Doc.",              width: 100 },
  { key: "fecha_entrega_doc",   label: "Fec. Entrega Doc.",        width: 130 },
];

const COLS_TRANSPORTISTA_PRINCIPAL = [
  "nro_spot", "fecha_carga", "estado_final", "cd_origen", "cd_destino",
  "placa", "rutas", "texto_detectado_ia", "estado_validacion_ia",
];

const COLS_TRANSPORTISTA_DETALLE = [
  "importe", "tipo_traslado", "cantidad", "requerimiento",
  "detalle_servicio", "realizado", "estado_doc", "fecha_entrega_doc",
];

// Ordena una lista de viajes por cualquier columna de COLS -- numérica si la columna
// está marcada como "right" (cantidad, importe, match_ia), alfabética/fecha en el resto
// (fecha_carga ordena bien como texto porque ya viene en formato ISO "AAAA-MM-DD").
function ordenarViajes(lista, col, dir) {
  if (!col) return lista;
  const colDef = COLS.find(c => c.key === col);
  const esNumerico = !!colDef?.right;
  const copia = [...lista];
  copia.sort((a, b) => {
    let av = a[col], bv = b[col];
    if (esNumerico) {
      av = Number(av) || 0; bv = Number(bv) || 0;
      return dir === "desc" ? bv - av : av - bv;
    }
    av = av === null || av === undefined ? "" : String(av);
    bv = bv === null || bv === undefined ? "" : String(bv);
    return dir === "desc" ? bv.localeCompare(av) : av.localeCompare(bv);
  });
  return copia;
}

function extraerNroSpotCorto(nroSpot) {
  if (!nroSpot) return "doc";
  const m = String(nroSpot).match(/(Nro\d+)/i);
  return m ? m[1] : nroSpot.replace(/\s+/g, "_").slice(0, 30);
}

async function calcularRangoAncla_(isAdminUser, empresaId) {
  let ancla = new Date(); 
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
      ancla = new Date(data[0].fecha_carga + "T00:00:00");
    }
  } catch { }

  const desde = new Date(ancla);
  desde.setDate(ancla.getDate() - (isAdminUser ? 6 : 30));
  const fmt = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return { desde: fmt(desde), hasta: fmt(ancla) };
}

async function comprimirImagen(file) {
  if (file.size > MAX_MB * 1024 * 1024) throw new Error(`El archivo supera los ${MAX_MB} MB.`);
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

  // Formato estricto DD/MM/AAAA
  const fmtLocal = (isoStr) => {
    if (!isoStr) return "";
    const m = isoStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : isoStr;
  };

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
        {rangoValido ? `${fmtLocal(fmt(selStart))} – ${fmtLocal(fmt(selEnd))}` : "Selecciona el rango de fechas"}
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

  // Generamos la fecha de ayer (D-1) como default para el dashboard
  const d1 = new Date();
  d1.setDate(d1.getDate() - 1);
  const D1_STR = `${d1.getFullYear()}-${String(d1.getMonth() + 1).padStart(2, "0")}-${String(d1.getDate()).padStart(2, "0")}`;

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
  
  // Asignamos D-1 por defecto al dashboard
  const [dashDesde, setDashDesde] = useState("");
  const [dashHasta, setDashHasta] = useState(D1_STR);
  const [dashFiltroAbierto, setDashFiltroAbierto] = useState(false);

  const [kpis, setKpis] = useState(null);
  const [ranking, setRanking] = useState([]);
  const [dashLoading, setDashLoading] = useState(false);
  const [cascadaDetalleModal, setCascadaDetalleModal] = useState(null);
  const [rankingSortCol, setRankingSortCol] = useState(null);
  const [tablaSortCol, setTablaSortCol] = useState(null);
  const [tablaSortDir, setTablaSortDir] = useState("asc");
  const [rankingSortDir, setRankingSortDir] = useState("desc");
  
  const [pwModalOpen,  setPwModalOpen]  = useState(false);
  const [pwNueva,      setPwNueva]      = useState("");
  const [pwConfirma,   setPwConfirma]   = useState("");
  const [pwErr,        setPwErr]        = useState("");
  const [pwSaving,     setPwSaving]     = useState(false);
  const [pwOk,         setPwOk]         = useState(false);
  const [editPlaca,    setEditPlaca]    = useState("");
  const [editRutaInput,setEditRutaInput]= useState("");
  const [editPlaca2,   setEditPlaca2]   = useState("");
  const [editRuta2Input,setEditRuta2Input]= useState("");
  const [editPlaca3,   setEditPlaca3]   = useState("");
  const [editRuta3Input,setEditRuta3Input]= useState("");
  const [editDocsVisibles, setEditDocsVisibles] = useState(1);
  const [tipoVehiculo, setTipoVehiculo] = useState({ 1: false, 2: false, 3: false }); // false=camión (3-3), true=moto (4-2)
  const [editSaving,   setEditSaving]   = useState(false);
  const [editErr,      setEditErr]      = useState("");
  // Estado de subida por documento: uploadState[1|2|3] = {file, err, uploading, validando, resultado, success}
  const [uploadState,  setUploadState]  = useState({});

  const [validModal,   setValidModal]   = useState(null);
  const [validSaving,  setValidSaving]  = useState(false);
  const [validErr,     setValidErr]     = useState("");

  const fileRef1   = useRef();
  const fileRef2   = useRef();
  const fileRef3   = useRef();
  const fileRefFor = (d) => d === 1 ? fileRef1 : d === 2 ? fileRef2 : fileRef3;
  const userMenuRef= useRef();
  const fetchViajesRef = useRef(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setLoading(false); });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    const isAdminUser = session.user.user_metadata?.role === "admin";
    const rolActual = isAdminUser ? "admin" : "transportista";
    if (_saved.fDesde && _saved.fHasta && _saved._rol === rolActual) return;
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
    q = q.eq("realizado", "SI");
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

  function openEditModal(viaje) {
    setEditModal(viaje);
    setEditPlaca(viaje.placa || "");
    setEditRutaInput(viaje.rutas || "");
    setEditPlaca2(viaje.placa_2 || "");
    setEditRuta2Input(viaje.rutas_2 || "");
    setEditPlaca3(viaje.placa_3 || "");
    setEditRuta3Input(viaje.rutas_3 || "");
    setEditDocsVisibles(viaje.placa_3 ? 3 : viaje.placa_2 ? 2 : 1);
    setTipoVehiculo({ 1: inferirMoto(viaje.placa), 2: inferirMoto(viaje.placa_2), 3: inferirMoto(viaje.placa_3) });
    setEditErr("");
  }

  // Acepta 3-3 (camión, formato estándar para todos) o 4-2 (moto, solo aplica hoy para MUNDO)
  const PLACA_RE = /^[A-Za-z0-9]{3}-[A-Za-z0-9]{3}$|^[A-Za-z0-9]{4}-[A-Za-z0-9]{2}$/;

  // Va insertando el guión solo mientras el usuario escribe -- sin importar si el carácter es
  // letra o número, solo importa la posición. esMoto cambia el punto de corte (4 en vez de 3).
  function formatearPlaca(valor, esMoto) {
    const limpio = valor.toUpperCase().replace(/[^A-Z0-9]/g, "");
    const p1 = esMoto ? 4 : 3;
    const p2 = esMoto ? 2 : 3;
    const seg1 = limpio.slice(0, p1);
    const seg2 = limpio.slice(p1, p1 + p2);
    return limpio.length >= p1 ? `${seg1}-${seg2}` : seg1;
  }

  // Al abrir un ticket ya existente, se deduce si es moto con solo mirar el largo del primer
  // segmento de la placa guardada (4 caracteres = moto) -- no hace falta guardar esto en la BD.
  function inferirMoto(placa) {
    if (!placa) return false;
    return (placa.split("-")[0] || "").length === 4;
  }

  async function saveEdit() {
    if (editPlaca && !PLACA_RE.test(editPlaca)) { setEditErr("Formato inválido en Placa 1. Usa ABC-123."); return; }
    if (editDocsVisibles >= 2 && editPlaca2 && !PLACA_RE.test(editPlaca2)) { setEditErr("Formato inválido en Placa 2. Usa ABC-123."); return; }
    if (editDocsVisibles >= 3 && editPlaca3 && !PLACA_RE.test(editPlaca3)) { setEditErr("Formato inválido en Placa 3. Usa ABC-123."); return; }
    setEditSaving(true); setEditErr("");
    try {
      const payload = {
        placa: editPlaca || null, rutas: editRutaInput.trim() || null,
        placa_2: editDocsVisibles >= 2 ? (editPlaca2 || null) : null,
        rutas_2: editDocsVisibles >= 2 ? (editRuta2Input.trim() || null) : null,
        placa_3: editDocsVisibles >= 3 ? (editPlaca3 || null) : null,
        rutas_3: editDocsVisibles >= 3 ? (editRuta3Input.trim() || null) : null,
      };
      const { data, error } = await supabase
        .from("viajes")
        .update(payload)
        .eq("nro_spot", editModal.nro_spot)
        .select();
      if (error) throw error;
      setViajes(prev => prev.map(v =>
        v.nro_spot === editModal.nro_spot ? { ...v, ...payload } : v
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

  // Estricto DD/MM/AAAA para presentación
  const fmtFechaSolo = (val) => {
    if (!val) return "";
    const m = String(val).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[3]}/${m[2]}/${m[1]}`;
    return String(val);
  };

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
    h = h % 12 || 12; 
    return `${dd}/${mm}/${aaaa} ${h}:${min} ${ampm}`;
  };

  function exportarExcel() {
    if (!viajes.length) return;
    const colsExp = COLS.filter(c => !c.adminOnly || isAdmin);

    // En la tabla, N°GR/Placa/Texto Detectado IA/Validación IA se ven consolidados con " | "
    // cuando el ticket tiene varios documentos -- pero en el Excel cada uno debe quedar en
    // su propia columna, para no perder el detalle exacto de cada guía.
    const CAMPOS_MULTIDOC = ["rutas", "placa", "texto_detectado_ia", "estado_validacion_ia"];
    const columnasFinal = [];
    colsExp.forEach(c => {
      columnasFinal.push(c);
      if (CAMPOS_MULTIDOC.includes(c.key)) {
        columnasFinal.push({ key: `${c.key}_2`, label: `${c.label} 2` });
        columnasFinal.push({ key: `${c.key}_3`, label: `${c.label} 3` });
      }
    });

    const headers = columnasFinal.map(c => c.label);
    const rows = viajes.map(v => columnasFinal.map(c => {
      const val = v[c.key];
      if (val === null || val === undefined) return "";
      if (c.key === "fecha_carga")
        return fmtFechaSolo(val);
      if (c.key === "fecha_modif" || c.key === "fecha_entrega_doc")
        return fmtFecha(val);
      if (["realizado", "estado_final", "estado_procesamiento_ia", "estado_validacion_ia", "estado_validacion_ia_2", "estado_validacion_ia_3"].includes(c.key))
        return val ? String(val).toUpperCase() : "";
      if (c.key === "foto_versiones") return Array.isArray(val) ? val.length : 0;
      return String(val);
    }));

    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

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
      window.open(data.signedUrl, "_blank");
    }
  }

  const SUF = (d) => d === 1 ? "" : `_${d}`;
  const docState = (d) => uploadState[d] || { file: null, err: "", uploading: false, validando: false, resultado: null, success: false };
  const setDocState = (d, patch) => setUploadState(prev => ({ ...prev, [d]: { ...(prev[d] || { file: null, err: "", uploading: false, validando: false, resultado: null, success: false }), ...patch } }));

  function docsDeclarados(viaje) {
    return [1, viaje?.placa_2 ? 2 : null, viaje?.placa_3 ? 3 : null].filter(Boolean);
  }

  function openModal(viaje) {
    const faltaPlaca = !viaje.placa || viaje.placa.trim() === "";
    const faltaRutas = !viaje.rutas || viaje.rutas.trim() === "";
    if (faltaPlaca || faltaRutas) {
      const campos = [faltaPlaca && "N° Placa", faltaRutas && "N° GR"].filter(Boolean).join(" y ");
      setModal({ ...viaje, _bloqueado: true, _mensajeBloqueo: campos });
    } else {
      setModal(viaje);
    }
    setUploadState({});
  }

  function handleFileSelect(docIndex, file) {
    if (!file) return;
    if (file.size > MAX_MB * 1024 * 1024) { setDocState(docIndex, { err: `El archivo supera los ${MAX_MB} MB.`, file: null }); return; }
    setDocState(docIndex, { err: "", file });
  }

  async function handleUpload(docIndex) {
    const st = docState(docIndex);
    if (!st.file || !modal) return;
    const suf = SUF(docIndex);
    setDocState(docIndex, { uploading: true, err: "" });
    try {
      const compressed = await comprimirImagen(st.file);
      const versiones  = modal[`foto_versiones${suf}`] || [];
      const nv  = versiones.length + 1;
      const ext = compressed.name.split(".").pop();
      const nroCorto = extraerNroSpotCorto(modal.nro_spot);
      const fileName = `${nroCorto}_doc${docIndex}_v${nv}.${ext}`;
      const path = `${modal.proveedor}/${modal.nro_spot}/${fileName}`;
      const { error: upErr } = await supabase.storage.from("documentos").upload(path, compressed, { upsert: true });
      if (upErr) throw upErr;

      const nuevasVersiones = [...versiones, { v: nv, path, nombre: fileName, subido_por: session.user.email, subido_en: new Date().toISOString() }];
      const ahora = new Date().toISOString();
      const intentosPrevios = modal[`intentos_ia${suf}`] || 0;
      const payload = {
        [`foto_versiones${suf}`]: nuevasVersiones,
        [`foto_url${suf}`]: path,
        [`foto_nombre${suf}`]: fileName,
        [`estado_procesamiento_ia${suf}`]: null,
        [`intentos_ia${suf}`]: intentosPrevios + 1,
        subido_por: session.user.email,
        subido_en: ahora,
        fecha_entrega_doc: versiones.length === 0 ? ahora : modal.fecha_entrega_doc,
        estado_doc: "Completo",
      };
      const { error: dbErr } = await supabase.from("viajes").update(payload).eq("nro_spot", modal.nro_spot);
      if (dbErr) throw dbErr;

      setDocState(docIndex, { uploading: false, validando: true });
      setModal(prev => ({ ...prev, ...payload }));

      let resultado = null;
      try {
        const resp = await fetch("/api/procesar-ocr", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ nro_spot: modal.nro_spot, documento: docIndex }),
        });
        if (resp.ok) resultado = await resp.json();
      } catch { }

      let resPayload = {};
      if (resultado) {
        resPayload = {
          [`estado_procesamiento_ia${suf}`]: resultado.estado_procesamiento_ia,
          [`estado_validacion_ia${suf}`]: resultado.estado_validacion_ia,
          [`texto_detectado_ia${suf}`]: resultado.texto_detectado_ia,
          [`match_ia${suf}`]: resultado.match_ia,
        };
        setModal(prev => ({ ...prev, ...resPayload }));
      }

      setDocState(docIndex, { validando: false, resultado, success: true });
      setViajes(prev => prev.map(v => v.nro_spot === modal.nro_spot ? { ...v, ...payload, ...resPayload } : v));
      fetchViajes();

      const cierraSolo = !resultado || resultado.estado_validacion_ia !== "NO_COINCIDE";
      if (cierraSolo) {
        setTimeout(() => {
          if (docsDeclarados(modal).length === 1) {
            setModal(null);
          }
          setDocState(docIndex, { success: false, resultado: null, file: null });
        }, 2200);
      }
    } catch (err) {
      setDocState(docIndex, { err: err.message || "Error al subir el archivo.", uploading: false });
    }
  }

  async function handleValidarManual() {
    if (!validModal) return;
    setValidSaving(true); setValidErr("");
    try {
      const ahora = new Date().toISOString();
      // Marca MANUAL en TODOS los documentos que este ticket tenga declarados (no solo el 1) --
      // el trigger de la base de datos es quien decide estado_final según si ya todos quedaron OK,
      // así este botón no puede finalizar un ticket con un 2do/3er documento aún pendiente.
      const payload = { estado_validacion_ia: "MANUAL", usuario_modif: session.user.email, fecha_modif: ahora };
      if (validModal.placa_2) payload.estado_validacion_ia_2 = "MANUAL";
      if (validModal.placa_3) payload.estado_validacion_ia_3 = "MANUAL";

      const { data, error } = await supabase.from("viajes").update(payload).eq("nro_spot", validModal.nro_spot).select().single();
      if (error) throw error;
      setViajes(prev => prev.map(v => v.nro_spot === validModal.nro_spot ? { ...v, ...data } : v));
      setValidModal(null);
    } catch (err) { setValidErr(err.message || "Error al validar."); }
    finally { setValidSaving(false); }
  }

  const meta    = session?.user?.user_metadata;
  const isAdmin = meta?.role === "admin";
  const empresa = meta?.empresa_id || "";
  const initials= (session?.user?.email || "U").substring(0, 2).toUpperCase();
  const inp     = { padding: "6px 10px", fontSize: 12, border: `0.5px solid ${BORDER}`, borderRadius: 8, background: "white", color: GRAY_900, outline: "none" };

  const fetchDashboard = useCallback(async () => {
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
    } catch (err) {
      console.error("Error cargando dashboard:", err);
    } finally {
      setDashLoading(false);
    }
  }, [isAdmin, dashProveedor, dashDesde, dashHasta]);

  // Condiciones de cada bloque de la cascada, para el detalle por clic (todos menos "Total")
  const CASCADA_CONDICIONES = {
    "No confirmados":    q => q.is("realizado", null),
    "No realizados":     q => q.eq("realizado", "NO"),
    "Pendiente escaneo": q => q.eq("realizado", "SI").is("foto_url", null).eq("estado_final", "PENDIENTE"),
    "Observado IA":      q => q.eq("realizado", "SI").eq("estado_final", "OBSERVADO"),
    "Listos para migrar":q => q.eq("realizado", "SI").eq("estado_final", "FINALIZADO"),
  };

  async function abrirDetalleCascada(categoria) {
    const condicion = CASCADA_CONDICIONES[categoria];
    if (!condicion) return; // "Total tickets" no tiene popup -- se ve desfiltrando la tabla general
    setCascadaDetalleModal({ categoria, cargando: true, filas: [] });
    try {
      let q = supabase.from("viajes").select("*").order("fecha_carga", { ascending: false }).limit(300);
      if (dashProveedor) q = q.eq("proveedor", dashProveedor);
      if (dashDesde) q = q.gte("fecha_carga", dashDesde);
      if (dashHasta) q = q.lte("fecha_carga", dashHasta);
      q = condicion(q);
      const { data, error } = await q;
      if (error) throw error;
      setCascadaDetalleModal({ categoria, cargando: false, filas: data || [] });
    } catch (err) {
      console.error("Error cargando detalle de cascada:", err);
      setCascadaDetalleModal({ categoria, cargando: false, filas: [], error: true });
    }
  }

  useEffect(() => {
    if (vista !== "dashboard" || !isAdmin) return;
    fetchDashboard();
    const intervalo = setInterval(() => fetchDashboard(), 30000);
    return () => clearInterval(intervalo);
  }, [vista, isAdmin, dashProveedor, dashDesde, dashHasta]);

  function limpiarFiltrosDashboard() {
    setDashProveedor(""); 
    setDashDesde(""); 
    setDashHasta(D1_STR);
  }

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

      <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>
        {/* SIDEBAR */}
        {isAdmin && (
          <div style={{ width: 190, background: "white", borderRight: `0.5px solid ${BORDER}`, flexShrink: 0, display: "flex", flexDirection: "column", padding: "14px 10px" }}>
            <button onClick={() => setVista("tabla")}
              style={{ display: "flex", alignItems: "center", gap: 9, padding: "9px 12px", borderRadius: 8, border: "none", background: vista === "tabla" ? RED_LIGHT : "transparent", color: vista === "tabla" ? RED_DARK : GRAY_900, fontSize: 12, fontWeight: vista === "tabla" ? 600 : 500, cursor: "pointer", textAlign: "left", marginBottom: 4 }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="3" y1="9" x2="21" y2="9"></line>
                <line x1="9" y1="21" x2="9" y2="9"></line>
              </svg> 
              Seguimiento Adicionales
            </button>
            <button onClick={() => setVista("dashboard")}
              style={{ display: "flex", alignItems: "center", gap: 9, padding: "9px 12px", borderRadius: 8, border: "none", background: vista === "dashboard" ? RED_LIGHT : "transparent", color: vista === "dashboard" ? RED_DARK : GRAY_900, fontSize: 12, fontWeight: vista === "dashboard" ? 600 : 500, cursor: "pointer", textAlign: "left" }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
                <line x1="18" y1="20" x2="18" y2="10"></line>
                <line x1="12" y1="20" x2="12" y2="4"></line>
                <line x1="6" y1="20" x2="6" y2="14"></line>
              </svg> 
              Dashboard
            </button>
          </div>
        )}

        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

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
        {vista === "tabla" && (
          <>
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
                <div style={{ fontSize: 11, fontWeight: 500, color: GRAY_900, marginBottom: 6 }}>Estado de Viaje Final</div>
                <select value={dEstFinal} onChange={e => setDEstFinal(e.target.value)} style={{ ...inp, width: "100%", boxSizing: "border-box" }}>
                  <option value="">Todos</option>
                  <option value="FINALIZADO">FINALIZADO</option>
                  <option value="PENDIENTE">PENDIENTE</option>
                  <option value="OBSERVADO">OBSERVADO</option>
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
                  <th key={c.key} onClick={() => {
                    if (tablaSortCol === c.key) setTablaSortDir(d => d === "asc" ? "desc" : "asc");
                    else { setTablaSortCol(c.key); setTablaSortDir("asc"); }
                  }} style={{ padding: "8px 11px", textAlign: c.right ? "right" : "left", fontSize: 10, fontWeight: 500,
                    color: c.headerGroup === "ia" ? BLUE : c.headerGroup === "transportista" ? AMBER : GRAY_500,
                    background: c.headerGroup === "ia" ? "#EEF4FC" : c.headerGroup === "transportista" ? "#FFF8EC" : GRAY_50,
                    borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", position: "sticky", top: 0, zIndex: 3, textTransform: "uppercase", letterSpacing: ".03em", cursor: "pointer", userSelect: "none" }}>
                    {c.label} {tablaSortCol === c.key && (tablaSortDir === "desc" ? "▼" : "▲")}
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
              ) : ordenarViajes(viajes, tablaSortCol, tablaSortDir).map(v => {
                const completo = (v.foto_versiones?.length || 0) > 0;

                return (
                  <tr key={v.nro_spot} className="viaje-row">
                    {colsVisibles.map(c => {
                      const editable = !isAdmin && (c.key === "placa" || c.key === "rutas");
                      let content = v[c.key];
                      const CAMPOS_CONSOLIDADOS = ["rutas", "placa", "texto_detectado_ia", "estado_validacion_ia"];
                      const esConsolidado = CAMPOS_CONSOLIDADOS.includes(c.key);

                      if (c.key === "estado_final") {
                        const ef = v.estado_final;
                        const bg = ef === "FINALIZADO" ? GREEN_LIGHT : ef === "PENDIENTE" ? AMBER_LIGHT : ef === "OBSERVADO" ? RED_LIGHT : GRAY_100;
                        const fg = ef === "FINALIZADO" ? GREEN : ef === "PENDIENTE" ? AMBER : ef === "OBSERVADO" ? RED_DARK : GRAY_500;
                        content = ef
                          ? <span style={{ display: "inline-flex", padding: "2px 8px", borderRadius: 999, fontSize: 10, fontWeight: 600, background: bg, color: fg, whiteSpace: "nowrap" }}>{ef}</span>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (["realizado","estado_doc","estado_procesamiento_ia"].includes(c.key)) {
                        content = v[c.key]
                          ? <span style={{ fontSize: 11, color: GRAY_900 }}>{String(v[c.key]).toUpperCase()}</span>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (["rutas","placa","texto_detectado_ia"].includes(c.key)) {
                        // Documento 1/2/3 en líneas separadas dentro de la misma celda -- la data real
                        // sigue viviendo en columnas propias (rutas_2/placa_2/etc), esto es solo visual.
                        const partes = [v[c.key], v[`${c.key}_2`], v[`${c.key}_3`]].filter(Boolean);
                        content = partes.length
                          ? <div style={{ display: "flex", flexDirection: "column", gap: 1, padding: "2px 0" }}>
                              {partes.map((p, i) => <span key={i} style={i === 0 ? { fontFamily: c.mono ? "monospace" : "inherit" } : { fontSize: 10, color: GRAY_500, fontFamily: c.mono ? "monospace" : "inherit" }}>{p}</span>)}
                            </div>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "estado_validacion_ia") {
                        const partes = [v.estado_validacion_ia, v.estado_validacion_ia_2, v.estado_validacion_ia_3].filter(Boolean).map(s => String(s).toUpperCase());
                        content = partes.length
                          ? <div style={{ display: "flex", flexDirection: "column", gap: 1, padding: "2px 0" }}>
                              {partes.map((p, i) => <span key={i} style={i === 0 ? { fontSize: 11, color: GRAY_900 } : { fontSize: 10, color: GRAY_500 }}>{p}</span>)}
                            </div>
                          : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "fecha_modif" || c.key === "fecha_entrega_doc") {
                        content = v[c.key] ? fmtFechaHora(v[c.key]) : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "fecha_carga") {
                        content = v[c.key] ? fmtFechaSolo(v[c.key]) : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (c.key === "hora_cita") {
                        const m = v[c.key] ? String(v[c.key]).match(/(\d{1,2}:\d{2})/) : null;
                        content = m ? m[1] : <span style={{ color: GRAY_200 }}>—</span>;
                      } else if (content === null || content === undefined || content === "") {
                        content = <span style={{ color: GRAY_200 }}>—</span>;
                      }
                      const tituloCelda = esConsolidado
                        ? [v[c.key], v[`${c.key}_2`], v[`${c.key}_3`]].filter(Boolean).join(" | ") || undefined
                        : (v[c.key] !== null && v[c.key] !== undefined && v[c.key] !== "") ? String(v[c.key]) : undefined;
                      return (
                        <td key={c.key} title={tituloCelda}
                          data-editable={editable ? "1" : undefined}
                          style={{ padding: "8px 11px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", overflow: esConsolidado ? "visible" : "hidden", textOverflow: esConsolidado ? "clip" : "ellipsis", whiteSpace: esConsolidado ? "normal" : "nowrap", fontFamily: c.mono ? "monospace" : "inherit", fontSize: c.mono ? 10 : 11, textAlign: c.right ? "right" : "left", color: c.muted ? GRAY_500 : GRAY_900 }}>
                          {content}
                        </td>
                      );
                    })}

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
                      {isAdmin && v.realizado === "SI" && v.estado_final !== "FINALIZADO" && (
                        <button onClick={() => { setValidModal(v); setValidErr(""); }} title="Aprobar adicional"
                          style={{ width: 28, height: 28, borderRadius: 7, border: `1.5px solid ${BLUE}`, background: "white", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={BLUE} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                            <polyline points="20 6 9 17 4 12"/>
                          </svg>
                        </button>
                      )}
                    </td>

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
            
            {/* Calendario con Dropdown format DD/MM/AAAA */}
            <div style={{ position: "relative" }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 5 }}>Fecha de ejecución</div>
              <button onClick={() => setDashFiltroAbierto(!dashFiltroAbierto)} 
                style={{ ...inp, width: 220, textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                {dashDesde && dashHasta ? `${fmtFechaSolo(dashDesde)} al ${fmtFechaSolo(dashHasta)}` : "Seleccionar fechas"} 
                <span style={{ fontSize: 9, transform: dashFiltroAbierto ? "rotate(180deg)" : "none" }}>▼</span>
              </button>
              {dashFiltroAbierto && (
                <div style={{ position: "absolute", top: "100%", left: 0, marginTop: 6, background: "white", border: `0.5px solid ${BORDER}`, borderRadius: 10, padding: 12, zIndex: 50, boxShadow: "0 6px 20px rgba(0,0,0,0.12)" }}>
                  <RangePicker desde={dashDesde} hasta={dashHasta} maxDias={90}
                    onChange={({ desde, hasta }) => { setDashDesde(desde); setDashHasta(hasta); setDashFiltroAbierto(false); }} />
                </div>
              )}
            </div>

            <button onClick={limpiarFiltrosDashboard} style={{ height: 34, padding: "0 14px", borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 12 }}>
              Limpiar filtros
            </button>
            <button onClick={() => fetchDashboard()} title="Actualizar"
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
                <div style={{ background: GRAY_50, borderRadius: 10, padding: "14px 16px", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 12, color: GRAY_500, marginBottom: 6 }}>Confirmación de Ejecución</div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ fontSize: 24, fontWeight: 600, color: GRAY_900 }}>{kpis.tickets_validados_ejecucion} <span style={{ fontSize: 13, color: GRAY_500, fontWeight: 400 }}>/ {kpis.total_tickets}</span></div>
                    {(() => {
                      const p = kpis.total_tickets > 0 ? Math.round((kpis.tickets_validados_ejecucion / kpis.total_tickets) * 100) : 0;
                      const c = p >= 70 ? GREEN : p >= 40 ? AMBER : RED_DARK;
                      const bg = p >= 70 ? GREEN_LIGHT : p >= 40 ? AMBER_LIGHT : RED_LIGHT;
                      return <span style={{ display: "inline-flex", padding: "2px 9px", borderRadius: 999, fontSize: 11, fontWeight: 600, background: bg, color: c }}>{p}%</span>;
                    })()}
                  </div>
                </div>
                <div style={{ background: GRAY_50, borderRadius: 10, padding: "14px 16px", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 12, color: GRAY_500, marginBottom: 6 }}>Escaneo de Guías</div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ fontSize: 24, fontWeight: 600, color: GRAY_900 }}>{kpis.tickets_con_foto} <span style={{ fontSize: 13, color: GRAY_500, fontWeight: 400 }}>/ {kpis.tickets_requieren_escaneo}</span></div>
                    {(() => {
                      const p = kpis.tickets_requieren_escaneo > 0 ? Math.round((kpis.tickets_con_foto / kpis.tickets_requieren_escaneo) * 100) : 0;
                      const c = p >= 70 ? GREEN : p >= 40 ? AMBER : RED_DARK;
                      const bg = p >= 70 ? GREEN_LIGHT : p >= 40 ? AMBER_LIGHT : RED_LIGHT;
                      return <span style={{ display: "inline-flex", padding: "2px 9px", borderRadius: 999, fontSize: 11, fontWeight: 600, background: bg, color: c }}>{p}%</span>;
                    })()}
                  </div>
                </div>
                <div style={{ background: RED_LIGHT, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: RED_DARK, marginBottom: 6 }}>Observados por IA</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: RED_DARK }}>{kpis.tickets_rechazados_ia}</div>
                </div>
                <div style={{ background: GREEN_LIGHT, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12, color: GREEN, marginBottom: 6 }}>Listos para migrar</div>
                  <div style={{ fontSize: 24, fontWeight: 600, color: GREEN }}>{kpis.tickets_listos_migrar}</div>
                </div>
              </div>

              {/* Cascada — gráfico real: eje Y, gridlines, y conectores por esquina según corresponda */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: GRAY_900, marginBottom: 34 }}>Estatus de Tickets para Migrar</div>
                {(() => {
                  const total = kpis.total_tickets || 0;
                  const ALTO_PX = 170;
                  const pct = v => total > 0 ? Math.round((v / total) * 100) : 0;

                  // landing = altura (en unidades de ticket) donde esta barra "entrega" a la siguiente:
                  // - la barra Total entrega desde su TOPE (arriba)
                  // - cada resta entrega desde su BASE (abajo), porque la siguiente cuelga desde ahí
                  let acumulado = total;
                  const restas = [
                    { label: "No confirmados", val: kpis.cascada_no_validados || 0, color: "#ff7676" },
                    { label: "No realizados", val: kpis.cascada_no_realizados || 0, color: "#ff4040" },
                    { label: "Pendiente escaneo", val: kpis.cascada_pendiente_subir || 0, color: "#ff0000" },
                    { label: "Observado IA", val: kpis.cascada_observado || 0, color: "#e00000" },
                  ].map(r => {
                    const base = acumulado - r.val;
                    const barra = { ...r, base, tope: acumulado, landing: base };
                    acumulado = base;
                    return barra;
                  });

                  const barras = [
                    { label: "Total tickets", val: total, base: 0, color: "#c20000", landing: total },
                    ...restas,
                    { label: "Listos para migrar", val: kpis.tickets_listos_migrar || 0, base: 0, color: GREEN, landing: null },
                  ];

                  // Eje Y con "números redondos" Y techo ajustado: en vez de forzar siempre 4 marcas
                  // (lo que a veces deja hasta el doble de espacio vacío arriba, según dónde caiga el
                  // total), se calcula el múltiplo "lindo" más chico que alcance a cubrir el total —
                  // así la barra más alta siempre llega a un tamaño consistente, sin importar el valor.
                  const rawStep = total > 0 ? total / 4 : 1;
                  const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
                  const norm = rawStep / mag;
                  const niceMult = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10;
                  const step = niceMult * mag;
                  const numTicks = total > 0 ? Math.ceil(total / step) : 4;
                  const ejeMax = step * numTicks;
                  const ticks = Array.from({ length: numTicks + 1 }, (_, i) => step * i);
                  const escala = ejeMax > 0 ? ALTO_PX / ejeMax : 0;

                  const N = barras.length;
                  const colPct = 100 / N;
                  const margenPct = colPct * 0.14;

                  return (
                    <div style={{ display: "flex" }}>
                      {/* Eje Y */}
                      <div style={{ position: "relative", width: 32, height: ALTO_PX, flexShrink: 0 }}>
                        {ticks.map(t => (
                          <div key={t} style={{ position: "absolute", bottom: t * escala - 6, right: 6, fontSize: 9, color: GRAY_500 }}>{t}</div>
                        ))}
                      </div>

                      {/* Área del gráfico */}
                      <div style={{ position: "relative", flex: 1, height: ALTO_PX }}>
                        {/* Gridlines horizontales */}
                        {ticks.map(t => (
                          <div key={t} style={{ position: "absolute", left: 0, right: 0, bottom: t * escala, borderTop: `1px dashed ${GRAY_200}` }} />
                        ))}

                        {/* Conectores punteados:
                            - Barra "Total" (i=0): conecta desde su esquina SUPERIOR derecha hasta la
                              esquina SUPERIOR izquierda de la siguiente (ambas al tope del total).
                            - Barras "resta" (i>0): conectan desde su esquina INFERIOR derecha hasta la
                              esquina SUPERIOR izquierda de la siguiente (donde la siguiente cuelga). */}
                        {barras.slice(0, -1).map((b, i) => b.landing === null ? null : (
                          <div key={`conn-${i}`} style={{
                            position: "absolute", bottom: b.landing * escala,
                            left: `calc(${(i + 1) * colPct}% - ${margenPct}%)`,
                            width: `${margenPct * 2}%`,
                            borderTop: `1.5px dotted ${GRAY_500}`,
                          }} />
                        ))}

                        {/* Barras — todas clickeables excepto "Total tickets" (esa se ve desfiltrando la tabla general) */}
                        {barras.map((b, i) => (
                          <div key={b.label} title={i === 0 ? b.label : `${b.label}: ${b.val} (${pct(b.val)}%) — clic para ver el detalle`}
                            onClick={i === 0 ? undefined : () => abrirDetalleCascada(b.label)}
                            style={{
                            position: "absolute",
                            left: `calc(${i * colPct}% + ${margenPct}%)`,
                            width: `${colPct - margenPct * 2}%`,
                            bottom: b.base * escala,
                            height: Math.max(b.val * escala, b.val > 0 ? 3 : 0),
                            background: b.color, borderRadius: 3,
                            cursor: i === 0 ? "default" : "pointer",
                          }}>
                            <div style={{ position: "absolute", top: -20, left: "50%", transform: "translateX(-50%)", textAlign: "center", fontSize: 11, fontWeight: 600, color: GRAY_900, lineHeight: 1.3, whiteSpace: "nowrap" }}>
                              {b.val} ({pct(b.val)}%)
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}
                {(() => {
                  const barrasLabels = ["Total tickets", "No confirmados", "No realizados", "Pendiente escaneo", "Observado IA", "Listos para migrar"];
                  return (
                    <div style={{ display: "flex", marginLeft: 32 }}>
                      {barrasLabels.map(l => (
                        <div key={l} style={{ flex: 1, fontSize: 10, color: GRAY_500, textAlign: "center", marginTop: 8, lineHeight: 1.3 }}>{l}</div>
                      ))}
                    </div>
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
                    {[
                      { h: "Proveedor", k: "proveedor" },
                      { h: "Tickets Realizados", k: "total" },
                      { h: "Pend. Escaneo Guía", k: "pendiente_subir" },
                      { h: "Observado IA", k: "observado" },
                      { h: "Listos para migrar", k: "listos_migrar" },
                      { h: "% Avance de cumplimiento", k: "pct_avance" },
                    ].map(({ h, k }, i, arr) => (
                      <th key={h} onClick={() => {
                        if (rankingSortCol === k) setRankingSortDir(d => d === "desc" ? "asc" : "desc");
                        else { setRankingSortCol(k); setRankingSortDir("desc"); }
                      }} style={{ padding: "8px 10px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".03em", borderBottom: `0.5px solid ${BORDER}`, borderRight: i < arr.length - 1 ? `0.5px solid ${BORDER}` : "none", cursor: "pointer", userSelect: "none", whiteSpace: "nowrap" }}>
                        {h} {rankingSortCol === k && (rankingSortDir === "desc" ? "▼" : "▲")}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...ranking].sort((a, b) => {
                    if (!rankingSortCol) return 0;
                    const av = a[rankingSortCol], bv = b[rankingSortCol];
                    if (rankingSortCol === "proveedor") {
                      const as = av || "FALTA ASIGNAR", bs = bv || "FALTA ASIGNAR";
                      return rankingSortDir === "desc" ? bs.localeCompare(as) : as.localeCompare(bs);
                    }
                    const an = Number(av) || 0, bn = Number(bv) || 0;
                    return rankingSortDir === "desc" ? bn - an : an - bn;
                  }).map(r => {
                    const avance = r.pct_avance ?? 0;
                    const colorAvance = avance >= 70 ? GREEN : avance >= 40 ? AMBER : RED_DARK;
                    const bgAvance = avance >= 70 ? GREEN_LIGHT : avance >= 40 ? AMBER_LIGHT : RED_LIGHT;
                    return (
                      <tr key={r.proveedor || "sin-proveedor"}>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, color: r.proveedor ? GRAY_900 : GRAY_500, fontStyle: r.proveedor ? "normal" : "italic" }}>{r.proveedor || "FALTA ASIGNAR"}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}` }}>{r.total}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}` }}>{r.pendiente_subir}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, color: r.observado > 0 ? RED_DARK : GRAY_900 }}>{r.observado}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, color: GREEN }}>{r.listos_migrar}</td>
                        <td style={{ padding: "8px 10px", textAlign: "center", borderBottom: `0.5px solid ${BORDER}` }}>
                          <span style={{ display: "inline-flex", padding: "2px 10px", borderRadius: 999, fontSize: 10, fontWeight: 600, background: bgAvance, color: colorAvance }}>{avance}%</span>
                        </td>
                      </tr>
                    );
                  })}
                  {ranking.length === 0 && (
                    <tr><td colSpan={6} style={{ padding: 24, textAlign: "center", color: GRAY_500 }}>Sin datos para el rango seleccionado</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

        </div>
      </div>

      {/* MODAL DETALLE POR BLOQUE DE LA CASCADA */}
      {cascadaDetalleModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }}
          onClick={e => e.target === e.currentTarget && setCascadaDetalleModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 20, width: "min(94vw, 1100px)", maxHeight: "85vh", display: "flex", flexDirection: "column" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexShrink: 0 }}>
              <div>
                <div style={{ fontSize: 14, fontWeight: 600, color: GRAY_900 }}>{cascadaDetalleModal.categoria}</div>
                <div style={{ fontSize: 11, color: GRAY_500, marginTop: 2 }}>
                  {cascadaDetalleModal.cargando ? "Cargando…" : `${cascadaDetalleModal.filas.length} ticket${cascadaDetalleModal.filas.length === 1 ? "" : "s"}${cascadaDetalleModal.filas.length === 300 ? " (mostrando los primeros 300)" : ""}`}
                </div>
              </div>
              <button onClick={() => setCascadaDetalleModal(null)} style={{ width: 26, height: 26, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 13, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>✕</button>
            </div>
            <div style={{ overflow: "auto", border: `0.5px solid ${BORDER}`, borderRadius: 10 }}>
              {cascadaDetalleModal.cargando ? (
                <div style={{ padding: 40, textAlign: "center", color: GRAY_500, fontSize: 12 }}>Cargando tickets…</div>
              ) : cascadaDetalleModal.error ? (
                <div style={{ padding: 40, textAlign: "center", color: RED_DARK, fontSize: 12 }}>Error al cargar el detalle. Intenta de nuevo.</div>
              ) : cascadaDetalleModal.filas.length === 0 ? (
                <div style={{ padding: 40, textAlign: "center", color: GRAY_500, fontSize: 12 }}>No hay tickets en esta categoría para el rango filtrado.</div>
              ) : (
                <table style={{ borderCollapse: "collapse", fontSize: 11, whiteSpace: "nowrap" }}>
                  <thead>
                    <tr style={{ background: GRAY_50 }}>
                      {COLS.map(c => (
                        <th key={c.key} style={{ padding: "7px 9px", textAlign: "left", fontSize: 9, fontWeight: 500, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".03em", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, position: "sticky", top: 0, background: GRAY_50 }}>{c.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {cascadaDetalleModal.filas.map(fila => (
                      <tr key={fila.nro_spot}>
                        {COLS.map(c => {
                          let val = fila[c.key];
                          if (["rutas", "placa", "texto_detectado_ia", "estado_validacion_ia"].includes(c.key)) {
                            val = [fila[c.key], fila[`${c.key}_2`], fila[`${c.key}_3`]].filter(Boolean).join(" | ") || "—";
                          } else if (c.key === "fecha_carga") val = val ? fmtFechaSolo(val) : "—";
                          else if (val === null || val === undefined || val === "") val = "—";
                          return (
                            <td key={c.key} style={{ padding: "7px 9px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, color: val === "—" ? GRAY_200 : GRAY_900, fontFamily: c.mono ? "monospace" : "inherit" }}>{String(val)}</td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

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
            <div style={{ fontSize: 10, fontWeight: 600, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 8 }}>Guía 1</div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <div style={{ fontSize: 11, color: GRAY_500, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° Placa</div>
                {editModal.proveedor === "MUNDO" && (
                  <div style={{ display: "flex", gap: 4 }}>
                    <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 1: false }))} title="Camión (3-3)"
                      style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${!tipoVehiculo[1] ? RED : BORDER}`, background: !tipoVehiculo[1] ? RED_LIGHT : "white", color: !tipoVehiculo[1] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="1" y="7" width="13" height="9"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6" cy="18" r="2"/><circle cx="17" cy="18" r="2"/></svg>
                    </button>
                    <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 1: true }))} title="Moto (4-2)"
                      style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${tipoVehiculo[1] ? RED : BORDER}`, background: tipoVehiculo[1] ? RED_LIGHT : "white", color: tipoVehiculo[1] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="5" cy="17" r="3"/><circle cx="19" cy="17" r="3"/><path d="M8 17h5l3-6h3M13 11l-2-4H8"/></svg>
                    </button>
                  </div>
                )}
              </div>
              <input value={editPlaca} onChange={e => { setEditPlaca(formatearPlaca(e.target.value, tipoVehiculo[1])); setEditErr(""); }}
                maxLength={7}
                style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 13, letterSpacing: ".08em", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
            </div>
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° GR</div>
              <input value={editRutaInput} onChange={e => { setEditRutaInput(e.target.value); setEditErr(""); }}
                style={{ ...inp, width: "100%", boxSizing: "border-box", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
            </div>

            {editDocsVisibles >= 2 && (
              <>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".04em" }}>Guía 2</div>
                  {editDocsVisibles === 2 && (
                    <button onClick={() => { setEditPlaca2(""); setEditRuta2Input(""); setEditDocsVisibles(1); }}
                      style={{ background: "none", border: "none", color: RED_DARK, fontSize: 10, cursor: "pointer" }}>Quitar</button>
                  )}
                </div>
                <div style={{ marginBottom: 18 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                    <div style={{ fontSize: 11, color: GRAY_500, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° Placa</div>
                    {editModal.proveedor === "MUNDO" && (
                      <div style={{ display: "flex", gap: 4 }}>
                        <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 2: false }))} title="Camión (3-3)"
                          style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${!tipoVehiculo[2] ? RED : BORDER}`, background: !tipoVehiculo[2] ? RED_LIGHT : "white", color: !tipoVehiculo[2] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="1" y="7" width="13" height="9"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6" cy="18" r="2"/><circle cx="17" cy="18" r="2"/></svg>
                        </button>
                        <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 2: true }))} title="Moto (4-2)"
                          style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${tipoVehiculo[2] ? RED : BORDER}`, background: tipoVehiculo[2] ? RED_LIGHT : "white", color: tipoVehiculo[2] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="5" cy="17" r="3"/><circle cx="19" cy="17" r="3"/><path d="M8 17h5l3-6h3M13 11l-2-4H8"/></svg>
                        </button>
                      </div>
                    )}
                  </div>
                  <input value={editPlaca2} onChange={e => { setEditPlaca2(formatearPlaca(e.target.value, tipoVehiculo[2])); setEditErr(""); }}
                    maxLength={7}
                    style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 13, letterSpacing: ".08em", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
                </div>
                <div style={{ marginBottom: 18 }}>
                  <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° GR</div>
                  <input value={editRuta2Input} onChange={e => { setEditRuta2Input(e.target.value); setEditErr(""); }}
                    style={{ ...inp, width: "100%", boxSizing: "border-box", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
                </div>
              </>
            )}

            {editDocsVisibles >= 3 && (
              <>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: GRAY_500, textTransform: "uppercase", letterSpacing: ".04em" }}>Guía 3</div>
                  <button onClick={() => { setEditPlaca3(""); setEditRuta3Input(""); setEditDocsVisibles(2); }}
                    style={{ background: "none", border: "none", color: RED_DARK, fontSize: 10, cursor: "pointer" }}>Quitar</button>
                </div>
                <div style={{ marginBottom: 18 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                    <div style={{ fontSize: 11, color: GRAY_500, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° Placa</div>
                    {editModal.proveedor === "MUNDO" && (
                      <div style={{ display: "flex", gap: 4 }}>
                        <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 3: false }))} title="Camión (3-3)"
                          style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${!tipoVehiculo[3] ? RED : BORDER}`, background: !tipoVehiculo[3] ? RED_LIGHT : "white", color: !tipoVehiculo[3] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="1" y="7" width="13" height="9"/><path d="M14 10h4l3 3v3h-7z"/><circle cx="6" cy="18" r="2"/><circle cx="17" cy="18" r="2"/></svg>
                        </button>
                        <button type="button" onClick={() => setTipoVehiculo(prev => ({ ...prev, 3: true }))} title="Moto (4-2)"
                          style={{ width: 22, height: 22, borderRadius: 6, border: `1px solid ${tipoVehiculo[3] ? RED : BORDER}`, background: tipoVehiculo[3] ? RED_LIGHT : "white", color: tipoVehiculo[3] ? RED_DARK : GRAY_500, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", padding: 0 }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="5" cy="17" r="3"/><circle cx="19" cy="17" r="3"/><path d="M8 17h5l3-6h3M13 11l-2-4H8"/></svg>
                        </button>
                      </div>
                    )}
                  </div>
                  <input value={editPlaca3} onChange={e => { setEditPlaca3(formatearPlaca(e.target.value, tipoVehiculo[3])); setEditErr(""); }}
                    maxLength={7}
                    style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 13, letterSpacing: ".08em", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
                </div>
                <div style={{ marginBottom: 18 }}>
                  <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>N° GR</div>
                  <input value={editRuta3Input} onChange={e => { setEditRuta3Input(e.target.value); setEditErr(""); }}
                    onKeyDown={e => e.key === "Enter" && (e.preventDefault(), !editSaving && saveEdit())}
                    style={{ ...inp, width: "100%", boxSizing: "border-box", background: "#FFFBF0", border: `1px solid ${BORDER}` }} />
                </div>
              </>
            )}

            {editDocsVisibles < 3 && (
              <button onClick={() => setEditDocsVisibles(d => d + 1)}
                style={{ display: "flex", alignItems: "center", gap: 5, background: "none", border: `1px dashed ${BORDER}`, borderRadius: 8, padding: "8px 12px", fontSize: 11, color: GRAY_500, cursor: "pointer", marginBottom: 18, width: "100%", justifyContent: "center" }}>
                + Agregar guía
              </button>
            )}
            <div style={{ fontSize: 10, color: GRAY_500, marginTop: -10, marginBottom: 16 }}>Presiona Enter en el último campo para guardar.</div>
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

      {validModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && !validSaving && setValidModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 380, maxWidth: "94vw" }}>
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

      {modal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && setModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 420, maxWidth: "92vw", maxHeight: "88vh", overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
              <div style={{ fontSize: 14, fontWeight: 500 }}>{modal._bloqueado ? "Campos incompletos" : "Subir documento(s)"}</div>
              <button onClick={() => setModal(null)} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 14, fontFamily: "monospace" }}>{modal.nro_spot}</div>

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
            ) : (
              <>
                {docsDeclarados(modal).map((docIndex, i) => {
                  const suf = SUF(docIndex);
                  const st = docState(docIndex);
                  const intentos = modal[`intentos_ia${suf}`] || 0;
                  const yaCoincide = modal[`estado_validacion_ia${suf}`] === "COINCIDE";
                  const bloqueadoPorIntentos = intentos >= 3 && !yaCoincide;
                  const nDocs = docsDeclarados(modal).length;

                  return (
                    <div key={docIndex} style={{ marginBottom: i < nDocs - 1 ? 22 : 0, paddingBottom: i < nDocs - 1 ? 20 : 0, borderBottom: i < nDocs - 1 ? `0.5px solid ${BORDER}` : "none" }}>
                      {nDocs > 1 && (
                        <div style={{ fontSize: 11, fontWeight: 600, color: GRAY_900, marginBottom: 8 }}>
                          Documento {docIndex} — Placa {modal[`placa${suf}`] || "—"} / GR {modal[`rutas${suf}`] || "—"}
                        </div>
                      )}

                      {bloqueadoPorIntentos ? (
                        <div style={{ padding: "12px 14px", background: RED_LIGHT, borderRadius: 10, border: "0.5px solid #f7c1c1" }}>
                          <div style={{ fontSize: 12, fontWeight: 600, color: RED_DARK, marginBottom: 4 }}>Se alcanzó el máximo de intentos</div>
                          <div style={{ fontSize: 11, color: RED_DARK }}>Un administrador validará el ticket manualmente.</div>
                        </div>
                      ) : st.success ? (
                        <div style={{ textAlign: "center", padding: "16px 0" }}>
                          {!st.resultado ? (
                            <>
                              <div style={{ fontSize: 28, marginBottom: 6 }}>✅</div>
                              <div style={{ fontSize: 12, color: GREEN, fontWeight: 500 }}>Documento guardado correctamente</div>
                              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>La validación automática se completará en breve.</div>
                            </>
                          ) : st.resultado.estado_validacion_ia === "COINCIDE" ? (
                            <>
                              <div style={{ fontSize: 28, marginBottom: 6 }}>✅</div>
                              <div style={{ fontSize: 12, color: GREEN, fontWeight: 500 }}>Documento validado correctamente</div>
                              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>Coincidencia: {st.resultado.match_ia}%</div>
                            </>
                          ) : (
                            <>
                              <div style={{ fontSize: 28, marginBottom: 6 }}>⚠️</div>
                              <div style={{ fontSize: 12, color: AMBER, fontWeight: 500 }}>El documento no coincide con lo registrado</div>
                              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>Detectado: {st.resultado.texto_detectado_ia || "—"} ({st.resultado.match_ia ?? 0}%)</div>
                              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 2 }}>Intento {intentos} de 3. Un administrador revisará tu documento.</div>
                              <button onClick={() => setDocState(docIndex, { success: false, resultado: null, file: null })}
                                style={{ marginTop: 10, padding: "6px 16px", background: GRAY_100, border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 11, cursor: "pointer", color: GRAY_900 }}>
                                Entendido
                              </button>
                            </>
                          )}
                        </div>
                      ) : st.validando ? (
                        <div style={{ textAlign: "center", padding: "16px 0" }}>
                          <div style={{ fontSize: 20, marginBottom: 6 }}>⏳</div>
                          <div style={{ fontSize: 12, color: GRAY_500 }}>Validando documento...</div>
                        </div>
                      ) : (
                        <>
                          {(modal[`foto_versiones${suf}`]?.length || 0) > 0 && (
                            <div style={{ marginBottom: 10, padding: "8px 10px", background: GREEN_LIGHT, borderRadius: 8, fontSize: 11, color: GREEN }}>
                              ✓ Ya tienes {modal[`foto_versiones${suf}`].length} documento(s) subido(s) aquí. Intento {intentos} de 3.
                            </div>
                          )}
                          <div onClick={() => fileRefFor(docIndex).current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleFileSelect(docIndex, e.dataTransfer.files[0]); }}
                            style={{ border: `1.5px dashed ${GRAY_200}`, borderRadius: 10, padding: "18px 14px", textAlign: "center", cursor: "pointer", marginBottom: 10 }}>
                            <div style={{ fontSize: 20, color: GRAY_200, marginBottom: 4 }}>📷</div>
                            <div style={{ fontSize: 11, color: GRAY_500 }}>{st.file ? st.file.name : "Clic o arrastra tu foto aquí"}</div>
                          </div>
                          {st.file && (
                            <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 10px", background: GRAY_100, borderRadius: 8, fontSize: 10, color: GRAY_900, marginBottom: 10, border: `0.5px solid ${BORDER}` }}>
                              <span>📎</span><span style={{ flex: 1 }}>Archivo: {st.file.name}</span>
                            </div>
                          )}
                          {st.err && <div style={{ padding: "7px 10px", background: RED_LIGHT, borderRadius: 8, fontSize: 10, color: RED_DARK, marginBottom: 10, border: `0.5px solid #f7c1c1` }}>⚠ {st.err}</div>}
                          <input ref={fileRefFor(docIndex)} type="file" accept="image/*" style={{ display: "none" }} onChange={e => handleFileSelect(docIndex, e.target.files[0])} />
                          <div style={{ display: "flex", justifyContent: "flex-end" }}>
                            <button onClick={() => handleUpload(docIndex)} disabled={!st.file || st.uploading}
                              style={{ padding: "6px 14px", background: st.file && !st.uploading ? RED : GRAY_200, color: st.file && !st.uploading ? "white" : GRAY_500, border: "none", borderRadius: 8, fontSize: 11, fontWeight: 500, cursor: st.file && !st.uploading ? "pointer" : "default" }}>
                              {st.uploading ? "Subiendo..." : `Guardar documento ${nDocs > 1 ? docIndex : ""}`}
                            </button>
                          </div>
                        </>
                      )}
                    </div>
                  );
                })}
                <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
                  <button onClick={() => setModal(null)} style={{ padding: "7px 14px", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500 }}>Cerrar</button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}