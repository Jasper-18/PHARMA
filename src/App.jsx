import { useState, useEffect, useRef, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://zffuccirauheklpxagga.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpmZnVjY2lyYXVoZWtscHhhZ2dhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI1Mjc0MzIsImV4cCI6MjA5ODEwMzQzMn0.MH8hJSktS_G3_omz9y48Vsp6PIlPcWdg6s6zdQtUNBo";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);

const RED = "#C8102E";
const RED_DARK = "#A00D24";
const RED_LIGHT = "#FCEBEB";
const GREEN = "#0f6e56";
const GREEN_LIGHT = "#e1f5ee";
const GRAY_50 = "#F8F8F8";
const GRAY_100 = "#F0EEEA";
const GRAY_200 = "#D3D1C7";
const GRAY_500 = "#888780";
const GRAY_900 = "#1a1a18";
const BORDER = "#e5e2db";
const MAX_MB = 10;

const COLS = [
  { key: "id_correlativo", label: "ID Correlativo", mono: true, width: 110 },
  { key: "nro_id_spot", label: "Nro SPOT", width: 200 },
  { key: "rutas", label: "Rutas / GR", trunc: true, width: 220 },
  { key: "fecha_carga", label: "Fecha", width: 90 },
  { key: "tipo_traslado", label: "Tipo", width: 60 },
  { key: "tipo_oneway", label: "Modalidad", muted: true, width: 90 },
  { key: "nro_traslado", label: "Nro Traslado", trunc: true, width: 120 },
  { key: "cd_origen", label: "Origen", width: 130 },
  { key: "cd_destino", label: "Destino", width: 130 },
  { key: "placa", label: "Placa", mono: true, width: 90 },
  { key: "cantidad_paletas", label: "Paletas", right: true, width: 70 },
  { key: "total_bultos", label: "Bultos", right: true, width: 70 },
  { key: "hora_cita", label: "Hora cita", width: 80 },
  { key: "observaciones", label: "Observaciones", trunc: true, muted: true, width: 160 },
  { key: "centro_costo", label: "Centro costo", muted: true, width: 110 },
  { key: "status", label: "Status viaje", width: 110 },
  { key: "criterio_falla", label: "Criterio falla", trunc: true, muted: true, width: 110 },
  { key: "spot", label: "Detalle de Servicio", trunc: true, width: 160 },
  { key: "importe_servicio", label: "Importe S/", right: true, width: 90 },
  { key: "validador_ejecucion", label: "Validador", muted: true, width: 90 },
  { key: "proveedor", label: "Proveedor", width: 90 },
];

// Extrae el número final de un id_correlativo (ej: "RICPAL Nro4465" -> 4465)
// para poder ordenar por ese número sin alterar el campo en la base de datos.
function extraerNumeroId(id) {
  if (!id) return -1;
  const m = String(id).match(/(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : -1;
}

async function comprimirImagen(file) {
  if (file.size > MAX_MB * 1024 * 1024) {
    throw new Error(`El archivo supera los ${MAX_MB} MB. Por favor selecciona una imagen más pequeña.`);
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

export default function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginErr, setLoginErr] = useState("");
  const [viajes, setViajes] = useState([]);
  const [dataLoading, setDataLoading] = useState(false);
  const [fDesde, setFDesde] = useState("");
  const [fHasta, setFHasta] = useState("");
  const [fStatus, setFStatus] = useState("");
  const [fEstadoDoc, setFEstadoDoc] = useState("");
  const [fIdCorr, setFIdCorr] = useState("");
  const [fNroSpot, setFNroSpot] = useState("");
  const [fRutas, setFRutas] = useState("");
  const [filterOpen, setFilterOpen] = useState(false);
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [modal, setModal] = useState(null);
  const [editModal, setEditModal] = useState(null);   // { viaje } — modal edición placa/rutas
  const [editPlaca, setEditPlaca] = useState("");
  const [editRutas, setEditRutas] = useState([]);     // array de strings
  const [editRutaInput, setEditRutaInput] = useState("");
  const [editSaving, setEditSaving] = useState(false);
  const [editErr, setEditErr] = useState("");
  const [corteInfo, setCorteInfo] = useState(null); // se fija en cada fetch dentro de rango válido
  const [uploading, setUploading] = useState(false);
  const [uploadFile, setUploadFile] = useState(null);
  const [uploadErr, setUploadErr] = useState("");
  const [uploadSuccess, setUploadSuccess] = useState(false);
  const fileRef = useRef();
  const userMenuRef = useRef();

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setLoading(false); });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    const handleClick = (e) => { if (userMenuRef.current && !userMenuRef.current.contains(e.target)) setUserMenuOpen(false); };
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  const fetchViajesRef = useRef(null);

  const fetchViajes = useCallback(async () => {
    if (!session) return;
    setDataLoading(true);
    const meta = session.user.user_metadata;
    const isAdmin = meta?.role === "admin";
    const filters = fetchViajesRef.current || {};
    let q = supabase.from("viajes").select("*").order("fecha_carga", { ascending: false });
    if (!isAdmin && meta?.empresa_id) q = q.eq("empresa_id", meta.empresa_id);
    if (filters.fDesde) q = q.gte("fecha_carga", filters.fDesde);
    if (filters.fHasta) q = q.lte("fecha_carga", filters.fHasta);
    if (filters.fStatus) q = q.ilike("status", `%${filters.fStatus}%`);
    if (filters.fNroSpot) q = q.ilike("nro_id_spot", `%${filters.fNroSpot}%`);
    if (filters.fRutas) q = q.ilike("rutas", `%${filters.fRutas}%`);
    if (filters.fIdCorr) q = q.ilike("id_correlativo", `%${filters.fIdCorr}%`);
    const { data } = await q;
    let rows = data || [];
    if (filters.fEstadoDoc === "Completo") rows = rows.filter(r => (r.foto_versiones?.length || 0) > 0);
    if (filters.fEstadoDoc === "Pendiente") rows = rows.filter(r => !(r.foto_versiones?.length > 0));

    // Orden de filas: fecha_carga desc (ya viene de Supabase) + desempate secundario
    // por los dígitos finales del id_correlativo (desc), solo para la vista —
    // no modifica nada en la base de datos.
    rows = [...rows].sort((a, b) => {
      if (a.fecha_carga !== b.fecha_carga) return 0; // respeta el orden por fecha que ya trajo Supabase
      return extraerNumeroId(b.id_correlativo) - extraerNumeroId(a.id_correlativo);
    });

    setViajes(rows);
    setDataLoading(false);

    // Corte: se fija según la hora REAL en que se hizo este fetch, no en cada render.
    // 8:00am-10:00am -> "8:30 am" fijo. 4:00pm-7:00pm -> "4:00 pm" fijo.
    // Fuera de esos rangos: se deja el último corte calculado, sin cambiarlo.
    const ahora = new Date();
    const hora = ahora.getHours() + ahora.getMinutes() / 60;
    const fechaStr = ahora.toLocaleDateString("es-PE", { day: "2-digit", month: "2-digit", year: "2-digit" });
    if (hora >= 8 && hora < 10) {
      setCorteInfo({ label: `${fechaStr} · 8:30 am` });
    } else if (hora >= 16 && hora < 19) {
      setCorteInfo({ label: `${fechaStr} · 4:00 pm` });
    }
    // si está fuera de rango, no se toca corteInfo (queda el último valor)
  }, [session]);

  // Solo carga al iniciar sesión — no refresca al cambiar de pestaña
  useEffect(() => { if (session) fetchViajes(); }, [session]);

  // Actualiza los filtros en el ref sin disparar refetch
  useEffect(() => {
    fetchViajesRef.current = { fDesde, fHasta, fStatus, fEstadoDoc, fNroSpot, fRutas, fIdCorr };
  }, [fDesde, fHasta, fStatus, fEstadoDoc, fNroSpot, fRutas, fIdCorr]);

  async function doLogin(e) {
    e.preventDefault();
    setLoginErr("");
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) setLoginErr("Usuario o contraseña incorrectos");
  }

  function openEditModal(viaje) {
    setEditModal(viaje);
    setEditPlaca(viaje.placa || "");
    // rutas viene como "GR001 | GR002 | GR003" → separar por " | "
    const rutasArr = viaje.rutas ? viaje.rutas.split("|").map(r => r.trim()).filter(Boolean) : [];
    setEditRutas(rutasArr);
    setEditRutaInput("");
    setEditErr("");
  }

  function addRuta() {
    const val = editRutaInput.trim();
    if (!val) return;
    if (editRutas.includes(val)) { setEditErr("Esa ruta ya está en la lista."); return; }
    setEditRutas(prev => [...prev, val]);
    setEditRutaInput("");
    setEditErr("");
  }

  function removeRuta(idx) {
    setEditRutas(prev => prev.filter((_, i) => i !== idx));
  }

  const PLACA_RE = /^[A-Za-z0-9]{3}-[A-Za-z0-9]{3}$/;

  async function saveEdit() {
    if (editPlaca && !PLACA_RE.test(editPlaca)) {
      setEditErr("Formato de placa inválido. Usa el formato ABC-123.");
      return;
    }
    setEditSaving(true);
    setEditErr("");
    try {
      const rutasStr = editRutas.join(" | ");
      const { error } = await supabase.from("viajes").update({
        placa: editPlaca || null,
        rutas: rutasStr || null,
      }).eq("id_correlativo", editModal.id_correlativo);
      if (error) throw error;
      // Actualizar la fila en estado local sin refetch completo
      setViajes(prev => prev.map(v =>
        v.id_correlativo === editModal.id_correlativo
          ? { ...v, placa: editPlaca || null, rutas: rutasStr || null }
          : v
      ));
      setEditModal(null);
    } catch (err) {
      setEditErr(err.message || "Error al guardar.");
    } finally {
      setEditSaving(false);
    }
  }

  function openModal(viaje) {
    setModal(viaje);
    setUploadFile(null);
    setUploadErr("");
    setUploadSuccess(false);
  }

  function handleFileSelect(file) {
    if (!file) return;
    if (file.size > MAX_MB * 1024 * 1024) {
      setUploadErr(`El archivo supera los ${MAX_MB} MB. Selecciona una imagen más pequeña.`);
      setUploadFile(null);
      return;
    }
    setUploadErr("");
    setUploadFile(file);
  }

  async function handleUpload() {
    if (!uploadFile || !modal) return;
    setUploading(true);
    setUploadErr("");
    try {
      const compressed = await comprimirImagen(uploadFile);
      const versiones = modal.foto_versiones || [];
      const nv = versiones.length + 1;
      const ext = compressed.name.split(".").pop();
      const path = `${modal.empresa_id}/${modal.id_correlativo}/foto_v${nv}.${ext}`;
      const { error: upErr } = await supabase.storage.from("evidencias").upload(path, compressed, { upsert: true });
      if (upErr) throw upErr;
      const { data: urlData } = supabase.storage.from("evidencias").getPublicUrl(path);
      const nuevasVersiones = [...versiones, { v: nv, url: urlData?.publicUrl || path, nombre: `foto_v${nv}.${ext}`, subido_por: session.user.email, subido_en: new Date().toISOString() }];
      const { error: dbErr } = await supabase.from("viajes").update({
        foto_versiones: nuevasVersiones,
        foto_url: urlData?.publicUrl || path,
        foto_nombre: `foto_v${nv}.${ext}`,
        subido_por: session.user.email,
        subido_en: new Date().toISOString(),
      }).eq("id_correlativo", modal.id_correlativo);
      if (dbErr) throw dbErr;
      setUploadSuccess(true);
      fetchViajes();
      setTimeout(() => setModal(null), 1600);
    } catch (err) {
      setUploadErr(err.message || "Error al subir el archivo.");
    } finally {
      setUploading(false);
    }
  }

  const meta = session?.user?.user_metadata;
  const isAdmin = meta?.role === "admin";
  const empresa = meta?.empresa_id || "";
  const initials = (session?.user?.email || "U").substring(0, 2).toUpperCase();
  const inp = { padding: "6px 10px", fontSize: 12, border: `0.5px solid ${BORDER}`, borderRadius: 8, background: "white", color: GRAY_900, outline: "none" };

  // Corte: se muestra el último corte fijado por fetchViajes (hora real de actualización).
  // Antes del primer fetch, se calcula un valor inicial razonable como placeholder.
  let corteLabel = corteInfo?.label;
  if (!corteLabel) {
    const ahora = new Date();
    const hora = ahora.getHours() + ahora.getMinutes() / 60;
    const fechaCorte = new Date(ahora);
    if (hora < 8) fechaCorte.setDate(fechaCorte.getDate() - 1);
    const fechaCorteStr = fechaCorte.toLocaleDateString("es-PE", { day: "2-digit", month: "2-digit", year: "2-digit" });
    corteLabel = (hora >= 8 && hora < 16.5) ? `${fechaCorteStr} · 8:00 am` : `${fechaCorteStr} · 4:30 pm`;
  }

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

  return (
    <div style={{ minHeight: "100vh", background: GRAY_100, display: "flex", flexDirection: "column" }}>

      {/* TOPBAR */}
      <div style={{ height: 52, background: RED, display: "flex", alignItems: "center", padding: "0 18px", gap: 10, flexShrink: 0 }}>
        <img src="/logo_fape.png" alt="FP" style={{ width: 30, height: 30, borderRadius: 6, background: "white", objectFit: "contain", padding: 2, flexShrink: 0 }} />
        <span style={{ fontSize: 16, fontWeight: 500, color: "white", letterSpacing: "-.3px" }}>Pharma<span style={{ opacity: .6, fontWeight: 400 }}>SPOT</span></span>
        <div style={{ width: 1, height: 18, background: "rgba(255,255,255,.25)" }} />
        <span style={{ fontSize: 13, color: "rgba(255,255,255,.75)", fontWeight: 400 }}>Seguimiento Adicionales</span>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          {/* Tarjeta de corte */}
          <div style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,.12)", border: "1px solid rgba(255,255,255,.2)", borderRadius: 8, padding: "5px 10px" }}>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#5DCAA5", flexShrink: 0 }} />
            <div>
              <div style={{ fontSize: 9, color: "rgba(255,255,255,.6)", textTransform: "uppercase", letterSpacing: ".05em" }}>Último corte</div>
              <div style={{ fontSize: 11, color: "white", fontWeight: 500 }}>{corteLabel}</div>
            </div>
          </div>
          {/* Menú usuario */}
          <div style={{ position: "relative" }} ref={userMenuRef}>
          <button onClick={() => setUserMenuOpen(o => !o)}
            style={{ width: 34, height: 34, borderRadius: "50%", background: "rgba(255,255,255,.2)", border: "1.5px solid rgba(255,255,255,.35)", color: "white", fontSize: 12, fontWeight: 500, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {initials}
          </button>
          {userMenuOpen && (
            <div style={{ position: "absolute", right: 0, top: 42, background: "white", border: `0.5px solid ${BORDER}`, borderRadius: 10, minWidth: 210, zIndex: 50, boxShadow: "0 4px 16px rgba(0,0,0,.1)" }}>
              <div style={{ padding: "10px 14px", borderBottom: `0.5px solid ${BORDER}` }}>
                <div style={{ fontSize: 12, fontWeight: 500, color: GRAY_900 }}>{session.user.email}</div>
                <div style={{ fontSize: 11, color: GRAY_500, marginTop: 2 }}>{isAdmin ? "Administrador" : `${empresa} · Transportista`}</div>
              </div>
              <button onClick={() => supabase.auth.signOut()}
                style={{ width: "100%", padding: "9px 14px", textAlign: "left", fontSize: 12, color: RED, cursor: "pointer", background: "none", border: "none" }}>
                Cerrar sesión
              </button>
            </div>
          )}
          </div>
        </div>
      </div>

      {/* TOOLBAR */}
      <div style={{ background: "white", borderBottom: `0.5px solid ${BORDER}`, padding: "8px 18px", display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
        <span style={{ fontSize: 14, fontWeight: 500, color: GRAY_900, flex: 1 }}>{isAdmin ? "Todos los transportes" : empresa}</span>
        <button onClick={() => { setFDesde(""); setFHasta(""); setFStatus(""); setFEstadoDoc(""); setFNroSpot(""); setFRutas(""); setFIdCorr(""); setFilterOpen(false); }}
          style={{ width: 34, height: 34, borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 15, display: "flex", alignItems: "center", justifyContent: "center" }} title="Limpiar filtros">✕</button>
        <button onClick={fetchViajes}
          style={{ width: 34, height: 34, borderRadius: 999, border: `0.5px solid ${BORDER}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }} title="Actualizar">↻</button>
        <button onClick={() => setFilterOpen(o => !o)}
          style={{ height: 34, padding: "0 16px", borderRadius: 999, border: `0.5px solid ${filterOpen ? RED : BORDER}`, background: filterOpen ? RED : "white", color: filterOpen ? "white" : GRAY_900, cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}>
          ⚙ Filtrar
        </button>
      </div>

      {/* FILTER PANEL */}
      {filterOpen && (
        <div style={{ background: "white", borderBottom: `0.5px solid ${BORDER}`, padding: "12px 18px", display: "flex", gap: 14, flexWrap: "wrap", alignItems: "flex-end", flexShrink: 0 }}>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>ID Correlativo</div>
            <input value={fIdCorr} onChange={e => setFIdCorr(e.target.value)} placeholder="Ej: 652" style={{ ...inp, minWidth: 120 }} />
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Nro SPOT</div>
            <input value={fNroSpot} onChange={e => setFNroSpot(e.target.value)} placeholder="Ej: 4001234" style={{ ...inp, minWidth: 120 }} />
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Rutas / GR</div>
            <input value={fRutas} onChange={e => setFRutas(e.target.value)} style={{ ...inp, minWidth: 140 }} />
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Desde</div>
            <input type="date" value={fDesde} onChange={e => setFDesde(e.target.value)} style={inp} />
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Hasta</div>
            <input type="date" value={fHasta} onChange={e => setFHasta(e.target.value)} style={inp} />
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Status viaje</div>
            <select value={fStatus} onChange={e => setFStatus(e.target.value)} style={inp}>
              <option value="">Todos</option>
              <option>Realizado</option>
              <option>No Ejecutado</option>
            </select>
          </div>
          <div>
            <div style={{ fontSize: 10, color: GRAY_500, marginBottom: 4, textTransform: "uppercase", letterSpacing: ".04em" }}>Estado doc</div>
            <select value={fEstadoDoc} onChange={e => setFEstadoDoc(e.target.value)} style={inp}>
              <option value="">Todos</option>
              <option value="Completo">Completo</option>
              <option value="Pendiente">Pendiente</option>
            </select>
          </div>
          <button onClick={fetchViajes} style={{ height: 34, padding: "0 18px", background: RED, color: "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: "pointer" }}>Aplicar</button>
        </div>
      )}

      {/* TABLE */}
      <div style={{ flex: 1, padding: "14px 18px", overflow: "hidden", display: "flex", flexDirection: "column" }}>
        <div style={{ overflowX: "auto", overflowY: "auto", flex: 1, background: "white", borderRadius: 10, border: `0.5px solid ${BORDER}` }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, tableLayout: "fixed" }}>
              <colgroup>
                {COLS.map(c => <col key={c.key} style={{ width: c.width || 100 }} />)}
                <col style={{ width: 50 }} />
                <col style={{ width: 96 }} />
                <col style={{ width: 96 }} />
              </colgroup>
              <thead>
                <tr>
                  {COLS.map(c => (
                    <th key={c.key} style={{ padding: "8px 11px", textAlign: c.right ? "right" : "left", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", position: "sticky", top: 0, zIndex: 3, textTransform: "uppercase", letterSpacing: ".03em" }}>
                      {c.label}
                    </th>
                  ))}
                  <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: 192, zIndex: 4, textTransform: "uppercase", letterSpacing: ".03em", borderLeft: `0.5px solid ${BORDER}` }}>
                    {/* lápiz — sin label */}
                  </th>
                  <th style={{ padding: "8px 11px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: 96, zIndex: 4, textTransform: "uppercase", letterSpacing: ".03em", borderLeft: `0.5px solid ${BORDER}` }}>
                    Estado doc
                  </th>
                  <th style={{ padding: "8px 11px", textAlign: "center", fontSize: 10, fontWeight: 500, color: GRAY_500, background: GRAY_50, borderBottom: `0.5px solid ${BORDER}`, whiteSpace: "nowrap", position: "sticky", top: 0, right: 0, zIndex: 4, textTransform: "uppercase", letterSpacing: ".03em", borderLeft: `1.5px solid ${BORDER}` }}>
                    Doc. adjuntos
                  </th>
                </tr>
              </thead>
              <tbody>
                {viajes.length === 0 ? (
                  <tr><td colSpan={COLS.length + 2} style={{ padding: 40, textAlign: "center", color: GRAY_500, fontSize: 12 }}>No hay viajes que mostrar</td></tr>
                ) : viajes.map(v => {
                  const versiones = v.foto_versiones || [];
                  const nv = versiones.length;
                  const completo = nv > 0;
                  return (
                    <tr key={v.id_correlativo}
                      onMouseEnter={e => e.currentTarget.querySelectorAll("td").forEach(td => { if (td.dataset.editable) td.style.background = "#FFF5E8"; else td.style.background = "#FFF5F5"; })}
                      onMouseLeave={e => e.currentTarget.querySelectorAll("td").forEach(td => { if (td.dataset.editable) td.style.background = "#FFFBF0"; else td.style.background = "white"; })}>
                      {COLS.map(c => {
                        let content;
                        if (c.key === "status") {
                          const ok = v.status?.toLowerCase().includes("realizado");
                          const warn = v.status?.toLowerCase().includes("no");
                          content = <span style={{ display: "inline-flex", padding: "2px 8px", borderRadius: 999, fontSize: 10, fontWeight: 500, background: ok ? GREEN_LIGHT : warn ? "#faeeda" : GRAY_100, color: ok ? GREEN : warn ? "#854f0b" : GRAY_500 }}>{v[c.key] || "—"}</span>;
                        } else if (c.key === "tipo_traslado") {
                          content = <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 4, background: v[c.key] === "TEA" ? "#e6f1fb" : RED_LIGHT, color: v[c.key] === "TEA" ? "#185fa5" : RED_DARK, fontWeight: 500 }}>{v[c.key] || "—"}</span>;
                        } else if (c.key === "importe_servicio") {
                          content = v[c.key] && v[c.key] !== "—" ? `S/${parseFloat(v[c.key]).toLocaleString("es-PE")}` : "—";
                        } else {
                          content = v[c.key] || "—";
                        }
                        const editable = !isAdmin && (c.key === "placa" || c.key === "rutas");
                        return (
                          <td key={c.key} title={c.trunc ? (v[c.key] || "") : undefined}
                            data-editable={editable ? "1" : undefined}
                            style={{ padding: "8px 11px", borderBottom: `0.5px solid ${BORDER}`, borderRight: `0.5px solid ${BORDER}`, verticalAlign: "middle", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontFamily: c.mono ? "monospace" : "inherit", fontSize: c.mono ? 10 : 11, textAlign: c.right ? "right" : "left", color: c.muted ? GRAY_500 : GRAY_900, background: editable ? "#FFFBF0" : undefined }}>
                            {content}
                          </td>
                        );
                      })}
                      {/* Botón lápiz — solo visible para transportistas */}
                      <td style={{ padding: "4px 6px", borderBottom: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: 192, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center", overflow: "hidden" }}>
                        {!isAdmin && (
                          <button onClick={() => openEditModal(v)}
                            title="Editar Placa y Rutas/GR"
                            style={{ width: 26, height: 26, borderRadius: "50%", border: `1px solid ${GRAY_200}`, background: "white", color: GRAY_500, cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto" }}>
                            ✏️
                          </button>
                        )}
                      </td>
                      <td style={{ padding: "8px 11px", borderBottom: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: 96, background: "white", borderLeft: `0.5px solid ${BORDER}`, zIndex: 2, textAlign: "center", overflow: "hidden" }}>
                        <span style={{ display: "inline-flex", padding: "2px 9px", borderRadius: 999, fontSize: 10, fontWeight: 500, background: completo ? GREEN_LIGHT : RED_LIGHT, color: completo ? GREEN : RED }}>
                          {completo ? "Completo" : "Pendiente"}
                        </span>
                      </td>
                      <td style={{ padding: "8px 11px", borderBottom: `0.5px solid ${BORDER}`, verticalAlign: "middle", position: "sticky", right: 0, background: "white", borderLeft: `1.5px solid ${BORDER}`, zIndex: 2, textAlign: "center", overflow: "hidden" }}>
                        <div style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                          <button onClick={() => openModal(v)}
                            style={{ width: 28, height: 28, borderRadius: "50%", border: `1.5px solid ${completo ? GREEN : GRAY_200}`, cursor: "pointer", background: completo ? GREEN_LIGHT : "white", color: completo ? GREEN : GRAY_500, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center" }}
                            title={completo ? "Agregar nuevo documento" : "Subir documento"}>
                            {completo ? "✓" : "↑"}
                          </button>
                          {completo && (
                            <button onClick={() => openModal(v)}
                              style={{ width: 20, height: 20, borderRadius: "50%", border: `1px solid ${BORDER}`, cursor: "pointer", background: "white", color: GRAY_500, fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center" }}
                              title="Agregar nuevo documento">↑</button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
        </div>
        <div style={{ fontSize: 11, color: GRAY_500, marginTop: 8, textAlign: "right" }}>
          {viajes.length} viajes · {viajes.filter(v => (v.foto_versiones?.length || 0) > 0).length} con documento · {viajes.filter(v => !(v.foto_versiones?.length > 0)).length} pendientes
        </div>
      </div>

      {/* MODAL EDICIÓN PLACA / RUTAS */}
      {editModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
          onClick={e => e.target === e.currentTarget && !editSaving && setEditModal(null)}>
          <div style={{ background: "white", borderRadius: 14, padding: 24, width: 420, maxWidth: "94vw", maxHeight: "90vh", overflowY: "auto" }}>

            {/* Header */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
              <div style={{ fontSize: 14, fontWeight: 500 }}>Editar datos del viaje</div>
              <button onClick={() => setEditModal(null)} disabled={editSaving}
                style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 18 }}>
              {editModal.id_correlativo} · {editModal.cd_origen} → {editModal.cd_destino}
            </div>

            {/* Placa */}
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>Placa</div>
              <input
                value={editPlaca}
                onChange={e => { setEditPlaca(e.target.value.toUpperCase()); setEditErr(""); }}
                placeholder="ABC-123"
                maxLength={7}
                style={{ ...inp, width: "100%", boxSizing: "border-box", fontFamily: "monospace", fontSize: 13, letterSpacing: ".08em", background: "#FFFBF0", border: `1px solid ${BORDER}` }}
              />
              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>Formato: 3 letras o números, guion, 3 letras o números. Ej: ABC-123</div>
            </div>

            {/* Rutas / GR */}
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 6, fontWeight: 500, textTransform: "uppercase", letterSpacing: ".04em" }}>Rutas / GR</div>

              {/* Chips */}
              {editRutas.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
                  {editRutas.map((r, i) => (
                    <span key={i} style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 10px", background: "#FFFBF0", border: `1px solid ${BORDER}`, borderRadius: 999, fontSize: 11, color: GRAY_900 }}>
                      {r}
                      <button onClick={() => removeRuta(i)}
                        style={{ background: "none", border: "none", cursor: "pointer", color: GRAY_500, fontSize: 11, lineHeight: 1, padding: 0, display: "flex", alignItems: "center" }}>✕</button>
                    </span>
                  ))}
                </div>
              )}

              {/* Input + botón + */}
              <div style={{ display: "flex", gap: 6 }}>
                <input
                  value={editRutaInput}
                  onChange={e => { setEditRutaInput(e.target.value); setEditErr(""); }}
                  onKeyDown={e => e.key === "Enter" && (e.preventDefault(), addRuta())}
                  placeholder="Ej: GR001234 o Ruta Trujillo"
                  style={{ ...inp, flex: 1, background: "#FFFBF0", border: `1px solid ${BORDER}` }}
                />
                <button onClick={addRuta}
                  style={{ width: 34, height: 34, borderRadius: 8, border: `1px solid ${BORDER}`, background: "#FFFBF0", color: GRAY_900, fontSize: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 400 }}>+</button>
              </div>
              <div style={{ fontSize: 10, color: GRAY_500, marginTop: 4 }}>Presiona Enter o el botón + para agregar cada ruta o GR.</div>
            </div>

            {/* Error */}
            {editErr && (
              <div style={{ padding: "8px 12px", background: RED_LIGHT, borderRadius: 8, fontSize: 11, color: RED_DARK, marginBottom: 12, border: `0.5px solid #f7c1c1` }}>
                ⚠ {editErr}
              </div>
            )}

            {/* Botones */}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setEditModal(null)} disabled={editSaving}
                style={{ padding: "7px 14px", border: `0.5px solid ${BORDER}`, borderRadius: 8, fontSize: 12, cursor: "pointer", background: "none", color: GRAY_500 }}>
                Cancelar
              </button>
              <button onClick={saveEdit} disabled={editSaving}
                style={{ padding: "7px 16px", background: editSaving ? GRAY_200 : RED, color: editSaving ? GRAY_500 : "white", border: "none", borderRadius: 8, fontSize: 12, fontWeight: 500, cursor: editSaving ? "default" : "pointer" }}>
                {editSaving ? "Guardando..." : "Guardar"}
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
              <div style={{ fontSize: 14, fontWeight: 500 }}>Subir documento</div>
              <button onClick={() => setModal(null)} style={{ width: 22, height: 22, borderRadius: "50%", border: `0.5px solid ${BORDER}`, background: "none", cursor: "pointer", fontSize: 12, color: GRAY_500, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
            </div>
            <div style={{ fontSize: 11, color: GRAY_500, marginBottom: 14 }}>{modal.id_correlativo} · {modal.cd_origen} → {modal.cd_destino}</div>

            {uploadSuccess ? (
              <div style={{ textAlign: "center", padding: "24px 0" }}>
                <div style={{ fontSize: 32, marginBottom: 8 }}>✅</div>
                <div style={{ fontSize: 13, color: GREEN, fontWeight: 500 }}>Documento guardado correctamente</div>
              </div>
            ) : (
              <>
                {(modal.foto_versiones?.length || 0) > 0 && (
                  <div style={{ marginBottom: 14, padding: "10px 12px", background: GREEN_LIGHT, borderRadius: 8, fontSize: 12, color: GREEN }}>
                    ✓ Ya tienes {modal.foto_versiones.length} documento(s) subido(s). Puedes agregar otro si necesitas corregir.
                  </div>
                )}
                <div onClick={() => fileRef.current?.click()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => { e.preventDefault(); handleFileSelect(e.dataTransfer.files[0]); }}
                  style={{ border: `1.5px dashed ${GRAY_200}`, borderRadius: 10, padding: "22px 16px", textAlign: "center", cursor: "pointer", marginBottom: 12 }}>
                  <div style={{ fontSize: 22, color: GRAY_200, marginBottom: 6 }}>📷</div>
                  <div style={{ fontSize: 12, color: GRAY_500 }}>
                    {uploadFile ? uploadFile.name : "Clic o arrastra tu foto aquí"}
                  </div>
                </div>

                {uploadFile && (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", background: GRAY_100, borderRadius: 8, fontSize: 11, color: GRAY_900, marginBottom: 12, border: `0.5px solid ${BORDER}` }}>
                    <span>📎</span>
                    <span style={{ flex: 1 }}>Archivo adjunto: {uploadFile.name}</span>
                  </div>
                )}

                {uploadErr && (
                  <div style={{ padding: "8px 12px", background: RED_LIGHT, borderRadius: 8, fontSize: 11, color: RED_DARK, marginBottom: 12, border: `0.5px solid #f7c1c1` }}>
                    ⚠ {uploadErr}
                  </div>
                )}

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
