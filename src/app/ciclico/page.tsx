"use client";

import { useEffect, useRef, useState, useMemo } from "react";
import { supabase } from "@/lib/supabase/client";
import * as XLSX from "xlsx";
import { QrCode } from "lucide-react";

// ── Clave de sesión propia (no comparte con el sistema de inventario general) ──
const CYCLIC_SESSION_KEY = "cyclic_session_user";
const CYCLIC_STORE_KEY   = "cyclic_current_store_id";

// ── Types ─────────────────────────────────────────────────────────────────────
type Role = "Operario" | "Validador" | "Administrador";

type CyclicUser = {
    id: string;
    username: string;
    full_name: string;
    role: Role;
};

type CyclicUserRow = {
    id: string;
    username: string;
    password: string;
    full_name: string;
    role: Role;
    is_active: boolean;
};

type Store = {
    id: string;
    name: string;
    code: string;
    is_active: boolean;
};

type CyclicSession = {
    id: string;
    store_id: string;
    session_date: string;
    created_by_name: string;
    status: "activo" | "cerrado";
    created_at: string;
};

type Assignment = {
    id: string;
    session_id: string;
    store_id: string;
    sku: string;
    description: string;
    unit: string;
    cost: number;
    system_stock: number | null;
    assigned_to_user_id: string | null;
    assigned_to_name: string | null;
    counted_quantity: number | null;
    location: string | null;
    note: string | null;
    counted_at: string | null;
    counted_by_name: string | null;
    status: "pendiente" | "contado" | "validado";
    is_extra: boolean;
};

// Producto del maestro general (solo lectura compartida)
type SharedProduct = {
    id: string;
    sku: string;
    description: string;
    unit: string;
    cost: number;
    system_stock: number;
    inventory_id: string | null;
};

type TabKey = "operario" | "validador" | "admin";

// ── Helpers ──────────────────────────────────────────────────────────────────
function cleanCode(value: string | null | undefined): string {
    if (!value) return "";
    let s = String(value).trim();
    s = s.replace(/^['"''""\u2018\u2019\u201C\u201D]+/, "").replace(/['"''""\u2018\u2019\u201C\u201D]+$/, "").trim();
    if (/[Ee][+-]/.test(s) && !isNaN(Number(s))) {
        const n = Number(s);
        if (isFinite(n)) s = Math.round(n).toString();
    }
    s = s.replace(/\.0+$/, "");
    if (/^\d+$/.test(s)) { s = s.replace(/^0+/, ""); if (s === "") s = "0"; }
    return s;
}

function normalizeText(v: string | null | undefined) {
    return String(v || "").trim().toLowerCase();
}

function formatMoney(value: number) {
    return `S/ ${Number(value || 0).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function formatDateTime(value: string | null | undefined) {
    if (!value) return "-";
    const d = new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return d.toLocaleString("es-PE");
}

function diffBadge(diff: number | null) {
    if (diff === null) return <span className="text-slate-400 text-xs">-</span>;
    if (diff === 0)   return <span className="text-green-700 font-bold text-xs">0</span>;
    if (diff > 0)     return <span className="text-blue-700 font-bold text-xs">+{diff}</span>;
    return <span className="text-red-600 font-bold text-xs">{diff}</span>;
}

function statusBadge(status: Assignment["status"]) {
    const base = "inline-block px-2 py-0.5 rounded-full text-xs font-semibold";
    if (status === "validado") return `${base} bg-green-100 text-green-700`;
    if (status === "contado")  return `${base} bg-blue-100 text-blue-700`;
    return `${base} bg-slate-100 text-slate-600`;
}

// ── Main Component ────────────────────────────────────────────────────────────
export default function CiclicoDashboardPage() {
    const [user, setUser]           = useState<CyclicUser | null>(null);
    const [activeTab, setActiveTab] = useState<TabKey>("operario");
    const [message, setMessage]     = useState("");
    const [messageType, setMessageType] = useState<"info" | "success" | "error">("info");
    const messageTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

    // Stores
    const [stores, setStores]               = useState<Store[]>([]);
    const [allStores, setAllStores]         = useState<Store[]>([]);
    const [selectedStoreId, setSelectedStoreId] = useState("");

    // Session
    const [session, setSession]             = useState<CyclicSession | null>(null);
    const [assignments, setAssignments]     = useState<Assignment[]>([]);
    const [loadingSession, setLoadingSession] = useState(false);

    // Users list (para asignar)
    const [users, setUsers] = useState<CyclicUserRow[]>([]);

    // Operario state
    const [myAssignments, setMyAssignments] = useState<Assignment[]>([]);
    const [countInputs, setCountInputs]     = useState<Record<string, { qty: string; location: string; note: string }>>({});
    const [operarioSearch, setOperarioSearch] = useState("");

    // Extra product (búsqueda en maestro compartido)
    const [showAddExtra, setShowAddExtra]   = useState(false);
    const [extraSearch, setExtraSearch]     = useState("");
    const [extraResults, setExtraResults]   = useState<SharedProduct[]>([]);
    const [extraSelected, setExtraSelected] = useState<SharedProduct | null>(null);

    // Validador
    const [assignCount, setAssignCount]     = useState("80");
    const [assignUserId, setAssignUserId]   = useState("");
    const [validadorFilter, setValidadorFilter] = useState<"todos" | "pendiente" | "contado" | "validado">("todos");
    const [validadorUserFilter, setValidadorUserFilter] = useState("");
    const [validadorSearch, setValidadorSearch] = useState("");

    // Admin — crear/editar usuario
    const [newUsername, setNewUsername]     = useState("");
    const [newPassword, setNewPassword]     = useState("");
    const [newFullName, setNewFullName]     = useState("");
    const [newRole, setNewRole]             = useState<Role>("Operario");
    const [showNewPassword, setShowNewPassword] = useState(false);

    const [editingUser, setEditingUser]     = useState<CyclicUserRow | null>(null);
    const [editRole, setEditRole]           = useState<Role>("Operario");
    const [editActive, setEditActive]       = useState(true);
    const [editPassword, setEditPassword]   = useState("");
    const [showEditPassword, setShowEditPassword] = useState(false);

    // Admin — stores
    const [newStoreName, setNewStoreName]   = useState("");
    const [newStoreCode, setNewStoreCode]   = useState("");

    // Admin — subir lista
    const [assignFile, setAssignFile]       = useState<File | null>(null);
    const [assignFileName, setAssignFileName] = useState("");
    const [uploadProgress, setUploadProgress] = useState<{ step: string; pct: number } | null>(null);
    const assignInputRef = useRef<HTMLInputElement | null>(null);

    // Scanner
    const [scannerTarget, setScannerTarget] = useState<"extra" | null>(null);
    const [scannerRunning, setScannerRunning] = useState(false);
    const [torchAvailable, setTorchAvailable] = useState(false);
    const [torchOn, setTorchOn]             = useState(false);
    const scannerRef      = useRef<any>(null);
    const scanHandledRef  = useRef(false);
    const scannerContainerId = "cyclic-scanner";

    // ── Messages ─────────────────────────────────────────────────────────────
    function showMessage(msg: string, type: "info" | "success" | "error" = "info") {
        if (messageTimerRef.current) clearTimeout(messageTimerRef.current);
        setMessage(msg); setMessageType(type);
        if (type === "success") messageTimerRef.current = setTimeout(() => setMessage(""), 4000);
    }
    function clearMessage() { setMessage(""); }

    // ── Init ─────────────────────────────────────────────────────────────────
    useEffect(() => {
        const raw = localStorage.getItem(CYCLIC_SESSION_KEY);
        if (!raw) { window.location.href = "/ciclico/login"; return; }
        const parsed = JSON.parse(raw) as CyclicUser;
        setUser(parsed);
        if (parsed.role === "Validador")     setActiveTab("validador");
        if (parsed.role === "Administrador") setActiveTab("admin");
    }, []);

    useEffect(() => { if (user) loadStores(); }, [user]);
    useEffect(() => {
        if (user && selectedStoreId) {
            loadSession();
            loadUsers();
        }
    }, [user, selectedStoreId]);

    // ── Realtime ─────────────────────────────────────────────────────────────
    useEffect(() => {
        if (!selectedStoreId) return;
        const channel = supabase
            .channel(`cyclic-rt-${selectedStoreId}`)
            .on("postgres_changes", { event: "*", schema: "public", table: "cyclic_assignments" }, () => {
                if (session) loadAssignments(session.id);
            })
            .on("postgres_changes", { event: "*", schema: "public", table: "cyclic_sessions" }, () => {
                loadSession();
            })
            .subscribe();
        return () => { supabase.removeChannel(channel); };
    }, [selectedStoreId, session?.id]);

    // ── Stores ────────────────────────────────────────────────────────────────
    async function loadStores() {
        if (!user) return;
        const { data: all } = await supabase.from("cyclic_stores").select("*").order("name");
        setAllStores((all || []) as Store[]);
        const { data } = await supabase.from("cyclic_stores").select("*").eq("is_active", true).order("name");
        const list = (data || []) as Store[];
        setStores(list);
        if (!list.length) return;
        const saved = localStorage.getItem(CYCLIC_STORE_KEY);
        const exists = list.find(s => s.id === saved);
        const sel = exists ? exists.id : list[0].id;
        setSelectedStoreId(sel);
        localStorage.setItem(CYCLIC_STORE_KEY, sel);
    }

    function handleStoreChange(id: string) {
        setSelectedStoreId(id);
        localStorage.setItem(CYCLIC_STORE_KEY, id);
        setSession(null); setAssignments([]); setMyAssignments([]);
    }

    // ── Sessions ─────────────────────────────────────────────────────────────
    async function loadSession() {
        if (!selectedStoreId) return;
        setLoadingSession(true);
        const today = new Date().toISOString().split("T")[0];
        const { data: sess } = await supabase
            .from("cyclic_sessions").select("*")
            .eq("store_id", selectedStoreId)
            .eq("session_date", today)
            .eq("status", "activo")
            .order("created_at", { ascending: false })
            .limit(1).maybeSingle();
        if (sess) {
            setSession(sess as CyclicSession);
            await loadAssignments(sess.id);
        } else {
            setSession(null); setAssignments([]); setMyAssignments([]);
        }
        setLoadingSession(false);
    }

    async function loadAssignments(sessionId: string) {
        const { data } = await supabase
            .from("cyclic_assignments").select("*")
            .eq("session_id", sessionId)
            .order("is_extra").order("sku");
        const all = (data || []) as Assignment[];
        setAssignments(all);
        if (user) setMyAssignments(all.filter(a => a.assigned_to_user_id === user.id));
    }

    async function loadUsers() {
        const { data } = await supabase.from("cyclic_users").select("*").order("full_name");
        setUsers((data || []) as CyclicUserRow[]);
    }

    async function createSession() {
        if (!user || !selectedStoreId) return;
        const today = new Date().toISOString().split("T")[0];
        const { data, error } = await supabase.from("cyclic_sessions").insert({
            store_id: selectedStoreId,
            session_date: today,
            created_by_name: user.full_name,
            status: "activo",
        }).select().single();
        if (error) { showMessage("Error al crear sesión: " + error.message, "error"); return; }
        setSession(data as CyclicSession);
        setAssignments([]); setMyAssignments([]);
        showMessage("✅ Sesión del día creada correctamente.", "success");
    }

    async function closeSession() {
        if (!session) return;
        const pending = assignments.filter(a => a.status === "pendiente").length;
        if (pending > 0) {
            const ok = window.confirm(`Hay ${pending} asignaciones pendientes. ¿Cerrar igualmente?`);
            if (!ok) return;
        }
        await supabase.from("cyclic_sessions").update({ status: "cerrado" }).eq("id", session.id);
        showMessage("✅ Sesión cerrada.", "success");
        setSession(null); setAssignments([]); setMyAssignments([]);
    }

    // ── Asignar códigos a operario ────────────────────────────────────────────
    async function assignCodes() {
        if (!session || !assignUserId) { showMessage("Selecciona un operario.", "error"); return; }
        const qty = parseInt(assignCount) || 80;
        const assignedUser = users.find(u => u.id === assignUserId);
        if (!assignedUser) return;
        const unassigned = assignments.filter(a => !a.assigned_to_user_id && !a.is_extra);
        if (unassigned.length === 0) { showMessage("No hay códigos sin asignar.", "error"); return; }
        const toAssign = unassigned.slice(0, qty);
        await Promise.all(toAssign.map(a =>
            supabase.from("cyclic_assignments").update({
                assigned_to_user_id: assignedUser.id,
                assigned_to_name: assignedUser.full_name,
            }).eq("id", a.id)
        ));
        showMessage(`✅ ${toAssign.length} códigos asignados a ${assignedUser.full_name}.`, "success");
        await loadAssignments(session.id);
    }

    async function revokeAssignment(assignmentId: string) {
        await supabase.from("cyclic_assignments").update({
            assigned_to_user_id: null,
            assigned_to_name: null,
        }).eq("id", assignmentId);
        if (session) await loadAssignments(session.id);
    }

    // ── Subir lista de SKU del día (con stock sistema) ────────────────────────
    async function uploadAssignmentList() {
        if (!session) { showMessage("Primero crea la sesión del día.", "error"); return; }
        if (!assignFile) { showMessage("Selecciona un archivo.", "error"); return; }
        const ok = window.confirm("Se reemplazarán los códigos de la sesión actual (no se tocan los extras). ¿Continuar?");
        if (!ok) return;
        try {
            const raw = await assignFile.arrayBuffer();
            const wb = XLSX.read(raw);
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

            // Eliminar solo los no-extras
            await supabase.from("cyclic_assignments").delete()
                .eq("session_id", session.id).eq("is_extra", false);

            const toInsert: any[] = [];
            setUploadProgress({ step: "Leyendo archivo...", pct: 10 });

            for (const row of rows) {
                const rawSku = row.SKU || row.sku || row.CODIGO || row.codigo || "";
                const sku = cleanCode(rawSku);
                if (!sku) continue;

                // Buscar en maestro compartido (solo lectura)
                const { data: prodData } = await supabase
                    .from("products").select("*")
                    .ilike("sku", sku)
                    .limit(1).maybeSingle();
                const prod = prodData as SharedProduct | null;

                // El stock sistema viene del Excel o del maestro
                const excelStock = row.STOCK ?? row.stock ?? row.STOCK_SISTEMA ?? row.stock_sistema ?? "";
                const systemStock = excelStock !== "" ? Number(excelStock) : (prod?.system_stock ?? null);

                toInsert.push({
                    session_id:   session.id,
                    store_id:     selectedStoreId,
                    sku,
                    description:  prod?.description || String(row.DESCRIPCION || row.description || "").trim() || sku,
                    unit:         prod?.unit        || String(row["UNIDAD DE MEDIDA"] || row.unit || "").trim(),
                    cost:         prod?.cost        || Number(row.COSTO || row.cost || 0),
                    system_stock: systemStock,
                    status:       "pendiente",
                    is_extra:     false,
                });
            }

            if (!toInsert.length) { showMessage("El archivo no tiene filas válidas.", "error"); setUploadProgress(null); return; }

            for (let i = 0; i < toInsert.length; i += 200) {
                const pct = Math.round(20 + ((i / toInsert.length) * 75));
                setUploadProgress({ step: `Insertando ${Math.min(i + 200, toInsert.length)} de ${toInsert.length}...`, pct });
                const { error } = await supabase.from("cyclic_assignments").insert(toInsert.slice(i, i + 200));
                if (error) { showMessage("Error al insertar: " + error.message, "error"); setUploadProgress(null); return; }
            }
            setUploadProgress(null);
            showMessage(`✅ ${toInsert.length} SKUs cargados para la sesión.`, "success");
            setAssignFile(null); setAssignFileName("");
            if (assignInputRef.current) assignInputRef.current.value = "";
            await loadAssignments(session.id);
        } catch (err: any) {
            setUploadProgress(null);
            showMessage("Error: " + (err?.message || "desconocido"), "error");
        }
    }

    // ── Operario: guardar conteo ──────────────────────────────────────────────
    async function saveCount(assignmentId: string) {
        if (!user) return;
        const inp = countInputs[assignmentId];
        if (!inp?.qty || !inp?.location) { showMessage("Ingresa cantidad y ubicación.", "error"); return; }
        const qty = Number(inp.qty);
        if (isNaN(qty) || qty < 0) { showMessage("Cantidad inválida.", "error"); return; }
        const { error } = await supabase.from("cyclic_assignments").update({
            counted_quantity: qty,
            location:         inp.location.trim(),
            note:             inp.note?.trim() || null,
            counted_at:       new Date().toISOString(),
            counted_by_name:  user.full_name,
            status:           "contado",
        }).eq("id", assignmentId);
        if (error) { showMessage("Error: " + error.message, "error"); return; }
        showMessage("✅ Conteo guardado.", "success");
        setCountInputs(prev => { const n = { ...prev }; delete n[assignmentId]; return n; });
        if (session) await loadAssignments(session.id);
    }

    // ── Operario: buscar producto extra (en maestro compartido + código barras) ─
    async function searchExtraProduct(text: string) {
        setExtraSearch(text);
        if (!text.trim()) { setExtraResults([]); return; }
        const words = text.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").split(/\s+/).filter(Boolean);

        let q1 = supabase.from("products").select("*");
        for (const w of words) q1 = q1.ilike("description", `%${w}%`);
        const { data: byDesc } = await q1.limit(100);

        let q2 = supabase.from("products").select("*");
        for (const w of words) q2 = q2.ilike("sku", `%${w}%`);
        const { data: bySku } = await q2.limit(100);

        const combined = [...(byDesc || []), ...(bySku || [])];
        const seen = new Set<string>();
        const deduped = combined.filter(p => { if (seen.has(p.id)) return false; seen.add(p.id); return true; });
        setExtraResults(deduped.filter(p => {
            const h = (p.sku + " " + p.description).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            return words.every(w => h.includes(w));
        }).slice(0, 20) as SharedProduct[]);
    }

    async function findExtraByBarcode(rawValue: string) {
        const value = cleanCode(rawValue);
        if (!value) return;

        // 1. Buscar en catálogo global de barcodes (compartido)
        const { data: gb } = await supabase.from("global_barcodes").select("sku").eq("barcode", value).maybeSingle();
        const skuToFind = gb?.sku || value;

        // 2. Buscar el producto en el maestro compartido
        const { data: prod } = await supabase.from("products").select("*").ilike("sku", skuToFind).limit(1).maybeSingle();
        if (prod) {
            setExtraSelected(prod as SharedProduct);
            setExtraSearch(prod.description);
            setExtraResults([]);
            showMessage("Producto encontrado: " + prod.sku, "success");
        } else {
            // Intentar sin ceros a la izquierda
            const trimmed = value.replace(/^0+/, "");
            if (trimmed !== value) {
                const { data: gb2 } = await supabase.from("global_barcodes").select("sku").eq("barcode", trimmed).maybeSingle();
                if (gb2?.sku) {
                    const { data: prod2 } = await supabase.from("products").select("*").ilike("sku", gb2.sku).limit(1).maybeSingle();
                    if (prod2) { setExtraSelected(prod2 as SharedProduct); setExtraSearch(prod2.description); setExtraResults([]); return; }
                }
            }
            showMessage(`Código "${value}" no encontrado en el catálogo.`, "error");
        }
    }

    async function addExtraAssignment() {
        if (!session || !user || !extraSelected) { showMessage("Selecciona un producto primero.", "error"); return; }
        const exists = assignments.find(a => normalizeText(a.sku) === normalizeText(extraSelected.sku));
        if (exists) { showMessage("Este código ya está en la sesión.", "error"); return; }
        const { error } = await supabase.from("cyclic_assignments").insert({
            session_id:           session.id,
            store_id:             selectedStoreId,
            sku:                  extraSelected.sku,
            description:          extraSelected.description,
            unit:                 extraSelected.unit,
            cost:                 extraSelected.cost,
            system_stock:         extraSelected.system_stock ?? null,
            assigned_to_user_id:  user.id,
            assigned_to_name:     user.full_name,
            status:               "pendiente",
            is_extra:             true,
        });
        if (error) { showMessage("Error: " + error.message, "error"); return; }
        showMessage("✅ Código agregado a tu lista.", "success");
        setShowAddExtra(false); setExtraSelected(null); setExtraSearch(""); setExtraResults([]);
        await loadAssignments(session.id);
    }

    // ── Validador: validar / desvalidar ───────────────────────────────────────
    async function validateCount(assignmentId: string) {
        await supabase.from("cyclic_assignments").update({ status: "validado" }).eq("id", assignmentId);
        if (session) await loadAssignments(session.id);
    }

    async function unvalidateCount(assignmentId: string) {
        await supabase.from("cyclic_assignments").update({ status: "contado" }).eq("id", assignmentId);
        if (session) await loadAssignments(session.id);
    }

    // ── Export ────────────────────────────────────────────────────────────────
    function exportSession() {
        if (!assignments.length) return;
        const rows = assignments.map(a => ({
            SKU:              a.sku,
            DESCRIPCION:      a.description,
            UM:               a.unit,
            COSTO:            a.cost,
            STOCK_SISTEMA:    a.system_stock ?? "",
            DIFERENCIA:       (a.counted_quantity != null && a.system_stock != null) ? a.counted_quantity - a.system_stock : "",
            ASIGNADO_A:       a.assigned_to_name || "",
            CANTIDAD_CONTADA: a.counted_quantity ?? "",
            UBICACION:        a.location || "",
            NOTA:             a.note || "",
            CONTADO_POR:      a.counted_by_name || "",
            FECHA_CONTEO:     formatDateTime(a.counted_at),
            ESTADO:           a.status,
            ES_EXTRA:         a.is_extra ? "Sí" : "No",
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        ws["!cols"] = [
            { wch: 15 }, { wch: 40 }, { wch: 8 }, { wch: 10 }, { wch: 12 },
            { wch: 10 }, { wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 20 },
            { wch: 20 }, { wch: 18 }, { wch: 10 }, { wch: 8 },
        ];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Conteo");
        const store = allStores.find(s => s.id === selectedStoreId);
        XLSX.writeFile(wb, `ciclico_${store?.name || "tienda"}_${new Date().toISOString().split("T")[0]}.xlsx`);
    }

    // ── Admin: crear usuario ──────────────────────────────────────────────────
    async function createUser() {
        if (!newUsername || !newPassword || !newFullName) { showMessage("Completa todos los campos.", "error"); return; }
        // Inserta en cyclic_users (tabla propia, sin FK a inventories)
        const { error } = await supabase.from("cyclic_users").insert({
            username:  newUsername.trim(),
            password:  newPassword.trim(),
            full_name: newFullName.trim(),
            role:      newRole,
            is_active: true,
        });
        if (error) { showMessage("Error al crear usuario: " + error.message, "error"); return; }
        setNewUsername(""); setNewPassword(""); setNewFullName(""); setShowNewPassword(false);
        showMessage("✅ Usuario creado.", "success");
        await loadUsers();
    }

    function openEditUser(u: CyclicUserRow) {
        setEditingUser(u);
        setEditRole(u.role);
        setEditActive(u.is_active);
        setEditPassword("");
    }

    async function saveEditUser() {
        if (!editingUser) return;
        const payload: any = { role: editRole, is_active: editActive };
        if (editPassword.trim()) payload.password = editPassword.trim();
        const { error } = await supabase.from("cyclic_users").update(payload).eq("id", editingUser.id);
        if (error) { showMessage("Error al guardar: " + error.message, "error"); return; }
        setEditingUser(null);
        showMessage("✅ Usuario actualizado.", "success");
        await loadUsers();
    }

    async function deleteUser(u: CyclicUserRow) {
        if (u.id === user?.id) { showMessage("No puedes eliminar tu propio usuario.", "error"); return; }
        const ok = window.confirm(`¿Eliminar al usuario "${u.username}"?`);
        if (!ok) return;
        const { error } = await supabase.from("cyclic_users").delete().eq("id", u.id);
        if (error) { showMessage("Error: " + error.message, "error"); return; }
        showMessage("✅ Usuario eliminado.", "success");
        await loadUsers();
    }

    // ── Admin: crear tienda ───────────────────────────────────────────────────
    async function createStore() {
        if (!newStoreName.trim()) { showMessage("Escribe el nombre de la tienda.", "error"); return; }
        const code = newStoreCode.trim() || newStoreName.trim().toLowerCase().normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
        const { error } = await supabase.from("cyclic_stores").insert({ name: newStoreName.trim(), code, is_active: true });
        if (error) { showMessage("Error: " + error.message, "error"); return; }
        setNewStoreName(""); setNewStoreCode("");
        showMessage("✅ Tienda creada.", "success");
        await loadStores();
    }

    // ── Logout ────────────────────────────────────────────────────────────────
    function logout() {
        localStorage.removeItem(CYCLIC_SESSION_KEY);
        localStorage.removeItem(CYCLIC_STORE_KEY);
        window.location.href = "/ciclico/login";
    }

    // ── Scanner ───────────────────────────────────────────────────────────────
    useEffect(() => {
        if (!scannerTarget) return;
        let cancelled = false;
        async function startScanner() {
            try {
                const module = await import("html5-qrcode");
                const Html5Qrcode = module.Html5Qrcode;
                if (cancelled) return;
                const qr = new Html5Qrcode(scannerContainerId);
                scannerRef.current = qr;
                setScannerRunning(true);
                await qr.start(
                    { facingMode: "environment" },
                    { fps: 8, qrbox: { width: 220, height: 120 }, aspectRatio: 1.6 },
                    (decoded: string) => {
                        if (scanHandledRef.current) return;
                        scanHandledRef.current = true;
                        closeScanner();
                        findExtraByBarcode(decoded);
                    },
                    () => {}
                );
                try {
                    const cap: any = (qr as any).getRunningTrackCapabilities?.();
                    setTorchAvailable(!!cap?.torch);
                } catch { setTorchAvailable(false); }
            } catch (err: any) {
                showMessage("No se pudo iniciar la cámara: " + (err?.message || ""), "error");
                setScannerRunning(false); setScannerTarget(null);
            }
        }
        const t = setTimeout(startScanner, 150);
        return () => { cancelled = true; clearTimeout(t); stopScanner(); };
    }, [scannerTarget]);

    async function stopScanner() {
        try { if (scannerRef.current) { await scannerRef.current.stop(); await scannerRef.current.clear(); scannerRef.current = null; } }
        catch { scannerRef.current = null; }
        finally { setScannerRunning(false); }
    }
    function closeScanner() { scanHandledRef.current = false; setTorchOn(false); setTorchAvailable(false); stopScanner(); setScannerTarget(null); }
    async function toggleTorch() {
        try { const n = !torchOn; await (scannerRef.current as any).applyVideoConstraints?.({ advanced: [{ torch: n }] }); setTorchOn(n); }
        catch { showMessage("Linterna no disponible.", "error"); }
    }

    // ── Derived ───────────────────────────────────────────────────────────────
    const currentStore = allStores.find(s => s.id === selectedStoreId) || null;

    const filteredMyAssignments = useMemo(() => {
        const text = operarioSearch.toLowerCase();
        if (!text) return myAssignments;
        return myAssignments.filter(a =>
            (a.sku + " " + a.description + " " + (a.location || "")).toLowerCase().includes(text)
        );
    }, [myAssignments, operarioSearch]);

    const filteredAssignments = useMemo(() => {
        return assignments.filter(a => {
            const textOk   = validadorSearch ? (a.sku + " " + a.description + " " + (a.assigned_to_name || "")).toLowerCase().includes(validadorSearch.toLowerCase()) : true;
            const statusOk = validadorFilter === "todos" || a.status === validadorFilter;
            const userOk   = validadorUserFilter ? a.assigned_to_user_id === validadorUserFilter : true;
            return textOk && statusOk && userOk;
        });
    }, [assignments, validadorSearch, validadorFilter, validadorUserFilter]);

    const sessionStats = useMemo(() => ({
        total:    assignments.length,
        assigned: assignments.filter(a => a.assigned_to_user_id).length,
        contado:  assignments.filter(a => a.status === "contado" || a.status === "validado").length,
        validado: assignments.filter(a => a.status === "validado").length,
        pending:  assignments.filter(a => a.status === "pendiente").length,
    }), [assignments]);

    if (!user) return null;

    // ─────────────────────────────────────────────────────────────────────────
    // RENDER
    // ─────────────────────────────────────────────────────────────────────────
    return (
        <main className="min-h-screen bg-slate-100 p-4 md:p-6">
            <div className="max-w-7xl mx-auto space-y-6">

                {/* ── Header operario ─────────────────────────────────────── */}
                {user.role === "Operario" && (
                    <section className="bg-white rounded-2xl p-4 shadow border border-slate-200">
                        <div className="flex items-center justify-between gap-3">
                            <div>
                                <div className="text-xs text-slate-500">Conteo Cíclico — {currentStore?.name || "-"}</div>
                                <div className="text-lg font-bold text-slate-900">{user.full_name}</div>
                            </div>
                            <button className="px-4 py-2 rounded-xl bg-slate-900 text-white text-sm font-semibold" onClick={logout}>Salir</button>
                        </div>
                    </section>
                )}

                {/* ── Header validador / admin ─────────────────────────── */}
                {user.role !== "Operario" && (
                    <section className="rounded-3xl bg-gradient-to-r from-slate-950 via-slate-900 to-slate-700 text-white shadow-xl p-6">
                        <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                            <div>
                                <div className="inline-flex rounded-full bg-white/10 px-4 py-1 text-sm mb-3">🔄 Sistema de Conteos Cíclicos</div>
                                <h1 className="text-2xl md:text-3xl font-bold">Bienvenido, {user.full_name}</h1>
                            </div>
                            <div className="flex flex-wrap gap-3 items-end">
                                <div className="bg-white/10 rounded-2xl p-3 min-w-[220px]">
                                    <div className="text-xs text-slate-300 mb-1">Tienda activa</div>
                                    <select
                                        className="w-full rounded-xl border border-white/20 bg-white text-slate-900 px-3 py-2"
                                        value={selectedStoreId}
                                        onChange={e => handleStoreChange(e.target.value)}
                                    >
                                        {stores.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                                    </select>
                                </div>
                                <button className="px-5 py-2.5 rounded-2xl bg-white text-slate-900 font-semibold" onClick={logout}>Salir</button>
                            </div>
                        </div>
                    </section>
                )}

                {/* ── Mensaje global ───────────────────────────────────────── */}
                {message && (
                    <div className={`rounded-2xl p-4 shadow text-sm border flex items-start justify-between gap-3 ${
                        messageType === "success" ? "bg-green-50 border-green-200 text-green-800"
                        : messageType === "error" ? "bg-red-50 border-red-200 text-red-800"
                        : "bg-white border-slate-200 text-slate-800"
                    }`}>
                        <span>{message}</span>
                        <button className="text-xs opacity-60 hover:opacity-100 shrink-0" onClick={clearMessage}>✕</button>
                    </div>
                )}

                {/* ── Stats validador / admin ──────────────────────────────── */}
                {user.role !== "Operario" && session && (
                    <section className="grid grid-cols-2 md:grid-cols-4 gap-3">
                        {[
                            { label: "Total códigos", val: sessionStats.total },
                            { label: "Asignados",     val: sessionStats.assigned },
                            { label: "Contados",      val: sessionStats.contado },
                            { label: "Validados",     val: sessionStats.validado },
                        ].map(s => (
                            <div key={s.label} className="bg-white rounded-2xl shadow p-4 border border-slate-200">
                                <div className="text-xs text-slate-500">{s.label}</div>
                                <div className="text-2xl font-bold text-slate-900 mt-1">{s.val}</div>
                                {sessionStats.total > 0 && (
                                    <div className="w-full bg-slate-200 rounded-full h-1.5 mt-2">
                                        <div className="bg-slate-900 h-1.5 rounded-full" style={{ width: `${Math.round((s.val / sessionStats.total) * 100)}%` }} />
                                    </div>
                                )}
                            </div>
                        ))}
                    </section>
                )}

                {/* ── Tabs ─────────────────────────────────────────────────── */}
                {user.role !== "Operario" && (
                    <section className="bg-white rounded-2xl p-3 shadow">
                        <div className="flex flex-wrap gap-2">
                            {user.role === "Administrador" && (
                                <button className={`px-4 py-2 rounded-xl border font-semibold text-sm ${activeTab === "operario" ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-900 border-slate-300"}`} onClick={() => setActiveTab("operario")}>Operario</button>
                            )}
                            <button className={`px-4 py-2 rounded-xl border font-semibold text-sm ${activeTab === "validador" ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-900 border-slate-300"}`} onClick={() => setActiveTab("validador")}>Validador</button>
                            {user.role === "Administrador" && (
                                <button className={`px-4 py-2 rounded-xl border font-semibold text-sm ${activeTab === "admin" ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-900 border-slate-300"}`} onClick={() => setActiveTab("admin")}>Administrador</button>
                            )}
                        </div>
                    </section>
                )}

                {/* ════════════════════════════════════════════════════════════
                    TAB OPERARIO
                ════════════════════════════════════════════════════════════ */}
                {(activeTab === "operario" || user.role === "Operario") && (user.role === "Operario" || user.role === "Administrador") && (
                    <>
                        {loadingSession && <div className="bg-white rounded-2xl p-4 shadow text-sm text-slate-500">Cargando sesión...</div>}

                        {!loadingSession && !session && (
                            <section className="bg-white rounded-3xl p-6 shadow text-center space-y-3">
                                <div className="text-4xl">📋</div>
                                <h2 className="text-xl font-bold text-slate-900">Sin sesión activa hoy</h2>
                                <p className="text-slate-500 text-sm">El validador debe crear la sesión del día y asignarte códigos.</p>
                            </section>
                        )}

                        {!loadingSession && session && myAssignments.length === 0 && (
                            <section className="bg-white rounded-3xl p-6 shadow text-center space-y-3">
                                <div className="text-4xl">⏳</div>
                                <h2 className="text-xl font-bold text-slate-900">Sin códigos asignados</h2>
                                <p className="text-slate-500 text-sm">El validador aún no te ha asignado códigos para hoy.</p>
                            </section>
                        )}

                        {!loadingSession && session && myAssignments.length > 0 && (
                            <>
                                {/* Progreso */}
                                <section className="bg-white rounded-3xl p-5 shadow space-y-2">
                                    <div className="flex justify-between items-center">
                                        <div>
                                            <div className="font-bold text-slate-900">Mi progreso hoy</div>
                                            <div className="text-sm text-slate-500">{myAssignments.filter(a => a.status !== "pendiente").length} de {myAssignments.length} contados</div>
                                        </div>
                                        <div className="text-xl font-bold text-slate-900">
                                            {Math.round((myAssignments.filter(a => a.status !== "pendiente").length / myAssignments.length) * 100)}%
                                        </div>
                                    </div>
                                    <div className="w-full bg-slate-200 rounded-full h-3">
                                        <div className="bg-slate-900 h-3 rounded-full transition-all"
                                            style={{ width: `${(myAssignments.filter(a => a.status !== "pendiente").length / myAssignments.length) * 100}%` }} />
                                    </div>
                                </section>

                                {/* Buscador + agregar extra */}
                                <section className="bg-white rounded-2xl p-4 shadow">
                                    <div className="flex gap-3">
                                        <input
                                            className="flex-1 border rounded-2xl p-3 text-sm"
                                            placeholder="Buscar en mis códigos..."
                                            value={operarioSearch}
                                            onChange={e => setOperarioSearch(e.target.value)}
                                        />
                                        <button
                                            className="px-4 py-2 rounded-2xl bg-indigo-700 text-white font-semibold text-sm"
                                            onClick={() => { setShowAddExtra(true); setExtraSelected(null); setExtraSearch(""); setExtraResults([]); }}
                                        >
                                            + Agregar código
                                        </button>
                                    </div>
                                </section>

                                {/* Cards de asignaciones */}
                                {filteredMyAssignments.map(a => {
                                    const inp   = countInputs[a.id] || { qty: "", location: "", note: "" };
                                    const isDone = a.status !== "pendiente";
                                    return (
                                        <section key={a.id} className={`bg-white rounded-3xl p-5 shadow space-y-4 border-l-4 ${isDone ? "border-green-500" : a.is_extra ? "border-indigo-400" : "border-slate-300"}`}>
                                            <div className="flex items-start justify-between gap-3">
                                                <div>
                                                    <div className="flex gap-2 flex-wrap mb-1">
                                                        <span className={statusBadge(a.status)}>{a.status}</span>
                                                        {a.is_extra && <span className="inline-block px-2 py-0.5 rounded-full text-xs font-semibold bg-indigo-100 text-indigo-700">Extra</span>}
                                                    </div>
                                                    <div className="font-bold text-slate-900">{a.sku}</div>
                                                    <div className="text-sm text-slate-600">{a.description}</div>
                                                    <div className="text-xs text-slate-400">{a.unit}</div>
                                                </div>
                                                {isDone && <span className="text-green-600 font-bold text-xl">✓</span>}
                                            </div>

                                            {isDone ? (
                                                <div className="bg-green-50 rounded-2xl p-3 text-sm space-y-1 border border-green-200">
                                                    <div><span className="font-semibold">Cantidad:</span> {a.counted_quantity}</div>
                                                    <div><span className="font-semibold">Ubicación:</span> {a.location}</div>
                                                    {a.note && <div><span className="font-semibold">Nota:</span> {a.note}</div>}
                                                    <div className="text-xs text-slate-500">{formatDateTime(a.counted_at)}</div>
                                                </div>
                                            ) : (
                                                <div className="space-y-3">
                                                    <div className="grid grid-cols-2 gap-3">
                                                        <div>
                                                            <label className="block text-xs font-semibold text-slate-600 mb-1">Cantidad</label>
                                                            <input
                                                                type="number" inputMode="numeric"
                                                                className="w-full border rounded-xl p-2.5 text-sm font-bold"
                                                                placeholder="0"
                                                                value={inp.qty}
                                                                onChange={e => setCountInputs(prev => ({ ...prev, [a.id]: { ...inp, qty: e.target.value } }))}
                                                            />
                                                        </div>
                                                        <div>
                                                            <label className="block text-xs font-semibold text-slate-600 mb-1">Ubicación</label>
                                                            <input
                                                                className="w-full border rounded-xl p-2.5 text-sm"
                                                                placeholder="Ej: A-12"
                                                                value={inp.location}
                                                                onChange={e => setCountInputs(prev => ({ ...prev, [a.id]: { ...inp, location: e.target.value } }))}
                                                            />
                                                        </div>
                                                    </div>
                                                    <input
                                                        className="w-full border rounded-xl p-2.5 text-sm"
                                                        placeholder="Nota (opcional)"
                                                        value={inp.note}
                                                        onChange={e => setCountInputs(prev => ({ ...prev, [a.id]: { ...inp, note: e.target.value } }))}
                                                    />
                                                    <button className="w-full py-3 rounded-2xl bg-slate-900 text-white font-bold" onClick={() => saveCount(a.id)}>
                                                        Guardar conteo
                                                    </button>
                                                </div>
                                            )}
                                        </section>
                                    );
                                })}
                            </>
                        )}
                    </>
                )}

                {/* ════════════════════════════════════════════════════════════
                    TAB VALIDADOR
                ════════════════════════════════════════════════════════════ */}
                {activeTab === "validador" && (user.role === "Validador" || user.role === "Administrador") && (
                    <>
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                                <div>
                                    <h2 className="text-2xl font-bold text-slate-900">Sesión del día</h2>
                                    <p className="text-slate-500 text-sm">{new Date().toLocaleDateString("es-PE", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}</p>
                                </div>
                                <div className="flex gap-3">
                                    {!session ? (
                                        <button className="px-5 py-2.5 rounded-2xl bg-slate-900 text-white font-semibold" onClick={createSession}>Crear sesión de hoy</button>
                                    ) : (
                                        <>
                                            <button className="px-4 py-2 rounded-2xl border font-semibold text-sm" onClick={exportSession}>📥 Exportar</button>
                                            <button className="px-4 py-2 rounded-2xl bg-red-600 text-white font-semibold text-sm" onClick={closeSession}>Cerrar sesión</button>
                                        </>
                                    )}
                                </div>
                            </div>

                            {session && (
                                <>
                                    {/* Subir lista */}
                                    <div className="border rounded-2xl p-4 bg-slate-50 space-y-3">
                                        <div className="font-semibold text-slate-800 text-sm">📋 Cargar lista de SKUs del día</div>
                                        <p className="text-xs text-slate-500">
                                            Columnas requeridas: <b>SKU</b>. Opcionales: DESCRIPCION, COSTO, <b>STOCK</b> (stock sistema para comparar diferencias).
                                            Si no incluyes STOCK, se toma del maestro de productos.
                                        </p>
                                        <div className="flex flex-wrap gap-3 items-center">
                                            <input ref={assignInputRef} type="file" accept=".xlsx,.xls"
                                                onChange={e => { const f = e.target.files?.[0] || null; setAssignFile(f); setAssignFileName(f ? f.name : ""); }} />
                                            <div className="text-sm text-slate-500">{assignFileName ? `📄 ${assignFileName}` : "Ningún archivo"}</div>
                                            {uploadProgress && (
                                                <div className="w-full space-y-1">
                                                    <div className="flex justify-between text-xs text-slate-600"><span>{uploadProgress.step}</span><span>{uploadProgress.pct}%</span></div>
                                                    <div className="w-full bg-slate-200 rounded-full h-2"><div className="bg-slate-900 h-2 rounded-full transition-all" style={{ width: `${uploadProgress.pct}%` }} /></div>
                                                </div>
                                            )}
                                            <button className={`px-4 py-2 rounded-xl font-semibold text-sm text-white ${uploadProgress ? "bg-slate-400 cursor-not-allowed" : "bg-slate-900"}`}
                                                onClick={uploadAssignmentList} disabled={!!uploadProgress}>
                                                {uploadProgress ? "Cargando..." : "Cargar lista"}
                                            </button>
                                        </div>
                                    </div>

                                    {/* Asignar a operario */}
                                    {assignments.length > 0 && (
                                        <div className="border rounded-2xl p-4 bg-amber-50 space-y-3">
                                            <div className="font-semibold text-slate-800 text-sm">👤 Asignar códigos a operario</div>
                                            <div className="grid sm:grid-cols-3 gap-3">
                                                <div>
                                                    <label className="text-xs font-semibold text-slate-600 block mb-1">Operario</label>
                                                    <select className="w-full border rounded-xl p-2 text-sm" value={assignUserId} onChange={e => setAssignUserId(e.target.value)}>
                                                        <option value="">Selecciona operario</option>
                                                        {users.filter(u => u.is_active).map(u => (
                                                            <option key={u.id} value={u.id}>{u.full_name} ({u.role})</option>
                                                        ))}
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="text-xs font-semibold text-slate-600 block mb-1">Cantidad de códigos</label>
                                                    <input type="number" className="w-full border rounded-xl p-2 text-sm" value={assignCount} onChange={e => setAssignCount(e.target.value)} min={1} />
                                                </div>
                                                <div className="flex items-end">
                                                    <button className="w-full px-4 py-2 rounded-xl bg-amber-600 text-white font-semibold text-sm" onClick={assignCodes}>Asignar</button>
                                                </div>
                                            </div>
                                            <p className="text-xs text-slate-500">Sin asignar: <b>{assignments.filter(a => !a.assigned_to_user_id && !a.is_extra).length}</b> códigos disponibles</p>
                                        </div>
                                    )}
                                </>
                            )}
                        </section>

                        {/* Tabla de asignaciones */}
                        {session && assignments.length > 0 && (
                            <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                                <h3 className="text-xl font-bold text-slate-900">Asignaciones de la sesión</h3>
                                <div className="flex flex-wrap gap-3">
                                    <input className="border rounded-xl p-2 text-sm flex-1 min-w-[180px]"
                                        placeholder="Buscar SKU, descripción, operario..."
                                        value={validadorSearch} onChange={e => setValidadorSearch(e.target.value)} />
                                    <select className="border rounded-xl p-2 text-sm" value={validadorFilter} onChange={e => setValidadorFilter(e.target.value as any)}>
                                        <option value="todos">Todos los estados</option>
                                        <option value="pendiente">Pendiente</option>
                                        <option value="contado">Contado</option>
                                        <option value="validado">Validado</option>
                                    </select>
                                    <select className="border rounded-xl p-2 text-sm" value={validadorUserFilter} onChange={e => setValidadorUserFilter(e.target.value)}>
                                        <option value="">Todos los operarios</option>
                                        {users.map(u => <option key={u.id} value={u.id}>{u.full_name}</option>)}
                                    </select>
                                </div>
                                <div className="overflow-auto rounded-2xl border">
                                    <table className="w-full text-xs md:text-sm">
                                        <thead className="bg-slate-100 sticky top-0">
                                            <tr>
                                                <th className="p-2 border text-left">SKU</th>
                                                <th className="p-2 border text-left">Descripción</th>
                                                <th className="p-2 border">UM</th>
                                                <th className="p-2 border">Asignado a</th>
                                                <th className="p-2 border">Stock Sist.</th>
                                                <th className="p-2 border">Contado</th>
                                                <th className="p-2 border">Diferencia</th>
                                                <th className="p-2 border">Ubicación</th>
                                                <th className="p-2 border">Nota</th>
                                                <th className="p-2 border">Fecha conteo</th>
                                                <th className="p-2 border">Estado</th>
                                                <th className="p-2 border">Acción</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filteredAssignments.map(a => {
                                                const diff = (a.counted_quantity != null && a.system_stock != null)
                                                    ? a.counted_quantity - a.system_stock : null;
                                                const rowColor = a.status === "validado" ? "bg-green-50"
                                                    : a.status === "contado" ? (diff !== null && diff !== 0 ? "bg-red-50" : "bg-blue-50")
                                                    : "";
                                                return (
                                                    <tr key={a.id} className={rowColor}>
                                                        <td className="p-2 border font-medium">{a.sku}</td>
                                                        <td className="p-2 border max-w-[160px] truncate">{a.description}</td>
                                                        <td className="p-2 border text-center">{a.unit}</td>
                                                        <td className="p-2 border">{a.assigned_to_name || <span className="text-slate-400 italic text-xs">Sin asignar</span>}</td>
                                                        <td className="p-2 border text-center font-mono">{a.system_stock ?? <span className="text-slate-400">-</span>}</td>
                                                        <td className="p-2 border text-center font-bold">{a.counted_quantity ?? "-"}</td>
                                                        <td className="p-2 border text-center">{diffBadge(diff)}</td>
                                                        <td className="p-2 border">{a.location || "-"}</td>
                                                        <td className="p-2 border text-xs max-w-[100px] truncate">{a.note || "-"}</td>
                                                        <td className="p-2 border text-xs whitespace-nowrap">{formatDateTime(a.counted_at)}</td>
                                                        <td className="p-2 border"><span className={statusBadge(a.status)}>{a.status}</span></td>
                                                        <td className="p-2 border whitespace-nowrap">
                                                            <div className="flex gap-1">
                                                                {a.status === "contado" && (
                                                                    <button className="px-2 py-1 rounded-lg bg-green-600 text-white text-xs font-semibold" onClick={() => validateCount(a.id)}>Validar</button>
                                                                )}
                                                                {a.status === "validado" && (
                                                                    <button className="px-2 py-1 rounded-lg border text-xs" onClick={() => unvalidateCount(a.id)}>Desvalidar</button>
                                                                )}
                                                                {a.assigned_to_user_id && a.status === "pendiente" && (
                                                                    <button className="px-2 py-1 rounded-lg border text-xs text-red-600" onClick={() => revokeAssignment(a.id)}>Revocar</button>
                                                                )}
                                                            </div>
                                                        </td>
                                                    </tr>
                                                );
                                            })}
                                            {filteredAssignments.length === 0 && (
                                                <tr><td colSpan={12} className="p-4 text-center text-slate-400">Sin resultados.</td></tr>
                                            )}
                                        </tbody>
                                    </table>
                                </div>
                            </section>
                        )}

                        {!session && !loadingSession && (
                            <section className="bg-white rounded-3xl p-6 shadow text-center space-y-3">
                                <div className="text-4xl">📅</div>
                                <p className="text-slate-500">No hay sesión activa hoy para esta tienda. Crea una para comenzar.</p>
                            </section>
                        )}
                    </>
                )}

                {/* ════════════════════════════════════════════════════════════
                    TAB ADMIN
                ════════════════════════════════════════════════════════════ */}
                {activeTab === "admin" && user.role === "Administrador" && (
                    <>
                        {/* Stats */}
                        <section className="grid md:grid-cols-2 xl:grid-cols-4 gap-4">
                            {[
                                { label: "Tiendas activas",  val: stores.length },
                                { label: "Usuarios",         val: users.length },
                                { label: "Sesión hoy",       val: session ? "Activa" : "Sin sesión" },
                                { label: "Códigos hoy",      val: assignments.length },
                            ].map(s => (
                                <div key={s.label} className="bg-white rounded-2xl shadow p-5 border border-slate-200">
                                    <div className="text-sm text-slate-500">{s.label}</div>
                                    <div className="text-2xl font-bold text-slate-900 mt-1">{s.val}</div>
                                </div>
                            ))}
                        </section>

                        <section className="grid lg:grid-cols-2 gap-6">
                            {/* Crear tienda */}
                            <div className="bg-white rounded-3xl p-6 shadow space-y-4">
                                <h3 className="text-xl font-bold text-slate-900">Crear tienda</h3>
                                <input className="w-full border rounded-2xl p-3" placeholder="Nombre de la tienda" value={newStoreName} onChange={e => setNewStoreName(e.target.value)} />
                                <input className="w-full border rounded-2xl p-3" placeholder="Código corto (opcional)" value={newStoreCode} onChange={e => setNewStoreCode(e.target.value)} />
                                <button className="px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold w-full" onClick={createStore}>Crear tienda</button>
                            </div>

                            {/* Crear usuario */}
                            <div className="bg-white rounded-3xl p-6 shadow space-y-4">
                                <h3 className="text-xl font-bold text-slate-900">Crear usuario</h3>
                                <input className="w-full border rounded-2xl p-3" placeholder="ID de usuario" value={newUsername} onChange={e => setNewUsername(e.target.value)} />
                                <input className="w-full border rounded-2xl p-3" placeholder="Nombre completo" value={newFullName} onChange={e => setNewFullName(e.target.value)} />
                                <div className="flex gap-2">
                                    <input className="flex-1 border rounded-2xl p-3" placeholder="Contraseña" type={showNewPassword ? "text" : "password"} value={newPassword} onChange={e => setNewPassword(e.target.value)} />
                                    <button type="button" className="px-4 rounded-2xl border text-sm" onClick={() => setShowNewPassword(!showNewPassword)}>{showNewPassword ? "Ocultar" : "Ver"}</button>
                                </div>
                                <select className="w-full border rounded-2xl p-3" value={newRole} onChange={e => setNewRole(e.target.value as Role)}>
                                    <option value="Operario">Operario</option>
                                    <option value="Validador">Validador</option>
                                    <option value="Administrador">Administrador</option>
                                </select>
                                <button className="px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold w-full" onClick={createUser}>Crear usuario</button>
                            </div>
                        </section>

                        {/* Tiendas */}
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <h3 className="text-xl font-bold text-slate-900">Tiendas</h3>
                            <div className="overflow-auto rounded-2xl border">
                                <table className="w-full text-sm">
                                    <thead className="bg-slate-100">
                                        <tr>
                                            <th className="p-3 border text-left">Nombre</th>
                                            <th className="p-3 border text-left">Código</th>
                                            <th className="p-3 border text-center">Estado</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {allStores.map(s => (
                                            <tr key={s.id}>
                                                <td className="p-3 border">{s.name}</td>
                                                <td className="p-3 border">{s.code}</td>
                                                <td className="p-3 border text-center">
                                                    <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${s.is_active ? "bg-green-100 text-green-700" : "bg-slate-100 text-slate-500"}`}>
                                                        {s.is_active ? "Activa" : "Inactiva"}
                                                    </span>
                                                </td>
                                            </tr>
                                        ))}
                                        {allStores.length === 0 && <tr><td colSpan={3} className="p-4 text-center text-slate-400">No hay tiendas.</td></tr>}
                                    </tbody>
                                </table>
                            </div>
                        </section>

                        {/* Usuarios */}
                        <section className="bg-white rounded-3xl p-6 shadow">
                            <div className="mb-4"><h3 className="text-xl font-bold text-slate-900">Usuarios</h3></div>
                            <div className="overflow-auto rounded-2xl border">
                                <table className="w-full text-sm">
                                    <thead className="bg-slate-100">
                                        <tr>
                                            <th className="p-3 border text-left">Usuario</th>
                                            <th className="p-3 border text-left">Nombre</th>
                                            <th className="p-3 border text-center">Rol</th>
                                            <th className="p-3 border text-center">Estado</th>
                                            <th className="p-3 border text-center">Acciones</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {users.map(u => (
                                            <tr key={u.id} className="hover:bg-slate-50">
                                                <td className="p-3 border font-medium">{u.username}</td>
                                                <td className="p-3 border">{u.full_name}</td>
                                                <td className="p-3 border text-center">
                                                    <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${
                                                        u.role === "Administrador" ? "bg-purple-100 text-purple-700"
                                                        : u.role === "Validador"    ? "bg-blue-100 text-blue-700"
                                                        : "bg-slate-100 text-slate-700"
                                                    }`}>{u.role}</span>
                                                </td>
                                                <td className="p-3 border text-center">
                                                    <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${u.is_active ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                                                        {u.is_active ? "Activo" : "Inactivo"}
                                                    </span>
                                                </td>
                                                <td className="p-3 border text-center">
                                                    <div className="flex gap-2 justify-center">
                                                        <button className="px-3 py-1 rounded-lg bg-slate-900 text-white text-xs font-semibold" onClick={() => openEditUser(u)}>✏️ Editar</button>
                                                        <button className={`px-3 py-1 rounded-lg text-white text-xs font-semibold ${u.id === user.id ? "bg-slate-300 cursor-not-allowed" : "bg-red-600"}`}
                                                            onClick={() => deleteUser(u)} disabled={u.id === user.id}>Eliminar</button>
                                                    </div>
                                                </td>
                                            </tr>
                                        ))}
                                        {users.length === 0 && <tr><td colSpan={5} className="p-4 text-center text-slate-400">No hay usuarios.</td></tr>}
                                    </tbody>
                                </table>
                            </div>
                        </section>
                    </>
                )}

                {/* ════════════════════════════════════════════════════════════
                    MODAL EDITAR USUARIO
                ════════════════════════════════════════════════════════════ */}
                {editingUser && (
                    <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
                        <div className="bg-white rounded-3xl p-6 w-full max-w-md shadow-2xl space-y-5 max-h-[90vh] overflow-y-auto">
                            <div>
                                <h3 className="text-2xl font-bold text-slate-900">Editar usuario</h3>
                                <p className="text-slate-500 text-sm mt-1">{editingUser.username} — {editingUser.full_name}</p>
                            </div>

                            <div>
                                <label className="block font-semibold mb-2 text-sm">Rol</label>
                                <select className="w-full border rounded-2xl p-3" value={editRole} onChange={e => setEditRole(e.target.value as Role)}>
                                    <option value="Operario">Operario</option>
                                    <option value="Validador">Validador</option>
                                    <option value="Administrador">Administrador</option>
                                </select>
                            </div>

                            <div>
                                <label className="block font-semibold mb-2 text-sm">Nueva contraseña <span className="font-normal text-slate-400">(dejar vacío para no cambiar)</span></label>
                                <div className="flex gap-2">
                                    <input className="flex-1 border rounded-2xl p-3" type={showEditPassword ? "text" : "password"}
                                        placeholder="Nueva contraseña..." value={editPassword} onChange={e => setEditPassword(e.target.value)} />
                                    <button type="button" className="px-4 rounded-2xl border text-sm" onClick={() => setShowEditPassword(!showEditPassword)}>{showEditPassword ? "Ocultar" : "Ver"}</button>
                                </div>
                            </div>

                            <div>
                                <label className="block font-semibold mb-2 text-sm">Estado</label>
                                <div className="flex gap-3">
                                    <button className={`flex-1 py-2.5 rounded-xl font-semibold text-sm border ${editActive ? "bg-green-600 text-white border-green-600" : "bg-white text-slate-700 border-slate-300"}`} onClick={() => setEditActive(true)}>✓ Activo</button>
                                    <button className={`flex-1 py-2.5 rounded-xl font-semibold text-sm border ${!editActive ? "bg-red-500 text-white border-red-500" : "bg-white text-slate-700 border-slate-300"}`} onClick={() => setEditActive(false)}>Inactivo</button>
                                </div>
                            </div>

                            <div className="flex gap-3 pt-1">
                                <button className="flex-1 px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={saveEditUser}>Guardar cambios</button>
                                <button className="flex-1 px-4 py-3 rounded-2xl border font-semibold" onClick={() => setEditingUser(null)}>Cancelar</button>
                            </div>
                        </div>
                    </div>
                )}

                {/* ════════════════════════════════════════════════════════════
                    MODAL AGREGAR CÓDIGO EXTRA
                ════════════════════════════════════════════════════════════ */}
                {showAddExtra && (
                    <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
                        <div className="bg-white rounded-3xl p-6 w-full max-w-lg shadow-2xl space-y-4">
                            <h3 className="text-xl font-bold text-slate-900">Agregar código extra</h3>
                            <p className="text-slate-500 text-sm">Busca por SKU o descripción en el maestro de productos, o escanea el código de barra.</p>
                            <div className="relative">
                                <input
                                    className="w-full border rounded-2xl p-3 pr-12"
                                    placeholder="Buscar SKU o descripción..."
                                    value={extraSearch}
                                    onChange={e => searchExtraProduct(e.target.value)}
                                />
                                <button type="button"
                                    className="absolute right-2 top-1/2 -translate-y-1/2 h-9 w-9 rounded-xl bg-slate-900 text-white flex items-center justify-center"
                                    onClick={() => { scanHandledRef.current = false; setScannerTarget("extra"); }}>
                                    <QrCode size={16} />
                                </button>
                            </div>
                            {extraResults.length > 0 && (
                                <div className="max-h-48 overflow-auto rounded-xl border">
                                    {extraResults.map(p => (
                                        <button key={p.id} type="button" className="w-full text-left p-3 border-b last:border-b-0 hover:bg-slate-50"
                                            onClick={() => { setExtraSelected(p); setExtraResults([]); setExtraSearch(p.description); }}>
                                            <div className="font-semibold text-slate-900 text-sm">{p.sku}</div>
                                            <div className="text-xs text-slate-600">{p.description}</div>
                                        </button>
                                    ))}
                                </div>
                            )}
                            {extraSelected && (
                                <div className="bg-green-50 border border-green-200 rounded-2xl p-3 text-sm">
                                    <div className="font-bold text-slate-900">{extraSelected.sku}</div>
                                    <div className="text-slate-600">{extraSelected.description}</div>
                                    <div className="text-xs text-slate-400">{extraSelected.unit}</div>
                                </div>
                            )}
                            <div className="flex gap-3">
                                <button className="flex-1 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={addExtraAssignment} disabled={!extraSelected}>Agregar</button>
                                <button className="flex-1 py-3 rounded-2xl border font-semibold"
                                    onClick={() => { setShowAddExtra(false); setExtraSelected(null); setExtraSearch(""); setExtraResults([]); }}>
                                    Cancelar
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                {/* ════════════════════════════════════════════════════════════
                    SCANNER
                ════════════════════════════════════════════════════════════ */}
                {scannerTarget && (
                    <div className="fixed inset-0 bg-black/70 flex items-center justify-center p-4 z-[60]">
                        <div className="bg-white w-full max-w-lg rounded-3xl p-5 shadow-2xl space-y-4">
                            <h3 className="text-xl font-bold text-slate-900">Escanear código de barra</h3>
                            <p className="text-sm text-slate-500">Apunta al código del producto.</p>
                            <div className="rounded-2xl overflow-hidden border bg-black min-h-[260px] flex items-center justify-center">
                                <div id={scannerContainerId} className="w-full" />
                            </div>
                            <div className="text-sm text-slate-500">{scannerRunning ? "Cámara activa. Apunta al código." : "Iniciando cámara..."}</div>
                            {torchAvailable && (
                                <button type="button" onClick={toggleTorch} className="w-full px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold">
                                    {torchOn ? "Apagar linterna 🔦" : "Prender linterna 🔦"}
                                </button>
                            )}
                            <button type="button" onClick={closeScanner} className="w-full px-4 py-2 rounded-2xl border text-sm font-semibold">Cancelar</button>
                        </div>
                    </div>
                )}

            </div>
        </main>
    );
}
