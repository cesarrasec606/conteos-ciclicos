"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { supabase } from "@/lib/supabase/client";
import * as XLSX from "xlsx";
import { QrCode } from "lucide-react";

type Role = "Operario" | "Validador" | "Administrador";
type TabKey = "operario" | "validador" | "maestro" | "admin";

type UserPermissions = {
    can_see_system_stock: boolean;
    can_see_cost: boolean;
    can_see_valued_difference: boolean;
};

type User = {
    id: string;
    username: string;
    full_name: string;
    role: Role;
    roles?: Role[];
    inventory_id: string | null;
    permissions?: UserPermissions;
    can_access_any_inventory?: boolean;
};

type Inventory = {
    id: string;
    name: string;
    code: string;
    is_active: boolean;
};

type Product = {
    id: string;
    inventory_id: string | null;
    sku: string;
    description: string;
    unit: string;
    cost: number;
    system_stock: number;
};

type RecordRow = {
    id: string;
    inventory_id: string | null;
    product_id: string | null;
    sku: string;
    barcode: string | null;
    description: string;
    unit: string;
    counted_quantity: number;
    system_stock: number;
    difference: number;
    location: string;
    user_id: string | null;
    user_name: string;
    validator_name: string | null;
    status: "Pendiente" | "Diferencia" | "Validado" | "Corregido";
    note: string | null;
    cost: number;
    counted_at: string;
};

type CountingSession = {
    id: string;
    inventory_id: string;
    name: string;
    description: string | null;
    created_by: string;
    created_by_name: string;
    created_at: string;
    closed_at: string | null;
    status: "open" | "closed";
};

type SessionSummary = {
    session: CountingSession;
    inventory_name: string;
    total_records: number;
    total_skus: number;
    ok_count: number;
    faltantes_count: number;
    sobrantes_count: number;
    no_contado_count: number;
    valued_difference: number;
    total_counted_value: number;
    avance_pct: number;
};

type AppUser = {
    id: string;
    username: string;
    password: string;
    full_name: string;
    role: Role;
    roles?: Role[];
    is_active: boolean;
    inventory_id: string | null;
    can_see_system_stock?: boolean;
    can_see_cost?: boolean;
    can_see_valued_difference?: boolean;
    can_access_any_inventory?: boolean;
};

type SidebarItem = { key: TabKey; label: string; icon: string; show: boolean };

// ── OFFLINE QUEUE TYPES ─────────────────────────────────────────────────────
type OfflineRecord = {
    local_id: string;       // UUID generado en cliente — clave de deduplicación
    synced: boolean;
    created_at_local: string;
    inventory_id: string;
    product_id: string;
    sku: string;
    barcode: string | null;
    description: string;
    unit: string;
    counted_quantity: number;
    system_stock: number;
    difference: number;
    location: string;
    user_id: string;
    user_name: string;
    status: "Pendiente" | "Diferencia";
    note: string;
    cost: number;
    counted_at: string;
};

// ── INDEXEDDB HELPERS ───────────────────────────────────────────────────────
const DB_NAME = "wms_offline";
const DB_VERSION = 1;
const STORE_NAME = "pending_records";

function openOfflineDB(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
            const db = (e.target as IDBOpenDBRequest).result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const store = db.createObjectStore(STORE_NAME, { keyPath: "local_id" });
                store.createIndex("synced", "synced", { unique: false });
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function saveOfflineRecord(record: OfflineRecord): Promise<void> {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        tx.objectStore(STORE_NAME).put(record);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

async function getPendingOfflineRecords(): Promise<OfflineRecord[]> {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readonly");
        const req = tx.objectStore(STORE_NAME).index("synced").getAll(IDBKeyRange.only(0));
        req.onsuccess = () => resolve(req.result as OfflineRecord[]);
        req.onerror = () => reject(req.error);
    });
}

async function markOfflineRecordSynced(local_id: string): Promise<void> {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        const store = tx.objectStore(STORE_NAME);
        const getReq = store.get(local_id);
        getReq.onsuccess = () => {
            const rec = getReq.result;
            if (rec) { rec.synced = true; store.put(rec); }
            resolve();
        };
        getReq.onerror = () => reject(getReq.error);
    });
}

async function getAllOfflineRecords(): Promise<OfflineRecord[]> {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readonly");
        const req = tx.objectStore(STORE_NAME).getAll();
        req.onsuccess = () => resolve(req.result as OfflineRecord[]);
        req.onerror = () => reject(req.error);
    });
}

function generateLocalId(): string {
    return `local_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

// ── CONSTANTS ───────────────────────────────────────────────────────────────
const ALL_ROLES: Role[] = ["Operario", "Validador", "Administrador"];

const DEFAULT_PERMISSIONS: UserPermissions = {
    can_see_system_stock: false,
    can_see_cost: false,
    can_see_valued_difference: false,
};

const CURRENT_INVENTORY_KEY = "current_inventory_id";

function getPrimaryRole(u: AppUser): Role {
    if (u.roles && u.roles.length > 0) {
        if (u.roles.includes("Administrador")) return "Administrador";
        if (u.roles.includes("Validador")) return "Validador";
        return u.roles[0];
    }
    return u.role;
}

function getUserPermissions(u: AppUser | User): UserPermissions {
    return {
        can_see_system_stock: (u as any).can_see_system_stock ?? false,
        can_see_cost: (u as any).can_see_cost ?? false,
        can_see_valued_difference: (u as any).can_see_valued_difference ?? false,
    };
}

function hasRole(u: User | null, role: Role): boolean {
    if (!u) return false;
    if (u.roles && u.roles.length > 0) return u.roles.includes(role);
    return u.role === role;
}

function canCount(u: User | null): boolean {
    if (!u) return false;
    if (hasRole(u, "Administrador")) return true;
    if (hasRole(u, "Operario")) return true;
    return false;
}

function canValidate(u: User | null): boolean {
    if (!u) return false;
    return hasRole(u, "Administrador") || hasRole(u, "Validador");
}

function canAccessAnyInventory(u: User | null): boolean {
    if (!u) return false;
    if (hasRole(u, "Administrador")) return true;
    if ((u as any).can_access_any_inventory === true) return true;
    return false;
}

function canSeeSystemStock(u: User | null): boolean {
    if (!u) return false;
    if (hasRole(u, "Administrador") || hasRole(u, "Validador")) return true;
    return u.permissions?.can_see_system_stock ?? false;
}

function canSeeCost(u: User | null): boolean {
    if (!u) return false;
    if (hasRole(u, "Administrador") || hasRole(u, "Validador")) return true;
    return u.permissions?.can_see_cost ?? false;
}

function canSeeValuedDifference(u: User | null): boolean {
    if (!u) return false;
    if (hasRole(u, "Administrador") || hasRole(u, "Validador")) return true;
    return u.permissions?.can_see_valued_difference ?? false;
}

function statusBadge(status: RecordRow["status"]) {
    const base = "inline-block px-2 py-0.5 rounded-full text-xs font-semibold";
    switch (status) {
        case "Pendiente":  return `${base} bg-slate-100 text-slate-700`;
        case "Diferencia": return `${base} bg-red-100 text-red-700`;
        case "Validado":   return `${base} bg-green-100 text-green-700`;
        case "Corregido":  return `${base} bg-blue-100 text-blue-700`;
        default:           return `${base} bg-slate-100 text-slate-700`;
    }
}

function diffBadge(diff: number) {
    if (diff === 0) return <span className="text-green-700 font-semibold">0</span>;
    if (diff > 0)   return <span className="text-blue-700 font-semibold">+{diff}</span>;
    return <span className="text-red-600 font-semibold">{diff}</span>;
}

function cleanCode(value: string | null | undefined): string {
    if (!value) return "";
    let s = String(value).trim();
    s = s.replace(/^['"''""\u2018\u2019\u201C\u201D]+/, "").replace(/['"''""\u2018\u2019\u201C\u201D]+$/, "").trim();
    if (/[Ee][+-]/.test(s) && !isNaN(Number(s))) {
        const n = Number(s);
        if (isFinite(n)) s = Math.round(n).toString();
    }
    s = s.replace(/\.0+$/, "");
    if (/^\d+$/.test(s)) {
        s = s.replace(/^0+/, "");
        if (s === "") s = "0";
    }
    return s;
}

function normalizeText(value: string | null | undefined) {
    return String(value || "").trim().toLowerCase();
}

function formatMoney(value: number) {
    return `S/ ${Number(value || 0).toLocaleString("es-PE", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    })}`;
}

function formatDateTime(value: string) {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleString("es-PE");
}

async function fetchAllProducts(inventoryId: string): Promise<Product[]> {
    const PAGE = 1000;
    const all: Product[] = [];
    let page = 0;
    let hasMore = true;
    while (hasMore) {
        const { data, error } = await supabase
            .from("products")
            .select("*")
            .eq("inventory_id", inventoryId)
            .order("sku")
            .range(page * PAGE, (page + 1) * PAGE - 1);
        if (error) break;
        if (data && data.length > 0) {
            all.push(...(data as Product[]));
            page++;
        }
        if (!data || data.length < PAGE) hasMore = false;
    }
    return all;
}

function RoleCheckboxes({
    selected,
    onChange,
}: {
    selected: Role[];
    onChange: (roles: Role[]) => void;
}) {
    function toggle(role: Role) {
        if (selected.includes(role)) {
            const next = selected.filter((r) => r !== role);
            if (next.length === 0) return;
            onChange(next);
        } else {
            onChange([...selected, role]);
        }
    }
    return (
        <div className="flex flex-wrap gap-3">
            {ALL_ROLES.map((role) => {
                const active = selected.includes(role);
                const color =
                    role === "Administrador"
                        ? active ? "bg-purple-600 text-white border-purple-600" : "bg-white text-slate-700 border-slate-300"
                        : role === "Validador"
                        ? active ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-700 border-slate-300"
                        : active ? "bg-slate-800 text-white border-slate-800" : "bg-white text-slate-700 border-slate-300";
                return (
                    <button
                        key={role}
                        type="button"
                        className={`px-4 py-2 rounded-xl border font-semibold text-sm transition ${color}`}
                        onClick={() => toggle(role)}
                    >
                        {active ? "✓ " : ""}{role}
                    </button>
                );
            })}
        </div>
    );
}

function PermissionsCheckboxes({
    perms,
    onChange,
    isAdmin,
}: {
    perms: UserPermissions;
    onChange: (perms: UserPermissions) => void;
    isAdmin: boolean;
}) {
    if (isAdmin) {
        return (
            <div className="text-xs text-slate-500 italic bg-slate-50 rounded-xl p-3">
                Los administradores tienen acceso total a todos los datos. No se requiere configurar permisos adicionales.
            </div>
        );
    }
    return (
        <div className="space-y-2">
            <label className="flex items-center gap-3 cursor-pointer">
                <input
                    type="checkbox"
                    className="w-4 h-4 accent-slate-800"
                    checked={perms.can_see_system_stock}
                    onChange={(e) => onChange({ ...perms, can_see_system_stock: e.target.checked })}
                />
                <span className="text-sm font-medium text-slate-700">Ver stock sistémico</span>
            </label>
            <label className="flex items-center gap-3 cursor-pointer">
                <input
                    type="checkbox"
                    className="w-4 h-4 accent-slate-800"
                    checked={perms.can_see_cost}
                    onChange={(e) => onChange({ ...perms, can_see_cost: e.target.checked })}
                />
                <span className="text-sm font-medium text-slate-700">Ver costo unitario</span>
            </label>
            <label className="flex items-center gap-3 cursor-pointer">
                <input
                    type="checkbox"
                    className="w-4 h-4 accent-slate-800"
                    checked={perms.can_see_valued_difference}
                    onChange={(e) => onChange({ ...perms, can_see_valued_difference: e.target.checked })}
                />
                <span className="text-sm font-medium text-slate-700">Ver diferencia valorizada (diferencia × costo)</span>
            </label>
        </div>
    );
}

export default function DashboardPage() {
    const [user, setUser] = useState<User | null>(null);
    const [inventories, setInventories] = useState<Inventory[]>([]);
    const [allInventories, setAllInventories] = useState<Inventory[]>([]);
    const [selectedInventoryId, setSelectedInventoryId] = useState("");
    const [products, setProducts] = useState<Product[]>([]);
    const [records, setRecords] = useState<RecordRow[]>([]);
    const [users, setUsers] = useState<AppUser[]>([]);
    const [message, setMessage] = useState("");
    const [messageType, setMessageType] = useState<"info" | "success" | "error">("info");
    const [loading, setLoading] = useState(true);
    const [activeTab, setActiveTab] = useState<TabKey>("operario");
    const [processingInventoryId, setProcessingInventoryId] = useState<string | null>(null);
    const [totalProductCount, setTotalProductCount] = useState(0);
    const [totalProductWithStockCount, setTotalProductWithStockCount] = useState(0);
    const [totalInventoryValue, setTotalInventoryValue] = useState(0);
    const [uploadProgress, setUploadProgress] = useState<{ step: string; pct: number } | null>(null);

    // ── OFFLINE STATE ────────────────────────────────────────────────────────
    const [isOnline, setIsOnline] = useState(typeof navigator !== "undefined" ? navigator.onLine : true);
    const [pendingCount, setPendingCount] = useState(0);
    const [syncing, setSyncing] = useState(false);
    const syncLockRef = useRef(false);

    const [searchValue, setSearchValue] = useState("");
    const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);
    const [searchResults, setSearchResults] = useState<Product[]>([]);
    const lastScanRef = useRef<string | null>(null);
    const [location, setLocation] = useState("");
    const [quantity, setQuantity] = useState("");
    const [searchText, setSearchText] = useState("");
    const [operarioHistorySearch, setOperarioHistorySearch] = useState("");
    const [statusFilter, setStatusFilter] = useState("todos");
    const [auditSearchText, setAuditSearchText] = useState("");
    const [auditStatusFilter, setAuditStatusFilter] = useState("todos");
    const [validadorSubTab, setValidadorSubTab] = useState<"registros" | "resumen">("registros");
    const [recordsPage, setRecordsPage] = useState(1);
    const RECORDS_PER_PAGE = 100;

    // ── OPERARIO MOBILE SUBTAB ───────────────────────────────────────────────
    const [operarioSubTab, setOperarioSubTab] = useState<"conteo" | "historial">("conteo");

    type AuditRow = {
        sku: string;
        description: string;
        unit: string;
        cost: number;
        system_stock: number;
        total_counted: number;
        difference: number;
        valued_difference: number;
        record_count: number;
        not_counted: boolean;
        status_resumen: "OK" | "FALTANTE" | "SOBRANTE" | "NO CONTADO";
    };
    const [auditByCode, setAuditByCode] = useState<AuditRow[]>([]);
    const [auditLoading, setAuditLoading] = useState(false);

    // ── INFORME MODAL ────────────────────────────────────────────────────────
    const [showReportModal, setShowReportModal] = useState(false);
    const [reportStoreName, setReportStoreName] = useState("");
    const [reportStoreLeader, setReportStoreLeader] = useState("");
    const [reportWarehouseAdvisor, setReportWarehouseAdvisor] = useState("");
    const [reportAuditorName, setReportAuditorName] = useState("");

    // ── SESIONES ────────────────────────────────────────────────────────────
    const [sessions, setSessions] = useState<CountingSession[]>([]);
    const [allSessionsSummary, setAllSessionsSummary] = useState<SessionSummary[]>([]);
    const [sessionsLoading, setSessionsLoading] = useState(false);
    const [allSessionsLoading, setAllSessionsLoading] = useState(false);
    const [newSessionName, setNewSessionName] = useState("");
    const [newSessionDescription, setNewSessionDescription] = useState("");
    const [activeSessionId, setActiveSessionId] = useState<string | null>(null);
    const [validadorSubTabMain, setValidadorSubTabMain] = useState<"sesiones" | "registros_y_resumen">("sesiones");

    const [skuProgress, setSkuProgress] = useState<{ total: number; counted: number; pct: number }>({ total: 0, counted: 0, pct: 0 });

    const [newUsername, setNewUsername] = useState("");
    const [newPassword, setNewPassword] = useState("");
    const [showNewPassword, setShowNewPassword] = useState(false);
    const [newFullName, setNewFullName] = useState("");
    const [newRoles, setNewRoles] = useState<Role[]>(["Operario"]);
    const [newPermissions, setNewPermissions] = useState<UserPermissions>({ ...DEFAULT_PERMISSIONS });
    const [newCanAccessAnyInventory, setNewCanAccessAnyInventory] = useState(false);

    const [editingUser, setEditingUser] = useState<AppUser | null>(null);
    const [editUserRoles, setEditUserRoles] = useState<Role[]>(["Operario"]);
    const [editUserInventoryId, setEditUserInventoryId] = useState("");
    const [editUserActive, setEditUserActive] = useState(true);
    const [editUserPermissions, setEditUserPermissions] = useState<UserPermissions>({ ...DEFAULT_PERMISSIONS });
    const [editUserCanAccessAnyInventory, setEditUserCanAccessAnyInventory] = useState(false);

    const [newInventoryName, setNewInventoryName] = useState("");
    const [newInventoryCode, setNewInventoryCode] = useState("");

    const [masterFile, setMasterFile] = useState<File | null>(null);
    const [usersFile, setUsersFile] = useState<File | null>(null);
    const [globalBarcodesFile, setGlobalBarcodesFile] = useState<File | null>(null);
    const [masterFileName, setMasterFileName] = useState("");
    const [usersFileName, setUsersFileName] = useState("");
    const [globalBarcodesFileName, setGlobalBarcodesFileName] = useState("");
    const [globalBarcodesCount, setGlobalBarcodesCount] = useState<number | null>(null);
    const globalBarcodesInputRef = useRef<HTMLInputElement | null>(null);

    const [editingRecord, setEditingRecord] = useState<RecordRow | null>(null);
    const [editSku, setEditSku] = useState("");
    const [editQty, setEditQty] = useState("");
    const [editLocation, setEditLocation] = useState("");
    const [editStatus, setEditStatus] = useState<RecordRow["status"]>("Pendiente");
    const [editNote, setEditNote] = useState("");
    const [editMatchedProduct, setEditMatchedProduct] = useState<Product | null>(null);

    const [scannerTarget, setScannerTarget] = useState<"product" | "location" | null>(null);
    const [scannerRunning, setScannerRunning] = useState(false);
    const [torchAvailable, setTorchAvailable] = useState(false);
    const [torchOn, setTorchOn] = useState(false);
    const masterInputRef = useRef<HTMLInputElement | null>(null);
    const usersInputRef = useRef<HTMLInputElement | null>(null);
    const scannerRef = useRef<any>(null);
    const scanHandledRef = useRef(false);
    const html5QrCodeModuleRef = useRef<any>(null);
    const overlayOpenedRef = useRef(false);
    const messageTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
    const scannerContainerId = "scanner-reader";

    function showMessage(msg: string, type: "info" | "success" | "error" = "info") {
        if (messageTimerRef.current) clearTimeout(messageTimerRef.current);
        setMessage(msg);
        setMessageType(type);
        if (type === "success") {
            messageTimerRef.current = setTimeout(() => setMessage(""), 4000);
        } else if (type === "error") {
            messageTimerRef.current = setTimeout(() => setMessage(""), 7000);
        }
    }

    // ── OFFLINE: detectar conexión y sincronizar ─────────────────────────────
    useEffect(() => {
        function handleOnline() {
            setIsOnline(true);
            syncPendingRecords();
        }
        function handleOffline() { setIsOnline(false); }
        window.addEventListener("online", handleOnline);
        window.addEventListener("offline", handleOffline);
        return () => {
            window.removeEventListener("online", handleOnline);
            window.removeEventListener("offline", handleOffline);
        };
    }, []);

    // Refrescar contador de pendientes periódicamente
    useEffect(() => {
        refreshPendingCount();
        const interval = setInterval(refreshPendingCount, 5000);
        return () => clearInterval(interval);
    }, []);

    async function refreshPendingCount() {
        try {
            const pending = await getPendingOfflineRecords();
            setPendingCount(pending.length);
        } catch (_) {}
    }

    /**
     * Sincroniza los registros guardados offline con Supabase.
     * Lógica anti-duplicados:
     *   1. Antes de insertar, busca en count_records si ya existe un registro
     *      con el mismo local_id en la columna "note" (campo que usamos para guardar el local_id).
     *      Si ya existe, simplemente lo marca como sincronizado sin insertar de nuevo.
     *   2. Si no existe, inserta y marca como sincronizado.
     */
    async function syncPendingRecords() {
        if (syncLockRef.current) return;
        syncLockRef.current = true;
        setSyncing(true);
        try {
            const pending = await getPendingOfflineRecords();
            if (pending.length === 0) return;

            let synced = 0;
            let failed = 0;
            for (const rec of pending) {
                try {
                    // Verificar si ya fue subido (deduplicación por local_id en nota)
                    const { data: existing } = await supabase
                        .from("count_records")
                        .select("id")
                        .eq("note", `__offline__${rec.local_id}`)
                        .maybeSingle();

                    if (existing) {
                        // Ya existe en DB — solo marcar como synced localmente
                        await markOfflineRecordSynced(rec.local_id);
                        synced++;
                        continue;
                    }

                    // Insertar en Supabase
                    const { error } = await supabase.from("count_records").insert({
                        inventory_id: rec.inventory_id,
                        product_id: rec.product_id,
                        sku: rec.sku,
                        barcode: rec.barcode,
                        description: rec.description,
                        unit: rec.unit,
                        counted_quantity: rec.counted_quantity,
                        system_stock: rec.system_stock,
                        difference: rec.difference,
                        location: rec.location,
                        user_id: rec.user_id,
                        user_name: rec.user_name,
                        status: rec.status,
                        note: `__offline__${rec.local_id}`,
                        cost: rec.cost,
                        counted_at: rec.counted_at,
                    });

                    if (!error) {
                        await markOfflineRecordSynced(rec.local_id);
                        synced++;
                    } else {
                        failed++;
                    }
                } catch (_) {
                    failed++;
                }
            }

            await refreshPendingCount();
            if (synced > 0) showMessage(`✅ Sincronizados ${synced} registros offline.`, "success");
            if (failed > 0) showMessage(`⚠️ ${failed} registros no pudieron sincronizarse. Se reintentará.`, "error");
        } finally {
            syncLockRef.current = false;
            setSyncing(false);
        }
    }

    useEffect(() => {
        const raw = localStorage.getItem("session_user");
        if (!raw) { window.location.href = "/"; return; }
        const parsed = JSON.parse(raw) as User;

        function applyUser(fresh: any) {
            const u: User = {
                id: fresh.id,
                username: fresh.username,
                full_name: fresh.full_name,
                role: fresh.role as Role,
                roles: fresh.roles && fresh.roles.length > 0 ? (fresh.roles as Role[]) : [fresh.role as Role],
                inventory_id: fresh.inventory_id ?? null,
                can_access_any_inventory: fresh.can_access_any_inventory ?? false,
                permissions: {
                    can_see_system_stock: fresh.can_see_system_stock ?? false,
                    can_see_cost: fresh.can_see_cost ?? false,
                    can_see_valued_difference: fresh.can_see_valued_difference ?? false,
                },
            };
            localStorage.setItem("session_user", JSON.stringify(u));
            setUser(u);
            const roles: Role[] = u.roles ?? [u.role];
            if (roles.includes("Administrador")) { setActiveTab("admin"); }
            else if (roles.includes("Operario")) { setActiveTab("operario"); }
            else if (roles.includes("Validador")) { setActiveTab("validador"); }
        }

        (async () => {
            try {
                if (navigator.onLine) {
                    const { data, error } = await supabase
                        .from("app_users")
                        .select("*")
                        .eq("id", parsed.id)
                        .maybeSingle();
                    applyUser((!error && data) ? data : parsed);
                } else {
                    applyUser(parsed);
                }
            } catch {
                applyUser(parsed);
            }
        })();
    }, []);

    useEffect(() => { if (user) loadInventories(); }, [user]);

    useEffect(() => {
        if (user && selectedInventoryId) {
            if (navigator.onLine) {
                loadAll();
                loadSessions();
            } else {
                setLoading(false);
            }
            setSelectedProduct(null);
            setSearchValue("");
            setLocation("");
            setQuantity("");
            setSearchResults([]);
        }
    }, [user, selectedInventoryId]);

    useEffect(() => {
        if (!selectedInventoryId) return;

        const channel = supabase
            .channel(`realtime-inventory-${selectedInventoryId}`)
            .on(
                "postgres_changes",
                { event: "INSERT", schema: "public", table: "count_records", filter: `inventory_id=eq.${selectedInventoryId}` },
                (payload) => {
                    setRecords((prev) => {
                        const newRecord = payload.new as RecordRow;
                        if (prev.some((r) => r.id === newRecord.id)) return prev;
                        return [newRecord, ...prev];
                    });
                }
            )
            .on(
                "postgres_changes",
                { event: "UPDATE", schema: "public", table: "count_records", filter: `inventory_id=eq.${selectedInventoryId}` },
                (payload) => {
                    setRecords((prev) => prev.map((r) => r.id === payload.new.id ? (payload.new as RecordRow) : r));
                }
            )
            .on(
                "postgres_changes",
                { event: "DELETE", schema: "public", table: "count_records", filter: `inventory_id=eq.${selectedInventoryId}` },
                (payload) => {
                    setRecords((prev) => prev.filter((r) => r.id !== payload.old.id));
                }
            )
            .on(
                "postgres_changes",
                { event: "*", schema: "public", table: "products", filter: `inventory_id=eq.${selectedInventoryId}` },
                () => { if (navigator.onLine) loadAll(); }
            )
            .subscribe();

        return () => { supabase.removeChannel(channel); };
    }, [selectedInventoryId]);

    useEffect(() => {
        import("html5-qrcode").then((mod) => { html5QrCodeModuleRef.current = mod; }).catch(() => {});
    }, []);

    useEffect(() => {
        if (!scannerTarget) return;
        let cancelled = false;
        async function startScanner() {
            try {
                const module = html5QrCodeModuleRef.current ?? await import("html5-qrcode");
                html5QrCodeModuleRef.current = module;
                const Html5Qrcode = module.Html5Qrcode;
                if (cancelled) return;
                const html5QrCode = new Html5Qrcode(scannerContainerId);
                scannerRef.current = html5QrCode;
                setScannerRunning(true);
                await html5QrCode.start(
                    { facingMode: "environment" },
                    { fps: 15, qrbox: { width: 280, height: 140 }, aspectRatio: 1.7 },
                    (decodedText: string) => { applyScannedValue(decodedText); },
                    () => {}
                );
                try {
                    const capabilities: any = (html5QrCode as any).getRunningTrackCapabilities?.();
                    setTorchAvailable(!!capabilities?.torch);
                } catch (_e) { setTorchAvailable(false); }
            } catch (error: any) {
                showMessage("No se pudo iniciar la cámara. Acepta el permiso si es la primera vez. " + (error?.message || ""), "error");
                setScannerRunning(false);
                setScannerTarget(null);
            }
        }
        startScanner();
        return () => { cancelled = true; stopScanner(); };
    }, [scannerTarget]);

    useEffect(() => {
        const hasOverlayOpen = !!scannerTarget || !!editingRecord;
        if (hasOverlayOpen && !overlayOpenedRef.current) {
            window.history.pushState({ overlay: true }, "");
            overlayOpenedRef.current = true;
        }
        if (!hasOverlayOpen) overlayOpenedRef.current = false;
        const handlePopState = () => {
            if (scannerTarget) { closeScanner(); return; }
            if (editingRecord) { closeEdit(); return; }
        };
        window.addEventListener("popstate", handlePopState);
        return () => { window.removeEventListener("popstate", handlePopState); };
    }, [scannerTarget, editingRecord]);

    useEffect(() => {
        if (validadorSubTab === "resumen" && selectedInventoryId) {
            buildAuditFromDB();
        }
    }, [validadorSubTab, selectedInventoryId, records]);

    function clearMessage() { setMessage(""); }

    // ── FUNCIONES DE SESIONES ────────────────────────────────────────────────
    async function loadSessions() {
        if (!selectedInventoryId) return;
        setSessionsLoading(true);
        try {
            const { data, error } = await supabase
                .from("counting_sessions")
                .select("*")
                .eq("inventory_id", selectedInventoryId)
                .order("created_at", { ascending: false });
            if (error) { showMessage("No se pudieron cargar las sesiones: " + error.message, "error"); return; }
            setSessions((data || []) as CountingSession[]);
            const openSession = ((data || []) as CountingSession[]).find(s => s.status === "open");
            if (openSession && !activeSessionId) setActiveSessionId(openSession.id);
        } finally {
            setSessionsLoading(false);
        }
    }

    async function createSession() {
        if (!user) return;
        if (!selectedInventoryId) { showMessage("Selecciona un inventario primero.", "error"); return; }
        if (!newSessionName.trim()) { showMessage("Escribe un nombre para la sesión.", "error"); return; }
        const { error } = await supabase.from("counting_sessions").insert({
            inventory_id: selectedInventoryId,
            name: newSessionName.trim(),
            description: newSessionDescription.trim() || null,
            created_by: user.id,
            created_by_name: user.full_name,
            created_at: new Date().toISOString(),
            closed_at: null,
            status: "open",
        });
        if (error) { showMessage("No se pudo crear la sesión: " + error.message, "error"); return; }
        setNewSessionName("");
        setNewSessionDescription("");
        showMessage("✅ Sesión creada correctamente.", "success");
        await loadSessions();
    }

    async function closeSession(session: CountingSession) {
        const confirmClose = window.confirm(`¿Cerrar la sesión "${session.name}"? No se podrán agregar más registros a ella.`);
        if (!confirmClose) return;
        const { error } = await supabase.from("counting_sessions")
            .update({ status: "closed", closed_at: new Date().toISOString() })
            .eq("id", session.id);
        if (error) { showMessage("No se pudo cerrar la sesión: " + error.message, "error"); return; }
        if (activeSessionId === session.id) setActiveSessionId(null);
        showMessage("✅ Sesión cerrada.", "success");
        await loadSessions();
    }

    async function reopenSession(session: CountingSession) {
        const confirmReopen = window.confirm(`¿Reabrir la sesión "${session.name}"?`);
        if (!confirmReopen) return;
        const { error } = await supabase.from("counting_sessions")
            .update({ status: "open", closed_at: null })
            .eq("id", session.id);
        if (error) { showMessage("No se pudo reabrir la sesión: " + error.message, "error"); return; }
        showMessage("✅ Sesión reabierta.", "success");
        await loadSessions();
    }

    async function deleteSession(session: CountingSession) {
        const confirmDel = window.confirm(`¿Eliminar permanentemente la sesión "${session.name}"?`);
        if (!confirmDel) return;
        const { error } = await supabase.from("counting_sessions").delete().eq("id", session.id);
        if (error) { showMessage("No se pudo eliminar la sesión: " + error.message, "error"); return; }
        if (activeSessionId === session.id) setActiveSessionId(null);
        showMessage("✅ Sesión eliminada.", "success");
        await loadSessions();
    }

    async function loadAllSessionsSummary() {
        setAllSessionsLoading(true);
        try {
            const { data: sessData } = await supabase
                .from("counting_sessions")
                .select("*")
                .order("created_at", { ascending: false });
            const { data: invData } = await supabase.from("inventories").select("id,name");
            const { data: recData } = await supabase
                .from("count_records")
                .select("session_id, sku, counted_quantity, system_stock, cost, difference, status");

            const invMap = new Map<string, string>((invData || []).map((i: any) => [i.id, i.name]));
            const recsBySess = new Map<string, any[]>();
            for (const r of (recData || [])) {
                if (!r.session_id) continue;
                if (!recsBySess.has(r.session_id)) recsBySess.set(r.session_id, []);
                recsBySess.get(r.session_id)!.push(r);
            }

            const summaries: SessionSummary[] = ((sessData || []) as CountingSession[]).map(session => {
                const recs = recsBySess.get(session.id) || [];
                const skuMap = new Map<string, { counted: number; system_stock: number; cost: number }>();
                for (const r of recs) {
                    const key = String(r.sku || "").toLowerCase().trim();
                    if (!skuMap.has(key)) skuMap.set(key, { counted: 0, system_stock: Number(r.system_stock || 0), cost: Number(r.cost || 0) });
                    skuMap.get(key)!.counted += Number(r.counted_quantity || 0);
                }
                let ok_count = 0, faltantes_count = 0, sobrantes_count = 0;
                let valued_difference = 0, total_counted_value = 0;
                for (const entry of skuMap.values()) {
                    const diff = entry.counted - entry.system_stock;
                    if (diff === 0) ok_count++;
                    else if (diff < 0) faltantes_count++;
                    else sobrantes_count++;
                    valued_difference += diff * entry.cost;
                    total_counted_value += entry.counted * entry.cost;
                }
                const total_skus_with_stock = [...skuMap.values()].filter(e => e.system_stock > 0).length;
                return {
                    session,
                    inventory_name: invMap.get(session.inventory_id) || session.inventory_id,
                    total_records: recs.length,
                    total_skus: skuMap.size,
                    ok_count,
                    faltantes_count,
                    sobrantes_count,
                    no_contado_count: 0,
                    valued_difference,
                    total_counted_value,
                    avance_pct: total_skus_with_stock > 0 ? Math.round((ok_count / total_skus_with_stock) * 100) : 0,
                };
            });

            setAllSessionsSummary(summaries);
        } finally {
            setAllSessionsLoading(false);
        }
    }

    async function loadInventories() {
        if (!user) return;
        const { data: allData, error: allError } = await supabase.from("inventories").select("*").order("name");
        if (allError && navigator.onLine) { showMessage("No se pudieron cargar todos los inventarios: " + allError.message, "error"); return; }
        setAllInventories((allData || []) as Inventory[]);
        const userRoles: Role[] = user.roles || [user.role];
        const isAdmin = userRoles.includes("Administrador");
        const freeAccess = isAdmin || (user as any).can_access_any_inventory === true;
        let query = supabase.from("inventories").select("*").eq("is_active", true).order("name");
        if (!freeAccess) {
            if (!user.inventory_id) {
                if (navigator.onLine) {
                    const { data: freshUser } = await supabase
                        .from("app_users").select("*").eq("id", user.id).maybeSingle();
                    if (freshUser?.can_access_any_inventory === true) {
                        const updated = { ...user, can_access_any_inventory: true };
                        localStorage.setItem("session_user", JSON.stringify(updated));
                        setUser(updated as User);
                        const { data: allActive } = await supabase.from("inventories").select("*").eq("is_active", true).order("name");
                        const list = (allActive || []) as Inventory[];
                        setInventories(list);
                        if (!list.length) { showMessage("No hay inventarios activos.", "error"); return; }
                        const saved = localStorage.getItem(CURRENT_INVENTORY_KEY);
                        const exists = list.find((x) => x.id === saved);
                        const selected = exists ? exists.id : list[0].id;
                        setSelectedInventoryId(selected);
                        localStorage.setItem(CURRENT_INVENTORY_KEY, selected);
                        return;
                    }
                    showMessage("Tu usuario no tiene inventario asignado. Contacta al administrador.", "error");
                }
                return;
            }
            query = query.eq("id", user.inventory_id);
        }
        const { data, error } = await query;
        if (error && navigator.onLine) { showMessage("No se pudieron cargar los inventarios: " + error.message, "error"); return; }
        const list = (data || []) as Inventory[];
        setInventories(list);
        if (!list.length && navigator.onLine) { showMessage("No hay inventarios disponibles para este usuario.", "error"); return; }
        if (freeAccess) {
            const saved = localStorage.getItem(CURRENT_INVENTORY_KEY);
            const exists = list.find((x) => x.id === saved);
            const selected = exists ? exists.id : list[0]?.id || "";
            setSelectedInventoryId(selected);
            localStorage.setItem(CURRENT_INVENTORY_KEY, selected);
            return;
        }
        const selected = list[0]?.id || "";
        setSelectedInventoryId(selected);
        localStorage.setItem(CURRENT_INVENTORY_KEY, selected);
    }

    async function loadAll() {
        if (!selectedInventoryId) return;
        setLoading(true);

        const { count: totalProducts } = await supabase
            .from("products").select("*", { count: "exact", head: true })
            .eq("inventory_id", selectedInventoryId);

        const [rRes, uRes, stockRes, valueRes] = await Promise.all([
            supabase.from("count_records").select("*").eq("inventory_id", selectedInventoryId).order("counted_at", { ascending: false }),
            supabase.from("app_users").select("*").order("username"),
            supabase.from("products").select("*", { count: "exact", head: true }).eq("inventory_id", selectedInventoryId).gt("system_stock", 0),
            supabase.rpc("get_inventory_value", { inv_id: selectedInventoryId }),
        ]);

        const allProducts = await fetchAllProducts(selectedInventoryId);

        const realInventoryValue = valueRes.data || 0;
        setProducts(allProducts);
        setRecords((rRes.data || []) as RecordRow[]);
        setUsers((uRes.data || []) as AppUser[]);
        setTotalProductCount(totalProducts || allProducts.length);
        setTotalProductWithStockCount(stockRes.count || 0);
        setTotalInventoryValue(realInventoryValue);

        const eligibleProducts = allProducts.filter((p) => Number(p.system_stock || 0) > 0);
        const eligibleSkuSet = new Set(eligibleProducts.map((p) => normalizeText(p.sku)));
        const recs = (rRes.data || []) as RecordRow[];
        const countedEligible = new Set(recs.map((r) => normalizeText(r.sku)).filter((sku) => eligibleSkuSet.has(sku))).size;
        const totalEligible = stockRes.count || eligibleProducts.length;
        setSkuProgress({
            total: totalEligible,
            counted: countedEligible,
            pct: totalEligible ? Number(((countedEligible / totalEligible) * 100).toFixed(2)) : 0,
        });

        setLoading(false);
    }

    async function buildAuditFromDB() {
        if (!selectedInventoryId) return;
        setAuditLoading(true);
        try {
            const allProducts = products.length > 0 ? products : await fetchAllProducts(selectedInventoryId);

            const { data: rData } = await supabase
                .from("count_records")
                .select("*")
                .eq("inventory_id", selectedInventoryId);
            const allRecords = (rData || []) as RecordRow[];

            const map = new Map<string, AuditRow>();
            for (const r of allRecords) {
                const key = normalizeText(r.sku);
                if (!map.has(key)) {
                    map.set(key, {
                        sku: r.sku,
                        description: r.description,
                        unit: r.unit,
                        cost: Number(r.cost || 0),
                        system_stock: Number(r.system_stock || 0),
                        total_counted: 0,
                        difference: 0,
                        valued_difference: 0,
                        record_count: 0,
                        not_counted: false,
                        status_resumen: "OK",
                    });
                }
                const entry = map.get(key)!;
                entry.total_counted += Number(r.counted_quantity || 0);
                entry.record_count += 1;
            }

            for (const p of allProducts) {
                const sysStock = Number(p.system_stock || 0);
                if (sysStock <= 0) continue;
                const key = normalizeText(p.sku);
                if (!map.has(key)) {
                    map.set(key, {
                        sku: p.sku,
                        description: p.description,
                        unit: p.unit,
                        cost: Number(p.cost || 0),
                        system_stock: sysStock,
                        total_counted: 0,
                        difference: -sysStock,
                        valued_difference: -sysStock * Number(p.cost || 0),
                        record_count: 0,
                        not_counted: true,
                        status_resumen: "NO CONTADO",
                    });
                }
            }

            const rows: AuditRow[] = [];
            for (const entry of map.values()) {
                if (!entry.not_counted) {
                    entry.difference = entry.total_counted - entry.system_stock;
                    entry.valued_difference = entry.difference * entry.cost;
                    if (entry.difference === 0) entry.status_resumen = "OK";
                    else if (entry.difference < 0) entry.status_resumen = "FALTANTE";
                    else entry.status_resumen = "SOBRANTE";
                }
                rows.push(entry);
            }

            rows.sort((a, b) => {
                if (a.not_counted && !b.not_counted) return 1;
                if (!a.not_counted && b.not_counted) return -1;
                return a.sku.localeCompare(b.sku);
            });

            setAuditByCode(rows);
        } finally {
            setAuditLoading(false);
        }
    }

    function handleInventoryChange(value: string) {
        const userRoles: Role[] = user?.roles || (user ? [user.role] : []);
        const isAdmin = userRoles.includes("Administrador");
        const freeAccess = isAdmin || (user as any)?.can_access_any_inventory === true;
        if (!freeAccess) { showMessage("Tu usuario solo puede trabajar en el inventario asignado.", "error"); return; }
        setSelectedInventoryId(value);
        localStorage.setItem(CURRENT_INVENTORY_KEY, value);
        setSelectedProduct(null);
        setSearchValue("");
        setLocation("");
        setQuantity("");
        setAuditByCode([]);
    }

    function logout() {
        localStorage.removeItem("session_user");
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
        window.location.href = "/";
    }

    async function findProductForScanner(rawValue: string) {
        const value = cleanCode(rawValue);
        if (!value) return null;

        if (!navigator.onLine) {
            const foundInMemory = products.find(p =>
                normalizeText(p.sku) === normalizeText(value)
            );
            return foundInMemory || null;
        }

        const { data: barcodeMatch, error: barcodeError } = await supabase
            .from("product_barcodes").select("product_id").eq("barcode", value).maybeSingle();
        if (!barcodeError && barcodeMatch?.product_id) {
            const foundInMemory = products.find(p => p.id === barcodeMatch.product_id);
            if (foundInMemory) return foundInMemory;
            const { data: directProduct } = await supabase.from("products").select("*").eq("id", barcodeMatch.product_id).maybeSingle();
            if (directProduct) return directProduct as Product;
        }

        const { data: globalMatch } = await supabase
            .from("global_barcodes").select("sku").eq("barcode", value).maybeSingle();
        if (globalMatch?.sku) {
            const { data: productBySku } = await supabase.from("products").select("*")
                .eq("inventory_id", selectedInventoryId).eq("sku", globalMatch.sku).maybeSingle();
            if (productBySku) return productBySku as Product;
            const foundInMemory = products.find(p => normalizeText(p.sku) === normalizeText(globalMatch.sku));
            if (foundInMemory) return foundInMemory;
        }

        if (/^\d+$/.test(value)) {
            const trimmed = value.replace(/^0+/, "");
            if (trimmed !== value) {
                const { data: barcodeMatchTrim } = await supabase
                    .from("product_barcodes").select("product_id").eq("barcode", trimmed).maybeSingle();
                if (barcodeMatchTrim?.product_id) {
                    const foundInMemory = products.find(p => p.id === barcodeMatchTrim.product_id);
                    if (foundInMemory) return foundInMemory;
                    const { data: directProduct } = await supabase.from("products").select("*").eq("id", barcodeMatchTrim.product_id).maybeSingle();
                    if (directProduct) return directProduct as Product;
                }
                const { data: globalMatchTrim } = await supabase
                    .from("global_barcodes").select("sku").eq("barcode", trimmed).maybeSingle();
                if (globalMatchTrim?.sku) {
                    const { data: productBySku } = await supabase.from("products").select("*")
                        .eq("inventory_id", selectedInventoryId).eq("sku", globalMatchTrim.sku).maybeSingle();
                    if (productBySku) return productBySku as Product;
                }
            }
        }

        const { data: skuExact } = await supabase.from("products").select("*")
            .eq("inventory_id", selectedInventoryId).eq("sku", value).maybeSingle();
        if (skuExact) return skuExact as Product;

        const foundBySku = products.find(p => normalizeText(p.sku) === normalizeText(value));
        if (foundBySku) return foundBySku;
        return null;
    }

    async function searchProductsAdvanced(text: string) {
        if (!text) return [];
        const words = text.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").split(/\s+/).filter(Boolean);
        if (words.length === 0) return [];

        // Offline: buscar en productos cargados en memoria
        if (!navigator.onLine || products.length > 0) {
            const results = products.filter(p => {
                const haystack = (p.sku + " " + p.description).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                return words.every(w => haystack.includes(w));
            }).slice(0, 20);
            return results;
        }

        let descQuery = supabase.from("products").select("*").eq("inventory_id", selectedInventoryId);
        for (const w of words) { descQuery = descQuery.ilike("description", `%${w}%`); }
        const { data: byDesc } = await descQuery.limit(300);
        let skuQuery = supabase.from("products").select("*").eq("inventory_id", selectedInventoryId);
        for (const w of words) { skuQuery = skuQuery.ilike("sku", `%${w}%`); }
        const { data: bySku } = await skuQuery.limit(300);
        const combined = [...(byDesc || []), ...(bySku || [])];
        const seen = new Set<string>();
        const deduped = combined.filter((p) => { if (seen.has(p.id)) return false; seen.add(p.id); return true; });
        return deduped.filter((p) => {
            const haystack = (p.sku + " " + p.description).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            return words.every((w) => haystack.includes(w));
        }).slice(0, 20);
    }

    const searchDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

    function handleSearchInputChange(value: string) {
        setSearchValue(value);
        const cleanValue = value.trim();
        if (!cleanValue) { setSelectedProduct(null); setSearchResults([]); setMessage(""); return; }
        if (searchDebounceRef.current) clearTimeout(searchDebounceRef.current);
        searchDebounceRef.current = setTimeout(async () => {
            const results = await searchProductsAdvanced(cleanValue);
            if (results.length === 0) { setSelectedProduct(null); setSearchResults([]); showMessage("No se encontró el producto.", "error"); return; }
            if (results.length === 1) { setSelectedProduct(results[0]); setSearchResults([]); showMessage("Producto encontrado correctamente.", "success"); return; }
            setSelectedProduct(null);
            setSearchResults(results);
            showMessage(`Se encontraron ${results.length} productos. Selecciona uno.`, "info");
        }, 350);
    }

    async function saveCount() {
        clearMessage();
        if (!user) return;
        if (!selectedInventoryId) { showMessage("Primero selecciona un inventario.", "error"); return; }
        if (!selectedProduct) { showMessage("Primero busca un producto del maestro.", "error"); return; }
        if (!location.trim() || quantity === "") { showMessage("Completa ubicación y cantidad.", "error"); return; }
        const counted = Number(quantity);
        if (Number.isNaN(counted) || counted < 0) { showMessage("La cantidad no es válida.", "error"); return; }
        const systemStock = Number(selectedProduct.system_stock || 0);
        const difference = counted - systemStock;
        const status: RecordRow["status"] = difference === 0 ? "Pendiente" : "Diferencia";

        if (!navigator.onLine) {
            // Guardar offline
            const offlineRec: OfflineRecord = {
                local_id: generateLocalId(),
                synced: false,
                created_at_local: new Date().toISOString(),
                inventory_id: selectedInventoryId,
                product_id: selectedProduct.id,
                sku: selectedProduct.sku,
                barcode: null,
                description: selectedProduct.description,
                unit: selectedProduct.unit,
                counted_quantity: counted,
                system_stock: systemStock,
                difference,
                location: location.trim(),
                user_id: user.id,
                user_name: user.full_name,
                status,
                note: "",
                cost: Number(selectedProduct.cost || 0),
                counted_at: new Date().toISOString(),
            };
            await saveOfflineRecord(offlineRec);
            await refreshPendingCount();
            setSearchValue("");
            setSelectedProduct(null);
            setLocation("");
            setQuantity("");
            showMessage("📦 Guardado offline. Se subirá cuando tengas internet.", "success");
            try { if (navigator.vibrate) navigator.vibrate([60, 30, 60]); } catch (_) {}
            const searchInput = document.querySelector("input[placeholder*='SKU']") as HTMLInputElement;
            if (searchInput) searchInput.focus();
            return;
        }

        const { error } = await supabase.from("count_records").insert({
            inventory_id: selectedInventoryId,
            product_id: selectedProduct.id,
            sku: selectedProduct.sku,
            barcode: null,
            description: selectedProduct.description,
            unit: selectedProduct.unit,
            counted_quantity: counted,
            system_stock: systemStock,
            difference,
            location: location.trim(),
            user_id: user.id,
            user_name: user.full_name,
            status,
            note: "",
            cost: Number(selectedProduct.cost || 0),
            counted_at: new Date().toISOString(),
        });
        if (error) { showMessage("No se pudo guardar el conteo: " + error.message, "error"); return; }
        setSearchValue("");
        setSelectedProduct(null);
        setLocation("");
        setQuantity("");
        showMessage("✅ Conteo guardado correctamente.", "success");
        try { if (navigator.vibrate) navigator.vibrate([60, 30, 60]); } catch (_) {}
        const searchInput = document.querySelector("input[placeholder*='SKU']") as HTMLInputElement;
        if (searchInput) searchInput.focus();
    }

    async function createUser() {
        clearMessage();
        if (!newUsername || !newPassword || !newFullName || newRoles.length === 0) {
            showMessage("Completa todos los datos del nuevo usuario.", "error"); return;
        }
        const primaryRole: Role = newRoles.includes("Administrador")
            ? "Administrador"
            : newRoles.includes("Validador")
            ? "Validador"
            : "Operario";

        const payload: any = {
            username: newUsername.trim(),
            password: newPassword.trim(),
            full_name: newFullName.trim(),
            role: primaryRole,
            roles: newRoles,
            is_active: true,
            can_see_system_stock: newPermissions.can_see_system_stock,
            can_see_cost: newPermissions.can_see_cost,
            can_see_valued_difference: newPermissions.can_see_valued_difference,
        };
        if (primaryRole === "Administrador") {
            payload.inventory_id = null;
            payload.can_access_any_inventory = true;
        } else if (newCanAccessAnyInventory) {
            payload.inventory_id = null;
            payload.can_access_any_inventory = true;
        } else {
            if (!selectedInventoryId) { showMessage("Primero selecciona un inventario para asignar al usuario.", "error"); return; }
            payload.inventory_id = selectedInventoryId;
            payload.can_access_any_inventory = false;
        }

        const { error } = await supabase.from("app_users").insert(payload);
        if (error) { showMessage("No se pudo crear el usuario: " + error.message, "error"); return; }
        setNewUsername(""); setNewPassword(""); setNewFullName("");
        setNewRoles(["Operario"]); setShowNewPassword(false);
        setNewPermissions({ ...DEFAULT_PERMISSIONS });
        setNewCanAccessAnyInventory(false);
        showMessage("✅ Usuario creado correctamente.", "success");
        loadAll();
    }

    async function deleteUser(userToDelete: AppUser) {
        clearMessage();
        if (user?.role !== "Administrador") { showMessage("Solo el administrador puede eliminar usuarios.", "error"); return; }
        if (userToDelete.id === user.id) { showMessage("No puedes eliminar tu propio usuario mientras estás conectado.", "error"); return; }
        const confirmDelete = window.confirm(`¿Seguro que deseas eliminar al usuario "${userToDelete.username}"?`);
        if (!confirmDelete) return;
        const { error } = await supabase.from("app_users").delete().eq("id", userToDelete.id);
        if (error) { showMessage("No se pudo eliminar el usuario: " + error.message, "error"); return; }
        showMessage("✅ Usuario eliminado correctamente.", "success");
        loadAll();
    }

    function openEditUser(u: AppUser) {
        setEditingUser(u);
        const roles: Role[] = u.roles && u.roles.length > 0 ? u.roles : [u.role];
        setEditUserRoles(roles);
        const isAdmin = roles.includes("Administrador");
        const freeAccess = (u.can_access_any_inventory === true) || (!isAdmin && !u.inventory_id);
        setEditUserCanAccessAnyInventory(freeAccess);
        setEditUserInventoryId(u.inventory_id || "");
        setEditUserActive(u.is_active);
        setEditUserPermissions({
            can_see_system_stock: u.can_see_system_stock ?? false,
            can_see_cost: u.can_see_cost ?? false,
            can_see_valued_difference: u.can_see_valued_difference ?? false,
        });
    }

    async function saveEditUser() {
        if (!editingUser) return;
        const primaryRole: Role = editUserRoles.includes("Administrador")
            ? "Administrador"
            : editUserRoles.includes("Validador")
            ? "Validador"
            : "Operario";

        const payload: any = {
            role: primaryRole,
            roles: editUserRoles,
            is_active: editUserActive,
            can_see_system_stock: editUserPermissions.can_see_system_stock,
            can_see_cost: editUserPermissions.can_see_cost,
            can_see_valued_difference: editUserPermissions.can_see_valued_difference,
        };
        if (primaryRole === "Administrador") {
            payload.inventory_id = null;
            payload.can_access_any_inventory = true;
        } else if (editUserCanAccessAnyInventory) {
            payload.inventory_id = null;
            payload.can_access_any_inventory = true;
        } else {
            payload.inventory_id = editUserInventoryId || null;
            payload.can_access_any_inventory = false;
        }

        const { error } = await supabase.from("app_users").update(payload).eq("id", editingUser.id);
        if (error) { showMessage("Error al actualizar: " + error.message, "error"); return; }
        showMessage("✅ Usuario actualizado correctamente.", "success");
        setEditingUser(null);
        loadAll();
    }

    async function createInventory() {
        clearMessage();
        if (user?.role !== "Administrador") { showMessage("Solo el administrador puede crear inventarios.", "error"); return; }
        if (!newInventoryName.trim()) { showMessage("Escribe el nombre del inventario.", "error"); return; }
        const code = newInventoryCode.trim() ||
            newInventoryName.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
        const codeExists = allInventories.find((inv) => normalizeText(inv.code) === normalizeText(code));
        if (codeExists) { showMessage("Ya existe un inventario con ese código.", "error"); return; }
        const { error } = await supabase.from("inventories").insert({ name: newInventoryName.trim(), code, is_active: true });
        if (error) { showMessage("No se pudo crear el inventario: " + error.message, "error"); return; }
        setNewInventoryName(""); setNewInventoryCode("");
        showMessage("✅ Inventario creado correctamente.", "success");
        await loadInventories();
    }

    async function uploadMaster() {
        clearMessage();
        if (!selectedInventoryId) { showMessage("Primero selecciona un inventario.", "error"); return; }
        if (!masterFile) { showMessage("Primero selecciona un archivo del maestro.", "error"); return; }
        const confirmReplace = window.confirm("Se reemplazará todo el maestro del inventario seleccionado. ¿Deseas continuar?");
        if (!confirmReplace) return;
        try {
            const data = await masterFile.arrayBuffer();
            const workbook = XLSX.read(data);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawMatrix: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
            if (rawMatrix.length < 2) { showMessage("El maestro no tiene filas válidas.", "error"); return; }

            let dataStartRow = 0;
            for (let i = 0; i < rawMatrix.length; i++) {
                const colA = String(rawMatrix[i][0] ?? "").trim();
                const colB = String(rawMatrix[i][1] ?? "").trim();
                const looksLikeHeader = /^(cod|sku|codigo|c[oó]digo|item|art[íi]culo|producto|descripci[oó]n|descripcion)/i.test(colA);
                const isEmpty = colA.length === 0;
                if (!isEmpty && !looksLikeHeader && colB.length > 0) { dataStartRow = i; break; }
            }

            const skuMap = new Map<string, any>();
            for (let i = dataStartRow; i < rawMatrix.length; i++) {
                const row = rawMatrix[i];
                const rawSku = row[0];
                const sku = cleanCode(String(rawSku ?? ""));
                if (!sku) continue;
                const description = String(row[1] ?? "").trim();
                if (!description) continue;
                const skuKey = normalizeText(sku);
                if (!skuMap.has(skuKey)) {
                    skuMap.set(skuKey, {
                        inventory_id: selectedInventoryId,
                        sku,
                        description,
                        unit: String(row[2] ?? "").trim(),
                        cost: Number(row[3] ?? 0) || 0,
                        system_stock: Number(row[4] ?? 0) || 0,
                    });
                }
            }
            if (skuMap.size === 0) { showMessage("El maestro no tiene filas válidas. Verifica que el archivo tenga datos desde la fila 2 (col 1: Código, col 2: Descripción).", "error"); return; }

            setUploadProgress({ step: "Eliminando maestro anterior...", pct: 5 });
            const { error: rpcError } = await supabase.rpc("delete_inventory_products", { inv_id: selectedInventoryId });
            if (rpcError) {
                let hasMore = true;
                while (hasMore) {
                    const { data: ids } = await supabase.from("products").select("id").eq("inventory_id", selectedInventoryId).limit(1000);
                    if (!ids || ids.length === 0) { hasMore = false; break; }
                    for (let i = 0; i < ids.length; i += 100) {
                        await supabase.from("product_barcodes").delete().in("product_id", ids.slice(i, i + 100).map((p: any) => p.id));
                    }
                    await supabase.from("products").delete().in("id", ids.map((p: any) => p.id));
                }
            }

            const uniqueProducts = Array.from(skuMap.values());
            const totalUniqueProducts = uniqueProducts.length;
            let insertedCount = 0;
            const batchSize = 2000;

            for (let i = 0; i < uniqueProducts.length; i += batchSize) {
                const batch = uniqueProducts.slice(i, i + batchSize);
                const pct = Math.round(10 + ((i / totalUniqueProducts) * 85));
                setUploadProgress({
                    step: `Insertando... ${Math.min(i + batchSize, totalUniqueProducts).toLocaleString()} de ${totalUniqueProducts.toLocaleString()}`,
                    pct,
                });
                const { data: rpcResult, error: rpcInsertError } = await supabase.rpc("bulk_insert_products", {
                    products: JSON.stringify(batch),
                });
                if (rpcInsertError) {
                    setUploadProgress(null);
                    showMessage(`Error insertando productos: ${rpcInsertError.message}`, "error");
                    return;
                }
                insertedCount += rpcResult ?? batch.length;
            }

            setUploadProgress({ step: "Actualizando stock en conteos existentes...", pct: 95 });
            try {
                const { data: existingRecords } = await supabase
                    .from("count_records")
                    .select("id, sku, counted_quantity")
                    .eq("inventory_id", selectedInventoryId);

                if (existingRecords && existingRecords.length > 0) {
                    const stockMap = new Map<string, number>();
                    for (const [key, p] of skuMap.entries()) {
                        stockMap.set(key, p.system_stock);
                    }

                    const updates = existingRecords
                        .map((r: any) => {
                            const skuKey = normalizeText(r.sku);
                            const newStock = stockMap.get(skuKey) ?? 0;
                            const newDiff = Number(r.counted_quantity) - newStock;
                            const newStatus = newDiff === 0 ? "Pendiente" : "Diferencia";
                            return { id: r.id, system_stock: newStock, difference: newDiff, status: newStatus };
                        }) as { id: string; system_stock: number; difference: number; status: string }[];

                    for (let i = 0; i < updates.length; i += 100) {
                        const batch = updates.slice(i, i + 100);
                        for (const upd of batch) {
                            await supabase.from("count_records").update({
                                system_stock: upd.system_stock,
                                difference: upd.difference,
                                status: upd.status,
                            }).eq("id", upd.id);
                        }
                    }
                }
            } catch (_syncErr) {
                console.warn("No se pudo sincronizar system_stock en conteos:", _syncErr);
            }

            setUploadProgress({ step: "Finalizando...", pct: 99 });
            setUploadProgress(null);
            showMessage(`✅ Maestro cargado: ${insertedCount.toLocaleString()} productos únicos. Stock actualizado en conteos existentes.`, "success");
            setMasterFile(null); setMasterFileName("");
            if (masterInputRef.current) masterInputRef.current.value = "";
            await loadAll();
        } catch (err: any) {
            console.error(err);
            setUploadProgress(null);
            showMessage("Error: " + (err?.message || "desconocido"), "error");
        }
    }

    async function uploadGlobalBarcodes() {
        clearMessage();
        if (!globalBarcodesFile) { showMessage("Primero selecciona un archivo de códigos de barra.", "error"); return; }
        const confirm = window.confirm(
            "Se actualizará el catálogo global de códigos de barra.\n" +
            "Los códigos existentes con el mismo barcode se actualizarán con el nuevo SKU.\n\n" +
            "Columnas requeridas: SKU, CODIGO_BARRA\n\n¿Deseas continuar?"
        );
        if (!confirm) return;
        try {
            const data = await globalBarcodesFile.arrayBuffer();
            const workbook = XLSX.read(data);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

            const barcodeMap = new Map<string, string>();
            for (const row of rows) {
                const rawSku = row.SKU || row.sku;
                const sku = cleanCode(rawSku);
                if (!sku) continue;
                const rawBarcode = row.CODIGO_BARRA || row.barcode || row.BARCODE || "";
                const barcode = rawBarcode ? cleanCode(rawBarcode) : "";
                if (!barcode) continue;
                barcodeMap.set(barcode, sku);
            }

            if (barcodeMap.size === 0) { showMessage("El archivo no tiene filas válidas. Necesita columnas SKU y CODIGO_BARRA.", "error"); return; }

            setUploadProgress({ step: "Subiendo catálogo global...", pct: 10 });
            const rows2insert = Array.from(barcodeMap.entries()).map(([barcode, sku]) => ({
                sku,
                barcode,
                updated_at: new Date().toISOString(),
            }));

            const batchSize = 500;
            let inserted = 0;
            for (let i = 0; i < rows2insert.length; i += batchSize) {
                const batch = rows2insert.slice(i, i + batchSize);
                const pct = Math.round(10 + ((i / rows2insert.length) * 85));
                setUploadProgress({ step: `Insertando... ${Math.min(i + batchSize, rows2insert.length).toLocaleString()} de ${rows2insert.length.toLocaleString()}`, pct });
                const { error } = await supabase.from("global_barcodes").upsert(batch, { onConflict: "barcode" });
                if (error) { setUploadProgress(null); showMessage("Error al insertar: " + error.message, "error"); return; }
                inserted += batch.length;
            }

            setUploadProgress(null);
            setGlobalBarcodesCount(inserted);
            showMessage(`✅ Catálogo global actualizado: ${inserted.toLocaleString()} códigos de barra.`, "success");
            setGlobalBarcodesFile(null); setGlobalBarcodesFileName("");
            if (globalBarcodesInputRef.current) globalBarcodesInputRef.current.value = "";
        } catch (err: any) {
            setUploadProgress(null);
            showMessage("Error: " + (err?.message || "desconocido"), "error");
        }
    }

    async function deleteMaster() {
        clearMessage();
        if (!selectedInventoryId) { showMessage("Primero selecciona un inventario.", "error"); return; }
        if (user?.role !== "Administrador") { showMessage("Solo el administrador puede eliminar el maestro.", "error"); return; }
        const currentCount = totalProductCount > 0 ? totalProductCount : products.length;
        const confirmDelete = window.confirm(`⚠️ ¿Estás SEGURO? Vas a eliminar TODOS los ${currentCount.toLocaleString()} productos del inventario "${currentInventory?.name}".\n\nEsta acción no se puede deshacer.`);
        if (!confirmDelete) return;
        setUploadProgress({ step: "Eliminando productos vía RPC...", pct: 30 });
        try {
            const { error } = await supabase.rpc("delete_inventory_products", { inv_id: selectedInventoryId });
            if (error) {
                console.warn("RPC no disponible, fallback manual:", error.message);
                setUploadProgress({ step: "Obteniendo lista de productos...", pct: 2 });
                let allProductIds: string[] = [];
                let page = 0; const pageSize = 1000; let hasMore = true;
                while (hasMore) {
                    const { data: productPage } = await supabase.from("products").select("id").eq("inventory_id", selectedInventoryId).range(page * pageSize, (page + 1) * pageSize - 1);
                    if (productPage && productPage.length > 0) { allProductIds.push(...productPage.map(p => p.id)); page++; }
                    if (!productPage || productPage.length < pageSize) hasMore = false;
                    setUploadProgress({ step: `Relevando productos... ${allProductIds.length.toLocaleString()}`, pct: Math.min(5 + page * 2, 20) });
                }
                const totalIds = allProductIds.length;
                let deletedBarcodes = 0;
                for (let i = 0; i < allProductIds.length; i += 100) {
                    await supabase.from("product_barcodes").delete().in("product_id", allProductIds.slice(i, i + 100));
                    deletedBarcodes += Math.min(100, allProductIds.length - i);
                    setUploadProgress({ step: `Eliminando códigos de barra... ${deletedBarcodes.toLocaleString()} de ${totalIds.toLocaleString()}`, pct: Math.round(20 + (deletedBarcodes / totalIds) * 40) });
                }
                for (let i = 0; i < allProductIds.length; i += 500) {
                    await supabase.from("products").delete().in("id", allProductIds.slice(i, i + 500));
                    setUploadProgress({ step: `Eliminando productos... ${Math.min(i + 500, totalIds).toLocaleString()} de ${totalIds.toLocaleString()}`, pct: Math.round(60 + ((i + 500) / totalIds) * 38) });
                }
            }
            setUploadProgress({ step: "Finalizando...", pct: 99 });
            setUploadProgress(null);
            showMessage(`✅ Maestro eliminado: ${currentCount.toLocaleString()} productos eliminados.`, "success");
            setMasterFile(null); setMasterFileName("");
            setTotalProductCount(0); setTotalProductWithStockCount(0); setTotalInventoryValue(0);
            if (masterInputRef.current) masterInputRef.current.value = "";
            await loadAll();
        } catch (err: any) {
            setUploadProgress(null);
            showMessage("Error al eliminar: " + (err?.message || "desconocido"), "error");
        }
    }

    async function reactivateInventory(inv: Inventory) {
        clearMessage();
        if (user?.role !== "Administrador") { showMessage("Solo el administrador puede reactivar inventarios.", "error"); return; }
        const confirmReactivate = window.confirm(`¿Reactivar el inventario "${inv.name}"?`);
        if (!confirmReactivate) return;
        const { error } = await supabase.from("inventories").update({ is_active: true }).eq("id", inv.id);
        if (error) showMessage("No se pudo reactivar: " + error.message, "error");
        else { showMessage(`✅ Inventario "${inv.name}" reactivado.`, "success"); await loadInventories(); }
    }

    async function manageInventory(inv: Inventory) {
        clearMessage();
        if (user?.role !== "Administrador") { showMessage("Solo el administrador puede eliminar o archivar inventarios.", "error"); return; }
        setProcessingInventoryId(inv.id);
        try {
            const { count: recordCount } = await supabase.from("count_records").select("*", { count: "exact", head: true }).eq("inventory_id", inv.id);
            const { count: productCount } = await supabase.from("products").select("*", { count: "exact", head: true }).eq("inventory_id", inv.id);
            const hasData = (recordCount || 0) > 0 || (productCount || 0) > 0;
            if (hasData) {
                const confirmArchive = window.confirm(`El inventario "${inv.name}" tiene datos (${productCount || 0} productos, ${recordCount || 0} registros).\n\nSe archivará (quedará inactivo). ¿Deseas continuar?`);
                if (!confirmArchive) { setProcessingInventoryId(null); return; }
                const { error } = await supabase.from("inventories").update({ is_active: false }).eq("id", inv.id);
                if (error) showMessage("No se pudo archivar el inventario: " + error.message, "error");
                else { showMessage(`✅ Inventario "${inv.name}" archivado correctamente.`, "success"); await loadInventories(); }
            } else {
                const confirmDelete = window.confirm(`El inventario "${inv.name}" está vacío y se eliminará permanentemente. ¿Deseas continuar?`);
                if (!confirmDelete) { setProcessingInventoryId(null); return; }
                const { error } = await supabase.from("inventories").delete().eq("id", inv.id);
                if (error) showMessage("No se pudo eliminar el inventario: " + error.message, "error");
                else { showMessage(`✅ Inventario "${inv.name}" eliminado correctamente.`, "success"); await loadInventories(); }
            }
        } catch (err: any) {
            showMessage("Error inesperado: " + (err?.message || "desconocido"), "error");
        } finally {
            setProcessingInventoryId(null);
        }
    }

    async function uploadUsers() {
        clearMessage();
        if (!usersFile) { showMessage("Primero selecciona un archivo de usuarios.", "error"); return; }
        const confirmUpload = window.confirm("Se insertarán los usuarios del archivo. Los usuarios con ID duplicado serán omitidos. ¿Deseas continuar?");
        if (!confirmUpload) return;
        try {
            const data = await usersFile.arrayBuffer();
            const workbook = XLSX.read(data);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            const usersToInsert = [];
            for (const row of rows) {
                const username = String(row.ID || row.USERNAME || row.usuario || "").trim();
                const password = String(row.CLAVE || row.PASSWORD || row.clave || "").trim();
                const full_name = String(row.NOMBRE || row.NOMBRE_COMPLETO || row.nombre || "").trim();
                const rawRole = String(row.ROL || row.ROLE || row.rol || "Operario").trim();
                const validRoles: Role[] = ["Operario", "Validador", "Administrador"];
                const role: Role = validRoles.includes(rawRole as Role) ? (rawRole as Role) : "Operario";
                if (username && password && full_name) {
                    const payload: any = {
                        username, password, full_name, role, roles: [role], is_active: true,
                        can_access_any_inventory: role === "Administrador" ? true : false,
                    };
                    if (role !== "Administrador" && selectedInventoryId) payload.inventory_id = selectedInventoryId;
                    usersToInsert.push(payload);
                }
            }
            if (!usersToInsert.length) { showMessage("El archivo no tiene filas válidas. Verifica que tenga columnas: ID, CLAVE, NOMBRE, ROL.", "error"); return; }
            let insertedCount = 0; let skippedCount = 0;
            for (const u of usersToInsert) {
                const { error } = await supabase.from("app_users").insert(u);
                if (error) skippedCount++; else insertedCount++;
            }
            showMessage(`✅ Usuarios insertados: ${insertedCount}. Omitidos (duplicados u error): ${skippedCount}.`, "success");
            setUsersFile(null); setUsersFileName("");
            if (usersInputRef.current) usersInputRef.current.value = "";
            await loadAll();
        } catch (err: any) {
            console.error(err);
            showMessage("Error al leer el archivo: " + (err?.message || "desconocido"), "error");
        }
    }

    async function deleteRecord(record: RecordRow) {
        clearMessage();
        if (!canValidate(user)) { showMessage("Solo el validador o administrador pueden eliminar registros.", "error"); return; }
        const confirmDelete = window.confirm(`¿Seguro que deseas eliminar el registro del SKU "${record.sku}"?`);
        if (!confirmDelete) return;
        const { error } = await supabase.from("count_records").delete().eq("id", record.id);
        if (error) { showMessage("No se pudo eliminar el registro: " + error.message, "error"); return; }
        showMessage("✅ Registro eliminado correctamente.", "success");
        loadAll();
    }

    function exportRecords() {
        const rows = filteredRecords.map((r) => ({
            SKU: r.sku, DESCRIPCION: r.description, UNIDAD_MEDIDA: r.unit,
            CANTIDAD_CONTADA: r.counted_quantity, STOCK_SISTEMA: r.system_stock,
            DIFERENCIA: r.difference, UBICACION: r.location, USUARIO: r.user_name,
            VALIDADOR: r.validator_name || "", FECHA_HORA: formatDateTime(r.counted_at),
            ESTADO: r.status, OBSERVACION: r.note || "",
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Registros");
        XLSX.writeFile(wb, `registros_${currentInventory?.name || "inventario"}.xlsx`);
    }

    function openEdit(record: RecordRow) {
        setEditingRecord(record);
        setEditSku(record.sku);
        setEditQty(String(record.counted_quantity));
        setEditLocation(record.location);
        setEditStatus(record.status);
        setEditNote(record.note || "");
        setEditMatchedProduct(null);
        supabase.from("products").select("*")
            .eq("inventory_id", record.inventory_id || selectedInventoryId)
            .ilike("sku", record.sku)
            .maybeSingle()
            .then(({ data }) => { if (data) setEditMatchedProduct(data as Product); });
    }

    function closeEdit() {
        setEditingRecord(null); setEditSku(""); setEditQty(""); setEditLocation("");
        setEditStatus("Pendiente"); setEditNote(""); setEditMatchedProduct(null);
    }

    async function handleEditSkuChange(value: string) {
        setEditSku(value);
        if (!value.trim()) { setEditMatchedProduct(null); return; }
        const { data } = await supabase.from("products").select("*")
            .eq("inventory_id", editingRecord?.inventory_id || selectedInventoryId)
            .ilike("sku", value.trim())
            .maybeSingle();
        setEditMatchedProduct(data ? (data as Product) : null);
    }

    async function saveEdit() {
        if (!editingRecord) return;
        if (!editSku.trim()) { showMessage("El SKU no puede estar vacío.", "error"); return; }
        const qty = Number(editQty);
        if (Number.isNaN(qty) || qty < 0) { showMessage("La cantidad editada no es válida.", "error"); return; }

        const { data: productData } = await supabase.from("products").select("*")
            .eq("inventory_id", editingRecord.inventory_id || selectedInventoryId)
            .ilike("sku", editSku.trim())
            .maybeSingle();

        const productMatch = productData as Product | null;
        if (!productMatch) { showMessage("El SKU editado no existe en el maestro de este inventario.", "error"); return; }

        const difference = qty - Number(productMatch.system_stock || 0);
        let finalStatus: RecordRow["status"] = editStatus;
        if (user?.role === "Operario") { finalStatus = difference === 0 ? "Pendiente" : "Diferencia"; }
        if (canValidate(user)) {
            if (editStatus === "Pendiente" || editStatus === "Diferencia") { finalStatus = difference === 0 ? "Validado" : "Corregido"; }
        }
        const payload: any = {
            inventory_id: selectedInventoryId, product_id: productMatch.id, sku: productMatch.sku,
            barcode: editingRecord.barcode || null, description: productMatch.description, unit: productMatch.unit,
            cost: productMatch.cost, system_stock: productMatch.system_stock, counted_quantity: qty,
            location: editLocation.trim(), difference, status: finalStatus, note: editNote.trim(),
            counted_at: new Date().toISOString(),
        };
        if (canValidate(user) && user) { payload.validator_name = user.full_name; }
        const { error } = await supabase.from("count_records").update(payload).eq("id", editingRecord.id);
        if (error) { showMessage("No se pudo editar el registro: " + error.message, "error"); return; }
        showMessage("✅ Registro actualizado correctamente.", "success");
        closeEdit();
        loadAll();
    }

    async function stopScanner() {
        try {
            if (scannerRef.current) { await scannerRef.current.stop(); await scannerRef.current.clear(); scannerRef.current = null; }
        } catch (_e) { scannerRef.current = null; }
        finally { setScannerRunning(false); }
    }

    function closeScanner() {
        scanHandledRef.current = false; setTorchOn(false); setTorchAvailable(false);
        stopScanner(); setScannerTarget(null);
    }

    async function toggleTorch() {
        try {
            if (!scannerRef.current) return;
            const nextTorch = !torchOn;
            await (scannerRef.current as any).applyVideoConstraints?.({ advanced: [{ torch: nextTorch }] });
            setTorchOn(nextTorch);
        } catch (_e) { showMessage("La linterna no está disponible en este dispositivo.", "error"); }
    }

    async function applyScannedValue(decodedText: string) {
        const cleanText = String(decodedText || "").trim();
        if (!cleanText) return;
        if (scanHandledRef.current) return;
        scanHandledRef.current = true;

        closeScanner();

        if (scannerTarget === "product") {
            setSearchValue(cleanText);
            const found = await findProductForScanner(cleanText);
            if (!found) {
                setSelectedProduct(null); setSearchResults([]);
                showMessage(`⚠️ Código escaneado: "${cleanText}" — no existe en el maestro.`, "error");
                return;
            }
            setSelectedProduct(found); setSearchResults([]);
            showMessage(`✅ ${found.sku} — ${found.description}`, "success");
            return;
        }
        if (scannerTarget === "location") {
            setLocation(cleanText);
            showMessage(`📍 Ubicación: ${cleanText}`, "success");
            return;
        }
    }

    function openScanner(target: "product" | "location") {
        clearMessage(); scanHandledRef.current = false; setScannerTarget(target);
    }

    const currentInventory = useMemo(() => allInventories.find((x) => x.id === selectedInventoryId) || null, [allInventories, selectedInventoryId]);
    const operarioRecords = useMemo(() => records.filter((r) => r.user_id === user?.id), [records, user]);
    const _operarioValidatedCount = useMemo(() => operarioRecords.filter((r) => r.status === "Validado" || r.status === "Corregido").length, [operarioRecords]);

    const filteredOperarioRecords = useMemo(() => {
        const text = operarioHistorySearch.trim().toLowerCase();
        if (!text) return operarioRecords;
        return operarioRecords.filter((r) => [r.sku, r.description, r.location, r.status, formatDateTime(r.counted_at)].join(" ").toLowerCase().includes(text));
    }, [operarioRecords, operarioHistorySearch]);

    const filteredRecords = useMemo(() => {
        return records.filter((r) => {
            const roleOk = canValidate(user) ? true : r.user_id === user?.id;
            const text = [r.sku, r.description, r.location, r.user_name, r.barcode || "", r.validator_name || ""].join(" ").toLowerCase();
            const textOk = text.includes(searchText.toLowerCase());
            const statusOk = statusFilter === "todos" ? true : r.status.toLowerCase() === statusFilter;
            return roleOk && textOk && statusOk;
        });
    }, [records, user, searchText, statusFilter]);

    const inventoryValue = totalInventoryValue;

    const countedValue = useMemo(() => {
        const skuMap = new Map<string, { cost: number; qty: number }>();
        for (const r of records) {
            const key = normalizeText(r.sku);
            if (!skuMap.has(key)) { skuMap.set(key, { cost: Number(r.cost || 0), qty: 0 }); }
            skuMap.get(key)!.qty += Number(r.counted_quantity || 0);
        }
        let total = 0;
        for (const item of skuMap.values()) total += item.cost * item.qty;
        return total;
    }, [records]);

    const valorizadoPct = useMemo(() => {
        if (!inventoryValue) return 0;
        return Number(((countedValue / inventoryValue) * 100).toFixed(2));
    }, [countedValue, inventoryValue]);

    const speedByUser = useMemo(() => {
        const map = new Map<string, { user: string; recordCount: number; units: number; firstAt: number | null; lastAt: number | null }>();
        records.forEach((r) => {
            const key = r.user_name || "Sin nombre";
            const current = map.get(key) || { user: key, recordCount: 0, units: 0, firstAt: null, lastAt: null };
            const time = new Date(r.counted_at).getTime();
            current.recordCount += 1;
            current.units += Number(r.counted_quantity || 0);
            if (!Number.isNaN(time)) {
                if (current.firstAt === null || time < current.firstAt) current.firstAt = time;
                if (current.lastAt === null || time > current.lastAt) current.lastAt = time;
            }
            map.set(key, current);
        });
        return Array.from(map.values()).map((x) => {
            let minutes = 1;
            if (x.firstAt !== null && x.lastAt !== null && x.lastAt >= x.firstAt) { minutes = Math.max(1, (x.lastAt - x.firstAt) / 60000); }
            return { user: x.user, skuCount: x.recordCount, units: x.units, minutes: Number(minutes.toFixed(2)), skuPerMin: Number((x.recordCount / minutes).toFixed(2)), unitsPerMin: Number((x.units / minutes).toFixed(2)) };
        }).sort((a, b) => b.skuPerMin - a.skuPerMin);
    }, [records]);

    const maxBarValue = useMemo(() => {
        if (!speedByUser.length) return 1;
        return Math.max(...speedByUser.flatMap((x) => [x.skuPerMin, x.unitsPerMin]), 1);
    }, [speedByUser]);

    const filteredAudit = useMemo(() => {
        return auditByCode.filter((row) => {
            const textOk = auditSearchText.trim()
                ? [row.sku, row.description].join(" ").toLowerCase().includes(auditSearchText.toLowerCase())
                : true;
            const statusOk = auditStatusFilter === "todos"
                ? true
                : auditStatusFilter === "no_contado"
                ? row.status_resumen === "NO CONTADO"
                : row.status_resumen.toLowerCase() === auditStatusFilter.toLowerCase();
            return textOk && statusOk;
        });
    }, [auditByCode, auditSearchText, auditStatusFilter]);

    const auditTotals = useMemo(() => {
        const totalValuedDiff = filteredAudit.reduce((s, r) => s + r.valued_difference, 0);
        const totalFaltantes = filteredAudit.filter((r) => r.status_resumen === "FALTANTE").length;
        const totalSobrantes = filteredAudit.filter((r) => r.status_resumen === "SOBRANTE").length;
        const totalOk = filteredAudit.filter((r) => r.status_resumen === "OK").length;
        const totalNoContado = filteredAudit.filter((r) => r.status_resumen === "NO CONTADO").length;
        return { totalValuedDiff, totalFaltantes, totalSobrantes, totalOk, totalNoContado };
    }, [filteredAudit]);

    function exportAudit() {
        type AuditExportRow = {
            SKU: string; DESCRIPCION: string; UNIDAD_MEDIDA: string; N_REGISTROS: number;
            STOCK_SISTEMA: number; STOCK_CONTADO: number; DIFERENCIA_UNIDAD: number;
            COSTO_UNITARIO: number; DIFERENCIA_VALORIZADA: number; STATUS: string;
        };
        const dataRows: AuditExportRow[] = filteredAudit.map((r) => ({
            SKU: r.sku, DESCRIPCION: r.description, UNIDAD_MEDIDA: r.unit, N_REGISTROS: r.record_count,
            STOCK_SISTEMA: r.system_stock, STOCK_CONTADO: r.total_counted, DIFERENCIA_UNIDAD: r.difference,
            COSTO_UNITARIO: r.cost, DIFERENCIA_VALORIZADA: r.valued_difference, STATUS: r.status_resumen,
        }));
        dataRows.push({
            SKU: "TOTAL", DESCRIPCION: `${filteredAudit.length} SKUs`, UNIDAD_MEDIDA: "",
            N_REGISTROS: filteredAudit.reduce((s, r) => s + r.record_count, 0),
            STOCK_SISTEMA: filteredAudit.reduce((s, r) => s + r.system_stock, 0),
            STOCK_CONTADO: filteredAudit.reduce((s, r) => s + r.total_counted, 0),
            DIFERENCIA_UNIDAD: filteredAudit.reduce((s, r) => s + r.difference, 0),
            COSTO_UNITARIO: 0, DIFERENCIA_VALORIZADA: auditTotals.totalValuedDiff,
            STATUS: `OK:${auditTotals.totalOk} | FALTANTE:${auditTotals.totalFaltantes} | SOBRANTE:${auditTotals.totalSobrantes} | NO CONTADO:${auditTotals.totalNoContado}`,
        });
        const ws = XLSX.utils.json_to_sheet(dataRows);
        ws["!cols"] = [
            { wch: 18 }, { wch: 40 }, { wch: 12 }, { wch: 12 },
            { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 16 },
            { wch: 22 }, { wch: 12 },
        ];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Resumen Auditoria");
        XLSX.writeFile(wb, `resumen_auditoria_${currentInventory?.name || "inventario"}.xlsx`);
    }

    function generateAndOpenReport() {
        // Determinar fecha de auditoría: última fecha de counted_at de los registros, o hoy
        let auditDate = new Date();
        if (records.length > 0) {
            const lastRecord = records.reduce((latest, r) => {
                const d = new Date(r.counted_at);
                return d > latest ? d : latest;
            }, new Date(records[0].counted_at));
            auditDate = lastRecord;
        }
        const dateStr = auditDate.toLocaleDateString("es-PE", { year: "numeric", month: "long", day: "numeric" });
        const invName = currentInventory?.name || "Inventario";

        const totalOk = auditByCode.filter(r => r.status_resumen === "OK").length;
        const totalFaltantes = auditByCode.filter(r => r.status_resumen === "FALTANTE").length;
        const totalSobrantes = auditByCode.filter(r => r.status_resumen === "SOBRANTE").length;
        const totalNoContado = auditByCode.filter(r => r.status_resumen === "NO CONTADO").length;
        const totalValuedDiff = auditByCode.reduce((s, r) => s + r.valued_difference, 0);
        const avancePct = skuProgress.pct;

        // Top 10 mayores diferencias valorizadas negativas (faltantes)
        const top10Faltantes = [...auditByCode]
            .filter(r => r.valued_difference < 0)
            .sort((a, b) => a.valued_difference - b.valued_difference)
            .slice(0, 10);

        // Top 10 sobrantes
        const top10Sobrantes = [...auditByCode]
            .filter(r => r.valued_difference > 0)
            .sort((a, b) => b.valued_difference - a.valued_difference)
            .slice(0, 10);

        const barWidth = (val: number, max: number) => max === 0 ? 0 : Math.round(Math.abs(val) / max * 100);
        const maxFalt = top10Faltantes.length > 0 ? Math.abs(top10Faltantes[0].valued_difference) : 1;
        const maxSobr = top10Sobrantes.length > 0 ? top10Sobrantes[0].valued_difference : 1;

        const statusPieTotal = totalOk + totalFaltantes + totalSobrantes + totalNoContado || 1;
        const okPct = Math.round(totalOk / statusPieTotal * 100);
        const faltPct = Math.round(totalFaltantes / statusPieTotal * 100);
        const sobrPct = Math.round(totalSobrantes / statusPieTotal * 100);
        const ncPct = Math.round(totalNoContado / statusPieTotal * 100);

        const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Informe de Auditoría — ${invName}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, Helvetica, sans-serif; background: #f1f5f9; color: #1e293b; }
  .container { max-width: 860px; margin: 0 auto; background: #fff; }
  /* HEADER */
  .header { background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%); color: #fff; padding: 36px 40px 28px; }
  .header h1 { font-size: 22px; font-weight: 800; letter-spacing: -0.5px; margin-bottom: 4px; }
  .header .sub { font-size: 13px; color: #94a3b8; margin-bottom: 18px; }
  .header-meta { display: flex; gap: 24px; flex-wrap: wrap; }
  .header-meta .item { font-size: 12px; color: #cbd5e1; }
  .header-meta .item strong { color: #fff; display: block; font-size: 13px; }
  /* KPI GRID */
  .kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 0; border-bottom: 1px solid #e2e8f0; }
  .kpi { padding: 20px 16px; border-right: 1px solid #e2e8f0; text-align: center; }
  .kpi:last-child { border-right: none; }
  .kpi .val { font-size: 28px; font-weight: 800; margin-bottom: 4px; }
  .kpi .label { font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi.green .val { color: #16a34a; } .kpi.red .val { color: #dc2626; }
  .kpi.blue .val { color: #2563eb; } .kpi.orange .val { color: #ea580c; }
  /* SECTION */
  .section { padding: 28px 40px; border-bottom: 1px solid #f1f5f9; }
  .section h2 { font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; }
  /* AVANCE */
  .progress-bar-bg { background: #e2e8f0; border-radius: 99px; height: 18px; overflow: hidden; margin-bottom: 6px; }
  .progress-bar-fill { height: 100%; border-radius: 99px; background: linear-gradient(90deg, #3b82f6, #1d4ed8); }
  .progress-label { font-size: 12px; color: #64748b; margin-bottom: 14px; }
  /* STATUS BARS */
  .status-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
  .status-row .label { width: 90px; font-size: 12px; font-weight: 600; color: #475569; }
  .status-row .bar-bg { flex: 1; background: #f1f5f9; border-radius: 4px; height: 22px; overflow: hidden; }
  .status-row .bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 8px; font-size: 11px; font-weight: 700; color: #fff; }
  .bar-ok { background: #16a34a; } .bar-falt { background: #dc2626; }
  .bar-sobr { background: #2563eb; } .bar-nc { background: #ea580c; }
  .status-row .count { font-size: 12px; font-weight: 700; width: 38px; text-align: right; }
  /* DIFFERENCE BANNER */
  .diff-banner { border-radius: 12px; padding: 18px 22px; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; }
  .diff-banner.neg { background: #fef2f2; border: 1px solid #fecaca; }
  .diff-banner.pos { background: #eff6ff; border: 1px solid #bfdbfe; }
  .diff-banner.zero { background: #f0fdf4; border: 1px solid #bbf7d0; }
  .diff-banner .amount { font-size: 26px; font-weight: 800; }
  .diff-banner.neg .amount { color: #dc2626; }
  .diff-banner.pos .amount { color: #2563eb; }
  .diff-banner.zero .amount { color: #16a34a; }
  .diff-banner .desc { font-size: 12px; color: #64748b; margin-top: 2px; }
  .diff-banner .emoji { font-size: 36px; }
  /* TOP TABLE */
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { background: #f8fafc; color: #475569; font-weight: 700; padding: 8px 10px; text-align: left; border-bottom: 2px solid #e2e8f0; }
  td { padding: 7px 10px; border-bottom: 1px solid #f1f5f9; color: #334155; }
  tr:last-child td { border-bottom: none; }
  .bar-mini-bg { background: #f1f5f9; border-radius: 3px; height: 8px; margin-top: 3px; }
  .bar-mini-fill { height: 100%; border-radius: 3px; }
  /* SIGN SECTION */
  .sign-section { padding: 28px 40px 36px; }
  .sign-section h2 { font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 20px; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; }
  .sign-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 24px; }
  .sign-box { text-align: center; }
  .sign-line { border-top: 2px solid #334155; margin-bottom: 8px; margin-top: 48px; }
  .sign-name { font-size: 13px; font-weight: 700; color: #0f172a; }
  .sign-role { font-size: 11px; color: #64748b; margin-top: 2px; }
  /* FOOTER */
  .footer { background: #0f172a; color: #64748b; font-size: 11px; text-align: center; padding: 14px; }
  @media print { body { background: #fff; } }
</style>
</head>
<body>
<div class="container">

  <!-- HEADER -->
  <div class="header">
    <h1>📊 Informe de Auditoría de Inventario</h1>
    <div class="sub">Generado automáticamente por WMS Conteo</div>
    <div class="header-meta">
      <div class="item"><strong>${reportStoreName || "—"}</strong>Nombre de tienda</div>
      <div class="item"><strong>${invName}</strong>Inventario</div>
      <div class="item"><strong>${dateStr}</strong>Fecha de auditoría</div>
      <div class="item"><strong>${reportAuditorName || user?.full_name || "—"}</strong>Auditor</div>
    </div>
  </div>

  <!-- KPI -->
  <div class="kpi-grid">
    <div class="kpi green"><div class="val">${totalOk}</div><div class="label">SKUs OK</div></div>
    <div class="kpi red"><div class="val">${totalFaltantes}</div><div class="label">Faltantes</div></div>
    <div class="kpi blue"><div class="val">${totalSobrantes}</div><div class="label">Sobrantes</div></div>
    <div class="kpi orange"><div class="val">${totalNoContado}</div><div class="label">No contados</div></div>
  </div>

  <!-- AVANCE -->
  <div class="section">
    <h2>📈 Avance del conteo</h2>
    <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${avancePct}%"></div></div>
    <div class="progress-label">${avancePct}% completado — ${skuProgress.counted} de ${skuProgress.total} SKUs contados</div>
    <div class="status-row">
      <div class="label">OK</div>
      <div class="bar-bg"><div class="bar-fill bar-ok" style="width:${barWidth(totalOk, statusPieTotal)}%">${okPct}%</div></div>
      <div class="count" style="color:#16a34a">${totalOk}</div>
    </div>
    <div class="status-row">
      <div class="label">Faltantes</div>
      <div class="bar-bg"><div class="bar-fill bar-falt" style="width:${barWidth(totalFaltantes, statusPieTotal)}%">${faltPct}%</div></div>
      <div class="count" style="color:#dc2626">${totalFaltantes}</div>
    </div>
    <div class="status-row">
      <div class="label">Sobrantes</div>
      <div class="bar-bg"><div class="bar-fill bar-sobr" style="width:${barWidth(totalSobrantes, statusPieTotal)}%">${sobrPct}%</div></div>
      <div class="count" style="color:#2563eb">${totalSobrantes}</div>
    </div>
    <div class="status-row">
      <div class="label">No contados</div>
      <div class="bar-bg"><div class="bar-fill bar-nc" style="width:${barWidth(totalNoContado, statusPieTotal)}%">${ncPct}%</div></div>
      <div class="count" style="color:#ea580c">${totalNoContado}</div>
    </div>
  </div>

  <!-- DIFERENCIA VALORIZADA -->
  <div class="section">
    <h2>💰 Diferencia valorizada</h2>
    <div class="diff-banner ${totalValuedDiff < 0 ? "neg" : totalValuedDiff > 0 ? "pos" : "zero"}">
      <div>
        <div class="amount">${formatMoney(totalValuedDiff)}</div>
        <div class="desc">Diferencia valorizada total (contado vs sistema)</div>
      </div>
      <div class="emoji">${totalValuedDiff < 0 ? "📉" : totalValuedDiff > 0 ? "📈" : "✅"}</div>
    </div>
  </div>

  <!-- TOP FALTANTES -->
  ${top10Faltantes.length > 0 ? `
  <div class="section">
    <h2>⚠️ Top faltantes por valor</h2>
    <table>
      <thead><tr><th>SKU</th><th>Descripción</th><th>Sist.</th><th>Contado</th><th>Dif.</th><th>Dif. Valor.</th><th style="width:80px">Impacto</th></tr></thead>
      <tbody>
        ${top10Faltantes.map(r => `
        <tr>
          <td><strong>${r.sku}</strong></td>
          <td>${r.description}</td>
          <td>${r.system_stock}</td>
          <td>${r.total_counted}</td>
          <td style="color:#dc2626;font-weight:700">${r.difference}</td>
          <td style="color:#dc2626;font-weight:700">${formatMoney(r.valued_difference)}</td>
          <td><div class="bar-mini-bg"><div class="bar-mini-fill bar-falt" style="width:${barWidth(r.valued_difference, maxFalt)}%"></div></div></td>
        </tr>`).join("")}
      </tbody>
    </table>
  </div>` : ""}

  <!-- TOP SOBRANTES -->
  ${top10Sobrantes.length > 0 ? `
  <div class="section">
    <h2>📦 Top sobrantes por valor</h2>
    <table>
      <thead><tr><th>SKU</th><th>Descripción</th><th>Sist.</th><th>Contado</th><th>Dif.</th><th>Dif. Valor.</th><th style="width:80px">Impacto</th></tr></thead>
      <tbody>
        ${top10Sobrantes.map(r => `
        <tr>
          <td><strong>${r.sku}</strong></td>
          <td>${r.description}</td>
          <td>${r.system_stock}</td>
          <td>${r.total_counted}</td>
          <td style="color:#2563eb;font-weight:700">+${r.difference}</td>
          <td style="color:#2563eb;font-weight:700">${formatMoney(r.valued_difference)}</td>
          <td><div class="bar-mini-bg"><div class="bar-mini-fill bar-sobr" style="width:${barWidth(r.valued_difference, maxSobr)}%"></div></div></td>
        </tr>`).join("")}
      </tbody>
    </table>
  </div>` : ""}

  <!-- FIRMAS -->
  <div class="sign-section">
    <h2>✍️ Firmas de conformidad</h2>
    <div class="sign-grid">
      <div class="sign-box">
        <div class="sign-line"></div>
        <div class="sign-name">${reportAuditorName || user?.full_name || "Auditor"}</div>
        <div class="sign-role">Auditor</div>
      </div>
      <div class="sign-box">
        <div class="sign-line"></div>
        <div class="sign-name">${reportStoreLeader || "Líder de tienda"}</div>
        <div class="sign-role">Líder de Tienda</div>
      </div>
      <div class="sign-box">
        <div class="sign-line"></div>
        <div class="sign-name">${reportWarehouseAdvisor || "Asesor de almacén"}</div>
        <div class="sign-role">Asesor de Almacén</div>
      </div>
    </div>
  </div>

  <div class="footer">WMS Conteo — Informe generado el ${new Date().toLocaleString("es-PE")} | ${invName}</div>
</div>
</body>
</html>`;

        const blob = new Blob([html], { type: "text/html;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `informe_auditoria_${currentInventory?.name || "inventario"}_${auditDate.toISOString().slice(0,10)}.html`;
        a.click();
        URL.revokeObjectURL(url);
        setShowReportModal(false);
    }

    const showSystemStock = canSeeSystemStock(user);
    const showCost = canSeeCost(user);
    const showValuedDiff = canSeeValuedDifference(user);

    const editingIsAdmin = editUserRoles.includes("Administrador");
    const newIsAdmin = newRoles.includes("Administrador");

    // ── SOLO OPERARIO — Vista WMS móvil ─────────────────────────────────────
    const isOnlyOperario = user && hasRole(user, "Operario") && !canValidate(user);

    // ── SIDEBAR NAV ITEMS (PC/Desktop) ─────────────────────────────────────────
    const sidebarItems: SidebarItem[] = !user ? [] : ([
        { key: "operario"  as TabKey, label: "Operario",      icon: "📦", show: canCount(user) },
        { key: "validador" as TabKey, label: "Validador",     icon: "✅", show: canValidate(user) },
        { key: "maestro"   as TabKey, label: "Maestro",       icon: "📋", show: canValidate(user) },
        { key: "admin"     as TabKey, label: "Administrador", icon: "🔧", show: hasRole(user, "Administrador") },
    ] as SidebarItem[]).filter(i => i.show);

    if (!user) return null;

    return (
        <>
        {/* ══════════════════════════════════════════════════════════════════════
            VISTA OPERARIO PURO — Layout móvil WMS (sin sidebar)
        ══════════════════════════════════════════════════════════════════════ */}
        {isOnlyOperario && (
        <main className="min-h-screen bg-slate-100 p-4 md:p-6">
            <div className="max-w-7xl mx-auto space-y-6">

                {/* ── HEADER OPERARIO MÓVIL (WMS style) ──────────────────────── */}
                <header className="bg-slate-900 text-white rounded-2xl shadow-lg overflow-hidden">
                    {/* Barra superior compacta */}
                    <div className="flex items-center justify-between px-4 py-3 border-b border-white/10">
                        <div className="flex items-center gap-2">
                            <div className="w-2 h-2 rounded-full bg-green-400 animate-pulse" />
                            <span className="text-xs font-semibold text-slate-300 uppercase tracking-wider">WMS Conteo</span>
                        </div>
                        <div className="flex items-center gap-2">
                            <span className={`text-xs px-2 py-0.5 rounded-full font-semibold ${isOnline ? "bg-green-500/20 text-green-300" : "bg-red-500/20 text-red-300"}`}>
                                {isOnline ? "● En línea" : "● Sin conexión"}
                            </span>
                            <button onClick={logout} className="text-xs text-slate-400 hover:text-white font-medium px-2 py-1 rounded-lg hover:bg-white/10 transition">
                                Salir
                            </button>
                        </div>
                    </div>
                    {/* Info del operario */}
                    <div className="px-4 py-3 flex items-center justify-between">
                        <div>
                            <div className="font-bold text-base leading-tight">{user.full_name}</div>
                            <div className="text-xs text-slate-400 mt-0.5">{currentInventory?.name || "Cargando..."}</div>
                        </div>
                        {pendingCount > 0 && (
                            <button
                                onClick={syncPendingRecords}
                                disabled={syncing || isOnline === false}
                                className="flex items-center gap-2 bg-amber-500/20 border border-amber-400/30 text-amber-300 text-xs font-semibold px-3 py-2 rounded-xl"
                            >
                                {syncing ? "⏫ Subiendo..." : `⚠ ${pendingCount} pendiente${pendingCount > 1 ? "s" : ""}`}
                            </button>
                        )}
                    </div>

                    {/* Mis registros de hoy */}
                    <div className="px-4 pb-3 flex gap-3">
                        <div className="flex-1 bg-white/5 rounded-xl px-3 py-2 text-center">
                            <div className="text-xs text-slate-400">Mis registros</div>
                            <div className="text-xl font-bold">{operarioRecords.length}</div>
                        </div>
                        <div className="flex-1 bg-white/5 rounded-xl px-3 py-2 text-center">
                            <div className="text-xs text-slate-400">Validados</div>
                            <div className="text-xl font-bold text-green-400">{_operarioValidatedCount}</div>
                        </div>
                        <div className="flex-1 bg-white/5 rounded-xl px-3 py-2 text-center">
                            <div className="text-xs text-slate-400">Offline</div>
                            <div className={`text-xl font-bold ${pendingCount > 0 ? "text-amber-400" : "text-slate-400"}`}>{pendingCount}</div>
                        </div>
                    </div>

                    {/* Sub-tabs WMS */}
                    <div className="flex border-t border-white/10">
                        <button
                            onClick={() => setOperarioSubTab("conteo")}
                            className={`flex-1 py-3 text-sm font-semibold transition-colors ${operarioSubTab === "conteo" ? "bg-white text-slate-900" : "text-slate-400 hover:text-white"}`}
                        >
                            📦 CONTEO
                        </button>
                        <button
                            onClick={() => setOperarioSubTab("historial")}
                            className={`flex-1 py-3 text-sm font-semibold transition-colors ${operarioSubTab === "historial" ? "bg-white text-slate-900" : "text-slate-400 hover:text-white"}`}
                        >
                            📋 MIS REGISTROS
                        </button>
                    </div>
                </header>

                {/* Mensaje global */}
                {message && (
                    <div className={`rounded-2xl p-4 shadow text-sm border flex items-start justify-between gap-3 ${messageType === "success" ? "bg-green-50 border-green-200 text-green-800" : messageType === "error" ? "bg-red-50 border-red-200 text-red-800" : "bg-white border-slate-200 text-slate-800"}`}>
                        <span>{message}</span>
                        <button className="text-xs opacity-60 hover:opacity-100 shrink-0" onClick={clearMessage}>✕</button>
                    </div>
                )}

                {loading && (
                    <div className="bg-white rounded-2xl p-4 shadow text-sm border border-slate-200">Cargando información del inventario...</div>
                )}

                {/* ══════════════════════════════════════════════════════════════
                    TAB OPERARIO — VISTA WMS MÓVIL (solo operarios puros)
                ══════════════════════════════════════════════════════════════ */}
                <>
                        {/* ── SUB-TAB: CONTEO ─────────────────────────────────── */}
                        {operarioSubTab === "conteo" && (
                            <div className="space-y-3">
                                {/* Banner sin conexión */}
                                {!isOnline && (
                                    <div className="bg-amber-50 border border-amber-300 rounded-2xl px-4 py-3 flex items-center gap-3">
                                        <span className="text-amber-600 text-lg">📡</span>
                                        <div>
                                            <div className="font-semibold text-amber-800 text-sm">Modo sin conexión</div>
                                            <div className="text-xs text-amber-600">Los conteos se guardan localmente y se subirán al recuperar internet.</div>
                                        </div>
                                    </div>
                                )}

                                {/* Paso 1 — Producto */}
                                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                                    <div className="bg-slate-800 text-white px-4 py-2.5 flex items-center gap-2">
                                        <span className="w-6 h-6 rounded-full bg-white/20 flex items-center justify-center text-xs font-bold">1</span>
                                        <span className="font-semibold text-sm uppercase tracking-wide">Producto / Código de barra</span>
                                    </div>
                                    <div className="p-4 space-y-3">
                                        <div className="relative">
                                            <input
                                                className="w-full border-2 rounded-xl p-3 pr-14 text-slate-900 bg-white text-base focus:border-slate-700 focus:outline-none"
                                                value={searchValue}
                                                onChange={(e) => handleSearchInputChange(e.target.value)}
                                                placeholder="SKU o descripción..."
                                                autoFocus
                                            />
                                            <button type="button" onClick={() => openScanner("product")}
                                                className="absolute right-2 top-1/2 -translate-y-1/2 h-10 w-10 rounded-xl bg-slate-900 text-white flex items-center justify-center shadow">
                                                <QrCode size={20} />
                                            </button>
                                        </div>

                                        {/* Lista de resultados */}
                                        {searchResults.length > 0 && (
                                            <div className="rounded-xl border border-slate-200 overflow-hidden">
                                                {searchResults.map((p) => (
                                                    <button key={p.id} type="button"
                                                        className="w-full text-left px-4 py-3 border-b last:border-b-0 hover:bg-slate-50 active:bg-slate-100"
                                                        onClick={() => { setSelectedProduct(p); setSearchResults([]); showMessage("Producto seleccionado correctamente.", "success"); }}>
                                                        <div className="font-bold text-slate-900">{p.sku}</div>
                                                        <div className="text-sm text-slate-600">{p.description}</div>
                                                        <div className="text-xs text-slate-400">UM: {p.unit || "-"}</div>
                                                    </button>
                                                ))}
                                            </div>
                                        )}

                                        {/* Producto seleccionado */}
                                        {selectedProduct && (
                                            <div className="rounded-xl bg-green-50 border-2 border-green-400 p-4">
                                                <div className="flex items-start justify-between gap-2">
                                                    <div className="space-y-1 min-w-0">
                                                        <div className="font-bold text-green-900 text-lg leading-tight">{selectedProduct.sku}</div>
                                                        <div className="text-sm text-green-800 font-medium leading-snug">{selectedProduct.description}</div>
                                                        <div className="text-xs text-green-700 bg-green-100 inline-block px-2 py-0.5 rounded-full">{selectedProduct.unit || "—"}</div>
                                                    </div>
                                                    <span className="text-green-500 text-2xl shrink-0">✓</span>
                                                </div>
                                                {showSystemStock && (
                                                    <div className="mt-2 pt-2 border-t border-green-200 text-xs text-green-700">
                                                        Stock sistema: <span className="font-bold">{selectedProduct.system_stock}</span>
                                                    </div>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                </div>

                                {/* Paso 2 — Ubicación */}
                                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                                    <div className="bg-slate-700 text-white px-4 py-2.5 flex items-center gap-2">
                                        <span className="w-6 h-6 rounded-full bg-white/20 flex items-center justify-center text-xs font-bold">2</span>
                                        <span className="font-semibold text-sm uppercase tracking-wide">Ubicación</span>
                                    </div>
                                    <div className="p-4">
                                        <div className="relative">
                                            <input
                                                className="w-full border-2 rounded-xl p-3 pr-14 text-slate-900 bg-white text-base focus:border-slate-700 focus:outline-none"
                                                value={location}
                                                onChange={(e) => setLocation(e.target.value)}
                                                placeholder="Ej: A-01-02"
                                            />
                                            <button type="button" onClick={() => openScanner("location")}
                                                className="absolute right-2 top-1/2 -translate-y-1/2 h-10 w-10 rounded-xl bg-slate-600 text-white flex items-center justify-center shadow">
                                                <QrCode size={20} />
                                            </button>
                                        </div>
                                    </div>
                                </div>

                                {/* Paso 3 — Cantidad */}
                                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                                    <div className="bg-slate-600 text-white px-4 py-2.5 flex items-center gap-2">
                                        <span className="w-6 h-6 rounded-full bg-white/20 flex items-center justify-center text-xs font-bold">3</span>
                                        <span className="font-semibold text-sm uppercase tracking-wide">Cantidad física</span>
                                    </div>
                                    <div className="p-4">
                                        <input
                                            className="w-full border-2 rounded-xl p-3 text-slate-900 bg-white text-2xl font-bold text-center focus:border-slate-700 focus:outline-none"
                                            type="number"
                                            inputMode="numeric"
                                            value={quantity}
                                            onChange={(e) => setQuantity(e.target.value)}
                                            placeholder="0"
                                        />
                                    </div>
                                </div>

                                {/* Botón guardar */}
                                <button
                                    onClick={saveCount}
                                    className="w-full py-5 rounded-2xl font-bold text-lg shadow-lg transition-all active:scale-95 bg-slate-900 text-white"
                                >
                                    {isOnline ? "✓ GUARDAR CONTEO" : "📦 GUARDAR OFFLINE"}
                                </button>

                                {/* Botón limpiar */}
                                <button
                                    onClick={() => { setSearchValue(""); setSelectedProduct(null); setSearchResults([]); setLocation(""); setQuantity(""); setMessage(""); }}
                                    className="w-full py-3 rounded-2xl font-semibold text-slate-700 border-2 border-slate-300 bg-white transition-all active:scale-95"
                                >
                                    Limpiar campos
                                </button>

                                {/* Sincronizar manualmente si hay pendientes y hay conexión */}
                                {pendingCount > 0 && isOnline && (
                                    <button
                                        onClick={syncPendingRecords}
                                        disabled={syncing}
                                        className="w-full py-3 rounded-2xl font-semibold text-amber-800 border-2 border-amber-400 bg-amber-50 transition-all active:scale-95"
                                    >
                                        {syncing ? "⏫ Sincronizando..." : `⬆ Subir ${pendingCount} registro${pendingCount > 1 ? "s" : ""} offline`}
                                    </button>
                                )}
                            </div>
                        )}

                        {/* ── SUB-TAB: MIS REGISTROS ──────────────────────────── */}
                        {operarioSubTab === "historial" && (
                            <div className="space-y-3">
                                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                                    <div className="bg-slate-800 text-white px-4 py-3 flex items-center justify-between">
                                        <span className="font-semibold text-sm uppercase tracking-wide">Mis registros</span>
                                        <span className="text-xs bg-white/20 px-2 py-0.5 rounded-full">{filteredOperarioRecords.length}</span>
                                    </div>
                                    <div className="p-3">
                                        <input
                                            className="w-full border rounded-xl p-3 text-sm"
                                            value={operarioHistorySearch}
                                            onChange={(e) => setOperarioHistorySearch(e.target.value)}
                                            placeholder="Buscar por SKU, descripción, ubicación..."
                                        />
                                    </div>

                                    {filteredOperarioRecords.length === 0 ? (
                                        <div className="px-4 pb-6 text-center text-slate-400 text-sm">No hay registros todavía.</div>
                                    ) : (
                                        <div className="divide-y divide-slate-100">
                                            {filteredOperarioRecords.map((r) => (
                                                <div key={r.id} className="px-4 py-3 flex items-start justify-between gap-3">
                                                    <div className="min-w-0 flex-1">
                                                        <div className="flex items-center gap-2 flex-wrap">
                                                            <span className="font-bold text-slate-900 text-sm">{r.sku}</span>
                                                            <span className={statusBadge(r.status)}>{r.status}</span>
                                                        </div>
                                                        <div className="text-xs text-slate-500 truncate mt-0.5">{r.description}</div>
                                                        <div className="flex items-center gap-3 mt-1 text-xs text-slate-600">
                                                            <span className="font-semibold">Cant: {r.counted_quantity}</span>
                                                            {r.location && <span>📍 {r.location}</span>}
                                                        </div>
                                                        <div className="text-xs text-slate-400 mt-0.5">{formatDateTime(r.counted_at)}</div>
                                                    </div>
                                                    <button
                                                        onClick={() => openEdit(r)}
                                                        className="shrink-0 px-3 py-2 rounded-xl border text-xs font-semibold hover:bg-slate-50 active:bg-slate-100"
                                                    >
                                                        Editar
                                                    </button>
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </div>
                            </div>
                        )}
                    </>

            </div>
        </main>
        )}

        {/* ══════════════════════════════════════════════════════════════════════
            VISTA ADMIN / VALIDADOR — Layout WMS Desktop con Sidebar
        ══════════════════════════════════════════════════════════════════════ */}
        {!isOnlyOperario && (
        <div className="min-h-screen bg-slate-100 flex">

            {/* ── SIDEBAR FIJO ─────────────────────────────────────────────── */}
            <aside className="hidden lg:flex flex-col w-56 bg-slate-950 text-white fixed top-0 left-0 h-screen z-30 shadow-2xl">
                {/* Logo / Brand */}
                <div className="px-5 py-5 border-b border-white/10">
                    <div className="flex items-center gap-2 mb-1">
                        <div className="w-2 h-2 rounded-full bg-green-400 animate-pulse" />
                        <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">WMS</span>
                    </div>
                    <div className="text-base font-bold text-white leading-tight">Sistema Conteo</div>
                    <div className="text-xs text-slate-400 mt-0.5 truncate">{currentInventory?.name || "—"}</div>
                </div>

                {/* Inventario selector */}
                {canAccessAnyInventory(user) && (
                    <div className="px-4 py-3 border-b border-white/10">
                        <div className="text-xs text-slate-500 mb-1.5 uppercase tracking-wider">Inventario</div>
                        <select
                            className="w-full rounded-lg border border-white/10 bg-white/10 text-white text-xs px-2 py-2 focus:outline-none focus:border-white/30"
                            value={selectedInventoryId}
                            onChange={(e) => handleInventoryChange(e.target.value)}
                        >
                            {inventories.map((inv) => (<option key={inv.id} value={inv.id} className="text-slate-900">{inv.name}</option>))}
                        </select>
                    </div>
                )}

                {/* Nav Items */}
                <nav className="flex-1 py-4 space-y-1 px-3">
                    {sidebarItems.map(item => (
                        <button
                            key={item.key}
                            onClick={() => setActiveTab(item.key)}
                            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition-all ${
                                activeTab === item.key
                                    ? "bg-white text-slate-900 shadow"
                                    : "text-slate-400 hover:text-white hover:bg-white/10"
                            }`}
                        >
                            <span className="text-base w-5 text-center">{item.icon}</span>
                            <span>{item.label}</span>
                            {activeTab === item.key && <span className="ml-auto w-1.5 h-1.5 rounded-full bg-slate-900" />}
                        </button>
                    ))}
                </nav>

                {/* Footer del sidebar */}
                <div className="px-4 py-4 border-t border-white/10 space-y-3">
                    {/* Stats rápidos */}
                    <div className="grid grid-cols-2 gap-2">
                        <div className="bg-white/5 rounded-lg px-2 py-2 text-center">
                            <div className="text-lg font-bold text-white">{records.length}</div>
                            <div className="text-xs text-slate-500">Registros</div>
                        </div>
                        <div className="bg-white/5 rounded-lg px-2 py-2 text-center">
                            <div className="text-lg font-bold text-white">{totalProductCount.toLocaleString()}</div>
                            <div className="text-xs text-slate-500">Productos</div>
                        </div>
                    </div>
                    {/* Usuario info */}
                    <div className="flex items-center gap-2">
                        <div className="w-7 h-7 rounded-full bg-white/10 flex items-center justify-center text-xs font-bold text-white shrink-0">
                            {user.full_name.charAt(0).toUpperCase()}
                        </div>
                        <div className="min-w-0">
                            <div className="text-xs font-semibold text-white truncate">{user.full_name}</div>
                            <div className="text-xs text-slate-500">{user.roles?.join(", ") || user.role}</div>
                        </div>
                    </div>
                    <button
                        onClick={logout}
                        className="w-full py-2 rounded-xl bg-white/10 text-white text-xs font-semibold hover:bg-red-600/30 hover:text-red-300 transition"
                    >
                        → Cerrar sesión
                    </button>
                </div>
            </aside>

            {/* ── MOBILE TOPBAR (visible solo en <lg) ─────────────────────── */}
            <div className="lg:hidden fixed top-0 left-0 right-0 z-20 bg-slate-950 text-white px-4 py-3 flex items-center justify-between shadow-lg">
                <div>
                    <div className="text-xs text-slate-400 uppercase tracking-wider">WMS Conteo</div>
                    <div className="text-sm font-bold">{currentInventory?.name || "—"}</div>
                </div>
                <div className="flex items-center gap-2">
                    {canAccessAnyInventory(user) && (
                        <select
                            className="rounded-lg border border-white/20 bg-white/10 text-white text-xs px-2 py-1.5 max-w-[120px]"
                            value={selectedInventoryId}
                            onChange={(e) => handleInventoryChange(e.target.value)}
                        >
                            {inventories.map((inv) => (<option key={inv.id} value={inv.id} className="text-slate-900">{inv.name}</option>))}
                        </select>
                    )}
                    <button className="text-xs bg-white/10 px-3 py-1.5 rounded-lg font-semibold" onClick={logout}>Salir</button>
                </div>
            </div>

            {/* ── MOBILE BOTTOM NAV (visible solo en <lg) ──────────────────── */}
            <div className="lg:hidden fixed bottom-0 left-0 right-0 z-20 bg-slate-950 border-t border-white/10 flex">
                {sidebarItems.map(item => (
                    <button
                        key={item.key}
                        onClick={() => setActiveTab(item.key)}
                        className={`flex-1 flex flex-col items-center py-2.5 text-xs font-semibold transition-colors ${
                            activeTab === item.key ? "text-white" : "text-slate-500"
                        }`}
                    >
                        <span className="text-lg">{item.icon}</span>
                        <span className="text-[10px] mt-0.5">{item.label}</span>
                    </button>
                ))}
            </div>

            {/* ── MAIN CONTENT AREA ─────────────────────────────────────────── */}
            <div className="flex-1 lg:ml-56 min-h-screen flex flex-col">

                {/* Top bar desktop */}
                <header className="hidden lg:flex items-center justify-between px-8 py-4 bg-white border-b border-slate-200 shadow-sm">
                    <div>
                        <div className="text-xs text-slate-400 uppercase tracking-widest font-semibold">
                            {sidebarItems.find(i => i.key === activeTab)?.icon} {sidebarItems.find(i => i.key === activeTab)?.label}
                        </div>
                        <h1 className="text-xl font-bold text-slate-900 leading-tight">
                            {activeTab === "operario" && "Módulo Operario"}
                            {activeTab === "validador" && "Módulo Validador"}
                            {activeTab === "maestro" && "Maestro de Productos"}
                            {activeTab === "admin" && "Administrador del Sistema"}
                        </h1>
                    </div>
                    <div className="flex items-center gap-4">
                        {/* KPI rápidos en topbar */}
                        <div className="flex items-center gap-3 text-sm">
                            <div className="text-center">
                                <div className="font-bold text-slate-900">{records.length}</div>
                                <div className="text-xs text-slate-400">Registros</div>
                            </div>
                            <div className="w-px h-8 bg-slate-200" />
                            <div className="text-center">
                                <div className="font-bold text-slate-900">{totalProductCount.toLocaleString()}</div>
                                <div className="text-xs text-slate-400">Productos</div>
                            </div>
                            <div className="w-px h-8 bg-slate-200" />
                            <div className="text-center">
                                <div className="font-bold text-slate-900">{users.length}</div>
                                <div className="text-xs text-slate-400">Usuarios</div>
                            </div>
                        </div>
                        {/* Online indicator */}
                        <span className={`text-xs px-2.5 py-1 rounded-full font-semibold ${isOnline ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                            {isOnline ? "● En línea" : "● Sin conexión"}
                        </span>
                    </div>
                </header>

                {/* Content */}
                <main className="flex-1 p-4 lg:p-7 space-y-6 mt-14 lg:mt-0 mb-16 lg:mb-0">

                {/* Mensaje global */}
                {message && (
                    <div className={`rounded-2xl p-4 shadow text-sm border flex items-start justify-between gap-3 ${messageType === "success" ? "bg-green-50 border-green-200 text-green-800" : messageType === "error" ? "bg-red-50 border-red-200 text-red-800" : "bg-white border-slate-200 text-slate-800"}`}>
                        <span>{message}</span>
                        <button className="text-xs opacity-60 hover:opacity-100 shrink-0" onClick={clearMessage}>✕</button>
                    </div>
                )}

                {loading && (
                    <div className="bg-white rounded-2xl p-4 shadow text-sm border border-slate-200">Cargando información del inventario...</div>
                )}

                {/* ══════════════════════════════════════════════════════════════
                    TAB OPERARIO — Vista para admin/validador con rol operario
                ══════════════════════════════════════════════════════════════ */}
                {!isOnlyOperario && activeTab === "operario" && canCount(user) && (
                    <>
                        <section className="bg-white rounded-3xl p-4 md:p-6 shadow space-y-4 md:space-y-6">
                            <div className="hidden md:block">
                                <h2 className="text-2xl font-bold text-slate-900">Módulo Operario</h2>
                                <p className="text-slate-600 mt-1">Inventario: <b>{currentInventory?.name || "-"}</b></p>
                            </div>
                            <div className="grid lg:grid-cols-3 gap-4">
                                <div className="lg:col-span-1">
                                    <label className="block font-semibold mb-2 text-slate-800">Código de barras / SKU</label>
                                    <div className="relative">
                                        <input
                                            className="w-full border rounded-2xl p-3 pr-12 text-slate-900 bg-white"
                                            value={searchValue}
                                            onChange={(e) => handleSearchInputChange(e.target.value)}
                                            placeholder="Digita SKU o descripción / Escanea Cód. Barra"
                                            autoFocus
                                        />
                                        <button type="button" onClick={() => openScanner("product")} className="absolute right-2 top-1/2 -translate-y-1/2 h-9 w-9 rounded-xl bg-slate-900 text-white text-sm flex items-center justify-center" title="Escanear producto">
                                            <QrCode size={18} />
                                        </button>
                                    </div>
                                </div>
                                <div>
                                    <label className="block font-semibold mb-2 text-slate-800">Ubicación</label>
                                    <div className="relative">
                                        <input
                                            className="w-full border rounded-2xl p-3 pr-14 text-slate-900 bg-white"
                                            value={location}
                                            onChange={(e) => setLocation(e.target.value)}
                                            placeholder="Escanea o escribe la ubicación"
                                        />
                                        <button type="button" onClick={() => openScanner("location")} className="absolute right-2 top-1/2 -translate-y-1/2 h-9 w-9 rounded-xl bg-slate-700 text-white text-sm flex items-center justify-center" title="Escanear ubicación">
                                            <QrCode size={18} />
                                        </button>
                                    </div>
                                </div>
                                <div>
                                    <label className="block font-semibold mb-2 text-slate-800">Cantidad física</label>
                                    <input className="w-full border rounded-2xl p-3 text-slate-900 bg-white" type="number" value={quantity} onChange={(e) => setQuantity(e.target.value)} placeholder="Ej. 12" />
                                </div>
                            </div>

                            {searchResults.length > 0 && (
                                <div className="rounded-2xl border border-slate-200 p-4 bg-white space-y-2">
                                    <div className="text-sm font-semibold text-slate-800">Selecciona un producto:</div>
                                    <div className="max-h-64 overflow-auto rounded-xl border border-slate-200">
                                        {searchResults.map((p) => (
                                            <button key={p.id} type="button" className="w-full text-left p-3 border-b last:border-b-0 hover:bg-slate-50" onClick={() => { setSelectedProduct(p); setSearchResults([]); showMessage("Producto seleccionado correctamente.", "success"); }}>
                                                <div className="font-semibold text-slate-900">{p.sku}</div>
                                                <div className="text-sm text-slate-600">{p.description}</div>
                                                <div className="text-xs text-slate-500">UM: {p.unit || "-"}</div>
                                            </button>
                                        ))}
                                    </div>
                                </div>
                            )}

                            {selectedProduct && (
                                <div className="rounded-2xl border border-green-200 p-5 bg-green-50">
                                    <div className="grid md:grid-cols-3 gap-4 text-sm">
                                        <div><div className="text-slate-500">Código del producto</div><div className="font-semibold text-slate-900">{selectedProduct.sku}</div></div>
                                        <div><div className="text-slate-500">Descripción</div><div className="font-semibold text-slate-900">{selectedProduct.description}</div></div>
                                        <div><div className="text-slate-500">Unidad de medida</div><div className="font-semibold text-slate-900">{selectedProduct.unit || "-"}</div></div>
                                        {showSystemStock && (
                                            <div><div className="text-slate-500">Stock sistema</div><div className="font-semibold text-slate-900">{selectedProduct.system_stock}</div></div>
                                        )}
                                        {showCost && (
                                            <div><div className="text-slate-500">Costo unitario</div><div className="font-semibold text-slate-900">{formatMoney(selectedProduct.cost)}</div></div>
                                        )}
                                    </div>
                                </div>
                            )}

                            <div className="flex flex-wrap gap-3">
                                <button className="px-5 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={saveCount}>Guardar conteo</button>
                                <button className="px-5 py-3 rounded-2xl border font-semibold" onClick={() => { setSearchValue(""); setSelectedProduct(null); setSearchResults([]); setLocation(""); setQuantity(""); setMessage(""); }}>Limpiar</button>
                            </div>
                        </section>

                        {/* Mis registros (vista no-operario) */}
                        <section className="bg-white rounded-3xl p-4 md:p-6 shadow space-y-4">
                            <div className="space-y-3">
                                <div>
                                    <h3 className="text-lg md:text-xl font-bold text-slate-900">Mis registros</h3>
                                    <p className="text-slate-600 text-sm mt-1">Aquí puedes revisar todos tus registros y buscar rápido por SKU, descripción o ubicación.</p>
                                </div>
                                <input className="w-full border rounded-2xl p-3 text-sm" value={operarioHistorySearch} onChange={(e) => setOperarioHistorySearch(e.target.value)} placeholder="Buscar en mis registros..." />
                            </div>
                            <div className="rounded-2xl border overflow-hidden">
                                <div className="max-h-[320px] overflow-auto">
                                    <table className="w-full text-sm">
                                        <thead className="bg-slate-100 sticky top-0">
                                            <tr>
                                                <th className="p-2 border">SKU</th>
                                                <th className="p-2 border">Desc.</th>
                                                <th className="p-2 border">Cant.</th>
                                                {showSystemStock && <th className="p-2 border">Stock Sis.</th>}
                                                {showSystemStock && <th className="p-2 border">DIF</th>}
                                                {showCost && <th className="p-2 border">Costo</th>}
                                                {showValuedDiff && <th className="p-2 border">Dif. Valoriz.</th>}
                                                <th className="p-2 border">Ubic.</th>
                                                <th className="p-2 border">Fecha</th>
                                                <th className="p-2 border">Estado</th>
                                                <th className="p-2 border">Acción</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filteredOperarioRecords.map((r) => (
                                                <tr key={r.id}>
                                                    <td className="p-2 border font-medium">{r.sku}</td>
                                                    <td className="p-2 border">{r.description}</td>
                                                    <td className="p-2 border">{r.counted_quantity}</td>
                                                    {showSystemStock && <td className="p-2 border">{r.system_stock}</td>}
                                                    {showSystemStock && <td className="p-2 border">{diffBadge(r.difference)}</td>}
                                                    {showCost && <td className="p-2 border">{formatMoney(r.cost)}</td>}
                                                    {showValuedDiff && <td className="p-2 border">{formatMoney(r.difference * r.cost)}</td>}
                                                    <td className="p-2 border">{r.location}</td>
                                                    <td className="p-2 border">{formatDateTime(r.counted_at)}</td>
                                                    <td className="p-2 border"><span className={statusBadge(r.status)}>{r.status}</span></td>
                                                    <td className="p-2 border">
                                                        <button className="px-3 py-2 rounded-lg border text-xs font-semibold" onClick={() => openEdit(r)}>Editar</button>
                                                    </td>
                                                </tr>
                                            ))}
                                            {filteredOperarioRecords.length === 0 && (
                                                <tr><td className="p-4 border text-center text-slate-500" colSpan={showSystemStock ? (showCost ? (showValuedDiff ? 11 : 10) : 9) : 7}>No hay registros todavía.</td></tr>
                                            )}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </section>
                    </>
                )}

                {/* ── TAB MAESTRO ────────────────────────────────────────────── */}
                {activeTab === "maestro" && canValidate(user) && (
                    <>
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div>
                                <h2 className="text-2xl font-bold text-slate-900">Maestro de productos</h2>
                                <p className="text-slate-600 mt-1">Inventario: <b>{currentInventory?.name || "-"}</b></p>
                                <p className="text-xs text-slate-500 mt-1">Columnas requeridas: <b>SKU</b>, <b>DESCRIPCION</b>. Opcionales: UNIDAD DE MEDIDA, COSTO, STOCK.</p>
                                <p className="text-xs text-indigo-600 mt-0.5">Los códigos de barra vienen del catálogo global — no necesitas incluirlos aquí.</p>
                            </div>
                            <div className="grid md:grid-cols-2 gap-6">
                                <div className="space-y-3">
                                    <input ref={masterInputRef} type="file" accept=".xlsx,.xls" onChange={(e) => { const f = e.target.files?.[0] || null; setMasterFile(f); setMasterFileName(f ? f.name : ""); }} />
                                    <div className="text-sm text-slate-500">{masterFileName ? `📄 ${masterFileName}` : "Ningún archivo seleccionado"}</div>
                                    {uploadProgress && (
                                        <div className="space-y-2">
                                            <div className="flex justify-between text-sm font-semibold text-slate-700"><span>{uploadProgress?.step}</span><span>{uploadProgress?.pct ?? 0}%</span></div>
                                            <div className="w-full bg-slate-200 rounded-full h-4 overflow-hidden"><div className="bg-slate-900 h-4 rounded-full transition-all duration-300" style={{ width: `${uploadProgress?.pct ?? 0}%` }} /></div>
                                            <p className="text-xs text-slate-500">No cierres ni recargues la página... Con 32k productos tarda aprox. 5-8 minutos.</p>
                                        </div>
                                    )}
                                    <button className={`px-4 py-3 rounded-2xl font-semibold text-white w-full ${uploadProgress ? "bg-slate-400 cursor-not-allowed" : "bg-slate-900"}`} type="button" onClick={uploadMaster} disabled={!!uploadProgress}>
                                        {uploadProgress ? "Cargando..." : "Insertar maestro"}
                                    </button>
                                </div>
                                <div className="bg-slate-50 rounded-2xl p-4 space-y-2">
                                    <div className="font-semibold text-slate-700 text-sm">Estado actual</div>
                                    <div className="text-3xl font-bold text-slate-900">{totalProductCount.toLocaleString()}</div>
                                    <div className="text-xs text-slate-500">productos cargados en este inventario</div>
                                    {totalProductCount > 0 && user?.role === "Administrador" && (
                                        <button className={`mt-3 px-4 py-2 rounded-xl text-white text-sm font-semibold w-full ${uploadProgress ? "bg-slate-400 cursor-not-allowed" : "bg-red-600 hover:bg-red-700"}`} onClick={deleteMaster} disabled={!!uploadProgress}>
                                            {uploadProgress ? "Procesando..." : "⚡ Eliminar todo el maestro"}
                                        </button>
                                    )}
                                </div>
                            </div>
                        </section>
                    </>
                )}

                {/* ── TAB VALIDADOR ────────────────────────────────────────────── */}
                {activeTab === "validador" && canValidate(user) && (
                    <>
                        {/* ── REGISTROS Y RESUMEN ─────────────────────────── */}
                        {true && (
                            <>
                        <section className="bg-white rounded-3xl p-6 shadow space-y-6">
                            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                                <div>
                                    <h2 className="text-2xl font-bold text-slate-900">Módulo Validador</h2>
                                    <p className="text-slate-600 mt-1">Revisión, validación, avance y productividad por inventario.</p>
                                </div>
                            </div>
                            <div className="grid md:grid-cols-2 xl:grid-cols-4 gap-4">
                                <div className="rounded-2xl border p-5 bg-slate-50">
                                    <div className="text-sm text-slate-500">Valorizado inventario</div>
                                    <div className="text-2xl font-bold mt-1">{formatMoney(inventoryValue)}</div>
                                </div>
                                <div className="rounded-2xl border p-5 bg-slate-50">
                                    <div className="text-sm text-slate-500">Valorizado contado</div>
                                    <div className="text-2xl font-bold mt-1">{formatMoney(countedValue)}</div>
                                </div>
                                <div className="rounded-2xl border p-5 bg-slate-50">
                                    <div className="text-sm text-slate-500">Avance por SKU</div>
                                    <div className="text-2xl font-bold mt-1">{skuProgress.pct}%</div>
                                    <div className="text-sm text-slate-500 mt-1">{skuProgress.counted} de {skuProgress.total}</div>
                                    <div className="text-xs text-slate-500 mt-2">No considera productos con stock 0.</div>
                                </div>
                                <div className="rounded-2xl border p-5 bg-slate-50">
                                    <div className="text-sm text-slate-500">Avance por valorizado</div>
                                    <div className="text-2xl font-bold mt-1">{valorizadoPct}%</div>
                                    <div className="text-sm text-slate-500 mt-1">{formatMoney(countedValue)} de {formatMoney(inventoryValue)}</div>
                                </div>
                            </div>
                        </section>

                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div>
                                <h3 className="text-xl font-bold text-slate-900">Indicadores por usuario</h3>
                                <p className="text-slate-600 text-sm mt-1">El cálculo usa el tiempo real entre el primer y el último registro por usuario.</p>
                            </div>
                            <div className="grid lg:grid-cols-2 gap-6">
                                <div className="border rounded-2xl p-5">
                                    <h4 className="font-semibold mb-4">SKU contados por min</h4>
                                    <div className="space-y-4">
                                        {speedByUser.map((item) => (
                                            <div key={item.user}>
                                                <div className="flex justify-between text-sm mb-1 gap-3"><span className="font-medium">{item.user}</span><span>{item.skuPerMin}</span></div>
                                                <div className="w-full bg-slate-200 rounded-full h-4 overflow-hidden"><div className="bg-slate-900 h-4" style={{ width: `${(item.skuPerMin / maxBarValue) * 100}%` }} /></div>
                                            </div>
                                        ))}
                                        {speedByUser.length === 0 && <div className="text-sm text-slate-500">Sin datos.</div>}
                                    </div>
                                </div>
                                <div className="border rounded-2xl p-5">
                                    <h4 className="font-semibold mb-4">Unidades contadas por min</h4>
                                    <div className="space-y-4">
                                        {speedByUser.map((item) => (
                                            <div key={item.user + "-u"}>
                                                <div className="flex justify-between text-sm mb-1 gap-3"><span className="font-medium">{item.user}</span><span>{item.unitsPerMin}</span></div>
                                                <div className="w-full bg-slate-200 rounded-full h-4 overflow-hidden"><div className="bg-blue-700 h-4" style={{ width: `${(item.unitsPerMin / maxBarValue) * 100}%` }} /></div>
                                            </div>
                                        ))}
                                        {speedByUser.length === 0 && <div className="text-sm text-slate-500">Sin datos.</div>}
                                    </div>
                                </div>
                            </div>
                        </section>

                        {/* Tabla con pestañas */}
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div className="flex gap-2 border-b border-slate-200 pb-0">
                                <button
                                    className={`px-5 py-2.5 rounded-t-xl font-semibold text-sm border-b-2 transition-colors ${validadorSubTab === "registros" ? "border-slate-900 text-slate-900 bg-slate-50" : "border-transparent text-slate-500 hover:text-slate-700"}`}
                                    onClick={() => setValidadorSubTab("registros")}
                                >
                                    Registros totales
                                </button>
                                <button
                                    className={`px-5 py-2.5 rounded-t-xl font-semibold text-sm border-b-2 transition-colors ${validadorSubTab === "resumen" ? "border-slate-900 text-slate-900 bg-slate-50" : "border-transparent text-slate-500 hover:text-slate-700"}`}
                                    onClick={() => setValidadorSubTab("resumen")}
                                >
                                    Resumen por código
                                </button>
                            </div>

                            {/* Registros totales */}
                            {validadorSubTab === "registros" && (
                                <div className="space-y-4">
                                    <div className="flex flex-col sm:flex-row gap-3">
                                        <input className="flex-1 border rounded-2xl p-3 text-sm" placeholder="Buscar..." value={searchText} onChange={(e) => { setSearchText(e.target.value); setRecordsPage(1); }} />
                                        <select className="border rounded-2xl p-3 text-sm" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                                            <option value="todos">Todos</option>
                                            <option value="pendiente">Pendiente</option>
                                            <option value="diferencia">Diferencia</option>
                                            <option value="validado">Validado</option>
                                            <option value="corregido">Corregido</option>
                                        </select>
                                        <button className="px-4 py-3 rounded-2xl border font-semibold shrink-0" onClick={exportRecords}>Descargar registros</button>
                                    </div>
                                    <div className="overflow-auto rounded-2xl border">
                                        <table className="w-full text-xs md:text-sm">
                                            <thead className="bg-slate-100">
                                                <tr>
                                                    <th className="p-3 border">SKU</th>
                                                    <th className="p-3 border">Descripción</th>
                                                    <th className="p-3 border">UM</th>
                                                    <th className="p-3 border">Cant. contada</th>
                                                    <th className="p-3 border">Stock sistema</th>
                                                    <th className="p-3 border">DIF</th>
                                                    <th className="p-3 border">Costo</th>
                                                    <th className="p-3 border">Dif. Valoriz.</th>
                                                    <th className="p-3 border">Ubicación</th>
                                                    <th className="p-3 border">Usuario</th>
                                                    <th className="p-3 border">Fecha y hora</th>
                                                    <th className="p-3 border">Estado</th>
                                                    <th className="p-3 border">Acción</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {filteredRecords.slice((recordsPage - 1) * RECORDS_PER_PAGE, recordsPage * RECORDS_PER_PAGE).map((r) => (
                                                    <tr key={r.id} className={r.status === "Diferencia" ? "bg-red-50" : r.status === "Validado" || r.status === "Corregido" ? "bg-green-50" : ""}>
                                                        <td className="p-3 border font-medium">{r.sku}</td>
                                                        <td className="p-3 border">{r.description}</td>
                                                        <td className="p-3 border">{r.unit}</td>
                                                        <td className="p-3 border">{r.counted_quantity}</td>
                                                        <td className="p-3 border">{r.system_stock}</td>
                                                        <td className="p-3 border">{diffBadge(r.difference)}</td>
                                                        <td className="p-3 border">{formatMoney(r.cost)}</td>
                                                        <td className="p-3 border">{formatMoney(r.difference * r.cost)}</td>
                                                        <td className="p-3 border">{r.location}</td>
                                                        <td className="p-3 border">{r.user_name}</td>
                                                        <td className="p-3 border">{formatDateTime(r.counted_at)}</td>
                                                        <td className="p-3 border"><span className={statusBadge(r.status)}>{r.status}</span></td>
                                                        <td className="p-3 border">
                                                            <div className="flex gap-2">
                                                                <button className="px-3 py-2 rounded-xl border text-xs font-semibold" onClick={() => openEdit(r)}>Editar</button>
                                                                <button className="px-3 py-2 rounded-xl bg-red-600 text-white hover:bg-red-700 text-xs font-semibold" onClick={() => deleteRecord(r)}>Eliminar</button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                ))}
                                                {filteredRecords.length === 0 && (
                                                    <tr><td className="p-4 border text-center text-slate-500" colSpan={13}>No hay registros para mostrar.</td></tr>
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                    {filteredRecords.length > RECORDS_PER_PAGE && (
                                        <div className="flex items-center justify-between gap-3 pt-2">
                                            <span className="text-sm text-slate-500">
                                                Mostrando {Math.min((recordsPage - 1) * RECORDS_PER_PAGE + 1, filteredRecords.length)}–{Math.min(recordsPage * RECORDS_PER_PAGE, filteredRecords.length)} de {filteredRecords.length.toLocaleString()} registros
                                            </span>
                                            <div className="flex gap-2">
                                                <button className="px-3 py-2 rounded-xl border text-sm font-semibold disabled:opacity-40" disabled={recordsPage === 1} onClick={() => setRecordsPage((p) => p - 1)}>← Anterior</button>
                                                <span className="px-3 py-2 text-sm font-semibold">Pág. {recordsPage} / {Math.ceil(filteredRecords.length / RECORDS_PER_PAGE)}</span>
                                                <button className="px-3 py-2 rounded-xl border text-sm font-semibold disabled:opacity-40" disabled={recordsPage * RECORDS_PER_PAGE >= filteredRecords.length} onClick={() => setRecordsPage((p) => p + 1)}>Siguiente →</button>
                                            </div>
                                        </div>
                                    )}
                                </div>
                            )}

                            {/* Resumen por código */}
                            {validadorSubTab === "resumen" && (
                                <div className="space-y-4">
                                    <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                                        <p className="text-slate-600 text-sm">
                                            Agrupa por SKU sumando todos los conteos. Incluye todos los productos del maestro con stock &gt; 0. Consulta directa a base de datos.
                                        </p>
                                        <div className="flex gap-2 shrink-0">
                                            <button
                                                className="px-4 py-3 rounded-2xl bg-slate-600 hover:bg-slate-700 text-white font-semibold text-sm"
                                                onClick={buildAuditFromDB}
                                                disabled={auditLoading}
                                            >
                                                {auditLoading ? "Cargando..." : "🔄 Actualizar"}
                                            </button>
                                            <button
                                                className="px-4 py-3 rounded-2xl bg-slate-800 hover:bg-slate-900 text-white font-semibold text-sm"
                                                onClick={exportAudit}
                                            >
                                                ⬇ Descargar resumen
                                            </button>
                                            <button
                                                className="px-4 py-3 rounded-2xl bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-sm"
                                                onClick={() => { setReportAuditorName(user?.full_name || ""); setShowReportModal(true); }}
                                            >
                                                📊 Generar informe
                                            </button>
                                        </div>
                                    </div>

                                    {auditLoading && (
                                        <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-800">
                                            Consultando base de datos completa, por favor espera...
                                        </div>
                                    )}

                                    {/* KPI cards */}
                                    <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                                        <div className="rounded-2xl border p-4 bg-slate-50">
                                            <div className="text-xs text-slate-500">Total SKUs</div>
                                            <div className="text-2xl font-bold mt-1">{auditByCode.length}</div>
                                        </div>
                                        <div className="rounded-2xl border p-4 bg-green-50">
                                            <div className="text-xs text-green-700">OK (sin diferencia)</div>
                                            <div className="text-2xl font-bold text-green-800 mt-1">{auditTotals.totalOk}</div>
                                        </div>
                                        <div className="rounded-2xl border p-4 bg-red-50">
                                            <div className="text-xs text-red-700">Faltantes</div>
                                            <div className="text-2xl font-bold text-red-700 mt-1">{auditTotals.totalFaltantes}</div>
                                        </div>
                                        <div className="rounded-2xl border p-4 bg-blue-50">
                                            <div className="text-xs text-blue-700">Sobrantes</div>
                                            <div className="text-2xl font-bold text-blue-700 mt-1">{auditTotals.totalSobrantes}</div>
                                        </div>
                                        <div className="rounded-2xl border p-4 bg-orange-50">
                                            <div className="text-xs text-orange-700">No contados</div>
                                            <div className="text-2xl font-bold text-orange-700 mt-1">{auditTotals.totalNoContado}</div>
                                        </div>
                                    </div>

                                    <div className="rounded-2xl border p-4 bg-amber-50 flex items-center justify-between gap-3">
                                        <div>
                                            <div className="text-xs text-amber-700 font-semibold">DIFERENCIA VALORIZADA TOTAL (filtrado)</div>
                                            <div className={`text-2xl font-bold mt-1 ${auditTotals.totalValuedDiff < 0 ? "text-red-700" : auditTotals.totalValuedDiff > 0 ? "text-blue-700" : "text-green-700"}`}>
                                                {formatMoney(auditTotals.totalValuedDiff)}
                                            </div>
                                        </div>
                                        <div className="text-3xl">{auditTotals.totalValuedDiff < 0 ? "📉" : auditTotals.totalValuedDiff > 0 ? "📈" : "✅"}</div>
                                    </div>

                                    {/* Filtros */}
                                    <div className="flex flex-col sm:flex-row gap-3">
                                        <input
                                            className="flex-1 border rounded-2xl p-3 text-sm"
                                            placeholder="Buscar por SKU o descripción..."
                                            value={auditSearchText}
                                            onChange={(e) => setAuditSearchText(e.target.value)}
                                        />
                                        <select
                                            className="border rounded-2xl p-3 text-sm"
                                            value={auditStatusFilter}
                                            onChange={(e) => setAuditStatusFilter(e.target.value)}
                                        >
                                            <option value="todos">Todos los status</option>
                                            <option value="ok">OK</option>
                                            <option value="faltante">Faltante</option>
                                            <option value="sobrante">Sobrante</option>
                                            <option value="no_contado">No contado</option>
                                        </select>
                                    </div>

                                    {/* Tabla resumen */}
                                    <div className="overflow-auto rounded-2xl border">
                                        <table className="w-full text-xs md:text-sm">
                                            <thead className="bg-slate-100 sticky top-0">
                                                <tr>
                                                    <th className="p-3 border text-left">SKU</th>
                                                    <th className="p-3 border text-left">Descripción</th>
                                                    <th className="p-3 border text-center">UM</th>
                                                    <th className="p-3 border text-center">Registros</th>
                                                    <th className="p-3 border text-center">Stock sistema</th>
                                                    <th className="p-3 border text-center">Stock contado</th>
                                                    <th className="p-3 border text-center">Dif. unidad</th>
                                                    <th className="p-3 border text-center">Costo unit.</th>
                                                    <th className="p-3 border text-center">Dif. valorizada</th>
                                                    <th className="p-3 border text-center">Status</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {filteredAudit.map((row) => (
                                                    <tr
                                                        key={row.sku}
                                                        className={
                                                            row.status_resumen === "NO CONTADO"
                                                                ? "bg-orange-50"
                                                                : row.status_resumen === "FALTANTE"
                                                                ? "bg-red-50"
                                                                : row.status_resumen === "SOBRANTE"
                                                                ? "bg-blue-50"
                                                                : "bg-green-50"
                                                        }
                                                    >
                                                        <td className="p-3 border font-medium">{row.sku}</td>
                                                        <td className="p-3 border">{row.description}</td>
                                                        <td className="p-3 border text-center">{row.unit}</td>
                                                        <td className="p-3 border text-center">{row.record_count}</td>
                                                        <td className="p-3 border text-center">{row.system_stock}</td>
                                                        <td className="p-3 border text-center font-semibold">{row.total_counted}</td>
                                                        <td className="p-3 border text-center">{diffBadge(row.difference)}</td>
                                                        <td className="p-3 border text-center">{formatMoney(row.cost)}</td>
                                                        <td className={`p-3 border text-center font-semibold ${row.valued_difference < 0 ? "text-red-700" : row.valued_difference > 0 ? "text-blue-700" : "text-green-700"}`}>
                                                            {formatMoney(row.valued_difference)}
                                                        </td>
                                                        <td className="p-3 border text-center">
                                                            <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-bold ${
                                                                row.status_resumen === "OK"
                                                                    ? "bg-green-100 text-green-700"
                                                                    : row.status_resumen === "FALTANTE"
                                                                    ? "bg-red-100 text-red-700"
                                                                    : row.status_resumen === "SOBRANTE"
                                                                    ? "bg-blue-100 text-blue-700"
                                                                    : "bg-orange-100 text-orange-700"
                                                            }`}>
                                                                {row.status_resumen}
                                                            </span>
                                                        </td>
                                                    </tr>
                                                ))}
                                                {filteredAudit.length > 0 && (
                                                    <tr className="bg-slate-200 font-bold">
                                                        <td className="p-3 border" colSpan={3}>TOTAL ({filteredAudit.length} SKUs)</td>
                                                        <td className="p-3 border text-center">{filteredAudit.reduce((s, r) => s + r.record_count, 0)}</td>
                                                        <td className="p-3 border text-center">{filteredAudit.reduce((s, r) => s + r.system_stock, 0)}</td>
                                                        <td className="p-3 border text-center">{filteredAudit.reduce((s, r) => s + r.total_counted, 0)}</td>
                                                        <td className="p-3 border text-center">{diffBadge(filteredAudit.reduce((s, r) => s + r.difference, 0))}</td>
                                                        <td className="p-3 border text-center">—</td>
                                                        <td className={`p-3 border text-center ${auditTotals.totalValuedDiff < 0 ? "text-red-700" : auditTotals.totalValuedDiff > 0 ? "text-blue-700" : "text-green-700"}`}>
                                                            {formatMoney(auditTotals.totalValuedDiff)}
                                                        </td>
                                                        <td className="p-3 border text-center text-xs">
                                                            ✅{auditTotals.totalOk} 📉{auditTotals.totalFaltantes} 📈{auditTotals.totalSobrantes} 🟠{auditTotals.totalNoContado}
                                                        </td>
                                                    </tr>
                                                )}
                                                {filteredAudit.length === 0 && (
                                                    <tr>
                                                        <td className="p-6 border text-center text-slate-400" colSpan={10}>
                                                            {auditLoading
                                                                ? "Cargando datos..."
                                                                : auditByCode.length === 0
                                                                ? "No hay registros contados aún."
                                                                : "Sin resultados para el filtro seleccionado."}
                                                        </td>
                                                    </tr>
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            )}
                        </section>
                            </>
                        )}
                    </>
                )}

                {/* ── TAB ADMIN ────────────────────────────────────────────────── */}
                {activeTab === "admin" && user.role === "Administrador" && (
                    <>
                        <section className="bg-white rounded-3xl p-6 shadow space-y-6">
                            <div>
                                <h2 className="text-2xl font-bold text-slate-900">Módulo Administrador</h2>
                                <p className="text-slate-600 mt-1">Crea inventarios, carga maestros y administra usuarios.</p>
                            </div>
                            <div className="grid md:grid-cols-2 xl:grid-cols-4 gap-4">
                                <div className="rounded-2xl border p-5 bg-slate-50"><div className="text-sm text-slate-500">Inventarios</div><div className="text-2xl font-bold mt-1">{inventories.length}</div></div>
                                <div className="rounded-2xl border p-5 bg-slate-50"><div className="text-sm text-slate-500">Productos</div><div className="text-2xl font-bold mt-1">{totalProductCount.toLocaleString()}</div></div>
                                <div className="rounded-2xl border p-5 bg-slate-50"><div className="text-sm text-slate-500">Usuarios</div><div className="text-2xl font-bold mt-1">{users.length}</div></div>
                                <div className="rounded-2xl border p-5 bg-slate-50"><div className="text-sm text-slate-500">Registros</div><div className="text-2xl font-bold mt-1">{records.length}</div></div>
                            </div>
                        </section>

                        <section className="grid lg:grid-cols-2 gap-6">
                            <div className="bg-white rounded-3xl p-6 shadow space-y-4">
                                <h3 className="text-xl font-bold text-slate-900">Crear inventario</h3>
                                <input className="w-full border rounded-2xl p-3" placeholder="Nombre del inventario" value={newInventoryName} onChange={(e) => setNewInventoryName(e.target.value)} />
                                <input className="w-full border rounded-2xl p-3" placeholder="Código corto (opcional)" value={newInventoryCode} onChange={(e) => setNewInventoryCode(e.target.value)} />
                                <button className="px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={createInventory}>Crear inventario</button>
                            </div>

                            {/* Crear usuario individual */}
                            <div className="bg-white rounded-3xl p-6 shadow space-y-4">
                                <h3 className="text-xl font-bold text-slate-900">Crear usuario individual</h3>
                                <input className="w-full border rounded-2xl p-3" placeholder="ID de usuario" value={newUsername} onChange={(e) => setNewUsername(e.target.value)} />
                                <input className="w-full border rounded-2xl p-3" placeholder="Nombre completo" value={newFullName} onChange={(e) => setNewFullName(e.target.value)} />
                                <div className="flex gap-2">
                                    <input className="w-full border rounded-2xl p-3" placeholder="Clave" type={showNewPassword ? "text" : "password"} value={newPassword} onChange={(e) => setNewPassword(e.target.value)} />
                                    <button type="button" className="px-4 rounded-2xl border" onClick={() => setShowNewPassword(!showNewPassword)}>{showNewPassword ? "Ocultar" : "Ver"}</button>
                                </div>

                                <div>
                                    <label className="block font-semibold mb-2 text-slate-800">Roles <span className="text-slate-400 font-normal text-xs">(selecciona uno o más)</span></label>
                                    <RoleCheckboxes selected={newRoles} onChange={setNewRoles} />
                                    <p className="text-xs text-slate-400 mt-2">Rol principal: <b>{newRoles.includes("Administrador") ? "Administrador" : newRoles.includes("Validador") ? "Validador" : "Operario"}</b></p>
                                </div>

                                {!newIsAdmin && (
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block font-semibold mb-2 text-slate-800">Acceso a inventarios</label>
                                            <label className="flex items-center gap-3 cursor-pointer">
                                                <input
                                                    type="checkbox"
                                                    className="w-4 h-4 accent-slate-800"
                                                    checked={newCanAccessAnyInventory}
                                                    onChange={(e) => setNewCanAccessAnyInventory(e.target.checked)}
                                                />
                                                <span className="text-sm font-medium text-slate-700">Acceso libre a todos los inventarios</span>
                                            </label>
                                            <p className="text-xs text-slate-400 mt-1">Útil para auditores que trabajan en múltiples inventarios.</p>
                                        </div>
                                        {!newCanAccessAnyInventory && (
                                            <div>
                                                <label className="block font-semibold mb-2 text-slate-800">Inventario asignado</label>
                                                <select className="w-full border rounded-2xl p-3" value={selectedInventoryId} disabled>
                                                    {allInventories.filter(inv => inv.is_active).map(inv => (
                                                        <option key={inv.id} value={inv.id}>{inv.name}</option>
                                                    ))}
                                                </select>
                                                <p className="text-xs text-slate-400 mt-1">Se asignará el inventario activo seleccionado actualmente.</p>
                                            </div>
                                        )}
                                    </div>
                                )}

                                <div>
                                    <label className="block font-semibold mb-2 text-slate-800">Permisos adicionales <span className="text-slate-400 font-normal text-xs">(datos sensibles)</span></label>
                                    <PermissionsCheckboxes perms={newPermissions} onChange={setNewPermissions} isAdmin={newIsAdmin} />
                                </div>

                                <button className="px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold w-full" onClick={createUser}>Crear usuario</button>
                            </div>
                        </section>

                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div>
                                <h3 className="text-xl font-bold text-slate-900">Inventarios</h3>
                                <p className="text-slate-600 text-sm mt-1">Si tiene información, se archiva. Si está vacío, se elimina.</p>
                            </div>
                            <div className="overflow-auto rounded-2xl border">
                                <table className="w-full text-sm">
                                    <thead className="bg-slate-100">
                                        <tr>
                                            <th className="p-3 border text-left">Nombre</th>
                                            <th className="p-3 border text-left">Código</th>
                                            <th className="p-3 border text-left">Estado</th>
                                            <th className="p-3 border text-left">Acción</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {allInventories.map((inv) => {
                                            const isGeneral = normalizeText(inv.code) === "general";
                                            const isSelected = inv.id === selectedInventoryId;
                                            const isProcessing = processingInventoryId === inv.id;
                                            return (
                                                <tr key={inv.id} className={!inv.is_active ? "bg-slate-50 opacity-70" : ""}>
                                                    <td className="p-3 border">{inv.name}</td>
                                                    <td className="p-3 border">{inv.code}</td>
                                                    <td className="p-3 border">
                                                        {!inv.is_active
                                                            ? <span className="inline-block px-2 py-0.5 rounded-full text-xs font-semibold bg-slate-200 text-slate-600">Archivado</span>
                                                            : isSelected
                                                            ? <span className="inline-block px-2 py-0.5 rounded-full text-xs font-semibold bg-green-100 text-green-700">Activo seleccionado</span>
                                                            : <span className="inline-block px-2 py-0.5 rounded-full text-xs font-semibold bg-blue-100 text-blue-700">Activo</span>}
                                                    </td>
                                                    <td className="p-3 border">
                                                        {!inv.is_active ? (
                                                            <button
                                                                className={`px-4 py-2 rounded-xl text-white ${isProcessing ? "bg-slate-400 cursor-not-allowed" : "bg-green-600 hover:bg-green-700"}`}
                                                                onClick={() => reactivateInventory(inv)}
                                                                disabled={isProcessing}
                                                            >
                                                                {isProcessing ? "Procesando..." : "Reactivar"}
                                                            </button>
                                                        ) : (
                                                            <button className={`px-4 py-2 rounded-xl text-white ${isGeneral || isSelected || isProcessing ? "bg-slate-400 cursor-not-allowed" : "bg-red-600 hover:bg-red-700"}`} onClick={() => manageInventory(inv)} disabled={isGeneral || isSelected || isProcessing}>
                                                                {isProcessing ? "Procesando..." : "Eliminar / Archivar"}
                                                            </button>
                                                        )}
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                        {allInventories.length === 0 && (<tr><td className="p-4 border text-center text-slate-500" colSpan={4}>No hay inventarios.</td></tr>)}
                                    </tbody>
                                </table>
                            </div>
                        </section>

                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <h3 className="text-xl font-bold text-slate-900">Insertar usuarios masivos</h3>
                            <p className="text-slate-600 text-sm">Plantilla: ID, CLAVE, NOMBRE, ROL</p>
                            <input ref={usersInputRef} type="file" accept=".xlsx,.xls" onChange={(e) => { const f = e.target.files?.[0] || null; setUsersFile(f); setUsersFileName(f ? f.name : ""); }} />
                            <div className="text-sm text-slate-500">{usersFileName ? `📄 ${usersFileName}` : "Ningún archivo seleccionado"}</div>
                            <button className="px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold" type="button" onClick={uploadUsers}>Insertar usuarios</button>
                        </section>

                        {/* ── CATÁLOGO GLOBAL DE BARCODES ──────────────────── */}
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <div>
                                <h3 className="text-xl font-bold text-slate-900">📦 Catálogo global de códigos de barra</h3>
                                <p className="text-slate-600 text-sm mt-1">
                                    Aplica a <b>todos los inventarios</b>. Sube una sola vez y actualiza cuando lo necesites.
                                    Al escanear un código, el sistema lo busca aquí, obtiene el SKU y lo vincula al producto del inventario activo.
                                </p>
                                <p className="text-xs text-slate-400 mt-1">Columnas requeridas: <b>SKU</b> y <b>CODIGO_BARRA</b>. Puede tener múltiples filas por SKU.</p>
                            </div>
                            <div className="grid md:grid-cols-2 gap-6">
                                <div className="space-y-3">
                                    <input
                                        ref={globalBarcodesInputRef}
                                        type="file"
                                        accept=".xlsx,.xls"
                                        onChange={(e) => { const f = e.target.files?.[0] || null; setGlobalBarcodesFile(f); setGlobalBarcodesFileName(f ? f.name : ""); }}
                                    />
                                    <div className="text-sm text-slate-500">{globalBarcodesFileName ? `📄 ${globalBarcodesFileName}` : "Ningún archivo seleccionado"}</div>
                                    {uploadProgress && (
                                        <div className="space-y-2">
                                            <div className="flex justify-between text-sm font-semibold text-slate-700"><span>{uploadProgress?.step}</span><span>{uploadProgress?.pct ?? 0}%</span></div>
                                            <div className="w-full bg-slate-200 rounded-full h-3 overflow-hidden"><div className="bg-indigo-600 h-3 rounded-full transition-all duration-300" style={{ width: `${uploadProgress?.pct ?? 0}%` }} /></div>
                                        </div>
                                    )}
                                    <button
                                        className={`px-4 py-3 rounded-2xl font-semibold text-white ${uploadProgress ? "bg-slate-400 cursor-not-allowed" : "bg-indigo-700 hover:bg-indigo-800"}`}
                                        type="button"
                                        onClick={uploadGlobalBarcodes}
                                        disabled={!!uploadProgress}
                                    >
                                        {uploadProgress ? "Subiendo..." : "Subir / Actualizar catálogo global"}
                                    </button>
                                </div>
                                <div className="bg-slate-50 rounded-2xl p-4 space-y-2">
                                    <div className="font-semibold text-slate-700 text-sm">¿Cómo funciona?</div>
                                    <ul className="text-slate-600 space-y-1 text-xs list-disc list-inside">
                                        <li>Sube una vez con todos los SKUs y sus códigos de barra</li>
                                        <li>Al escanear, el sistema busca aquí, obtiene el SKU y lo vincula al producto del inventario activo</li>
                                        <li>Puedes actualizar sin afectar ningún inventario</li>
                                    </ul>
                                    {globalBarcodesCount !== null && (
                                        <div className="mt-2 text-indigo-700 font-semibold text-sm">✅ Último upload: {globalBarcodesCount?.toLocaleString()} códigos</div>
                                    )}
                                </div>
                            </div>
                        </section>

                        {/* Tabla usuarios */}
                        <section className="bg-white rounded-3xl p-6 shadow space-y-4">
                            <h3 className="text-xl font-bold text-slate-900">Usuarios registrados</h3>
                            <div className="overflow-auto rounded-2xl border">
                                <table className="w-full text-sm">
                                    <thead className="bg-slate-100">
                                        <tr>
                                            <th className="p-3 border">ID</th>
                                            <th className="p-3 border">Nombre</th>
                                            <th className="p-3 border">Roles</th>
                                            <th className="p-3 border">Inventario</th>
                                            <th className="p-3 border">Permisos extra</th>
                                            <th className="p-3 border">Activo</th>
                                            <th className="p-3 border">Acciones</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {users.map((u) => {
                                            const displayRoles: Role[] = u.roles && u.roles.length > 0 ? u.roles : [u.role];
                                            return (
                                                <tr key={u.id}>
                                                    <td className="p-3 border font-medium">{u.username}</td>
                                                    <td className="p-3 border">{u.full_name}</td>
                                                    <td className="p-3 border">
                                                        <div className="flex flex-wrap gap-1">
                                                            {displayRoles.map((r) => {
                                                                const roleBadge = r === "Administrador" ? "bg-purple-100 text-purple-700" :
                                                                                  r === "Validador"     ? "bg-blue-100 text-blue-700" :
                                                                                                         "bg-slate-100 text-slate-700";
                                                                return (
                                                                    <span key={r} className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${roleBadge}`}>{r}</span>
                                                                );
                                                            })}
                                                        </div>
                                                    </td>
                                                    <td className="p-3 border">{getPrimaryRole(u) === "Administrador" ? "Todos" : allInventories.find((inv) => inv.id === u.inventory_id)?.name || <span className="text-slate-400 italic">Sin asignar</span>}</td>
                                                    <td className="p-3 border">
                                                        <div className="flex flex-wrap gap-1 text-xs">
                                                            {u.can_access_any_inventory && <span className="bg-purple-100 text-purple-700 px-2 py-0.5 rounded-full">🌐 Todos inv.</span>}
                                                            {u.can_see_system_stock && <span className="bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full">Stock sis.</span>}
                                                            {u.can_see_cost && <span className="bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full">Costo</span>}
                                                            {u.can_see_valued_difference && <span className="bg-sky-100 text-sky-700 px-2 py-0.5 rounded-full">Dif. valor.</span>}
                                                            {!u.can_see_system_stock && !u.can_see_cost && !u.can_see_valued_difference && !u.can_access_any_inventory && getPrimaryRole(u) !== "Administrador" && getPrimaryRole(u) !== "Validador" && (
                                                                <span className="text-slate-400 italic">Ninguno</span>
                                                            )}
                                                            {(getPrimaryRole(u) === "Administrador" || getPrimaryRole(u) === "Validador") && (
                                                                <span className="text-slate-400 italic text-xs">Acceso total</span>
                                                            )}
                                                        </div>
                                                    </td>
                                                    <td className="p-3 border">
                                                        <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${u.is_active ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                                                            {u.is_active ? "Activo" : "Inactivo"}
                                                        </span>
                                                    </td>
                                                    <td className="p-3 border">
                                                        <div className="flex gap-2">
                                                            <button className="px-3 py-1.5 rounded-xl border text-xs font-semibold hover:bg-slate-50" onClick={() => openEditUser(u)}>
                                                                ✏️ Editar
                                                            </button>
                                                            <button className={`px-3 py-1.5 rounded-xl text-white text-xs font-semibold ${u.id === user.id ? "bg-slate-300 cursor-not-allowed" : "bg-red-600 hover:bg-red-700"}`} onClick={() => deleteUser(u)} disabled={u.id === user.id}>
                                                                Eliminar
                                                            </button>
                                                        </div>
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                        {users.length === 0 && (<tr><td className="p-4 border text-center text-slate-500" colSpan={7}>No hay usuarios todavía.</td></tr>)}
                                    </tbody>
                                </table>
                            </div>
                        </section>
                    </>
                )}

                {/* ── MODAL EDITAR USUARIO ─────────────────────────────────── */}
                {editingUser && (
                    <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
                        <div className="bg-white rounded-3xl p-6 w-full max-w-md shadow-2xl space-y-5 max-h-[90vh] overflow-y-auto">
                            <div>
                                <h3 className="text-2xl font-bold text-slate-900">Editar usuario</h3>
                                <p className="text-slate-500 text-sm mt-1">
                                    <b>{editingUser.username}</b> — {editingUser.full_name}
                                </p>
                            </div>

                            <div>
                                <label className="block font-semibold mb-2">Roles <span className="text-slate-400 font-normal text-xs">(selecciona uno o más)</span></label>
                                <RoleCheckboxes selected={editUserRoles} onChange={setEditUserRoles} />
                                <p className="text-xs text-slate-400 mt-2">Rol principal: <b>{editUserRoles.includes("Administrador") ? "Administrador" : editUserRoles.includes("Validador") ? "Validador" : "Operario"}</b></p>
                            </div>

                            {!editingIsAdmin && !editUserCanAccessAnyInventory && (
                                <div>
                                    <label className="block font-semibold mb-2">Inventario asignado</label>
                                    <select className="w-full border rounded-2xl p-3" value={editUserInventoryId} onChange={e => setEditUserInventoryId(e.target.value)}>
                                        <option value="">Sin asignar</option>
                                        {allInventories.filter(inv => inv.is_active).map(inv => (
                                            <option key={inv.id} value={inv.id}>{inv.name}</option>
                                        ))}
                                    </select>
                                </div>
                            )}
                            {!editingIsAdmin && editUserCanAccessAnyInventory && (
                                <div className="rounded-xl bg-blue-50 border border-blue-200 px-4 py-3 text-sm text-blue-700">
                                    ✓ Este usuario tiene acceso libre a todos los inventarios.
                                </div>
                            )}

                            <div>
                                <label className="block font-semibold mb-2">Permisos adicionales <span className="text-slate-400 font-normal text-xs">(datos sensibles)</span></label>
                                <PermissionsCheckboxes perms={editUserPermissions} onChange={setEditUserPermissions} isAdmin={editingIsAdmin} />
                            </div>

                            {!editingIsAdmin && (
                                <div>
                                    <label className="block font-semibold mb-2">Acceso a inventarios</label>
                                    <label className="flex items-center gap-3 cursor-pointer">
                                        <input
                                            type="checkbox"
                                            className="w-4 h-4 accent-slate-800"
                                            checked={editUserCanAccessAnyInventory}
                                            onChange={(e) => {
                                                setEditUserCanAccessAnyInventory(e.target.checked);
                                                if (e.target.checked) setEditUserInventoryId("");
                                            }}
                                        />
                                        <span className="text-sm font-medium text-slate-700">Acceso libre a todos los inventarios</span>
                                    </label>
                                    <p className="text-xs text-slate-400 mt-1">Permite a este usuario trabajar en cualquier inventario activo (útil para auditores).</p>
                                </div>
                            )}

                            <div>
                                <label className="block font-semibold mb-2">Estado</label>
                                <div className="flex gap-3">
                                    <button
                                        className={`flex-1 py-2.5 rounded-xl font-semibold text-sm border ${editUserActive ? "bg-green-600 text-white border-green-600" : "bg-white text-slate-700 border-slate-300"}`}
                                        onClick={() => setEditUserActive(true)}
                                    >
                                        ✓ Activo
                                    </button>
                                    <button
                                        className={`flex-1 py-2.5 rounded-xl font-semibold text-sm border ${!editUserActive ? "bg-red-500 text-white border-red-500" : "bg-white text-slate-700 border-slate-300"}`}
                                        onClick={() => setEditUserActive(false)}
                                    >
                                        Inactivo
                                    </button>
                                </div>
                            </div>
                            <div className="flex gap-3 pt-1">
                                <button className="flex-1 px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={saveEditUser}>
                                    Guardar cambios
                                </button>
                                <button className="flex-1 px-4 py-3 rounded-2xl border font-semibold" onClick={() => setEditingUser(null)}>
                                    Cancelar
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                {/* ── MODAL EDITAR REGISTRO ──────────────────────────────────── */}
                {editingRecord && (
                    <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-50">
                        <div className="bg-white rounded-3xl p-6 w-full max-w-2xl space-y-5 shadow-2xl max-h-[90vh] overflow-y-auto">
                            <div>
                                <h3 className="text-2xl font-bold text-slate-900">Editar registro</h3>
                                <p className="text-slate-600 text-sm mt-1">El SKU debe existir en el maestro del inventario seleccionado.</p>
                            </div>
                            <div className="grid md:grid-cols-2 gap-4">
                                <div>
                                    <label className="block font-semibold mb-2">SKU</label>
                                    <input className="w-full border rounded-2xl p-3" value={editSku} onChange={(e) => handleEditSkuChange(e.target.value)} />
                                </div>
                                {showSystemStock && (
                                    <div>
                                        <label className="block font-semibold mb-2">Stock sistema</label>
                                        <input className="w-full border rounded-2xl p-3 bg-slate-100" value={editMatchedProduct ? editMatchedProduct.system_stock : editingRecord.system_stock} disabled />
                                    </div>
                                )}
                                <div>
                                    <label className="block font-semibold mb-2">Descripción</label>
                                    <input className="w-full border rounded-2xl p-3 bg-slate-100" value={editMatchedProduct ? editMatchedProduct.description : editingRecord.description} disabled />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-2">Unidad de medida</label>
                                    <input className="w-full border rounded-2xl p-3 bg-slate-100" value={editMatchedProduct ? editMatchedProduct.unit : editingRecord.unit} disabled />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-2">Cantidad</label>
                                    <input className="w-full border rounded-2xl p-3" type="number" value={editQty} onChange={(e) => setEditQty(e.target.value)} />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-2">Ubicación</label>
                                    <input className="w-full border rounded-2xl p-3" value={editLocation} onChange={(e) => setEditLocation(e.target.value)} />
                                </div>
                                <div className="md:col-span-2">
                                    <label className="block font-semibold mb-2">Observación / Nota</label>
                                    <textarea className="w-full border rounded-2xl p-3 resize-none" rows={2} value={editNote} onChange={(e) => setEditNote(e.target.value)} placeholder="Opcional: agrega una observación..." />
                                </div>
                                {canValidate(user) && (
                                    <div className="md:col-span-2">
                                        <label className="block font-semibold mb-2">Estado</label>
                                        <select className="w-full border rounded-2xl p-3" value={editStatus} onChange={(e) => setEditStatus(e.target.value as RecordRow["status"])}>
                                            <option value="Pendiente">Pendiente</option>
                                            <option value="Diferencia">Diferencia</option>
                                            <option value="Validado">Validado</option>
                                            <option value="Corregido">Corregido</option>
                                        </select>
                                    </div>
                                )}
                                <div className="md:col-span-2 flex flex-col sm:flex-row gap-3">
                                    <button className="px-5 py-3 rounded-2xl bg-slate-900 text-white font-semibold" onClick={saveEdit}>Guardar cambios</button>
                                    <button className="px-5 py-3 rounded-2xl border font-semibold" onClick={closeEdit}>Cancelar</button>
                                </div>
                            </div>
                        </div>
                    </div>
                )}

                {/* ── SCANNER OVERLAY ───────────────────────────────────────── */}
                {scannerTarget && (
                    <div className="fixed inset-0 bg-black/70 flex items-center justify-center p-4 z-[60]">
                        <div className="bg-white w-full max-w-lg rounded-3xl p-5 shadow-2xl space-y-4">
                            <div>
                                <h3 className="text-xl font-bold text-slate-900">
                                    {scannerTarget === "product" ? "Escanear producto" : "Escanear ubicación"}
                                </h3>
                                <p className="text-sm text-slate-500">
                                    {scannerTarget === "product" ? "Busca primero por código de barra y si no existe por SKU." : "Escanea o digita la ubicación."}
                                </p>
                            </div>
                            <div className="rounded-2xl overflow-hidden border bg-black min-h-[260px] flex items-center justify-center">
                                <div id={scannerContainerId} className="w-full" />
                            </div>
                            <div className="text-sm text-slate-500">
                                {scannerRunning ? "Cámara activa. Apunta al código." : "Iniciando cámara..."}
                            </div>
                            {torchAvailable && (
                                <button type="button" onClick={toggleTorch} className="w-full px-4 py-3 rounded-2xl bg-slate-900 text-white font-semibold">
                                    {torchOn ? "Apagar linterna 🔦" : "Prender linterna 🔦"}
                                </button>
                            )}
                        </div>
                    </div>
                )}

                {/* ── MODAL GENERAR INFORME ─────────────────────────────────── */}
                {showReportModal && (
                    <div className="fixed inset-0 bg-black/60 flex items-center justify-center p-4 z-50">
                        <div className="bg-white rounded-3xl p-6 w-full max-w-lg space-y-5 shadow-2xl max-h-[90vh] overflow-y-auto">
                            <div>
                                <h3 className="text-2xl font-bold text-slate-900">📊 Generar informe completo</h3>
                                <p className="text-slate-500 text-sm mt-1">
                                    Se generará un informe HTML con gráficos y dashboards, compatible con Gmail y cualquier navegador.
                                </p>
                            </div>

                            <div className="rounded-2xl bg-indigo-50 border border-indigo-200 px-4 py-3 text-sm text-indigo-800">
                                <span className="font-semibold">Inventario:</span> {currentInventory?.name || "—"}
                                <br/>
                                <span className="font-semibold">Fecha de auditoría:</span>{" "}
                                {(() => {
                                    let d = new Date();
                                    if (records.length > 0) {
                                        d = records.reduce((latest, r) => {
                                            const dt = new Date(r.counted_at);
                                            return dt > latest ? dt : latest;
                                        }, new Date(records[0].counted_at));
                                    }
                                    return d.toLocaleDateString("es-PE", { year: "numeric", month: "long", day: "numeric" });
                                })()}
                            </div>

                            <div className="space-y-4">
                                <div>
                                    <label className="block font-semibold mb-1.5 text-sm text-slate-700">Nombre de la tienda <span className="text-red-500">*</span></label>
                                    <input
                                        className="w-full border rounded-2xl p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300"
                                        placeholder="Ej: Tienda Centro Lima"
                                        value={reportStoreName}
                                        onChange={(e) => setReportStoreName(e.target.value)}
                                    />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-1.5 text-sm text-slate-700">Nombre del líder de tienda <span className="text-red-500">*</span></label>
                                    <input
                                        className="w-full border rounded-2xl p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300"
                                        placeholder="Ej: Juan Pérez"
                                        value={reportStoreLeader}
                                        onChange={(e) => setReportStoreLeader(e.target.value)}
                                    />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-1.5 text-sm text-slate-700">Nombre del asesor de almacén <span className="text-red-500">*</span></label>
                                    <input
                                        className="w-full border rounded-2xl p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300"
                                        placeholder="Ej: María García"
                                        value={reportWarehouseAdvisor}
                                        onChange={(e) => setReportWarehouseAdvisor(e.target.value)}
                                    />
                                </div>
                                <div>
                                    <label className="block font-semibold mb-1.5 text-sm text-slate-700">Nombre del auditor</label>
                                    <input
                                        className="w-full border rounded-2xl p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300"
                                        placeholder="Ej: Carlos López"
                                        value={reportAuditorName}
                                        onChange={(e) => setReportAuditorName(e.target.value)}
                                    />
                                </div>
                            </div>

                            <div className="rounded-2xl bg-amber-50 border border-amber-200 px-4 py-3 text-xs text-amber-800">
                                💡 El informe se descargará como archivo <strong>.html</strong>. Puedes abrirlo en tu navegador, imprimirlo o adjuntarlo a un correo en Gmail.
                            </div>

                            <div className="flex gap-3 pt-1">
                                <button
                                    className="flex-1 px-4 py-3 rounded-2xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                    onClick={generateAndOpenReport}
                                    disabled={!reportStoreName.trim() || !reportStoreLeader.trim() || !reportWarehouseAdvisor.trim()}
                                >
                                    📥 Descargar informe
                                </button>
                                <button
                                    className="flex-1 px-4 py-3 rounded-2xl border font-semibold text-sm"
                                    onClick={() => setShowReportModal(false)}
                                >
                                    Cancelar
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                </main>
            </div>
        </div>
        )}
        </>
    );
}
