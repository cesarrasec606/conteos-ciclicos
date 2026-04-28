"use client";

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase/client";

type Role = "Operario" | "Validador" | "Administrador";

type SessionUser = {
  id: string;
  username: string;
  full_name: string;
  role: Role;
  inventory_id: string | null;
  can_access_any_inventory: boolean;
  login_at: number;
  expires_at: number;
};

type LoginUser = {
  id: string;
  username: string;
  password: string;
  full_name: string;
  role: Role;
  roles: Role[] | null;
  is_active: boolean;
  inventory_id: string | null;
  can_access_any_inventory: boolean;
  can_see_system_stock: boolean;
  can_see_cost: boolean;
  can_see_valued_difference: boolean;
};

const SESSION_KEY            = "session_user";
const CURRENT_INVENTORY_KEY  = "current_inventory_id";
const SESSION_HOURS          = 8;

export default function LoginPage() {
  const [username, setUsername]         = useState("");
  const [password, setPassword]         = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading]           = useState(false);
  const [checkingSession, setCheckingSession] = useState(true);
  const [error, setError]               = useState("");

  // Modal cambiar contraseña
  const [pendingUser, setPendingUser]         = useState<LoginUser | null>(null);
  const [showChangePass, setShowChangePass]   = useState(false);
  const [newPass, setNewPass]                 = useState("");
  const [confirmPass, setConfirmPass]         = useState("");
  const [showNewPass, setShowNewPass]         = useState(false);
  const [showConfirmPass, setShowConfirmPass] = useState(false);
  const [modalError, setModalError]           = useState("");
  const [modalLoading, setModalLoading]       = useState(false);

  useEffect(() => {
    checkExistingSession();
  }, []);

  // ── Verificar sesión activa ─────────────────────────────────────────────
  async function checkExistingSession() {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      if (!raw) { setCheckingSession(false); return; }

      const parsed = JSON.parse(raw) as SessionUser;
      if (!parsed?.id || !parsed?.expires_at) {
        clearSession(); setCheckingSession(false); return;
      }
      if (Date.now() > parsed.expires_at) {
        clearSession(); setCheckingSession(false); return;
      }

      const { data, error: dbError } = await supabase
        .from("app_users")
        .select("id, username, full_name, role, is_active, inventory_id, can_access_any_inventory")
        .eq("id", parsed.id)
        .eq("is_active", true)
        .single();

      if (dbError || !data) { clearSession(); setCheckingSession(false); return; }

      const freeAccess = data.role === "Administrador" || data.can_access_any_inventory === true;
      if (!freeAccess && !data.inventory_id) {
        setError("Este usuario no tiene inventario asignado. Contacta al administrador.");
        clearSession(); setCheckingSession(false); return;
      }

      const refreshed: SessionUser = {
        id: data.id,
        username: data.username,
        full_name: data.full_name,
        role: data.role,
        inventory_id: data.inventory_id,
        can_access_any_inventory: data.can_access_any_inventory ?? false,
        login_at: parsed.login_at || Date.now(),
        expires_at: Date.now() + SESSION_HOURS * 60 * 60 * 1000,
      };

      localStorage.setItem(SESSION_KEY, JSON.stringify(refreshed));
      if (!freeAccess && data.inventory_id) {
        localStorage.setItem(CURRENT_INVENTORY_KEY, data.inventory_id);
      } else {
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
      }

      window.location.href = "/dashboard";
    } catch {
      clearSession(); setCheckingSession(false);
    }
  }

  function clearSession() {
    localStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(CURRENT_INVENTORY_KEY);
  }

  // ── Login ───────────────────────────────────────────────────────────────
  async function handleLogin() {
    setError("");
    const cleanUser = username.trim();
    const cleanPass = password.trim();
    if (!cleanUser || !cleanPass) { setError("Completa usuario y contraseña."); return; }

    setLoading(true);

    const { data, error: dbError } = await supabase
      .from("app_users")
      .select("id, username, password, full_name, role, roles, is_active, inventory_id, can_access_any_inventory, can_see_system_stock, can_see_cost, can_see_valued_difference")
      .eq("username", cleanUser)
      .eq("is_active", true)
      .single();

    if (dbError || !data) {
      setError("Usuario no encontrado o inactivo.");
      setLoading(false); return;
    }

    const user = data as LoginUser;

    if (String(user.password) !== cleanPass) {
      setError("Contraseña incorrecta.");
      setLoading(false); return;
    }

    const freeAccess = user.role === "Administrador" || user.can_access_any_inventory === true;
    if (!freeAccess && !user.inventory_id) {
      setError("Este usuario no tiene inventario asignado. Contacta al administrador.");
      setLoading(false); return;
    }

    const sessionUser: SessionUser = {
      id: user.id,
      username: user.username,
      full_name: user.full_name,
      role: user.role,
      inventory_id: user.inventory_id,
      can_access_any_inventory: user.can_access_any_inventory ?? false,
      login_at: Date.now(),
      expires_at: Date.now() + SESSION_HOURS * 60 * 60 * 1000,
    };

    localStorage.setItem(SESSION_KEY, JSON.stringify(sessionUser));
    if (!freeAccess && user.inventory_id) {
      localStorage.setItem(CURRENT_INVENTORY_KEY, user.inventory_id);
    } else {
      localStorage.removeItem(CURRENT_INVENTORY_KEY);
    }

    window.location.href = "/dashboard";
  }

  // ── Abrir modal cambiar contraseña (verifica credenciales primero) ──────
  async function openChangePass() {
    setError("");
    const cleanUser = username.trim();
    if (!cleanUser || !password) {
      setError("Ingresa usuario y contraseña primero para cambiarla.");
      return;
    }
    setLoading(true);
    const { data, error: dbError } = await supabase
      .from("app_users")
      .select("id, username, password, full_name, role, roles, is_active, inventory_id, can_access_any_inventory, can_see_system_stock, can_see_cost, can_see_valued_difference")
      .eq("username", cleanUser)
      .eq("is_active", true)
      .single();

    setLoading(false);
    if (dbError || !data) { setError("Usuario o contraseña incorrectos."); return; }
    if (String((data as LoginUser).password) !== password) {
      setError("Contraseña incorrecta."); return;
    }
    setPendingUser(data as LoginUser);
    setNewPass(""); setConfirmPass(""); setModalError("");
    setShowChangePass(true);
  }

  // ── Guardar nueva contraseña ────────────────────────────────────────────
  async function handleChangePassword() {
    if (!pendingUser) return;
    setModalError("");
    if (!newPass || newPass.length < 4) {
      setModalError("La contraseña debe tener al menos 4 caracteres."); return;
    }
    if (newPass !== confirmPass) {
      setModalError("Las contraseñas no coinciden."); return;
    }
    setModalLoading(true);

    const { error: upErr } = await supabase
      .from("app_users")
      .update({ password: newPass })
      .eq("id", pendingUser.id);

    if (upErr) {
      setModalError("Error al actualizar. Intenta de nuevo.");
      setModalLoading(false); return;
    }

    // Iniciar sesión directamente con la nueva contraseña
    const freeAccess = pendingUser.role === "Administrador" || pendingUser.can_access_any_inventory === true;
    const sessionUser: SessionUser = {
      id: pendingUser.id,
      username: pendingUser.username,
      full_name: pendingUser.full_name,
      role: pendingUser.role,
      inventory_id: pendingUser.inventory_id,
      can_access_any_inventory: pendingUser.can_access_any_inventory ?? false,
      login_at: Date.now(),
      expires_at: Date.now() + SESSION_HOURS * 60 * 60 * 1000,
    };
    localStorage.setItem(SESSION_KEY, JSON.stringify(sessionUser));
    if (!freeAccess && pendingUser.inventory_id) {
      localStorage.setItem(CURRENT_INVENTORY_KEY, pendingUser.inventory_id);
    } else {
      localStorage.removeItem(CURRENT_INVENTORY_KEY);
    }
    setModalLoading(false);
    window.location.href = "/dashboard";
  }

  // ── Loading inicial ─────────────────────────────────────────────────────
  if (checkingSession) {
    return (
      <main
        className="min-h-screen flex items-center justify-center p-4 relative overflow-hidden"
        style={{ background: "linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%)" }}
      >
        <div
          className="relative z-10 w-full max-w-sm text-center space-y-4"
          style={{
            background: "rgba(255,255,255,0.05)",
            backdropFilter: "blur(20px)",
            border: "1px solid rgba(255,255,255,0.12)",
            borderRadius: "24px",
            padding: "40px 32px",
          }}
        >
          <svg className="animate-spin h-8 w-8 mx-auto text-orange-400" viewBox="0 0 24 24" fill="none">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
          </svg>
          <p className="text-slate-300 text-sm font-medium">Verificando sesión...</p>
        </div>
      </main>
    );
  }

  // ── UI principal ────────────────────────────────────────────────────────
  return (
    <main
      className="min-h-screen flex items-center justify-center p-4 relative overflow-hidden"
      style={{ background: "linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%)" }}
    >
      {/* Marca de agua RASECORP */}
      <div
        className="absolute inset-0 flex items-center justify-center pointer-events-none select-none"
        style={{ opacity: 0.10 }}
      >
        <svg viewBox="0 0 520 400" xmlns="http://www.w3.org/2000/svg" className="w-[90vw] max-w-2xl">
          <polygon points="130,20 230,20 280,105 230,190 130,190 80,105"
            fill="none" stroke="white" strokeWidth="8" />
          <polygon points="135,32 225,32 270,105 225,178 135,178 90,105"
            fill="none" stroke="#f97316" strokeWidth="3" />
          <text x="180" y="145" textAnchor="middle" fill="white"
            fontSize="110" fontWeight="900" fontFamily="Arial Black, Arial, sans-serif">R</text>
          <text x="310" y="115" textAnchor="start" fill="white"
            fontSize="72" fontWeight="900" fontFamily="Arial Black, Arial, sans-serif" letterSpacing="-2">RASE</text>
          <text x="310" y="182" textAnchor="start" fill="#f97316"
            fontSize="72" fontWeight="900" fontFamily="Arial Black, Arial, sans-serif" letterSpacing="-2">CORP</text>
          <text x="160" y="235" textAnchor="middle" fill="white"
            fontSize="20" fontWeight="400" fontFamily="Arial, sans-serif" letterSpacing="6">SOLUCIONES LOGÍSTICAS</text>
          <line x1="80" y1="248" x2="130" y2="248" stroke="#f97316" strokeWidth="2" />
          <line x1="190" y1="248" x2="240" y2="248" stroke="#f97316" strokeWidth="2" />
          <text x="60"  y="310" textAnchor="middle" fill="white" fontSize="28">🏭</text>
          <text x="145" y="310" textAnchor="middle" fill="white" fontSize="28">📦</text>
          <text x="230" y="310" textAnchor="middle" fill="white" fontSize="28">🚚</text>
          <text x="315" y="310" textAnchor="middle" fill="white" fontSize="28">🔗</text>
          <text x="180" y="370" textAnchor="middle" fill="white"
            fontSize="15" fontFamily="Arial, sans-serif" letterSpacing="1">EFICIENCIA · CONFIANZA · COMPROMISO</text>
        </svg>
      </div>

      {/* Puntos de luz decorativos */}
      <div className="absolute top-20 left-10 w-72 h-72 rounded-full pointer-events-none"
        style={{ background: "radial-gradient(circle, rgba(249,115,22,0.12) 0%, transparent 70%)" }} />
      <div className="absolute bottom-20 right-10 w-96 h-96 rounded-full pointer-events-none"
        style={{ background: "radial-gradient(circle, rgba(99,102,241,0.1) 0%, transparent 70%)" }} />

      {/* Card de login */}
      <div
        className="relative z-10 w-full max-w-sm space-y-5"
        style={{
          background: "rgba(255,255,255,0.05)",
          backdropFilter: "blur(20px)",
          WebkitBackdropFilter: "blur(20px)",
          border: "1px solid rgba(255,255,255,0.12)",
          borderRadius: "24px",
          padding: "36px 32px",
          boxShadow: "0 25px 50px rgba(0,0,0,0.5)",
        }}
      >
        {/* Header con logo */}
        <div className="text-center space-y-1 pb-2">
          <div className="flex items-center justify-center gap-3 mb-3">
            <div style={{
              background: "linear-gradient(135deg, #1e3a5f 0%, #0f2744 100%)",
              borderRadius: "14px",
              padding: "8px",
              boxShadow: "0 4px 16px rgba(249,115,22,0.35), 0 0 0 2px rgba(249,115,22,0.5)",
            }}>
              <svg viewBox="0 0 60 60" width="44" height="44">
                <defs>
                  <linearGradient id="hexGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stopColor="#f97316" />
                    <stop offset="100%" stopColor="#c2410c" />
                  </linearGradient>
                </defs>
                <polygon points="30,3 54,17 54,43 30,57 6,43 6,17" fill="url(#hexGrad)" />
                <polygon points="30,3 54,17 54,43 30,57 6,43 6,17"
                  fill="none" stroke="rgba(255,255,255,0.4)" strokeWidth="1.5" />
                <text x="30" y="42" textAnchor="middle" fill="white"
                  fontSize="32" fontWeight="900" fontFamily="Arial Black, sans-serif">R</text>
              </svg>
            </div>
            <div className="text-left">
              <p className="text-white font-black text-xl leading-none tracking-widest"
                style={{ textShadow: "0 2px 8px rgba(249,115,22,0.4)" }}>
                RASE<span style={{ color: "#f97316" }}>CORP</span>
              </p>
              <p className="text-slate-300 text-xs font-semibold tracking-widest leading-none mt-1">
                SOLUCIONES LOGÍSTICAS
              </p>
            </div>
          </div>
          <h1 className="text-2xl font-bold text-white">Conteos Cíclicos</h1>
          <p className="text-slate-400 text-sm">Ingresa con tus credenciales</p>
        </div>

        {/* Error */}
        {error && (
          <div
            className="rounded-2xl p-3 text-sm text-red-300 font-medium flex items-center gap-2"
            style={{ background: "rgba(239,68,68,0.15)", border: "1px solid rgba(239,68,68,0.3)" }}
          >
            <span>⚠️</span> {error}
          </div>
        )}

        {/* Campos */}
        <div className="space-y-3">
          <input
            className="w-full rounded-2xl p-3 text-sm text-white placeholder-slate-400 outline-none transition-all"
            style={{
              background: "rgba(255,255,255,0.08)",
              border: "1px solid rgba(255,255,255,0.15)",
            }}
            placeholder="Usuario"
            value={username}
            onChange={e => setUsername(e.target.value)}
            onKeyDown={e => e.key === "Enter" && handleLogin()}
            autoComplete="username"
          />

          <div className="relative">
            <input
              className="w-full rounded-2xl p-3 text-sm text-white placeholder-slate-400 outline-none pr-12 transition-all"
              style={{
                background: "rgba(255,255,255,0.08)",
                border: "1px solid rgba(255,255,255,0.15)",
              }}
              placeholder="Contraseña"
              type={showPassword ? "text" : "password"}
              value={password}
              onChange={e => setPassword(e.target.value)}
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              autoComplete="current-password"
            />
            <button
              type="button"
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white text-lg px-1 transition-colors"
              onClick={() => setShowPassword(!showPassword)}
            >
              {showPassword ? "🙈" : "👁️"}
            </button>
          </div>
        </div>

        {/* Cambiar contraseña */}
        <div className="text-center">
          <button
            type="button"
            className="text-xs text-orange-400 hover:text-orange-300 underline transition-colors"
            onClick={openChangePass}
          >
            Cambiar contraseña
          </button>
        </div>

        {/* Botón ingresar */}
        <button
          className="w-full rounded-2xl p-3 font-bold text-sm transition-all disabled:opacity-50"
          style={{
            background: loading
              ? "rgba(249,115,22,0.6)"
              : "linear-gradient(135deg, #f97316 0%, #ea580c 100%)",
            color: "white",
            boxShadow: loading ? "none" : "0 4px 15px rgba(249,115,22,0.4)",
          }}
          onClick={handleLogin}
          disabled={loading}
        >
          {loading ? (
            <span className="flex items-center justify-center gap-2">
              <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
              </svg>
              Ingresando...
            </span>
          ) : "Ingresar"}
        </button>

        {/* Footer */}
        <p className="text-center text-slate-500 text-xs pt-1">
          © 2025 RaseCorp · Soluciones Logísticas
        </p>
      </div>

      {/* ════════ MODAL — CAMBIAR CONTRASEÑA ════════ */}
      {showChangePass && pendingUser && (
        <div className="fixed inset-0 bg-black/70 flex items-center justify-center p-4 z-50">
          <div
            className="w-full max-w-sm space-y-5"
            style={{
              background: "rgba(15,23,42,0.98)",
              backdropFilter: "blur(20px)",
              border: "1px solid rgba(99,102,241,0.3)",
              borderRadius: "24px",
              padding: "32px 28px",
              boxShadow: "0 25px 50px rgba(0,0,0,0.7)",
            }}
          >
            <div className="text-center space-y-2">
              <div className="text-4xl">🔐</div>
              <h2 className="text-xl font-bold text-white">Cambiar contraseña</h2>
              <p className="text-slate-400 text-sm">
                Hola <b className="text-white">{pendingUser.full_name}</b>, elige tu nueva contraseña.
              </p>
            </div>

            {modalError && (
              <div
                className="rounded-2xl p-3 text-sm text-red-300 font-medium flex items-center gap-2"
                style={{ background: "rgba(239,68,68,0.15)", border: "1px solid rgba(239,68,68,0.3)" }}
              >
                <span>⚠️</span> {modalError}
              </div>
            )}

            <div className="space-y-3">
              {/* Nueva contraseña */}
              <div>
                <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider block mb-1">
                  Nueva contraseña
                </label>
                <div className="relative">
                  <input
                    className="w-full rounded-2xl p-3 text-sm text-white placeholder-slate-500 outline-none pr-12"
                    style={{
                      background: "rgba(255,255,255,0.08)",
                      border: "1px solid rgba(255,255,255,0.15)",
                    }}
                    placeholder="Mínimo 4 caracteres"
                    type={showNewPass ? "text" : "password"}
                    value={newPass}
                    onChange={e => setNewPass(e.target.value)}
                    autoFocus
                  />
                  <button
                    type="button"
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white text-lg px-1"
                    onClick={() => setShowNewPass(!showNewPass)}
                  >
                    {showNewPass ? "🙈" : "👁️"}
                  </button>
                </div>
              </div>

              {/* Confirmar contraseña */}
              <div>
                <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider block mb-1">
                  Confirmar nueva contraseña
                </label>
                <div className="relative">
                  <input
                    className="w-full rounded-2xl p-3 text-sm text-white placeholder-slate-500 outline-none pr-12"
                    style={{
                      background: "rgba(255,255,255,0.08)",
                      border: "1px solid rgba(255,255,255,0.15)",
                    }}
                    placeholder="Repite la contraseña"
                    type={showConfirmPass ? "text" : "password"}
                    value={confirmPass}
                    onChange={e => setConfirmPass(e.target.value)}
                    onKeyDown={e => e.key === "Enter" && handleChangePassword()}
                  />
                  <button
                    type="button"
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white text-lg px-1"
                    onClick={() => setShowConfirmPass(!showConfirmPass)}
                  >
                    {showConfirmPass ? "🙈" : "👁️"}
                  </button>
                </div>
              </div>
            </div>

            <div className="flex gap-3">
              <button
                className="flex-1 rounded-2xl p-3 font-bold text-sm transition-all disabled:opacity-50"
                style={{
                  background: modalLoading
                    ? "rgba(99,102,241,0.6)"
                    : "linear-gradient(135deg, #6366f1 0%, #4f46e5 100%)",
                  color: "white",
                  boxShadow: modalLoading ? "none" : "0 4px 15px rgba(99,102,241,0.4)",
                }}
                onClick={handleChangePassword}
                disabled={modalLoading}
              >
                {modalLoading ? (
                  <span className="flex items-center justify-center gap-2">
                    <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
                    </svg>
                    Guardando...
                  </span>
                ) : "Guardar y entrar"}
              </button>
              <button
                className="px-4 py-3 rounded-2xl border font-semibold text-slate-400 text-sm border-slate-600 hover:border-slate-400 transition-colors"
                onClick={() => {
                  setShowChangePass(false);
                  setPendingUser(null);
                  setModalError("");
                  setNewPass("");
                  setConfirmPass("");
                }}
                disabled={modalLoading}
              >
                Cancelar
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
