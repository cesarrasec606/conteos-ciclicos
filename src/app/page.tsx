"use client";

import Image from "next/image";
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

const SESSION_KEY = "session_user";
const CURRENT_INVENTORY_KEY = "current_inventory_id";
const SESSION_HOURS = 8;

export default function LoginPage() {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [checkingSession, setCheckingSession] = useState(true);
  const [message, setMessage] = useState("");

  useEffect(() => {
    checkExistingSession();
  }, []);

  async function checkExistingSession() {
    try {
      const raw = localStorage.getItem(SESSION_KEY);

      if (!raw) {
        setCheckingSession(false);
        return;
      }

      const parsed = JSON.parse(raw) as SessionUser;

      if (!parsed?.id || !parsed?.expires_at) {
        localStorage.removeItem(SESSION_KEY);
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
        setCheckingSession(false);
        return;
      }

      if (Date.now() > parsed.expires_at) {
        localStorage.removeItem(SESSION_KEY);
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
        setCheckingSession(false);
        return;
      }

      const { data, error } = await supabase
        .from("app_users")
        .select("id, username, full_name, role, roles, is_active, inventory_id, can_access_any_inventory, can_see_system_stock, can_see_cost, can_see_valued_difference")
        .eq("id", parsed.id)
        .eq("is_active", true)
        .single();

      if (error || !data) {
        localStorage.removeItem(SESSION_KEY);
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
        setCheckingSession(false);
        return;
      }

      const freeAccess = data.role === "Administrador" || data.can_access_any_inventory === true;

      if (!freeAccess && !data.inventory_id) {
        setMessage("Este usuario no tiene inventario asignado. Contacta al administrador.");
        localStorage.removeItem(SESSION_KEY);
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
        setCheckingSession(false);
        return;
      }

      const refreshedSession: SessionUser = {
        id: data.id,
        username: data.username,
        full_name: data.full_name,
        role: data.role,
        inventory_id: data.inventory_id,
        can_access_any_inventory: data.can_access_any_inventory ?? false,
        login_at: parsed.login_at || Date.now(),
        expires_at: Date.now() + SESSION_HOURS * 60 * 60 * 1000,
      };

      localStorage.setItem(SESSION_KEY, JSON.stringify(refreshedSession));

      if (!freeAccess && data.inventory_id) {
        localStorage.setItem(CURRENT_INVENTORY_KEY, data.inventory_id);
      } else {
        localStorage.removeItem(CURRENT_INVENTORY_KEY);
      }

      window.location.href = "/dashboard";
    } catch {
      localStorage.removeItem(SESSION_KEY);
      localStorage.removeItem(CURRENT_INVENTORY_KEY);
      setCheckingSession(false);
    }
  }

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setMessage("");

    const cleanUsername = username.trim();
    const cleanPassword = password.trim();

    if (!cleanUsername || !cleanPassword) {
      setMessage("Completa usuario y contraseña.");
      return;
    }

    setLoading(true);

    try {
      const { data, error } = await supabase
        .from("app_users")
        .select("id, username, password, full_name, role, roles, is_active, inventory_id, can_access_any_inventory, can_see_system_stock, can_see_cost, can_see_valued_difference")
        .eq("username", cleanUsername)
        .eq("is_active", true)
        .single();

      if (error || !data) {
        setMessage("Usuario no encontrado o inactivo.");
        setLoading(false);
        return;
      }

      const user = data as LoginUser;

      if (String(user.password) !== cleanPassword) {
        setMessage("Contraseña incorrecta.");
        setLoading(false);
        return;
      }

      const freeAccess = user.role === "Administrador" || user.can_access_any_inventory === true;

      if (!freeAccess && !user.inventory_id) {
        setMessage("Este usuario no tiene inventario asignado. Contacta al administrador.");
        setLoading(false);
        return;
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
    } catch (err: any) {
      setMessage("No se pudo iniciar sesión: " + (err?.message || "Error desconocido"));
      setLoading(false);
      return;
    }

    setLoading(false);
  }

  if (checkingSession) {
    return (
      <main className="min-h-screen bg-slate-100 flex items-center justify-center p-6">
        <div className="bg-white rounded-3xl shadow-xl p-8 w-full max-w-md text-center">
          <h1 className="text-2xl font-bold text-slate-900">Sistema de Conteos Cíclicos</h1>
          <p className="text-slate-500 mt-2">Verificando sesión...</p>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-slate-100 flex items-center justify-center p-6">
      <div className="w-full max-w-6xl grid lg:grid-cols-2 overflow-hidden rounded-3xl shadow-2xl bg-white">
        <div className="bg-slate-900 text-white p-8 md:p-10 flex flex-col justify-center">
          <div className="mx-auto lg:mx-0 w-full max-w-md">
            <div className="bg-white rounded-2xl p-3 shadow-lg">
              <Image
                src="/rasecorp-logo.png"
                alt="RASECORP"
                width={900}
                height={900}
                className="w-full h-auto rounded-xl"
                priority
              />
            </div>

            <h1 className="text-3xl md:text-4xl font-bold leading-tight mt-6 text-center lg:text-left">
              Sistema Web de
              <br />
              Conteos Cíclicos
            </h1>

            <p className="mt-4 text-slate-300 text-center lg:text-left">
              Plataforma de gestión para conteos cíclicos, validación y control operativo.
            </p>

            <div className="mt-6 space-y-2 text-sm text-slate-300 text-center lg:text-left">
              <div>📞 +51 980 683 349</div>
              <div>✉️ RasecorpSL@gmail.com</div>
              <div>IG: RaseCorpSL</div>
              <div>FB: RaseCorpSL</div>
            </div>
          </div>
        </div>

        <div className="p-8 md:p-10 flex items-center bg-slate-50">
          <div className="w-full max-w-md mx-auto">
            <h2 className="text-3xl font-bold text-slate-900">Iniciar sesión</h2>
            <p className="text-slate-500 mt-2">Ingresa tu usuario y contraseña.</p>

            {message && (
              <div className="mt-6 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {message}
              </div>
            )}

            <form onSubmit={handleLogin} className="mt-6 space-y-5">
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-2">
                  Usuario
                </label>
                <input
                  className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none focus:border-slate-900"
                  placeholder="Ej. admin"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  autoComplete="username"
                />
              </div>

              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-2">
                  Contraseña
                </label>
                <div className="flex gap-2">
                  <input
                    className="w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 outline-none focus:border-slate-900"
                    type={showPassword ? "text" : "password"}
                    placeholder="Ingresa tu contraseña"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    autoComplete="current-password"
                  />
                  <button
                    type="button"
                    className="rounded-2xl border border-slate-300 bg-white px-4"
                    onClick={() => setShowPassword((v) => !v)}
                  >
                    {showPassword ? "Ocultar" : "Ver"}
                  </button>
                </div>
              </div>

              <button
                type="submit"
                disabled={loading}
                className="w-full rounded-2xl bg-slate-900 text-white py-3 font-semibold disabled:opacity-60"
              >
                {loading ? "Ingresando..." : "Entrar"}
              </button>
            </form>
          </div>
        </div>
      </div>
    </main>
  );
}