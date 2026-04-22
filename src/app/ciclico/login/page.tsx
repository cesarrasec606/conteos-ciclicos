"use client";

import { useState } from "react";
import { supabase } from "@/lib/supabase/client";

// Clave de sesión propia — NO comparte nada con el sistema de inventario general
export const CYCLIC_SESSION_KEY = "cyclic_session_user";

export default function CiclicoLoginPage() {
    const [username, setUsername] = useState("");
    const [password, setPassword] = useState("");
    const [showPassword, setShowPassword] = useState(false);
    const [loading, setLoading] = useState(false);
    const [message, setMessage] = useState("");

    async function handleLogin() {
        setMessage("");
        if (!username.trim() || !password.trim()) {
            setMessage("Ingresa tu usuario y contraseña.");
            return;
        }
        setLoading(true);
        try {
            const { data, error } = await supabase
                .from("cyclic_users")
                .select("*")
                .eq("username", username.trim())
                .eq("password", password.trim())
                .eq("is_active", true)
                .maybeSingle();

            if (error) {
                setMessage("Error de conexión. Intenta de nuevo.");
                return;
            }
            if (!data) {
                setMessage("Usuario o contraseña incorrectos, o usuario inactivo.");
                return;
            }

            // Guardar sesión en clave propia del sistema cíclico
            localStorage.setItem(CYCLIC_SESSION_KEY, JSON.stringify({
                id: data.id,
                username: data.username,
                full_name: data.full_name,
                role: data.role,
            }));

            window.location.href = "/ciclico";
        } finally {
            setLoading(false);
        }
    }

    function handleKeyDown(e: React.KeyboardEvent) {
        if (e.key === "Enter") handleLogin();
    }

    return (
        <main className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-700 flex items-center justify-center p-4">
            <div className="w-full max-w-md space-y-6">

                {/* Logo / título */}
                <div className="text-center">
                    <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-white/10 mb-4">
                        <span className="text-3xl">🔄</span>
                    </div>
                    <h1 className="text-3xl font-bold text-white">Conteos Cíclicos</h1>
                    <p className="text-slate-400 mt-1 text-sm">Ingresa con tu usuario y contraseña</p>
                </div>

                {/* Card de login */}
                <div className="bg-white rounded-3xl p-8 shadow-2xl space-y-5">
                    <div>
                        <label className="block text-sm font-semibold text-slate-700 mb-2">
                            Usuario
                        </label>
                        <input
                            className="w-full border border-slate-200 rounded-2xl p-3 text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-400"
                            placeholder="Ej: jperez"
                            value={username}
                            onChange={e => setUsername(e.target.value)}
                            onKeyDown={handleKeyDown}
                            autoComplete="username"
                            autoFocus
                        />
                    </div>

                    <div>
                        <label className="block text-sm font-semibold text-slate-700 mb-2">
                            Contraseña
                        </label>
                        <div className="flex gap-2">
                            <input
                                className="flex-1 border border-slate-200 rounded-2xl p-3 text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-400"
                                placeholder="••••••••"
                                type={showPassword ? "text" : "password"}
                                value={password}
                                onChange={e => setPassword(e.target.value)}
                                onKeyDown={handleKeyDown}
                                autoComplete="current-password"
                            />
                            <button
                                type="button"
                                className="px-4 rounded-2xl border border-slate-200 text-sm text-slate-600 font-medium"
                                onClick={() => setShowPassword(!showPassword)}
                            >
                                {showPassword ? "Ocultar" : "Ver"}
                            </button>
                        </div>
                    </div>

                    {message && (
                        <div className="bg-red-50 border border-red-200 text-red-700 rounded-2xl p-3 text-sm">
                            {message}
                        </div>
                    )}

                    <button
                        className={`w-full py-3.5 rounded-2xl font-bold text-white transition ${
                            loading
                                ? "bg-slate-400 cursor-not-allowed"
                                : "bg-slate-900 hover:bg-slate-700"
                        }`}
                        onClick={handleLogin}
                        disabled={loading}
                    >
                        {loading ? "Ingresando..." : "Ingresar"}
                    </button>
                </div>

                <p className="text-center text-slate-500 text-xs">
                    Sistema de Conteos Cíclicos · Solicita acceso al administrador
                </p>
            </div>
        </main>
    );
}
