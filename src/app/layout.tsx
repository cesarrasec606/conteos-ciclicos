import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Sistema Web de Conteos Cíclicos",
  description: "RASECORP - Soluciones Logísticas",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}