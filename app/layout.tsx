import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = { title: "Famossul | Romaneios", description: "Gerador de romaneios e etiquetas" };
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="pt-BR"><body>{children}</body></html>}
