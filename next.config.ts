import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  experimental: {
    serverActions: { bodySizeLimit: "15mb" }
  },
  // Garante que os modelos oficiais sejam incluídos nas funções serverless do Vercel.
  outputFileTracingIncludes: {
    "/api/generate/romaneio": ["./public/templates/Romaneio.xlsx"],
    "/api/generate/etiquetas": ["./public/templates/etiqueta-base.png"]
  }
};

export default nextConfig;
