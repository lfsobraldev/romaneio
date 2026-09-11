import type { LogisticsConfig, MountType } from "./types";

export const MOUNT_LABELS: Record<MountType, string> = {
  MONTADO_HS: "MONTADO HS",
  MONTADO_TIMADEL: "MONTADO TIMADEL",
  COMPLEMENTO_OBRA: "COMPLEMENTO OBRA",
  REVENDA: "REVENDA",
  MONTADO_ESTANCIA: "MONTADO ESTÂNCIA",
};

export const DEFAULT_LOGISTICS_CONFIG: LogisticsConfig = {
  MONTADO_HS: { maxDoors: 32, maxM3: 2.0, maxTrimGames: 250, maxFrameGames: 250 },
  MONTADO_TIMADEL: { maxDoors: 34, maxM3: 2.0, maxTrimGames: 250, maxFrameGames: 250 },
  COMPLEMENTO_OBRA: { maxDoors: 34, maxM3: 2.0, maxTrimGames: 250, maxFrameGames: 250 },
  REVENDA: { maxDoors: 34, maxDoorsCardboard: 28, maxDoorsThickness41: 28, maxM3: 1.8, maxTrimGames: 100, maxFrameGames: 100 },
  MONTADO_ESTANCIA: { maxDoors: 34, maxM3: 2.0, maxTrimGames: 250, maxFrameGames: 250 },
  mixedMaxM3: 1.4,
  additionalItems: [],
  romaneioNote: "",
  labelNote: "",
};

export function sanitizeConfig(input?: Partial<LogisticsConfig>): LogisticsConfig {
  if (!input) return structuredClone(DEFAULT_LOGISTICS_CONFIG);
  const d = DEFAULT_LOGISTICS_CONFIG;
  const safe = (v: unknown, fallback: number) => typeof v === "number" && Number.isFinite(v) && v > 0 ? v : fallback;
  const rule = (key: Exclude<MountType, "REVENDA">) => ({
    maxDoors: safe(input[key]?.maxDoors, d[key].maxDoors),
    maxM3: safe(input[key]?.maxM3, d[key].maxM3),
    maxTrimGames: safe(input[key]?.maxTrimGames, d[key].maxTrimGames),
    maxFrameGames: safe(input[key]?.maxFrameGames, d[key].maxFrameGames),
  });
  return {
    MONTADO_HS: rule("MONTADO_HS"),
    MONTADO_TIMADEL: rule("MONTADO_TIMADEL"),
    COMPLEMENTO_OBRA: rule("COMPLEMENTO_OBRA"),
    MONTADO_ESTANCIA: rule("MONTADO_ESTANCIA"),
    REVENDA: {
      maxDoors: safe(input.REVENDA?.maxDoors, d.REVENDA.maxDoors),
      maxDoorsCardboard: safe(input.REVENDA?.maxDoorsCardboard, d.REVENDA.maxDoorsCardboard),
      maxDoorsThickness41: safe(input.REVENDA?.maxDoorsThickness41, d.REVENDA.maxDoorsThickness41),
      maxM3: safe(input.REVENDA?.maxM3, d.REVENDA.maxM3),
      maxTrimGames: safe(input.REVENDA?.maxTrimGames, d.REVENDA.maxTrimGames),
      maxFrameGames: safe(input.REVENDA?.maxFrameGames, d.REVENDA.maxFrameGames),
    },
    mixedMaxM3: safe(input.mixedMaxM3, d.mixedMaxM3),
    additionalItems: Array.isArray(input.additionalItems) ? input.additionalItems.filter(Boolean) as LogisticsConfig["additionalItems"] : [],
    romaneioNote: typeof input.romaneioNote === "string" ? input.romaneioNote : "",
    labelNote: typeof input.labelNote === "string" ? input.labelNote : "",
  };
}
