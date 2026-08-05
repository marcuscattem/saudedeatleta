import { useAuth } from "@/_core/hooks/useAuth";
import { Button } from "@/components/ui/button";
import { getLoginUrl } from "@/const";
import ExcelJS from "exceljs";
import {
  Activity,
  ArrowDown,
  ArrowRight,
  CalendarCheck2,
  Calculator,
  Check,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  Download,
  Dumbbell,
  ExternalLink,
  GraduationCap,
  Info,
  LogIn,
  Monitor,
  Ruler,
  UserPlus,
} from "lucide-react";
import { useEffect, useState } from "react";
import type * as React from "react";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";

const EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PERIMETER_ETM_LIMIT = 1;
const SKINFOLD_ETM_LIMIT = 5;

type MeasurementKind = "perimeter" | "skinfold";

/* ------------------------------------------------------------------ *
 *  Design tokens
 * ------------------------------------------------------------------ */
const FONTS_HREF =
  "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap";
const FONT_HEAD = "'Space Grotesk', system-ui, sans-serif";
const FONT_BODY = "'Manrope', system-ui, sans-serif";

const C = {
  navy: "#14203A",
  surface: "#F7F8FA",
  card: "#FFFFFF",
  border: "#E4E8EF",
  ink: "#14203A",
  muted: "#6B7585",
  faint: "#94A0B3",
  teal: "#0E9C8E",
  tealDark: "#0B7E72",
  tealSoft: "#E9F6F3",
  indigo: "#5B6CD6",
  indigoDark: "#4250B8",
  indigoSoft: "#EEF0FB",
  danger: "#C8453A",
  dangerSoft: "#FBE9E7",
};

const INJECTED_CSS = `
.sa-input::placeholder{color:#AAB3C2;}
.sa-input:focus{border-color:${C.teal}!important;background:#fff!important;}
.sa-input-indigo:focus{border-color:${C.indigo}!important;background:#fff!important;}
.sa-tap{transition:transform .08s ease;}
.sa-tap:active{transform:scale(.985);}
.sa-num::-webkit-inner-spin-button,.sa-num::-webkit-outer-spin-button{-webkit-appearance:none;margin:0;}
.sa-num{-moz-appearance:textfield;}
`;

/* ------------------------------------------------------------------ *
 *  Anthropometry (perimeter-only protocol)
 * ------------------------------------------------------------------ */
const antropoFields = [
  { key: "braco", label: "Braço direito", region: "Membro superior", excelHeader: "Braço (cm)", kind: "perimeter" },
  { key: "cintura", label: "Cintura", region: "Tronco", excelHeader: "Cintura (cm)", kind: "perimeter" },
  { key: "panturrilha", label: "Panturrilha direita", region: "Membro inferior", excelHeader: "Panturrilha (cm)", kind: "perimeter" },
] as const;

/* ------------------------------------------------------------------ *
 *  ISAK fields — collection order: cima -> baixo, dorsal -> ventral.
 *  Todas as DOBRAS primeiro, depois os PERÍMETROS.
 *  Dobras opcionais (torácica/axilar média) entram na posição
 *  anatômica correta (após o bíceps) quando expandido.
 * ------------------------------------------------------------------ */
const isakDobrasBase = [
  {
    key: "subscap",
    label: "Dobra subescapular (mm)",
    short: "Subescapular",
    region: "Dorsal · tronco superior",
    kind: "skinfold",
    description: "Dobra oblíqua a 45 graus, marcada cerca de 2 cm abaixo do Subscapulare.",
  },
  {
    key: "triceps",
    label: "Dobra de tríceps (mm)",
    short: "Tríceps",
    region: "Dorsal · braço",
    kind: "skinfold",
    description: "Dobra vertical na face posterior do braço, no nível do Ponto Médio Acromiale-Radiale.",
  },
  {
    key: "biceps",
    label: "Dobra de bíceps (mm)",
    short: "Bíceps",
    region: "Ventral · braço",
    kind: "skinfold",
    description: "Dobra vertical na face anterior do braço, no mesmo nível do Ponto Médio Acromiale-Radiale.",
  },
] as const;

const isakDobrasOptional = [
  {
    key: "toracica",
    label: "Dobra torácica (mm)",
    short: "Torácica",
    region: "Ventral · tronco",
    kind: "skinfold",
    optional: true,
    description: "Dobra diagonal na região torácica, alinhada entre a prega axilar anterior e o mamilo.",
  },
  {
    key: "axilar_media",
    label: "Dobra axilar média (mm)",
    short: "Axilar média",
    region: "Lateral · tronco",
    kind: "skinfold",
    optional: true,
    description: "Dobra vertical na linha axilar média, no nível do processo xifoide.",
  },
] as const;

const isakDobrasRest = [
  {
    key: "iliaca",
    label: "Dobra de crista ilíaca (mm)",
    short: "Crista ilíaca",
    region: "Lateral · quadril",
    kind: "skinfold",
    description: "Dobra logo acima do Iliocristale, na linha axilar média, levemente inclinada para baixo e para frente.",
  },
  {
    key: "supraesp",
    label: "Dobra supraespinhal (mm)",
    short: "Supraespinhal",
    region: "Ventral · quadril",
    kind: "skinfold",
    description: "Dobra oblíqua a 45 graus na interseção entre a linha do Iliospinale à axila anterior e a linha horizontal do Iliocristale.",
  },
  {
    key: "abdom",
    label: "Dobra abdominal (mm)",
    short: "Abdominal",
    region: "Ventral · abdome",
    kind: "skinfold",
    description: "Dobra vertical localizada 5 cm à direita da cicatriz umbilical.",
  },
  {
    key: "coxa",
    label: "Dobra de coxa anterior (mm)",
    short: "Coxa anterior",
    region: "Ventral · coxa",
    kind: "skinfold",
    description: "Dobra vertical na face anterior da coxa, no ponto médio entre a prega inguinal e o Patellare.",
  },
  {
    key: "pant_dobra",
    label: "Dobra de panturrilha medial (mm)",
    short: "Panturrilha medial",
    region: "Medial · perna",
    kind: "skinfold",
    description: "Dobra vertical na face medial da panturrilha, no nível do maior perímetro da perna.",
  },
] as const;

const isakPerimFields = [
  {
    key: "braco_rel",
    label: "Perímetro de braço relaxado (cm)",
    short: "Braço relaxado",
    region: "Membro superior",
    kind: "perimeter",
    description: "Fita no nível do Ponto Médio Acromiale-Radiale, com o braço relaxado ao lado do corpo.",
  },
  {
    key: "braco_flet",
    label: "Perímetro de braço contraído (cm)",
    short: "Braço contraído",
    region: "Membro superior",
    kind: "perimeter",
    description: "Maior perímetro do braço com cotovelo flexionado a 90 graus e musculatura contraída.",
  },
  {
    key: "torax",
    label: "Perímetro de tórax (cm)",
    short: "Tórax",
    region: "Tronco",
    kind: "perimeter",
    description: "Fita no nível do Mesosternale, passando sob as axilas, ao final de uma expiração normal.",
  },
  {
    key: "cintura",
    label: "Perímetro de cintura (cm)",
    short: "Cintura",
    region: "Tronco",
    kind: "perimeter",
    description: "Fita no ponto mais estreito entre a última costela e a crista ilíaca.",
  },
  {
    key: "abdome_perim",
    label: "Perímetro de abdome (cm)",
    short: "Abdome",
    region: "Tronco",
    kind: "perimeter",
    description: "Fita ao redor do abdome no nível da cicatriz umbilical, sem comprimir a pele.",
  },
  {
    key: "gluteo",
    label: "Perímetro de quadril (cm)",
    short: "Quadril",
    region: "Quadril",
    kind: "perimeter",
    description: "Maior perímetro da região glútea, com o avaliado em pé e pés unidos.",
  },
  {
    key: "coxa_media",
    label: "Perímetro de coxa média (cm)",
    short: "Coxa média",
    region: "Membro inferior",
    kind: "perimeter",
    description: "Fita posicionada na altura da medida da dobra cutânea de coxa anterior, perpendicular ao eixo da coxa.",
  },
  {
    key: "pant_perim",
    label: "Perímetro de panturrilha medial (cm)",
    short: "Panturrilha medial",
    region: "Membro inferior",
    kind: "perimeter",
    description: "Maior perímetro da panturrilha, com a fita perpendicular ao eixo da perna.",
  },
] as const;

type IsakField = {
  key: string;
  label: string;
  short: string;
  region: string;
  kind: MeasurementKind;
  optional?: boolean;
  description: string;
};

function buildIsakFields(expanded: boolean): IsakField[] {
  const dobras = expanded
    ? [...isakDobrasBase, ...isakDobrasOptional, ...isakDobrasRest]
    : [...isakDobrasBase, ...isakDobrasRest];
  return [...dobras, ...isakPerimFields] as unknown as IsakField[];
}

function isakDobrasFor(expanded: boolean): IsakField[] {
  return (expanded
    ? [...isakDobrasBase, ...isakDobrasOptional, ...isakDobrasRest]
    : [...isakDobrasBase, ...isakDobrasRest]) as unknown as IsakField[];
}

const isakTutorialPoints = [
  { name: "Acromiale", description: "Ponto na borda superior e lateral do processo acromial da escápula." },
  { name: "Radiale", description: "Ponto na borda proximal e lateral da cabeça do rádio." },
  { name: "Ponto Médio Acromiale-Radiale", description: "Ponto equidistante entre o Acromiale e o Radiale." },
  { name: "Subscapulare", description: "Ponto mais inferior do ângulo inferior da escápula." },
  { name: "Iliocristale", description: "Ponto mais lateral da borda superior da crista ilíaca." },
  { name: "Iliospinale", description: "Ponto mais inferior e proeminente da espinha ilíaca ântero-superior." },
  { name: "Mesosternale", description: "Ponto no plano sagital médio do esterno, no nível da quarta costela." },
  { name: "Ponto inguinal", description: "Ponto na linha inguinal imaginária entre a espinha ilíaca anterossuperior e a sínfise púbica." },
  { name: "Patellare", description: "Ponto médio na borda superior da patela." },
  { name: "Cicatriz umbilical", description: "Referência central do abdome para a dobra abdominal e o perímetro de abdome." },
] as const;

type Sex = "male" | "female";
type CompositionEquationKey =
  | "slaughter_1988"
  | "jackson_pollock_1978_men"
  | "jackson_pollock_ward_1980_women"
  | "petroski_1995_men"
  | "petroski_1995_women"
  | "guedes_1991_men"
  | "guedes_1991_women"
  | "faulkner_yuhasz_1968"
  | "durnin_womersley_1974";

type CompositionInputs = {
  sex: "" | Sex;
  age: string;
  mass: string;
  stature: string;
  equation: "" | CompositionEquationKey;
};

type CompositionEquation = {
  key: CompositionEquationKey;
  label: string;
  sexes: Sex[];
  ageRange: [number, number];
  foldKeys: string[];
  direct: boolean;
};

type CompositionMetric = {
  label: string;
  value: string;
};

type CompositionEquationComparison = {
  key: CompositionEquationKey;
  label: string;
  isSelected: boolean;
  isPertinent: boolean;
  bodyFatPercent: number;
  density: number;
  foldSum: number;
  method: string;
  requiredFolds: CompositionMetric[];
  warnings: string[];
};

type CompositionResult = {
  participantId: string;
  date: string;
  sex: Sex;
  age: number;
  mass: number;
  stature: number;
  selectedEquationKey: CompositionEquationKey;
  equationLabel: string;
  equationMethod: string;
  bodyFatPercent: number;
  density: number;
  foldSum: number;
  imc: number;
  imcClassification: string;
  rcq: number;
  rce: number;
  cmb: number;
  requiredFolds: CompositionMetric[];
  allMeasurements: CompositionMetric[];
  comparisons: CompositionEquationComparison[];
  warnings: string[];
};

const sexLabels: Record<Sex, string> = {
  male: "Masculino",
  female: "Feminino",
};

const compositionEquations: CompositionEquation[] = [
  {
    key: "slaughter_1988",
    label: "Slaughter et al. (1988)",
    sexes: ["male", "female"],
    ageRange: [8, 17],
    foldKeys: ["triceps", "pant_dobra"],
    direct: true,
  },
  {
    key: "jackson_pollock_1978_men",
    label: "Jackson & Pollock (1978) - homens",
    sexes: ["male"],
    ageRange: [18, 61],
    foldKeys: ["toracica", "abdom", "coxa"],
    direct: false,
  },
  {
    key: "jackson_pollock_ward_1980_women",
    label: "Jackson, Pollock & Ward (1980) - mulheres",
    sexes: ["female"],
    ageRange: [18, 55],
    foldKeys: ["triceps", "iliaca", "coxa"],
    direct: false,
  },
  {
    key: "petroski_1995_men",
    label: "Petroski (1995) - homens",
    sexes: ["male"],
    ageRange: [18, 66],
    foldKeys: ["triceps", "subscap", "iliaca", "pant_dobra"],
    direct: false,
  },
  {
    key: "petroski_1995_women",
    label: "Petroski (1995) - mulheres",
    sexes: ["female"],
    ageRange: [18, 51],
    foldKeys: ["axilar_media", "iliaca", "coxa", "pant_dobra"],
    direct: false,
  },
  {
    key: "guedes_1991_men",
    label: "Guedes (1991) - homens",
    sexes: ["male"],
    ageRange: [18, 30],
    foldKeys: ["triceps", "iliaca", "abdom"],
    direct: false,
  },
  {
    key: "guedes_1991_women",
    label: "Guedes (1991) - mulheres",
    sexes: ["female"],
    ageRange: [18, 30],
    foldKeys: ["subscap", "iliaca", "coxa"],
    direct: false,
  },
  {
    key: "faulkner_yuhasz_1968",
    label: "Faulkner-Yuhasz (1968)",
    sexes: ["male", "female"],
    ageRange: [18, 40],
    foldKeys: ["subscap", "triceps", "iliaca", "abdom"],
    direct: true,
  },
  {
    key: "durnin_womersley_1974",
    label: "Durnin & Womersley (1974)",
    sexes: ["male", "female"],
    ageRange: [16, 72],
    foldKeys: ["subscap", "triceps", "biceps", "iliaca"],
    direct: false,
  },
];

/* ------------------------------------------------------------------ *
 *  Helpers (unchanged math/export logic)
 * ------------------------------------------------------------------ */
function dateStamp(date = new Date()) {
  return date.toISOString().split("T")[0];
}

function safeFilenamePart(value: string) {
  return value.trim().replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "participante";
}

function collectionFilename(date: string, participantId: string) {
  return `${dateStamp(new Date(date))}_${safeFilenamePart(participantId)}.xlsx`;
}

function styleHeader(worksheet: ExcelJS.Worksheet, color: string) {
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
}

function getEtmLimit(kind: MeasurementKind) {
  return kind === "skinfold" ? SKINFOLD_ETM_LIMIT : PERIMETER_ETM_LIMIT;
}

function calculateMean(values: number[]) {
  const validValues = values.filter(Number.isFinite);
  if (validValues.length === 0) return Number.NaN;
  return validValues.reduce((sum, value) => sum + value, 0) / validValues.length;
}

function calculateMax(values: number[]) {
  const validValues = values.filter(Number.isFinite);
  return validValues.length > 0 ? Math.max(...validValues) : Number.NaN;
}

function calculateEtmPercent(values: number[]) {
  if (values.length < 2 || values.some((value) => !Number.isFinite(value))) {
    return Number.NaN;
  }
  const mean = calculateMean(values);
  if (!Number.isFinite(mean) || mean === 0) return Number.NaN;

  const pairwiseDiffs: number[] = [];
  for (let firstIndex = 0; firstIndex < values.length - 1; firstIndex++) {
    for (let secondIndex = firstIndex + 1; secondIndex < values.length; secondIndex++) {
      pairwiseDiffs.push(values[firstIndex] - values[secondIndex]);
    }
  }
  const sumSquaredDiffs = pairwiseDiffs.reduce((sum, diff) => sum + diff ** 2, 0);
  const etm = Math.sqrt(sumSquaredDiffs / (2 * pairwiseDiffs.length));
  return (etm / Math.abs(mean)) * 100;
}

function formatNumber(value: number, fractionDigits = 2) {
  return Number.isFinite(value) ? value.toFixed(fractionDigits) : "Pend.";
}

function parseNumericInput(value: string) {
  return Number(value.replace(",", "."));
}

function formatMaybeNumber(value: number, fractionDigits = 2, suffix = "") {
  return Number.isFinite(value) ? `${value.toFixed(fractionDigits)}${suffix}` : "Não calculado";
}

function classifyImc(imc: number) {
  if (!Number.isFinite(imc)) return "Não calculado";
  if (imc < 18.5) return "Baixo peso";
  if (imc < 25) return "Eutrofia";
  if (imc < 30) return "Sobrepeso";
  if (imc < 35) return "Obesidade grau I";
  if (imc < 40) return "Obesidade grau II";
  return "Obesidade grau III";
}

function siriFromDensity(density: number) {
  return (4.95 / density - 4.5) * 100;
}

function durninWomersleyDensity(sex: Sex, age: number, foldSum: number) {
  const logSum = Math.log10(foldSum);
  if (sex === "male") {
    if (age < 20) return 1.1620 - 0.0630 * logSum;
    if (age < 30) return 1.1631 - 0.0632 * logSum;
    if (age < 40) return 1.1422 - 0.0544 * logSum;
    if (age < 50) return 1.1620 - 0.0700 * logSum;
    return 1.1715 - 0.0779 * logSum;
  }
  if (age < 20) return 1.1549 - 0.0678 * logSum;
  if (age < 30) return 1.1599 - 0.0717 * logSum;
  if (age < 40) return 1.1423 - 0.0632 * logSum;
  if (age < 50) return 1.1333 - 0.0612 * logSum;
  return 1.1339 - 0.0645 * logSum;
}

function calculateBodyFatPercent(equationKey: CompositionEquationKey, sex: Sex, age: number, foldSum: number) {
  const logSum = Math.log10(foldSum);
  let density = Number.NaN;
  let bodyFatPercent = Number.NaN;
  let method = "Conversão por Siri";

  switch (equationKey) {
    case "slaughter_1988":
      method = "Equação direta";
      bodyFatPercent = sex === "male" ? 0.735 * foldSum + 1 : 0.61 * foldSum + 5.1;
      break;
    case "jackson_pollock_1978_men":
      density = 1.10938 - 0.0008267 * foldSum + 0.0000016 * foldSum ** 2 - 0.0002574 * age;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "jackson_pollock_ward_1980_women":
      density = 1.0994921 - 0.0009929 * foldSum + 0.0000023 * foldSum ** 2 - 0.0001392 * age;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "petroski_1995_men":
      density = 1.10726863 - 0.00081201 * foldSum + 0.00000212 * foldSum ** 2 - 0.00041761 * age;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "petroski_1995_women":
      density = 1.1954713 - 0.07513507 * logSum - 0.00041072 * age;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "guedes_1991_men":
      density = 1.1736 - 0.06706 * logSum;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "guedes_1991_women":
      density = 1.1665 - 0.07063 * logSum;
      bodyFatPercent = siriFromDensity(density);
      break;
    case "faulkner_yuhasz_1968":
      method = "Equação direta";
      bodyFatPercent = 0.153 * foldSum + 5.783;
      break;
    case "durnin_womersley_1974":
      density = durninWomersleyDensity(sex, age, foldSum);
      bodyFatPercent = siriFromDensity(density);
      break;
  }

  return { bodyFatPercent, density, method };
}

function getEquationApplicabilityWarnings(equation: CompositionEquation, sex: Sex, age: number) {
  const warnings: string[] = [];
  if (!equation.sexes.includes(sex)) {
    warnings.push(`Equação originalmente indicada para: ${equation.sexes.map((item) => sexLabels[item]).join(" ou ")}.`);
  }
  if (age < equation.ageRange[0] || age > equation.ageRange[1]) {
    warnings.push(`Idade fora da faixa original da equação (${equation.ageRange[0]}-${equation.ageRange[1]} anos).`);
  }
  return warnings;
}

function getFpmStats(rightMeasurements: number[], leftMeasurements: number[]) {
  const rightAverage = calculateMean(rightMeasurements);
  const leftAverage = calculateMean(leftMeasurements);
  const rightMax = calculateMax(rightMeasurements);
  const leftMax = calculateMax(leftMeasurements);
  const allMeasurements = [...rightMeasurements, ...leftMeasurements];
  return {
    rightAverage,
    leftAverage,
    generalAverage: calculateMean(allMeasurements),
    rightMax,
    leftMax,
    generalMax: calculateMax(allMeasurements),
  };
}

const EVAL_COUNT_KEY = "sa_eval_count";

function readTodayEvalCount() {
  const today = dateStamp();
  try {
    const raw = localStorage.getItem(EVAL_COUNT_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as { date?: string; count?: number };
      if (parsed && parsed.date === today) return parsed.count ?? 0;
    }
  } catch {
    /* ignore */
  }
  return 0;
}

/* ================================================================== */
export default function Home() {
  const { user, isAuthenticated, refresh } = useAuth();
  const isStaticPages =
    import.meta.env.BASE_URL === "/saudedeatleta/" ||
    (typeof window !== "undefined" && window.location.hostname.endsWith("github.io"));
  const canUseApp = isAuthenticated || isStaticPages;
  const [activeApp, setActiveApp] = useState("home");
  const [authMode, setAuthMode] = useState<"signin" | "signup">("signin");
  const [authName, setAuthName] = useState("");
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [participantId, setParticipantId] = useState("");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);

  // Daily evaluation counter
  const [evalCount, setEvalCount] = useState(0);

  // Antropometria state
  const [antropoRound, setAntropoRound] = useState(1);
  const [antropoData, setAntropoData] = useState({
    braco: [] as number[],
    cintura: [] as number[],
    panturrilha: [] as number[],
  });
  const [antropoInputs, setAntropoInputs] = useState({ braco: "", cintura: "", panturrilha: "" });
  const [antropoReview, setAntropoReview] = useState(false);
  const [antropoEtmOverrideConfirmed, setAntropoEtmOverrideConfirmed] = useState(false);

  // FPM state
  const [fpmRound, setFpmRound] = useState(1);
  const [fpmDominantHand, setFpmDominantHand] = useState("");
  const [fpmBestLeg, setFpmBestLeg] = useState("");
  const [fpmData, setFpmData] = useState({ right: [] as number[], left: [] as number[] });
  const [fpmInputs, setFpmInputs] = useState({ right: "", left: "" });

  // ISAK state
  const [isakRound, setIsakRound] = useState(1);
  const [isakData, setIsakData] = useState<Record<string, number[]>>({});
  const [isakInputs, setIsakInputs] = useState<Record<string, string>>({});
  const [isakReview, setIsakReview] = useState(false);
  const [isakEtmOverrideConfirmed, setIsakEtmOverrideConfirmed] = useState(false);
  const [isakExpandedSkinfolds, setIsakExpandedSkinfolds] = useState(false);
  const [isakTutorialStep, setIsakTutorialStep] = useState<"points" | "measurements">("measurements");
  const [isakTutorialCheckedPoints, setIsakTutorialCheckedPoints] = useState<Record<string, boolean>>({});
  const [isakCompositionStep, setIsakCompositionStep] = useState<"none" | "prompt" | "form" | "result">("none");
  const [compositionInputs, setCompositionInputs] = useState<CompositionInputs>({
    sex: "",
    age: "",
    mass: "",
    stature: "",
    equation: "",
  });
  const [compositionResult, setCompositionResult] = useState<CompositionResult | null>(null);

  // Mutations
  const saveAntropoMutation = trpc.evaluations.saveAntropometria.useMutation();
  const saveFpmMutation = trpc.evaluations.saveFpm.useMutation();
  const saveIsakMutation = trpc.evaluations.saveIsak.useMutation();
  const signInMutation = trpc.auth.signIn.useMutation();
  const signUpMutation = trpc.auth.signUp.useMutation();
  const isAuthPending = signInMutation.isPending || signUpMutation.isPending;

  /* ----- font + css injection ----- */
  useEffect(() => {
    if (!document.getElementById("sa-fonts")) {
      const link = document.createElement("link");
      link.id = "sa-fonts";
      link.rel = "stylesheet";
      link.href = FONTS_HREF;
      document.head.appendChild(link);
    }
    if (!document.getElementById("sa-css")) {
      const style = document.createElement("style");
      style.id = "sa-css";
      style.textContent = INJECTED_CSS;
      document.head.appendChild(style);
    }
  }, []);

  /* ----- load today's evaluation count ----- */
  useEffect(() => {
    setEvalCount(readTodayEvalCount());
  }, []);

  const incrementEvalCount = () => {
    const today = dateStamp();
    const base = readTodayEvalCount();
    const next = base + 1;
    try {
      localStorage.setItem(EVAL_COUNT_KEY, JSON.stringify({ date: today, count: next }));
    } catch {
      /* ignore */
    }
    setEvalCount(next);
  };

  const handleAuthSubmit = async () => {
    try {
      if (authMode === "signup") {
        await signUpMutation.mutateAsync({ name: authName.trim(), email: authEmail.trim(), password: authPassword });
        toast.success("Conta criada com sucesso!");
      } else {
        await signInMutation.mutateAsync({ email: authEmail.trim(), password: authPassword });
        toast.success("Login realizado!");
      }
      setAuthName("");
      setAuthPassword("");
      await refresh();
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Erro ao autenticar");
    }
  };

  const downloadExcel = (excelResult: { data: string; filename: string }) => {
    const link = document.createElement("a");
    link.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${excelResult.data}`;
    link.download = excelResult.filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
  };

  const downloadWorkbook = async (workbook: ExcelJS.Workbook, filename: string) => {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer as BlobPart], { type: EXCEL_CONTENT_TYPE });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const generateLocalAntropoExcel = async (data: {
    participantId: string;
    date: string;
    bracoMeasurements: number[];
    cinturaMeasurements: number[];
    panturrilhaMeasurements: number[];
  }) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Antropometria");
    worksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Rodada", key: "round", width: 10 },
      { header: "Braço (cm)", key: "braco", width: 15 },
      { header: "Cintura (cm)", key: "cintura", width: 15 },
      { header: "Panturrilha (cm)", key: "panturrilha", width: 15 },
    ];
    data.bracoMeasurements.forEach((braco, index) => {
      worksheet.addRow({
        participantId: data.participantId,
        date: new Date(data.date).toLocaleDateString("pt-BR"),
        round: index + 1,
        braco,
        cintura: data.cinturaMeasurements[index],
        panturrilha: data.panturrilhaMeasurements[index],
      });
    });
    styleHeader(worksheet, "FF4472C4");
    await downloadWorkbook(workbook, `antropometria_${safeFilenamePart(data.participantId)}_${dateStamp()}.xlsx`);
  };

  const generateLocalFpmExcel = async (data: {
    participantId: string;
    date: string;
    dominantHand: string;
    bestLeg: string;
    rightMeasurements: number[];
    leftMeasurements: number[];
  }) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("FPM");
    const stats = getFpmStats(data.rightMeasurements, data.leftMeasurements);
    worksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Mão Dominante", key: "dominantHand", width: 15 },
      { header: "Perna Melhor", key: "bestLeg", width: 15 },
      { header: "Direito 1 (kgf)", key: "right1", width: 15 },
      { header: "Direito 2 (kgf)", key: "right2", width: 15 },
      { header: "Direito 3 (kgf)", key: "right3", width: 15 },
      { header: "Esquerdo 1 (kgf)", key: "left1", width: 15 },
      { header: "Esquerdo 2 (kgf)", key: "left2", width: 15 },
      { header: "Esquerdo 3 (kgf)", key: "left3", width: 15 },
      { header: "Média força lado direito (kgf)", key: "rightAverage", width: 25 },
      { header: "Média força lado esquerdo (kgf)", key: "leftAverage", width: 26 },
      { header: "Média geral (kgf)", key: "generalAverage", width: 18 },
      { header: "Maior força lado direito (kgf)", key: "rightMax", width: 25 },
      { header: "Maior força lado esquerdo (kgf)", key: "leftMax", width: 26 },
      { header: "Maior força geral (kgf)", key: "generalMax", width: 22 },
    ];
    worksheet.addRow({
      participantId: data.participantId,
      date: new Date(data.date).toLocaleDateString("pt-BR"),
      dominantHand: data.dominantHand,
      bestLeg: data.bestLeg,
      right1: data.rightMeasurements[0],
      right2: data.rightMeasurements[1],
      right3: data.rightMeasurements[2],
      left1: data.leftMeasurements[0],
      left2: data.leftMeasurements[1],
      left3: data.leftMeasurements[2],
      rightAverage: stats.rightAverage,
      leftAverage: stats.leftAverage,
      generalAverage: stats.generalAverage,
      rightMax: stats.rightMax,
      leftMax: stats.leftMax,
      generalMax: stats.generalMax,
    });
    styleHeader(worksheet, "FF70AD47");
    await downloadWorkbook(workbook, collectionFilename(data.date, data.participantId));
  };

  function addIsakMeasurementSheetsToWorkbook(
    workbook: ExcelJS.Workbook,
    data: {
      participantId: string;
      date: string;
      measurements: Record<string, number[]>;
      expandedSkinfolds: boolean;
    },
  ) {
    const fields = buildIsakFields(data.expandedSkinfolds);
    const roundsWorksheet = workbook.addWorksheet("Rodadas ISAK");
    roundsWorksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Rodada", key: "round", width: 10 },
      ...fields.map((field) => ({ header: field.label, key: field.key, width: 18 })),
    ];
    const numRounds = data.measurements[fields[0].key]?.length || 0;
    for (let round = 0; round < numRounds; round++) {
      const row: Record<string, unknown> = {
        participantId: data.participantId,
        date: new Date(data.date).toLocaleDateString("pt-BR"),
        round: round + 1,
      };
      fields.forEach((field) => {
        row[field.key] = data.measurements[field.key]?.[round];
      });
      roundsWorksheet.addRow(row);
    }
    styleHeader(roundsWorksheet, "FFC00000");

    const summaryWorksheet = workbook.addWorksheet("Medias e ETM");
    summaryWorksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Medida", key: "measure", width: 28 },
      { header: "Tipo", key: "kind", width: 14 },
      { header: "Rodada 1", key: "round1", width: 12 },
      { header: "Rodada 2", key: "round2", width: 12 },
      { header: "Rodada 3", key: "round3", width: 12 },
      { header: "Média", key: "mean", width: 12 },
      { header: "ETM (%)", key: "etm", width: 12 },
      { header: "Alvo ETM (%)", key: "limit", width: 14 },
      { header: "Status", key: "status", width: 18 },
    ];
    fields.forEach((field) => {
      const values = data.measurements[field.key] ?? [];
      const etmPercent = calculateEtmPercent(values);
      const limit = getEtmLimit(field.kind);
      summaryWorksheet.addRow({
        participantId: data.participantId,
        date: new Date(data.date).toLocaleDateString("pt-BR"),
        measure: field.label,
        kind: field.kind === "skinfold" ? "Dobra cutânea" : "Perímetro",
        round1: values[0],
        round2: values[1],
        round3: values[2],
        mean: calculateMean(values),
        etm: Number.isFinite(etmPercent) ? etmPercent : undefined,
        limit,
        status: Number.isFinite(etmPercent) && etmPercent < limit ? "Dentro do alvo" : "Fora do alvo",
      });
    });
    styleHeader(summaryWorksheet, "FF0E9C8E");
  }

  function addCompositionSheetsToWorkbook(workbook: ExcelJS.Workbook, result: CompositionResult) {
    const compositionWorksheet = workbook.addWorksheet("Composicao");
    compositionWorksheet.columns = [
      { header: "Campo", key: "field", width: 30 },
      { header: "Valor", key: "value", width: 45 },
    ];
    [
      ["ID participante", result.participantId],
      ["Data", new Date(result.date).toLocaleDateString("pt-BR")],
      ["Sexo", sexLabels[result.sex]],
      ["Idade", result.age],
      ["Massa corporal (kg)", result.mass],
      ["Estatura (cm)", result.stature],
      ["Equação em destaque", result.equationLabel],
      ["Pertinência da equação em destaque", result.comparisons.find((comparison) => comparison.isSelected)?.isPertinent ? "Pertinente" : "Não pertinente"],
      ["Método", result.equationMethod],
      ["Percentual de gordura (%)", result.bodyFatPercent],
      ["Densidade corporal", result.density],
      ["Soma das dobras (mm)", result.foldSum],
      ["IMC", result.imc],
      ["Classificação IMC", result.imcClassification],
      ["RCQ", result.rcq],
      ["RCE", result.rce],
      ["CMB corrigida (cm)", result.cmb],
      ["Dobras usadas", result.requiredFolds.map((metric) => `${metric.label}: ${metric.value}`).join(" | ")],
      ["Avisos", result.warnings.join(" | ") || "Sem avisos"],
    ].forEach(([field, value]) => compositionWorksheet.addRow({ field, value }));
    styleHeader(compositionWorksheet, "FF5B6CD6");

    const equationsWorksheet = workbook.addWorksheet("Equacoes");
    equationsWorksheet.columns = [
      { header: "Equação", key: "equation", width: 38 },
      { header: "Destaque", key: "selected", width: 12 },
      { header: "Pertinente", key: "pertinent", width: 14 },
      { header: "Percentual de gordura (%)", key: "bodyFatPercent", width: 24 },
      { header: "Densidade", key: "density", width: 14 },
      { header: "Soma das dobras (mm)", key: "foldSum", width: 22 },
      { header: "Método", key: "method", width: 18 },
      { header: "Dobras usadas", key: "requiredFolds", width: 60 },
      { header: "Avisos", key: "warnings", width: 70 },
    ];
    result.comparisons.forEach((comparison) => {
      equationsWorksheet.addRow({
        equation: comparison.label,
        selected: comparison.isSelected ? "Sim" : "Não",
        pertinent: comparison.isPertinent ? "Sim" : "Não",
        bodyFatPercent: Number.isFinite(comparison.bodyFatPercent) ? comparison.bodyFatPercent : undefined,
        density: Number.isFinite(comparison.density) ? comparison.density : undefined,
        foldSum: Number.isFinite(comparison.foldSum) ? comparison.foldSum : undefined,
        method: comparison.method,
        requiredFolds: comparison.requiredFolds.map((metric) => `${metric.label}: ${metric.value}`).join(" | "),
        warnings: comparison.warnings.join(" | ") || "Sem avisos",
      });
    });
    styleHeader(equationsWorksheet, "FF14203A");
  }

  const generateLocalIsakExcel = async (data: {
    participantId: string;
    date: string;
    measurements: Record<string, number[]>;
    expandedSkinfolds: boolean;
    compositionResult?: CompositionResult | null;
  }) => {
    const workbook = new ExcelJS.Workbook();
    addIsakMeasurementSheetsToWorkbook(workbook, data);
    if (data.compositionResult) addCompositionSheetsToWorkbook(workbook, data.compositionResult);
    await downloadWorkbook(
      workbook,
      data.compositionResult
        ? `isak_completo_${safeFilenamePart(data.participantId)}_${dateStamp(new Date(data.date))}.xlsx`
        : `isak_${safeFilenamePart(data.participantId)}_${dateStamp()}.xlsx`,
    );
  };

  // Proteção contra atualização de página sem salvar
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (hasUnsavedChanges && activeApp !== "home") {
        e.preventDefault();
        e.returnValue = "";
        return "";
      }
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [hasUnsavedChanges, activeApp]);

  // Marcar como tendo mudanças não salvas
  useEffect(() => {
    if (activeApp !== "home") setHasUnsavedChanges(true);
  }, [
    activeApp,
    participantId,
    date,
    antropoInputs,
    fpmInputs,
    fpmDominantHand,
    fpmBestLeg,
    isakInputs,
    isakExpandedSkinfolds,
    compositionInputs,
  ]);

  const antropoEtmRows = antropoFields.map((field) => {
    const values = antropoData[field.key];
    const etmPercent = calculateEtmPercent(values);
    const limit = getEtmLimit(field.kind);
    return {
      key: field.key,
      label: field.label,
      values,
      etmPercent,
      limit,
      isValid: Number.isFinite(etmPercent) && etmPercent < limit,
    };
  });

  const activeIsakFields = buildIsakFields(isakExpandedSkinfolds);
  const activeDobras = isakDobrasFor(isakExpandedSkinfolds);
  const isakCanChangeExpandedSkinfolds = isakRound === 1 && Object.keys(isakData).length === 0 && !isakReview;

  const isakEtmRows = activeIsakFields.map((field) => {
    const values = isakData[field.key] ?? [];
    const etmPercent = calculateEtmPercent(values);
    const limit = getEtmLimit(field.kind);
    return {
      key: field.key,
      label: field.short,
      values,
      etmPercent,
      limit,
      isValid: Number.isFinite(etmPercent) && etmPercent < limit,
    };
  });

  const antropoHasInvalidEtm = antropoEtmRows.some((row) => !row.isValid);
  const isakHasInvalidEtm = isakEtmRows.some((row) => !row.isValid);
  const allTutorialPointsChecked = isakTutorialPoints.every((point) => isakTutorialCheckedPoints[point.name]);

  const toggleTutorialPoint = (pointName: string, checked: boolean) => {
    setIsakTutorialCheckedPoints((current) => ({ ...current, [pointName]: checked }));
  };

  const resetAntropo = () => {
    setAntropoRound(1);
    setAntropoData({ braco: [], cintura: [], panturrilha: [] });
    setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
    setAntropoReview(false);
    setAntropoEtmOverrideConfirmed(false);
  };

  const resetFpm = () => {
    setFpmRound(1);
    setFpmData({ right: [], left: [] });
    setFpmInputs({ right: "", left: "" });
    setFpmDominantHand("");
    setFpmBestLeg("");
  };

  const resetIsak = () => {
    setIsakRound(1);
    setIsakData({});
    setIsakInputs({});
    setIsakReview(false);
    setIsakEtmOverrideConfirmed(false);
    setIsakExpandedSkinfolds(false);
    setIsakTutorialStep("measurements");
    setIsakTutorialCheckedPoints({});
    setIsakCompositionStep("none");
    setCompositionInputs({ sex: "", age: "", mass: "", stature: "", equation: "" });
    setCompositionResult(null);
  };

  const startIsakTutorial = () => {
    resetIsak();
    setIsakTutorialStep("points");
  };

  const resetParticipant = () => {
    setParticipantId("");
    setDate(new Date().toISOString().split("T")[0]);
  };

  const resetAfterSave = (app: "antropo" | "fpm" | "isak") => {
    incrementEvalCount();
    setHasUnsavedChanges(false);
    setActiveApp("home");
    if (app === "antropo") resetAntropo();
    if (app === "fpm") resetFpm();
    if (app === "isak") resetIsak();
    resetParticipant();
  };

  const updateAntropoReviewValue = (key: (typeof antropoFields)[number]["key"], index: number, value: string) => {
    const nextValue = value === "" ? Number.NaN : Number(value);
    setAntropoEtmOverrideConfirmed(false);
    setAntropoData((current) => ({
      ...current,
      [key]: current[key].map((measurement, measurementIndex) => (measurementIndex === index ? nextValue : measurement)),
    }));
  };

  const updateIsakReviewValue = (key: string, index: number, value: string) => {
    const nextValue = value === "" ? Number.NaN : Number(value);
    setIsakEtmOverrideConfirmed(false);
    setIsakData((current) => ({
      ...current,
      [key]: (current[key] ?? []).map((measurement, measurementIndex) =>
        measurementIndex === index ? nextValue : measurement,
      ),
    }));
  };

  const buildCompositionResult = (): { result?: CompositionResult; error?: string } => {
    const sex = compositionInputs.sex;
    const age = parseNumericInput(compositionInputs.age);
    const mass = parseNumericInput(compositionInputs.mass);
    const stature = parseNumericInput(compositionInputs.stature);
    const equation = compositionEquations.find((item) => item.key === compositionInputs.equation);

    if (!sex) return { error: "Selecione o sexo para avaliar a pertinência das equações." };
    if (!Number.isFinite(age) || age <= 0) return { error: "Informe uma idade válida." };
    if (!Number.isFinite(mass) || mass <= 0) return { error: "Informe a massa corporal em kg." };
    if (!Number.isFinite(stature) || stature <= 0) return { error: "Informe a estatura em cm." };
    if (!equation) return { error: "Selecione uma equação de composição corporal." };

    const fullFieldList = buildIsakFields(true);
    const fullFieldByKey = new Map(fullFieldList.map((field) => [field.key, field]));
    const meanByKey = new Map(activeIsakFields.map((field) => [field.key, calculateMean(isakData[field.key] ?? [])]));

    const cintura = meanByKey.get("cintura") ?? Number.NaN;
    const quadril = meanByKey.get("gluteo") ?? Number.NaN;
    const bracoRelaxado = meanByKey.get("braco_rel") ?? Number.NaN;
    const triceps = meanByKey.get("triceps") ?? Number.NaN;
    const imc = mass / (stature / 100) ** 2;
    const rcq = Number.isFinite(cintura) && Number.isFinite(quadril) && quadril > 0 ? cintura / quadril : Number.NaN;
    const rce = Number.isFinite(cintura) && stature > 0 ? cintura / stature : Number.NaN;
    const cmb =
      Number.isFinite(bracoRelaxado) && Number.isFinite(triceps)
        ? bracoRelaxado - (Math.PI * triceps) / 10 - (sex === "male" ? 10 : 6.5)
        : Number.NaN;

    const comparisons = compositionEquations.map((item) => {
      const missingFoldLabels = item.foldKeys
        .filter((key) => !Number.isFinite(meanByKey.get(key)))
        .map((key) => fullFieldByKey.get(key)?.short ?? key);
      const applicabilityWarnings = getEquationApplicabilityWarnings(item, sex, age);
      let foldSum = Number.NaN;
      let bodyFatPercent = Number.NaN;
      let density = Number.NaN;
      let method = item.direct ? "Equação direta" : "Conversão por Siri";
      const warnings = [...applicabilityWarnings];

      if (missingFoldLabels.length > 0) {
        warnings.push(`Dobras ausentes: ${missingFoldLabels.join(", ")}.`);
      } else {
        foldSum = item.foldKeys.reduce((sum, key) => sum + (meanByKey.get(key) ?? 0), 0);
        if (!Number.isFinite(foldSum) || foldSum <= 0) {
          warnings.push("Soma das dobras inválida.");
        } else {
          const calculation = calculateBodyFatPercent(item.key, sex, age, foldSum);
          bodyFatPercent = calculation.bodyFatPercent;
          density = calculation.density;
          method = calculation.method;
          if (!Number.isFinite(bodyFatPercent)) warnings.push("Resultado não calculado com esses dados.");
        }
      }

      return {
        key: item.key,
        label: item.label,
        isSelected: item.key === equation.key,
        isPertinent: applicabilityWarnings.length === 0,
        bodyFatPercent,
        density,
        foldSum,
        method,
        requiredFolds: item.foldKeys.map((key) => ({
          label: fullFieldByKey.get(key)?.short ?? key,
          value: formatMaybeNumber(meanByKey.get(key) ?? Number.NaN, 1, " mm"),
        })),
        warnings,
      };
    });
    const selectedComparison = comparisons.find((item) => item.key === equation.key);
    const warnings: string[] = [];

    selectedComparison?.warnings.forEach((warning) => warnings.push(warning));
    if (age > 100) warnings.push("Idade acima de 100 anos: conferir cadastro.");
    if (mass < 25 || mass > 250) warnings.push("Massa corporal fora da faixa usual: conferir valor.");
    if (stature < 100 || stature > 230) warnings.push("Estatura fora da faixa usual: conferir valor.");
    activeIsakFields.forEach((field) => {
      const value = meanByKey.get(field.key) ?? Number.NaN;
      if (field.kind === "skinfold" && value > 60) warnings.push(`${field.short}: dobra acima de 60 mm.`);
      if (field.kind === "perimeter" && value > 250) warnings.push(`${field.short}: perímetro acima de 250 cm.`);
    });

    return {
      result: {
        participantId: participantId.trim(),
        date,
        sex,
        age,
        mass,
        stature,
        selectedEquationKey: equation.key,
        equationLabel: equation.label,
        equationMethod: selectedComparison?.method ?? (equation.direct ? "Equação direta" : "Conversão por Siri"),
        bodyFatPercent: selectedComparison?.bodyFatPercent ?? Number.NaN,
        density: selectedComparison?.density ?? Number.NaN,
        foldSum: selectedComparison?.foldSum ?? Number.NaN,
        imc,
        imcClassification: classifyImc(imc),
        rcq,
        rce,
        cmb,
        requiredFolds: selectedComparison?.requiredFolds ?? [],
        allMeasurements: activeIsakFields.map((field) => ({
          label: field.short,
          value: formatMaybeNumber(meanByKey.get(field.key) ?? Number.NaN, 1, field.kind === "skinfold" ? " mm" : " cm"),
        })),
        comparisons,
        warnings,
      },
    };
  };

  const handleCalculateComposition = () => {
    const calculation = buildCompositionResult();
    if (calculation.error || !calculation.result) {
      toast.error(calculation.error ?? "Não foi possível calcular a composição corporal");
      return;
    }
    setCompositionResult(calculation.result);
    setIsakCompositionStep("result");
  };

  const downloadBlobFile = (filename: string, blob: Blob) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const exportCompositionCsv = () => {
    if (!compositionResult) return;
    const csvRows: Array<[string, string]> = [
      ["ID participante", compositionResult.participantId],
      ["Data", new Date(compositionResult.date).toLocaleDateString("pt-BR")],
      ["Sexo", sexLabels[compositionResult.sex]],
      ["Idade", `${compositionResult.age}`],
      ["Equação", compositionResult.equationLabel],
      ["Método", compositionResult.equationMethod],
      ["Percentual de gordura", formatMaybeNumber(compositionResult.bodyFatPercent, 2, "%")],
      ["Densidade corporal", formatMaybeNumber(compositionResult.density, 4)],
      ["Soma das dobras", formatMaybeNumber(compositionResult.foldSum, 1, " mm")],
      ["IMC", formatMaybeNumber(compositionResult.imc, 2)],
      ["Classificação IMC", compositionResult.imcClassification],
      ["RCQ", formatMaybeNumber(compositionResult.rcq, 3)],
      ["RCE", formatMaybeNumber(compositionResult.rce, 3)],
      ["CMB corrigida", formatMaybeNumber(compositionResult.cmb, 1, " cm")],
      ["Massa corporal", formatMaybeNumber(compositionResult.mass, 1, " kg")],
      ["Estatura", formatMaybeNumber(compositionResult.stature, 1, " cm")],
      ["Avisos", compositionResult.warnings.join(" | ") || "Sem avisos"],
      ["", ""],
      ["Medidas usadas", ""],
      ...compositionResult.allMeasurements.map((metric) => [metric.label, metric.value] as [string, string]),
      ["", ""],
      ["Comparação entre equações", ""],
      ["Equação", "Pertinência | % gordura | Soma | Avisos"],
      ...compositionResult.comparisons.map((comparison) => [
        comparison.label,
        `${comparison.isPertinent ? "Pertinente" : "Não pertinente"} | ${formatMaybeNumber(comparison.bodyFatPercent, 2, "%")} | ${formatMaybeNumber(comparison.foldSum, 1, " mm")} | ${comparison.warnings.join(" | ") || "Sem avisos"}`,
      ] as [string, string]),
    ];
    const csv = csvRows
      .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
      .join("\n");
    downloadBlobFile(
      `composicao_isak_${safeFilenamePart(compositionResult.participantId)}_${dateStamp(new Date(compositionResult.date))}.csv`,
      new Blob([csv], { type: "text/csv;charset=utf-8" }),
    );
  };

  const exportCompositionExcel = async () => {
    if (!compositionResult) return;
    try {
      await generateLocalIsakExcel({
        participantId,
        date,
        measurements: isakData,
        expandedSkinfolds: isakExpandedSkinfolds,
        compositionResult,
      });
    } catch {
      toast.error("Não foi possível gerar o Excel completo");
    }
  };

  const exportCompositionJpeg = async () => {
    if (!compositionResult) return;
    const card = document.getElementById("isak-composition-card");
    if (!card) return;
    const width = Math.ceil(card.scrollWidth);
    const height = Math.ceil(card.scrollHeight);
    const clone = card.cloneNode(true) as HTMLElement;
    clone.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
    clone.style.width = `${width}px`;
    clone.style.boxSizing = "border-box";
    const serializedCard = new XMLSerializer().serializeToString(clone);
    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
        <foreignObject width="100%" height="100%">${serializedCard}</foreignObject>
      </svg>
    `;
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = width * 2;
      canvas.height = height * 2;
      const context = canvas.getContext("2d");
      if (!context) {
        toast.error("Não foi possível gerar o JPEG");
        return;
      }
      context.fillStyle = "#FFFFFF";
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.scale(2, 2);
      context.drawImage(image, 0, 0);
      canvas.toBlob(
        (blob) => {
          if (!blob) {
            toast.error("Não foi possível gerar o JPEG");
            return;
          }
          downloadBlobFile(
            `composicao_isak_${safeFilenamePart(compositionResult.participantId)}_${dateStamp(new Date(compositionResult.date))}.jpeg`,
            blob,
          );
        },
        "image/jpeg",
        0.95,
      );
    };
    image.onerror = () => toast.error("Não foi possível gerar o JPEG");
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  };

  const saveAntropoEvaluation = async (data = antropoData) => {
    const payload = {
      participantId,
      date,
      bracoMeasurements: data.braco,
      cinturaMeasurements: data.cintura,
      panturrilhaMeasurements: data.panturrilha,
    };
    try {
      if (isStaticPages) {
        await generateLocalAntropoExcel(payload);
      } else {
        const result = await saveAntropoMutation.mutateAsync({ ...payload, date: new Date(payload.date) });
        downloadExcel(result);
      }
      toast.success(isStaticPages ? "Antropometria salva em Excel!" : "Antropometria salva com Excel local e online!");
      resetAfterSave("antropo");
    } catch {
      try {
        await generateLocalAntropoExcel(payload);
        toast.success("Antropometria salva em Excel!");
        resetAfterSave("antropo");
      } catch {
        toast.error("Erro ao gerar Excel de antropometria");
      }
    }
  };

  const saveIsakEvaluation = async (data = isakData) => {
    const payload = { participantId, date, measurements: data, expandedSkinfolds: isakExpandedSkinfolds, compositionResult };
    try {
      if (isStaticPages) {
        await generateLocalIsakExcel(payload);
      } else {
        const result = await saveIsakMutation.mutateAsync({
          participantId: payload.participantId,
          measurements: payload.measurements,
          date: new Date(payload.date),
        });
        downloadExcel(result);
      }
      toast.success(isStaticPages ? "ISAK salva em Excel!" : "ISAK salva com Excel local e online!");
      resetAfterSave("isak");
    } catch {
      try {
        await generateLocalIsakExcel(payload);
        toast.success("ISAK salva em Excel!");
        resetAfterSave("isak");
      } catch {
        toast.error("Erro ao gerar Excel ISAK");
      }
    }
  };

  const handleSaveAntropoReview = async () => {
    if (antropoHasInvalidEtm && !antropoEtmOverrideConfirmed) {
      toast.error("Marque a confirmação para salvar com ETM fora do alvo");
      return;
    }
    await saveAntropoEvaluation();
  };

  const handleSaveIsakReview = async () => {
    if (isakHasInvalidEtm && !isakEtmOverrideConfirmed) {
      toast.error("Marque a confirmação para salvar com ETM fora do alvo");
      return;
    }
    if (activeApp === "isakTutorial" && isakCompositionStep === "none") {
      setIsakCompositionStep("prompt");
      return;
    }
    await saveIsakEvaluation();
  };

  const handleAntropoRound = async () => {
    if (!participantId.trim()) {
      toast.error("ID do participante é obrigatório");
      return;
    }
    const braco = parseFloat(antropoInputs.braco);
    const cintura = parseFloat(antropoInputs.cintura);
    const panturrilha = parseFloat(antropoInputs.panturrilha);
    if (isNaN(braco) || isNaN(cintura) || isNaN(panturrilha)) {
      toast.error("Preencha todos os valores");
      return;
    }
    const newData = {
      braco: [...antropoData.braco, braco],
      cintura: [...antropoData.cintura, cintura],
      panturrilha: [...antropoData.panturrilha, panturrilha],
    };
    if (antropoRound < 3) {
      setAntropoData(newData);
      setAntropoRound(antropoRound + 1);
      setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
      toast.success(`Rodada ${antropoRound} salva! Próxima rodada...`);
      focusInputSoon("a-braco");
    } else {
      setAntropoData(newData);
      setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
      setAntropoReview(true);
      setAntropoEtmOverrideConfirmed(false);
      toast.success("Revise o ETM antes de salvar");
    }
  };

  const handleFpmRound = async () => {
    if (!participantId.trim()) {
      toast.error("ID do participante é obrigatório");
      return;
    }
    if (!fpmDominantHand) {
      toast.error("Selecione a mão dominante");
      return;
    }
    if (!fpmBestLeg) {
      toast.error("Selecione a perna melhor");
      return;
    }
    const right = parseFloat(fpmInputs.right);
    const left = parseFloat(fpmInputs.left);
    if (isNaN(right) || isNaN(left)) {
      toast.error("Preencha todos os valores");
      return;
    }
    const newData = { right: [...fpmData.right, right], left: [...fpmData.left, left] };
    if (fpmRound < 3) {
      setFpmData(newData);
      setFpmRound(fpmRound + 1);
      setFpmInputs({ right: "", left: "" });
      toast.success(`Medida ${fpmRound} salva! Próxima medida...`);
      focusInputSoon("f-right");
    } else {
      try {
        const payload = {
          participantId,
          date,
          dominantHand: fpmDominantHand,
          bestLeg: fpmBestLeg,
          rightMeasurements: newData.right,
          leftMeasurements: newData.left,
        };
        if (isStaticPages) {
          await generateLocalFpmExcel(payload);
        } else {
          const result = await saveFpmMutation.mutateAsync({ ...payload, date: new Date(payload.date) });
          downloadExcel(result);
        }
        toast.success(isStaticPages ? "FPM salva em Excel!" : "FPM salva com Excel local e online!");
        incrementEvalCount();
        setHasUnsavedChanges(false);
        setActiveApp("home");
        resetFpm();
        resetParticipant();
      } catch {
        try {
          await generateLocalFpmExcel({
            participantId,
            date,
            dominantHand: fpmDominantHand,
            bestLeg: fpmBestLeg,
            rightMeasurements: newData.right,
            leftMeasurements: newData.left,
          });
          toast.success("FPM salva em Excel!");
          incrementEvalCount();
          setHasUnsavedChanges(false);
          setActiveApp("home");
          resetFpm();
          resetParticipant();
        } catch {
          toast.error("Erro ao gerar Excel de FPM");
        }
      }
    }
  };

  const handleIsakRound = async () => {
    if (!participantId.trim()) {
      toast.error("ID do participante é obrigatório");
      return;
    }
    let valid = true;
    const newData: Record<string, number[]> = { ...isakData };
    activeIsakFields.forEach((field) => {
      const val = parseFloat(isakInputs[field.key] || "");
      if (isNaN(val)) {
        valid = false;
        return;
      }
      if (!newData[field.key]) newData[field.key] = [];
      newData[field.key].push(val);
    });
    if (!valid) {
      toast.error("Preencha todos os valores");
      return;
    }
    if (isakRound < 3) {
      setIsakData(newData);
      setIsakRound(isakRound + 1);
      setIsakInputs({});
      toast.success(`Rodada ${isakRound} salva! Próxima rodada...`);
      focusInputSoon(`inp-${activeIsakFields[0].key}`);
    } else {
      setIsakData(newData);
      setIsakInputs({});
      setIsakReview(true);
      setIsakEtmOverrideConfirmed(false);
      toast.success("Revise o ETM antes de salvar");
    }
  };

  const handleCancel = () => {
    if (hasUnsavedChanges) {
      if (confirm("Você tem alterações não salvas. Deseja descartar?")) {
        setHasUnsavedChanges(false);
        setActiveApp("home");
        resetAntropo();
        resetFpm();
        resetIsak();
        resetParticipant();
      }
    } else {
      setActiveApp("home");
    }
  };

  /* ----- focus helpers (auto-advance field to field) ----- */
  function focusInputSoon(id: string) {
    window.setTimeout(() => {
      const el = document.getElementById(id) as HTMLInputElement | null;
      if (el) {
        el.focus();
        el.select?.();
      }
    }, 60);
  }

  function makeEnterAdvance(ids: string[], current: string, onLast: () => void) {
    return (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const idx = ids.indexOf(current);
      if (idx > -1 && idx < ids.length - 1) {
        const el = document.getElementById(ids[idx + 1]) as HTMLInputElement | null;
        if (el) {
          el.focus();
          el.select?.();
        }
      } else {
        onLast();
      }
    };
  }

  /* ================================================================ *
   *  Auth gate
   * ================================================================ */
  if (!canUseApp) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: C.surface,
          fontFamily: FONT_BODY,
          padding: 18,
        }}
      >
        <div
          style={{
            width: "100%",
            maxWidth: 400,
            background: C.card,
            border: `1px solid ${C.border}`,
            borderRadius: 22,
            padding: 24,
            boxShadow: "0 8px 30px rgba(20,32,58,.08)",
          }}
        >
          <div style={{ fontSize: 11, letterSpacing: ".16em", textTransform: "uppercase", color: C.teal, fontWeight: 700, fontFamily: FONT_HEAD }}>
            Portal Saúde de Atleta
          </div>
          <h1 style={{ margin: "4px 0 18px", fontSize: 24, fontWeight: 800, color: C.ink, letterSpacing: "-.02em" }}>
            {authMode === "signup" ? "Criar conta" : "Entrar"}
          </h1>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 16 }}>
            <Button type="button" variant={authMode === "signin" ? "default" : "outline"} onClick={() => setAuthMode("signin")} className="gap-2">
              <LogIn className="size-4" /> Entrar
            </Button>
            <Button type="button" variant={authMode === "signup" ? "default" : "outline"} onClick={() => setAuthMode("signup")} className="gap-2">
              <UserPlus className="size-4" /> Criar conta
            </Button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {authMode === "signup" && (
              <input className="sa-input" type="text" placeholder="Nome" value={authName} onChange={(e) => setAuthName(e.target.value)} style={inputStyle} />
            )}
            <input className="sa-input" type="email" placeholder="Email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} style={inputStyle} />
            <input className="sa-input" type="password" placeholder="Senha" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} style={inputStyle} />
            <Button onClick={handleAuthSubmit} disabled={isAuthPending} className="w-full gap-2">
              {authMode === "signup" ? <UserPlus className="size-4" /> : <LogIn className="size-4" />}
              {authMode === "signup" ? "Criar conta" : "Entrar"}
            </Button>
            <Button variant="outline" onClick={() => (window.location.href = getLoginUrl())} className="w-full">
              Entrar com login externo
            </Button>
          </div>
        </div>
      </div>
    );
  }

  /* ================================================================ *
   *  Derived view state
   * ================================================================ */
  const isIsak = activeApp === "isak" || activeApp === "isakTutorial";
  const isTutorial = activeApp === "isakTutorial";
  const tutorialPointsStep = isTutorial && isakTutorialStep === "points";
  const isakCompositionActive = isTutorial && isakCompositionStep !== "none";
  const isakCollect = (activeApp === "isak" || (isTutorial && isakTutorialStep === "measurements")) && !isakReview;
  const reviewScreen = ((activeApp === "antropo" && antropoReview) || (isIsak && isakReview)) && !isakCompositionActive;
  const antropoCollect = activeApp === "antropo" && !antropoReview;
  const selectedCompositionEquation = compositionEquations.find((item) => item.key === compositionInputs.equation);
  const selectedCompositionAge = parseNumericInput(compositionInputs.age);
  const selectedCompositionFieldByKey = new Map(buildIsakFields(true).map((field) => [field.key, field]));
  const selectedCompositionMeanByKey = new Map(activeIsakFields.map((field) => [field.key, calculateMean(isakData[field.key] ?? [])]));
  const selectedCompositionNotices =
    selectedCompositionEquation && compositionInputs.sex && Number.isFinite(selectedCompositionAge)
      ? [
          ...getEquationApplicabilityWarnings(selectedCompositionEquation, compositionInputs.sex, selectedCompositionAge),
          ...selectedCompositionEquation.foldKeys
            .filter((key) => !Number.isFinite(selectedCompositionMeanByKey.get(key)))
            .map((key) => `Dobra exigida ainda ausente: ${selectedCompositionFieldByKey.get(key)?.short ?? key}.`),
        ]
      : [];

  const headerInfo: Record<string, { kicker: string; title: string }> = {
    home: { kicker: "", title: "PORTAL SAÚDE DE ATLETA" },
    antropo: { kicker: "Coleta de campo", title: "Antropometria" },
    fpm: { kicker: "Coleta de campo", title: "Força de Preensão" },
    isak: { kicker: "Coleta de campo", title: "Antropometria ISAK 1" },
    isakTutorial: { kicker: "Modo guiado", title: "ISAK Tutorial" },
  };
  const header = headerInfo[activeApp] ?? headerInfo.home;

  const isakFilled = activeIsakFields.filter((f) => {
    const v = isakInputs[f.key];
    return v !== undefined && v !== "" && !isNaN(parseFloat(v));
  }).length;
  const isakProgressPct = Math.round((isakFilled / activeIsakFields.length) * 100);

  const isakInputIds = activeIsakFields.map((f) => `inp-${f.key}`);
  const antropoInputIds = antropoFields.map((f) => `a-${f.key}`);
  const fpmInputIds = ["f-right", "f-left"];

  const showActionBar = antropoCollect || activeApp === "fpm" || tutorialPointsStep || isakCollect || reviewScreen;
  const counterBottom = showActionBar ? 92 : 24;

  function roundDots(round: number, total = 3) {
    return Array.from({ length: total }, (_, i) => {
      const n = i + 1;
      return n < round ? C.teal : n === round ? C.navy : "#D7DDE6";
    });
  }

  /* ================================================================ */
  return (
    <div style={{ minHeight: "100vh", display: "flex", justifyContent: "center", background: "#E7EBF1", fontFamily: FONT_BODY, color: C.ink }}>
      <div
        style={{
          width: "100%",
          maxWidth: 480,
          minHeight: "100vh",
          background: C.surface,
          position: "relative",
          display: "flex",
          flexDirection: "column",
          boxShadow: "0 0 60px rgba(20,32,58,.10)",
        }}
      >
        {/* ---------- sticky top stack ---------- */}
        <div style={{ position: "sticky", top: 0, zIndex: 20 }}>
          <div style={{ background: C.navy, color: "#fff", padding: "18px 18px 16px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, minHeight: 34 }}>
              {activeApp !== "home" && (
                <button onClick={handleCancel} className="sa-tap" aria-label="Voltar" style={iconBtnStyle}>
                  <ChevronLeft size={18} />
                </button>
              )}
              <div style={{ flex: 1, minWidth: 0 }}>
                {header.kicker && (
                  <div style={{ fontSize: 11, letterSpacing: ".16em", textTransform: "uppercase", color: "#7FE3D6", fontWeight: 700, fontFamily: FONT_HEAD }}>
                    {header.kicker}
                  </div>
                )}
                <div style={{ fontSize: 18, fontWeight: 700, letterSpacing: "-.01em", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {header.title}
                </div>
              </div>
              <div
                style={{
                  width: 34,
                  height: 34,
                  borderRadius: 11,
                  background: "rgba(255,255,255,.08)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontFamily: FONT_HEAD,
                  fontWeight: 700,
                  fontSize: 13,
                  color: "#9FB0CC",
                }}
                title={user?.name ?? (isStaticPages ? "Modo público" : "")}
              >
                {(user?.name ?? "SA").slice(0, 2).toUpperCase()}
              </div>
            </div>
          </div>

          {/* ISAK / tutorial sticky progress */}
          {isakCollect && (
            <div style={{ background: C.surface, borderBottom: `1px solid #E7EBF1`, boxShadow: "0 8px 16px rgba(20,32,58,.06)", padding: "11px 18px 12px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 9 }}>
                <div style={{ fontSize: 14, fontWeight: 800 }}>
                  Rodada {isakRound} <span style={{ color: C.faint, fontWeight: 600 }}>de 3</span>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {roundDots(isakRound).map((color, i) => (
                    <span key={i} style={{ width: 24, height: 6, borderRadius: 99, background: color }} />
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 6 }}>
                <span style={{ fontSize: 12, fontWeight: 700, color: C.muted, display: "flex", alignItems: "center", gap: 6 }}>
                  <ArrowDown size={13} color={C.teal} /> Dorsal → ventral, de cima p/ baixo
                </span>
                <span style={{ fontSize: 12, fontWeight: 800, fontFamily: FONT_HEAD, color: C.teal }}>
                  {isakFilled}/{activeIsakFields.length}
                </span>
              </div>
              <div style={{ height: 7, borderRadius: 99, background: "#EEF1F5", overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${isakProgressPct}%`, background: "linear-gradient(90deg,#0E9C8E,#2BC4A8)", borderRadius: 99, transition: "width .3s ease" }} />
              </div>
            </div>
          )}
        </div>

        {/* ---------- body ---------- */}
        <div style={{ flex: 1 }}>
          {/* ===== HOME ===== */}
          {activeApp === "home" && (
            <div style={{ padding: "22px 18px 40px" }}>
              <div style={{ fontSize: 11, letterSpacing: ".14em", textTransform: "uppercase", color: C.faint, fontWeight: 700, margin: "6px 4px 10px" }}>
                Protocolos
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                <MenuCard
                  icon={<Ruler size={24} />}
                  accent="teal"
                  title="Antropometria"
                  subtitle="3 perímetros · 3 rodadas · ETM"
                  onClick={() => {
                    setActiveApp("antropo");
                    resetAntropo();
                    resetParticipant();
                  }}
                />
                <MenuCard
                  icon={<Dumbbell size={24} />}
                  accent="indigo"
                  title="Força de Preensão Manual"
                  subtitle="Direito e esquerdo · 3 medidas"
                  onClick={() => {
                    setActiveApp("fpm");
                    resetFpm();
                    resetParticipant();
                  }}
                />
                <MenuCard
                  icon={<ClipboardList size={24} />}
                  accent="teal"
                  title="Antropometria ISAK 1"
                  subtitle="Dobras + perímetros · sentido padronizado"
                  onClick={() => {
                    setActiveApp("isak");
                    resetIsak();
                    resetParticipant();
                  }}
                />
                <MenuCard
                  icon={<GraduationCap size={24} />}
                  accent="indigo"
                  title="ISAK Tutorial"
                  subtitle="Pontos anatômicos + guia passo a passo"
                  onClick={() => {
                    setActiveApp("isakTutorial");
                    startIsakTutorial();
                    resetParticipant();
                  }}
                />
              </div>

              <div style={{ fontSize: 11, letterSpacing: ".14em", textTransform: "uppercase", color: C.faint, fontWeight: 700, margin: "22px 4px 10px" }}>
                Ferramentas externas
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12 }}>
                <a href="https://ultrakcalc.github.io/UltraKcalc/index.html" target="_blank" rel="noopener noreferrer" className="sa-tap" style={toolCardStyle}>
                  <Calculator size={22} color={C.indigo} />
                  <span style={{ display: "block", fontSize: 14, fontWeight: 700, marginTop: 8 }}>
                    UltraKcalc <ExternalLink size={12} style={{ display: "inline", verticalAlign: "middle", color: C.faint }} />
                  </span>
                  <span style={{ display: "block", fontSize: 12, color: C.faint }}>Composição corporal</span>
                </a>
                <a href="https://marcuscattem.github.io/aiMET/" target="_blank" rel="noopener noreferrer" className="sa-tap" style={toolCardStyle}>
                  <Activity size={22} color={C.teal} />
                  <span style={{ display: "block", fontSize: 14, fontWeight: 700, marginTop: 8 }}>
                    METCalc <ExternalLink size={12} style={{ display: "inline", verticalAlign: "middle", color: C.faint }} />
                  </span>
                  <span style={{ display: "block", fontSize: 12, color: C.faint }}>Gasto energético</span>
                </a>
                <a href="https://marcuscattem.github.io/GUST/" target="_blank" rel="noopener noreferrer" className="sa-tap" style={toolCardStyle}>
                  <Monitor size={22} color={C.indigo} />
                  <span style={{ display: "block", fontSize: 14, fontWeight: 700, marginTop: 8 }}>
                    GUST <ExternalLink size={12} style={{ display: "inline", verticalAlign: "middle", color: C.faint }} />
                  </span>
                  <span style={{ display: "block", fontSize: 12, color: C.faint }}>Ultrassom</span>
                </a>
              </div>
            </div>
          )}

          {/* ===== PARTICIPANT (shared) ===== */}
          {activeApp !== "home" && !tutorialPointsStep && (
            <div style={{ padding: "18px 18px 0" }}>
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 18, padding: 14, boxShadow: "0 2px 10px rgba(20,32,58,.04)", display: "flex", gap: 12 }}>
                <label style={{ flex: 1.4, minWidth: 0 }}>
                  <span style={fieldLabelStyle}>ID do participante</span>
                  <input className="sa-input" value={participantId} onChange={(e) => setParticipantId(e.target.value)} placeholder="Ex.: ATL-014" style={inputStyle} />
                </label>
                <label style={{ flex: 1, minWidth: 0 }}>
                  <span style={fieldLabelStyle}>Data</span>
                  <input className="sa-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} style={inputStyle} />
                </label>
              </div>
            </div>
          )}

          {/* ===== ANTROPOMETRIA collect ===== */}
          {antropoCollect && (
            <div style={{ padding: "16px 18px 130px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", margin: "8px 2px 12px" }}>
                <div style={{ fontSize: 15, fontWeight: 800 }}>
                  Rodada {antropoRound} <span style={{ color: C.faint, fontWeight: 600 }}>de 3</span>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {roundDots(antropoRound).map((color, i) => (
                    <span key={i} style={{ width: 26, height: 6, borderRadius: 99, background: color }} />
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {antropoFields.map((f) => (
                  <SimpleRow
                    key={f.key}
                    label={f.label}
                    region={f.region}
                    unit="cm"
                    accent="teal"
                    inputId={`a-${f.key}`}
                    value={antropoInputs[f.key]}
                    onChange={(v) => setAntropoInputs({ ...antropoInputs, [f.key]: v })}
                    onKeyDown={makeEnterAdvance(antropoInputIds, `a-${f.key}`, handleAntropoRound)}
                  />
                ))}
              </div>
            </div>
          )}

          {/* ===== FPM ===== */}
          {activeApp === "fpm" && (
            <div style={{ padding: "16px 18px 130px" }}>
              <div style={{ display: "flex", gap: 10, margin: "6px 0 14px" }}>
                <label style={{ flex: 1 }}>
                  <span style={fieldLabelStyle}>Mão dominante</span>
                  <select value={fpmDominantHand} onChange={(e) => setFpmDominantHand(e.target.value)} className="sa-input" style={inputStyle}>
                    <option value="">Selecione</option>
                    <option value="Direita">Direita</option>
                    <option value="Esquerda">Esquerda</option>
                  </select>
                </label>
                <label style={{ flex: 1 }}>
                  <span style={fieldLabelStyle}>Perna melhor</span>
                  <select value={fpmBestLeg} onChange={(e) => setFpmBestLeg(e.target.value)} className="sa-input" style={inputStyle}>
                    <option value="">Selecione</option>
                    <option value="Direita">Direita</option>
                    <option value="Esquerda">Esquerda</option>
                  </select>
                </label>
              </div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", margin: "4px 2px 12px" }}>
                <div style={{ fontSize: 15, fontWeight: 800 }}>
                  {fpmRound}ª medida <span style={{ color: C.faint, fontWeight: 600 }}>de 3</span>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {roundDots(fpmRound).map((color, i) => (
                    <span key={i} style={{ width: 26, height: 6, borderRadius: 99, background: color }} />
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <SimpleRow
                  label="Lado direito"
                  region="Dinamômetro"
                  unit="kgf"
                  accent="indigo"
                  inputId="f-right"
                  value={fpmInputs.right}
                  onChange={(v) => setFpmInputs({ ...fpmInputs, right: v })}
                  onKeyDown={makeEnterAdvance(fpmInputIds, "f-right", handleFpmRound)}
                />
                <SimpleRow
                  label="Lado esquerdo"
                  region="Dinamômetro"
                  unit="kgf"
                  accent="indigo"
                  inputId="f-left"
                  value={fpmInputs.left}
                  onChange={(v) => setFpmInputs({ ...fpmInputs, left: v })}
                  onKeyDown={makeEnterAdvance(fpmInputIds, "f-left", handleFpmRound)}
                />
              </div>
            </div>
          )}

          {/* ===== TUTORIAL points ===== */}
          {tutorialPointsStep && (
            <div style={{ padding: "16px 18px 130px" }}>
              <div style={{ padding: "18px 18px 0" }} />
              <div style={{ background: C.indigoSoft, border: "1px solid #DCE0F6", borderRadius: 14, padding: "12px 14px", margin: "6px 0 14px", display: "flex", gap: 10 }}>
                <Info size={18} color={C.indigo} style={{ flex: "none", marginTop: 1 }} />
                <span style={{ fontSize: 13, color: "#3B4663", lineHeight: 1.5 }}>
                  Confirme a marcação de cada ponto anatômico antes de iniciar a coleta. Todos são obrigatórios.
                </span>
              </div>
              <div style={{ fontSize: 11, letterSpacing: ".14em", textTransform: "uppercase", color: C.faint, fontWeight: 700, margin: "0 2px 10px" }}>
                Pontos anatômicos · {isakTutorialPoints.filter((p) => isakTutorialCheckedPoints[p.name]).length}/{isakTutorialPoints.length}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                {isakTutorialPoints.map((p) => {
                  const checked = Boolean(isakTutorialCheckedPoints[p.name]);
                  return (
                    <button
                      key={p.name}
                      onClick={() => toggleTutorialPoint(p.name, !checked)}
                      className="sa-tap"
                      style={{
                        textAlign: "left",
                        cursor: "pointer",
                        background: checked ? C.indigoSoft : "#fff",
                        border: `1.5px solid ${checked ? "#C7CEF2" : C.border}`,
                        borderRadius: 14,
                        padding: "13px 14px",
                        display: "flex",
                        gap: 12,
                        alignItems: "flex-start",
                      }}
                    >
                      <span
                        style={{
                          flex: "none",
                          width: 24,
                          height: 24,
                          borderRadius: 8,
                          border: `2px solid ${checked ? C.indigo : "#CCD3DE"}`,
                          background: checked ? C.indigo : "#fff",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          marginTop: 1,
                        }}
                      >
                        {checked && <Check size={14} color="#fff" strokeWidth={3.2} />}
                      </span>
                      <span style={{ flex: 1, minWidth: 0 }}>
                        <span style={{ display: "block", fontSize: 15, fontWeight: 700, letterSpacing: "-.01em" }}>{p.name}</span>
                        <span style={{ display: "block", fontSize: 12.5, color: C.muted, lineHeight: 1.45, marginTop: 2 }}>{p.description}</span>
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* ===== ISAK / TUTORIAL measurements ===== */}
          {isakCollect && (
            <div style={{ padding: "16px 18px 130px" }}>
              {/* expand toggle */}
              {isakCanChangeExpandedSkinfolds && (
                <button
                  onClick={() => setIsakExpandedSkinfolds((v) => !v)}
                  className="sa-tap"
                  style={{
                    width: "100%",
                    textAlign: "left",
                    cursor: "pointer",
                    background: isakExpandedSkinfolds ? C.indigoSoft : "#fff",
                    border: `1.5px solid ${isakExpandedSkinfolds ? "#C7CEF2" : C.border}`,
                    borderRadius: 14,
                    padding: "12px 14px",
                    display: "flex",
                    gap: 11,
                    alignItems: "center",
                    marginBottom: 14,
                  }}
                >
                  <span style={{ flex: "none", width: 38, height: 22, borderRadius: 99, background: isakExpandedSkinfolds ? C.indigo : "#D7DDE6", position: "relative", transition: "background .2s" }}>
                    <span style={{ position: "absolute", top: 2, left: isakExpandedSkinfolds ? 18 : 2, width: 18, height: 18, borderRadius: 99, background: "#fff", boxShadow: "0 1px 3px rgba(0,0,0,.2)", transition: "left .2s" }} />
                  </span>
                  <span style={{ flex: 1 }}>
                    <span style={{ display: "block", fontSize: 14, fontWeight: 700 }}>Avaliação expandida</span>
                    <span style={{ display: "block", fontSize: 12, color: C.muted }}>Inclui dobras torácica e axilar média</span>
                  </span>
                </button>
              )}

              {/* DOBRAS */}
              <SectionHeader accent="teal" label="Dobras cutâneas" meta={`${activeDobras.length} dobras · mm`} />
              <div style={{ display: "flex", flexDirection: "column", gap: 9, marginBottom: 20 }}>
                {activeDobras.map((f, i) => (
                  <MeasureRow
                    key={f.key}
                    num={i + 1}
                    field={f}
                    accent="teal"
                    showDescription={isTutorial && isakRound === 1}
                    value={isakInputs[f.key]}
                    onChange={(v) => setIsakInputs({ ...isakInputs, [f.key]: v })}
                    onKeyDown={makeEnterAdvance(isakInputIds, `inp-${f.key}`, handleIsakRound)}
                  />
                ))}
              </div>

              {/* PERÍMETROS */}
              <SectionHeader accent="indigo" label="Perímetros" meta={`${isakPerimFields.length} perímetros · cm`} />
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                {isakPerimFields.map((f, i) => (
                  <MeasureRow
                    key={f.key}
                    num={activeDobras.length + i + 1}
                    field={f as unknown as IsakField}
                    accent="indigo"
                    showDescription={isTutorial && isakRound === 1}
                    value={isakInputs[f.key]}
                    onChange={(v) => setIsakInputs({ ...isakInputs, [f.key]: v })}
                    onKeyDown={makeEnterAdvance(isakInputIds, `inp-${f.key}`, handleIsakRound)}
                  />
                ))}
              </div>
            </div>
          )}

          {/* ===== REVIEW (antropo + isak) ===== */}
          {reviewScreen && (
            <div style={{ padding: "16px 18px 130px" }}>
              <div style={{ margin: "6px 2px 14px" }}>
                <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: "-.01em" }}>Revisão do ETM</div>
                <div style={{ fontSize: 13, color: C.muted, marginTop: 2 }}>
                  Confira as 3 rodadas. O erro técnico de medida (ETM%) deve ficar abaixo do critério.
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                {(activeApp === "antropo" ? antropoEtmRows : isakEtmRows).map((row) => {
                  const valid = row.isValid;
                  return (
                    <div key={row.key} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "12px 13px", boxShadow: "0 1px 6px rgba(20,32,58,.03)" }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 9 }}>
                        <span style={{ fontSize: 14, fontWeight: 700, letterSpacing: "-.01em" }}>{row.label}</span>
                        <span
                          style={{
                            fontSize: 11,
                            fontWeight: 800,
                            fontFamily: FONT_HEAD,
                            padding: "4px 9px",
                            borderRadius: 8,
                            background: valid ? C.tealSoft : C.dangerSoft,
                            color: valid ? C.tealDark : C.danger,
                          }}
                        >
                          {valid ? "OK" : "Ajustar"}
                        </span>
                      </div>
                      <div style={{ display: "flex", gap: 7, alignItems: "flex-end" }}>
                        {[0, 1, 2].map((idx) => (
                          <label key={idx} style={{ flex: 1 }}>
                            <span style={{ display: "block", fontSize: 10, fontWeight: 700, color: C.faint, letterSpacing: ".04em", marginBottom: 3 }}>R{idx + 1}</span>
                            <input
                              className="sa-input sa-num"
                              type="number"
                              step="0.1"
                              value={Number.isFinite(row.values[idx]) ? row.values[idx] : ""}
                              onChange={(e) =>
                                activeApp === "antropo"
                                  ? updateAntropoReviewValue(row.key as (typeof antropoFields)[number]["key"], idx, e.target.value)
                                  : updateIsakReviewValue(row.key, idx, e.target.value)
                              }
                              style={{ width: "100%", textAlign: "center", border: `1px solid #E0E5EC`, borderRadius: 10, padding: "9px 4px", fontSize: 15, fontWeight: 700, fontFamily: FONT_HEAD, color: C.ink, background: "#FBFCFD", outline: "none" }}
                            />
                          </label>
                        ))}
                        <div style={{ flex: "none", textAlign: "right", minWidth: 70, paddingBottom: 2 }}>
                          <span style={{ display: "block", fontSize: 10, fontWeight: 700, color: C.faint, letterSpacing: ".04em" }}>ETM · &lt;{row.limit}%</span>
                          <span style={{ display: "block", fontSize: 16, fontWeight: 800, fontFamily: FONT_HEAD, color: valid ? C.tealDark : C.danger }}>
                            {Number.isFinite(row.etmPercent) ? `${formatNumber(row.etmPercent)}%` : "—"}
                          </span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {(activeApp === "antropo" ? antropoHasInvalidEtm : isakHasInvalidEtm) && (
                <label
                  className="sa-tap"
                  style={{ display: "flex", cursor: "pointer", marginTop: 14, background: "#FEF6E7", border: "1.5px solid #F4D58A", borderRadius: 14, padding: "13px 14px", gap: 11, alignItems: "flex-start" }}
                >
                  <input
                    type="checkbox"
                    checked={activeApp === "antropo" ? antropoEtmOverrideConfirmed : isakEtmOverrideConfirmed}
                    onChange={(e) => (activeApp === "antropo" ? setAntropoEtmOverrideConfirmed(e.target.checked) : setIsakEtmOverrideConfirmed(e.target.checked))}
                    style={{ marginTop: 2, width: 18, height: 18, accentColor: "#E0A92E" }}
                  />
                  <span style={{ fontSize: 13, color: "#8A6A1F", lineHeight: 1.45, fontWeight: 600 }}>
                    Confirmo que desejo salvar esta avaliação mesmo com ETM fora do alvo.
                  </span>
                </label>
              )}
            </div>
          )}

          {/* ===== ISAK tutorial composition estimate ===== */}
          {isakCompositionActive && (
            <div style={{ padding: "16px 18px 130px" }}>
              {isakCompositionStep === "prompt" && (
                <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 18, padding: 18, boxShadow: "0 2px 10px rgba(20,32,58,.04)" }}>
                  <div style={{ width: 44, height: 44, borderRadius: 14, background: C.indigoSoft, color: C.indigo, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 12 }}>
                    <Activity size={23} />
                  </div>
                  <div style={{ fontSize: 19, fontWeight: 800, letterSpacing: "-.01em" }}>Estimar composição corporal?</div>
                  <div style={{ fontSize: 13, color: C.muted, lineHeight: 1.55, marginTop: 6 }}>
                    A avaliação ISAK foi concluída. Você pode usar as médias das 3 rodadas para estimar percentual de gordura, IMC, RCQ, RCE e CMB corrigida.
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 16 }}>
                    <button
                      type="button"
                      onClick={() => setIsakCompositionStep("form")}
                      className="sa-tap"
                      style={{ border: "none", background: C.indigo, color: "#fff", borderRadius: 14, padding: 14, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
                    >
                      Sim, estimar composição <ArrowRight size={18} />
                    </button>
                    <button
                      type="button"
                      onClick={() => saveIsakEvaluation()}
                      disabled={saveIsakMutation.isPending}
                      className="sa-tap"
                      style={secondaryWideBtnStyle}
                    >
                      Não, salvar avaliação
                    </button>
                  </div>
                </div>
              )}

              {isakCompositionStep === "form" && (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <div style={{ margin: "6px 2px 2px" }}>
                    <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: "-.01em" }}>Dados para estimativa</div>
                    <div style={{ fontSize: 13, color: C.muted, marginTop: 2 }}>
                      As fórmulas usam sexo, idade, massa corporal, estatura e dobras específicas da equação escolhida.
                    </div>
                  </div>
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 16, padding: 14, display: "grid", gap: 10 }}>
                    <label>
                      <span style={fieldLabelStyle}>Sexo</span>
                      <select
                        value={compositionInputs.sex}
                        onChange={(event) => setCompositionInputs((current) => ({ ...current, sex: event.target.value as "" | Sex }))}
                        className="sa-input"
                        style={inputStyle}
                      >
                        <option value="">Selecione</option>
                        <option value="male">Masculino</option>
                        <option value="female">Feminino</option>
                      </select>
                    </label>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <label>
                        <span style={fieldLabelStyle}>Idade</span>
                        <input
                          className="sa-input"
                          inputMode="decimal"
                          value={compositionInputs.age}
                          onChange={(event) => setCompositionInputs((current) => ({ ...current, age: event.target.value }))}
                          placeholder="anos"
                          style={inputStyle}
                        />
                      </label>
                      <label>
                        <span style={fieldLabelStyle}>Massa</span>
                        <input
                          className="sa-input"
                          inputMode="decimal"
                          value={compositionInputs.mass}
                          onChange={(event) => setCompositionInputs((current) => ({ ...current, mass: event.target.value }))}
                          placeholder="kg"
                          style={inputStyle}
                        />
                      </label>
                    </div>
                    <label>
                      <span style={fieldLabelStyle}>Estatura</span>
                      <input
                        className="sa-input"
                        inputMode="decimal"
                        value={compositionInputs.stature}
                        onChange={(event) => setCompositionInputs((current) => ({ ...current, stature: event.target.value }))}
                        placeholder="cm"
                        style={inputStyle}
                      />
                    </label>
                    <label>
                      <span style={fieldLabelStyle}>Equação</span>
                      <select
                        value={compositionInputs.equation}
                        onChange={(event) =>
                          setCompositionInputs((current) => ({
                            ...current,
                            equation: event.target.value as "" | CompositionEquationKey,
                          }))
                        }
                        className="sa-input"
                        style={inputStyle}
                      >
                        <option value="">Selecione</option>
                        {compositionEquations.map((equation) => (
                          <option key={equation.key} value={equation.key}>
                            {equation.label}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                  <div style={{ background: C.indigoSoft, border: "1px solid #DCE0F6", borderRadius: 14, padding: "12px 14px", display: "flex", gap: 10 }}>
                    <Info size={18} color={C.indigo} style={{ flex: "none", marginTop: 1 }} />
                    <span style={{ fontSize: 12.5, color: "#3B4663", lineHeight: 1.5 }}>
                      Jackson & Pollock homens e Petroski mulheres exigem as dobras torácica ou axilar média. Para essas equações, marque a avaliação expandida no início do ISAK Tutorial.
                    </span>
                  </div>
                  {selectedCompositionNotices.length > 0 && (
                    <div style={{ background: "#FEF6E7", border: "1px solid #F4D58A", borderRadius: 14, padding: "12px 14px", display: "flex", gap: 10 }}>
                      <Info size={18} color="#8A6A1F" style={{ flex: "none", marginTop: 1 }} />
                      <span style={{ fontSize: 12.5, color: "#8A6A1F", lineHeight: 1.5 }}>
                        <strong>Equação com cautela:</strong> {selectedCompositionNotices.join(" ")} Você poderá avançar mesmo assim.
                      </span>
                    </div>
                  )}
                  <div style={{ display: "flex", gap: 10 }}>
                    <button type="button" onClick={() => setIsakCompositionStep("prompt")} className="sa-tap" style={secondaryBtnStyle}>
                      Voltar
                    </button>
                    <button
                      type="button"
                      onClick={handleCalculateComposition}
                      className="sa-tap"
                      style={{ flex: 1, border: "none", background: C.indigo, color: "#fff", borderRadius: 14, padding: 14, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
                    >
                      Calcular <ArrowRight size={18} />
                    </button>
                  </div>
                </div>
              )}

              {isakCompositionStep === "result" && compositionResult && (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <div
                    id="isak-composition-card"
                    style={{
                      background: "#fff",
                      border: `1px solid ${C.border}`,
                      borderRadius: 18,
                      padding: 18,
                      boxShadow: "0 8px 24px rgba(20,32,58,.08)",
                      color: C.ink,
                      fontFamily: FONT_BODY,
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start" }}>
                      <div>
                        <div style={{ fontSize: 11, letterSpacing: ".14em", textTransform: "uppercase", color: C.indigo, fontWeight: 800, fontFamily: FONT_HEAD }}>
                          Resumo ISAK Tutorial
                        </div>
                        <div style={{ fontSize: 21, fontWeight: 800, letterSpacing: "-.02em", marginTop: 3 }}>
                          {compositionResult.participantId || "Participante"}
                        </div>
                        <div style={{ fontSize: 12.5, color: C.muted, marginTop: 2 }}>
                          {new Date(compositionResult.date).toLocaleDateString("pt-BR")} · {sexLabels[compositionResult.sex]} · {compositionResult.age} anos
                        </div>
                      </div>
                      <div style={{ textAlign: "right", fontSize: 11, color: C.faint, fontWeight: 700 }}>
                        {compositionResult.equationLabel}
                      </div>
                    </div>

                    <div style={{ marginTop: 18, background: C.navy, color: "#fff", borderRadius: 16, padding: "17px 16px" }}>
                      <div style={{ fontSize: 12, color: "#AFC0DC", fontWeight: 700 }}>Percentual de gordura estimado</div>
                      <div style={{ fontSize: 42, lineHeight: 1, fontWeight: 800, fontFamily: FONT_HEAD, marginTop: 5 }}>
                        {formatMaybeNumber(compositionResult.bodyFatPercent, 1, "%")}
                      </div>
                      <div style={{ fontSize: 12, color: "#7FE3D6", fontWeight: 700, marginTop: 8 }}>
                        {compositionResult.equationMethod}
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 8, marginTop: 12 }}>
                      {[
                        ["IMC", `${formatMaybeNumber(compositionResult.imc, 1)} · ${compositionResult.imcClassification}`],
                        ["Densidade", formatMaybeNumber(compositionResult.density, 4)],
                        ["Soma das dobras", formatMaybeNumber(compositionResult.foldSum, 1, " mm")],
                        ["RCQ", formatMaybeNumber(compositionResult.rcq, 3)],
                        ["RCE", formatMaybeNumber(compositionResult.rce, 3)],
                        ["CMB corrigida", formatMaybeNumber(compositionResult.cmb, 1, " cm")],
                        ["Massa", formatMaybeNumber(compositionResult.mass, 1, " kg")],
                        ["Estatura", formatMaybeNumber(compositionResult.stature, 1, " cm")],
                      ].map(([label, value]) => (
                        <div key={label} style={{ border: `1px solid ${C.border}`, borderRadius: 12, padding: 10, background: "#FBFCFD" }}>
                          <div style={{ fontSize: 10, color: C.faint, fontWeight: 800, letterSpacing: ".06em", textTransform: "uppercase" }}>{label}</div>
                          <div style={{ fontSize: 13, color: C.ink, fontWeight: 800, marginTop: 3 }}>{value}</div>
                        </div>
                      ))}
                    </div>

                    <div style={{ marginTop: 14 }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: C.indigo, marginBottom: 7 }}>Dobras usadas na equação</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
                        {compositionResult.requiredFolds.map((metric) => (
                          <span key={metric.label} style={{ background: C.indigoSoft, color: C.indigoDark, borderRadius: 9, padding: "6px 8px", fontSize: 11.5, fontWeight: 800 }}>
                            {metric.label}: {metric.value}
                          </span>
                        ))}
                      </div>
                    </div>

                    <div style={{ marginTop: 14 }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: C.tealDark, marginBottom: 7 }}>Medidas ISAK (médias)</div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 6 }}>
                        {compositionResult.allMeasurements.map((metric) => (
                          <div key={metric.label} style={{ display: "flex", justifyContent: "space-between", gap: 8, borderBottom: "1px solid #EEF1F5", paddingBottom: 4, fontSize: 11.5 }}>
                            <span style={{ color: C.muted, fontWeight: 700 }}>{metric.label}</span>
                            <span style={{ color: C.ink, fontWeight: 800 }}>{metric.value}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    {compositionResult.warnings.length > 0 && (
                      <div style={{ marginTop: 14, background: "#FEF6E7", border: "1px solid #F4D58A", borderRadius: 12, padding: 10 }}>
                        <div style={{ fontSize: 11, fontWeight: 800, color: "#8A6A1F", marginBottom: 4 }}>Avisos de cautela</div>
                        {compositionResult.warnings.map((warning) => (
                          <div key={warning} style={{ fontSize: 11.5, color: "#8A6A1F", lineHeight: 1.45 }}>
                            {warning}
                          </div>
                        ))}
                      </div>
                    )}

                    <div style={{ marginTop: 14 }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: C.indigo, marginBottom: 7 }}>Comparação entre equações</div>
                      <div style={{ display: "grid", gap: 7 }}>
                        {compositionResult.comparisons.map((comparison) => (
                          <div
                            key={comparison.key}
                            style={{
                              border: `1px solid ${comparison.isSelected ? C.indigo : C.border}`,
                              background: comparison.isSelected ? C.indigoSoft : "#FBFCFD",
                              borderRadius: 12,
                              padding: 10,
                            }}
                          >
                            <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
                              <div style={{ minWidth: 0 }}>
                                <div style={{ fontSize: 12.5, color: C.ink, fontWeight: 800 }}>
                                  {comparison.label}
                                  {comparison.isSelected ? " · destaque" : ""}
                                </div>
                                <div style={{ fontSize: 10.5, color: comparison.isPertinent ? C.tealDark : "#8A6A1F", fontWeight: 800, marginTop: 2 }}>
                                  {comparison.isPertinent ? "Pertinente para sexo/idade" : "Não pertinente ou com cautela"}
                                </div>
                              </div>
                              <div style={{ flex: "none", textAlign: "right" }}>
                                <div style={{ fontSize: 16, color: C.ink, fontWeight: 900, fontFamily: FONT_HEAD }}>
                                  {formatMaybeNumber(comparison.bodyFatPercent, 1, "%")}
                                </div>
                                <div style={{ fontSize: 10.5, color: C.faint, fontWeight: 700 }}>
                                  Σ {formatMaybeNumber(comparison.foldSum, 1, " mm")}
                                </div>
                              </div>
                            </div>
                            {comparison.warnings.length > 0 && (
                              <div style={{ fontSize: 11, color: "#8A6A1F", lineHeight: 1.45, marginTop: 6 }}>
                                {comparison.warnings.join(" ")}
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                    <button type="button" onClick={exportCompositionJpeg} className="sa-tap" style={secondaryWideBtnStyle}>
                      <Download size={16} /> JPEG
                    </button>
                    <button type="button" onClick={exportCompositionCsv} className="sa-tap" style={secondaryWideBtnStyle}>
                      <Download size={16} /> CSV
                    </button>
                    <button type="button" onClick={exportCompositionExcel} className="sa-tap" style={secondaryWideBtnStyle}>
                      <Download size={16} /> Excel
                    </button>
                  </div>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button type="button" onClick={() => setIsakCompositionStep("form")} className="sa-tap" style={secondaryBtnStyle}>
                      Editar
                    </button>
                    <button
                      type="button"
                      onClick={() => saveIsakEvaluation()}
                      disabled={saveIsakMutation.isPending}
                      className="sa-tap"
                      style={{ flex: 1, border: "none", background: C.teal, color: "#fff", borderRadius: 14, padding: 14, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
                    >
                      Salvar e finalizar <ArrowRight size={18} />
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* ---------- sticky action bar ---------- */}
        {showActionBar && (
          <div style={{ position: "sticky", bottom: 0, zIndex: 20, background: "linear-gradient(180deg, rgba(247,248,250,0) 0%, #F7F8FA 26%)", padding: "14px 18px 18px" }}>
            <div style={{ display: "flex", gap: 11 }}>
              <button onClick={handleCancel} className="sa-tap" style={secondaryBtnStyle}>
                Cancelar
              </button>
              {(() => {
                let label = "";
                let onClick: () => void = () => {};
                let color = C.teal;
                if (antropoCollect) {
                  label = antropoRound === 3 ? "Revisar ETM" : "Próxima rodada";
                  onClick = handleAntropoRound;
                  color = C.teal;
                } else if (activeApp === "fpm") {
                  label = fpmRound === 3 ? "Finalizar e salvar" : "Próxima medida";
                  onClick = handleFpmRound;
                  color = C.indigo;
                } else if (tutorialPointsStep) {
                  label = "Iniciar 1ª rodada";
                  onClick = () => {
                    if (!allTutorialPointsChecked) {
                      toast.error("Marque todos os pontos anatômicos");
                      return;
                    }
                    setIsakTutorialStep("measurements");
                    focusInputSoon(`inp-${activeIsakFields[0].key}`);
                  };
                  color = C.indigo;
                } else if (isakCollect) {
                  label = isakRound === 3 ? "Revisar ETM" : "Próxima rodada";
                  onClick = handleIsakRound;
                  color = C.teal;
                } else if (reviewScreen) {
                  label = "Salvar avaliação";
                  onClick = activeApp === "antropo" ? handleSaveAntropoReview : handleSaveIsakReview;
                  color = C.teal;
                }
                const disabled =
                  (reviewScreen &&
                    activeApp === "antropo" &&
                    antropoHasInvalidEtm &&
                    !antropoEtmOverrideConfirmed) ||
                  (reviewScreen && isIsak && isakHasInvalidEtm && !isakEtmOverrideConfirmed) ||
                  (tutorialPointsStep && !allTutorialPointsChecked);
                return (
                  <button
                    onClick={onClick}
                    disabled={disabled}
                    className="sa-tap"
                    style={{
                      flex: 1,
                      cursor: disabled ? "not-allowed" : "pointer",
                      border: "none",
                      background: disabled ? "#AEB8C7" : color,
                      color: "#fff",
                      fontSize: 15,
                      fontWeight: 800,
                      padding: 14,
                      borderRadius: 14,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: 8,
                      boxShadow: disabled ? "none" : "0 6px 16px rgba(14,156,142,.28)",
                    }}
                  >
                    {label}
                    <ArrowRight size={18} strokeWidth={2.6} />
                  </button>
                );
              })()}
            </div>
          </div>
        )}

        {/* ---------- floating daily counter ---------- */}
        <div
          style={{
            position: "fixed",
            left: "50%",
            transform: "translateX(-50%)",
            bottom: counterBottom,
            zIndex: 45,
            display: "flex",
            alignItems: "center",
            gap: 10,
            background: C.navy,
            color: "#fff",
            padding: "9px 9px 9px 14px",
            borderRadius: 99,
            boxShadow: "0 8px 24px rgba(20,32,58,.28)",
            border: "1px solid rgba(255,255,255,.08)",
          }}
        >
          <CalendarCheck2 size={15} color="#7FE3D6" />
          <span style={{ fontSize: 13, fontWeight: 600, color: "#C4CEDE", letterSpacing: "-.01em" }}>
            {evalCount === 1 ? "avaliação hoje" : "avaliações hoje"}
          </span>
          <span
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              minWidth: 28,
              height: 28,
              padding: "0 8px",
              borderRadius: 99,
              background: C.teal,
              color: "#fff",
              fontSize: 15,
              fontWeight: 800,
              fontFamily: FONT_HEAD,
            }}
          >
            {evalCount}
          </span>
        </div>
      </div>
    </div>
  );
}

/* ================================================================== *
 *  Presentational helpers
 * ================================================================== */
const inputStyle: React.CSSProperties = {
  width: "100%",
  border: "1px solid #E0E5EC",
  borderRadius: 11,
  padding: "11px 12px",
  fontSize: 15,
  fontWeight: 600,
  color: C.ink,
  background: "#FBFCFD",
  outline: "none",
};

const fieldLabelStyle: React.CSSProperties = {
  display: "block",
  fontSize: 11,
  fontWeight: 700,
  color: C.faint,
  letterSpacing: ".08em",
  textTransform: "uppercase",
  marginBottom: 5,
};

const iconBtnStyle: React.CSSProperties = {
  border: "none",
  background: "rgba(255,255,255,.10)",
  color: "#fff",
  width: 34,
  height: 34,
  borderRadius: 11,
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  cursor: "pointer",
  flex: "none",
};

const toolCardStyle: React.CSSProperties = {
  flex: 1,
  textDecoration: "none",
  color: "inherit",
  border: `1px solid ${C.border}`,
  background: "#fff",
  borderRadius: 16,
  padding: 14,
};

const secondaryBtnStyle: React.CSSProperties = {
  flex: "none",
  cursor: "pointer",
  border: "1.5px solid #DCE1E9",
  background: "#fff",
  color: "#3B4663",
  fontSize: 15,
  fontWeight: 700,
  padding: "14px 20px",
  borderRadius: 14,
};

const secondaryWideBtnStyle: React.CSSProperties = {
  cursor: "pointer",
  border: "1.5px solid #DCE1E9",
  background: "#fff",
  color: "#3B4663",
  fontSize: 14,
  fontWeight: 800,
  padding: 14,
  borderRadius: 14,
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 8,
};

function MenuCard({
  icon,
  title,
  subtitle,
  accent,
  onClick,
}: {
  icon: React.ReactNode;
  title: string;
  subtitle: string;
  accent: "teal" | "indigo";
  onClick: () => void;
}) {
  const soft = accent === "teal" ? C.tealSoft : C.indigoSoft;
  const fg = accent === "teal" ? C.teal : C.indigo;
  return (
    <button
      onClick={onClick}
      className="sa-tap"
      style={{
        display: "flex",
        alignItems: "center",
        gap: 15,
        textAlign: "left",
        cursor: "pointer",
        border: `1px solid ${C.border}`,
        background: "#fff",
        borderRadius: 18,
        padding: 16,
        boxShadow: "0 2px 10px rgba(20,32,58,.04)",
      }}
    >
      <span style={{ flex: "none", width: 48, height: 48, borderRadius: 14, background: soft, color: fg, display: "flex", alignItems: "center", justifyContent: "center" }}>
        {icon}
      </span>
      <span style={{ flex: 1, minWidth: 0 }}>
        <span style={{ display: "block", fontSize: 16, fontWeight: 700, letterSpacing: "-.01em" }}>{title}</span>
        <span style={{ display: "block", fontSize: 13, color: C.muted, marginTop: 1 }}>{subtitle}</span>
      </span>
      <ChevronRight size={20} color="#C2CAD6" strokeWidth={2.4} />
    </button>
  );
}

function SectionHeader({ accent, label, meta }: { accent: "teal" | "indigo"; label: string; meta: string }) {
  const soft = accent === "teal" ? C.tealSoft : C.indigoSoft;
  const fg = accent === "teal" ? C.tealDark : C.indigoDark;
  const dot = accent === "teal" ? C.teal : C.indigo;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 9, margin: "4px 2px 10px" }}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 6, background: soft, color: fg, fontSize: 11, fontWeight: 800, letterSpacing: ".1em", textTransform: "uppercase", padding: "5px 10px", borderRadius: 8, fontFamily: FONT_HEAD }}>
        <span style={{ width: 7, height: 7, borderRadius: 99, background: dot }} />
        {label}
      </span>
      <span style={{ fontSize: 12, color: "#AAB3C2", fontWeight: 600 }}>{meta}</span>
    </div>
  );
}

function SimpleRow({
  label,
  region,
  unit,
  accent,
  inputId,
  value,
  onChange,
  onKeyDown,
}: {
  label: string;
  region: string;
  unit: string;
  accent: "teal" | "indigo";
  inputId: string;
  value: string | undefined;
  onChange: (v: string) => void;
  onKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => void;
}) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${C.border}`, borderRadius: 16, padding: "13px 15px", display: "flex", alignItems: "center", gap: 13, boxShadow: "0 1px 6px rgba(20,32,58,.03)" }}>
      <span style={{ flex: 1, minWidth: 0 }}>
        <span style={{ display: "block", fontSize: 15, fontWeight: 700, letterSpacing: "-.01em" }}>{label}</span>
        <span style={{ display: "block", fontSize: 12, color: C.faint, marginTop: 1 }}>{region}</span>
      </span>
      <span style={{ position: "relative", flex: "none" }}>
        <input
          id={inputId}
          className={accent === "teal" ? "sa-input sa-num" : "sa-input-indigo sa-num"}
          type="number"
          inputMode="decimal"
          step="0.1"
          placeholder="0,0"
          value={value ?? ""}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={onKeyDown}
          style={{ width: 104, textAlign: "right", border: "1.5px solid #E0E5EC", borderRadius: 12, padding: "12px 42px 12px 12px", fontSize: 17, fontWeight: 700, fontFamily: FONT_HEAD, color: C.ink, background: "#FBFCFD", outline: "none" }}
        />
        <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", fontSize: 11, fontWeight: 700, color: "#AAB3C2", fontFamily: FONT_HEAD, pointerEvents: "none" }}>{unit}</span>
      </span>
    </div>
  );
}

function MeasureRow({
  num,
  field,
  accent,
  showDescription,
  value,
  onChange,
  onKeyDown,
}: {
  num: number;
  field: IsakField;
  accent: "teal" | "indigo";
  showDescription: boolean;
  value: string | undefined;
  onChange: (v: string) => void;
  onKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => void;
}) {
  const filled = (value ?? "") !== "" && !isNaN(parseFloat(value ?? ""));
  const soft = accent === "teal" ? C.tealSoft : C.indigoSoft;
  const fg = accent === "teal" ? C.tealDark : C.indigoDark;
  const barColor = filled ? (accent === "teal" ? C.teal : C.indigo) : "#E9EDF3";
  const unit = field.kind === "skinfold" ? "mm" : "cm";
  return (
    <div style={{ background: "#fff", border: `1px solid ${C.border}`, borderLeft: `3px solid ${barColor}`, borderRadius: 14, padding: "11px 13px", display: "flex", alignItems: "center", gap: 11, boxShadow: "0 1px 6px rgba(20,32,58,.03)" }}>
      <span style={{ flex: "none", width: 26, height: 26, borderRadius: 8, background: soft, color: fg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, fontFamily: FONT_HEAD }}>{num}</span>
      <span style={{ flex: 1, minWidth: 0 }}>
        <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 14.5, fontWeight: 700, letterSpacing: "-.01em" }}>{field.short}</span>
          {field.optional && (
            <span style={{ fontSize: 9.5, fontWeight: 800, letterSpacing: ".05em", color: C.indigo, background: C.indigoSoft, padding: "2px 5px", borderRadius: 5 }}>EXP</span>
          )}
        </span>
        <span style={{ display: "block", fontSize: 11.5, color: C.faint, marginTop: 1 }}>{field.region}</span>
        {showDescription && <span style={{ display: "block", fontSize: 11.5, color: "#8A93A3", lineHeight: 1.45, marginTop: 4 }}>{field.description}</span>}
      </span>
      <span style={{ position: "relative", flex: "none" }}>
        <input
          id={`inp-${field.key}`}
          className={accent === "teal" ? "sa-input sa-num" : "sa-input-indigo sa-num"}
          type="number"
          inputMode="decimal"
          step="0.1"
          placeholder="0,0"
          value={value ?? ""}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={onKeyDown}
          style={{ width: 92, textAlign: "right", border: "1.5px solid #E0E5EC", borderRadius: 11, padding: "11px 32px 11px 10px", fontSize: 16, fontWeight: 700, fontFamily: FONT_HEAD, color: C.ink, background: "#FBFCFD", outline: "none" }}
        />
        <span style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", fontSize: 11, fontWeight: 700, color: "#AAB3C2", fontFamily: FONT_HEAD, pointerEvents: "none" }}>{unit}</span>
      </span>
    </div>
  );
}
