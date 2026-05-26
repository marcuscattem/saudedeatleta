import { useAuth } from "@/_core/hooks/useAuth";
import { Button } from "@/components/ui/button";
import ExcelJS from "exceljs";
import { useState, useEffect } from "react";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";

const EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PERIMETER_ETM_LIMIT = 1;
const SKINFOLD_ETM_LIMIT = 5;

type MeasurementKind = "perimeter" | "skinfold";

const antropoFields = [
  { key: "braco", label: "Braço Direito (cm)", excelHeader: "Braço (cm)", kind: "perimeter" },
  { key: "cintura", label: "Cintura (cm)", excelHeader: "Cintura (cm)", kind: "perimeter" },
  { key: "panturrilha", label: "Panturrilha Direita (cm)", excelHeader: "Panturrilha (cm)", kind: "perimeter" },
] as const;

const isakFields = [
  { key: "subscap", label: "Dobra subescapular (mm)", kind: "skinfold" },
  { key: "triceps", label: "Dobra de tríceps (mm)", kind: "skinfold" },
  { key: "biceps", label: "Dobra de bíceps (mm)", kind: "skinfold" },
  { key: "iliaca", label: "Dobra de crista ilíaca (mm)", kind: "skinfold" },
  { key: "supraesp", label: "Dobra supraespinhal (mm)", kind: "skinfold" },
  { key: "abdom", label: "Dobra abdominal (mm)", kind: "skinfold" },
  { key: "coxa", label: "Dobra de coxa anterior (mm)", kind: "skinfold" },
  { key: "pant_dobra", label: "Dobra de panturrilha medial (mm)", kind: "skinfold" },
  { key: "torax", label: "Perímetro de tórax (cm)", kind: "perimeter" },
  { key: "braco_rel", label: "Perímetro de braço relaxado (cm)", kind: "perimeter" },
  { key: "braco_flet", label: "Perímetro de braço contraído (cm)", kind: "perimeter" },
  { key: "cintura", label: "Perímetro de cintura (cm)", kind: "perimeter" },
  { key: "gluteo", label: "Perímetro de quadril (cm)", kind: "perimeter" },
  { key: "coxa_media", label: "Perímetro de coxa média (cm)", kind: "perimeter" },
  { key: "pant_perim", label: "Perímetro de panturrilha medial (cm)", kind: "perimeter" },
] as const;

function dateStamp(date = new Date()) {
  return date.toISOString().split("T")[0];
}

function safeFilenamePart(value: string) {
  return value.trim().replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "participante";
}

function styleHeader(worksheet: ExcelJS.Worksheet, color: string) {
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };
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

function getFpmStats(rightMeasurements: number[], leftMeasurements: number[]) {
  const rightAverage = calculateMean(rightMeasurements);
  const leftAverage = calculateMean(leftMeasurements);
  const rightMax = calculateMax(rightMeasurements);
  const leftMax = calculateMax(leftMeasurements);

  return {
    rightAverage,
    leftAverage,
    rightMax,
    leftMax,
    totalMax: calculateMax([rightMax, leftMax]),
  };
}

export default function Home() {
  const { user, isAuthenticated } = useAuth();
  const isStaticPages =
    import.meta.env.BASE_URL === "/saudedeatleta/" ||
    (typeof window !== "undefined" && window.location.hostname.endsWith("github.io"));
  const canUseApp = isAuthenticated || isStaticPages;
  const [activeApp, setActiveApp] = useState("home");
  const [participantId, setParticipantId] = useState("");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);

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
  const [fpmData, setFpmData] = useState({
    right: [] as number[],
    left: [] as number[],
  });
  const [fpmInputs, setFpmInputs] = useState({ right: "", left: "" });

  // ISAK state
  const [isakRound, setIsakRound] = useState(1);
  const [isakData, setIsakData] = useState<Record<string, number[]>>({});
  const [isakInputs, setIsakInputs] = useState<Record<string, string>>({});
  const [isakReview, setIsakReview] = useState(false);
  const [isakEtmOverrideConfirmed, setIsakEtmOverrideConfirmed] = useState(false);

  // Mutations
  const saveAntropoMutation = trpc.evaluations.saveAntropometria.useMutation();
  const saveFpmMutation = trpc.evaluations.saveFpm.useMutation();
  const saveIsakMutation = trpc.evaluations.saveIsak.useMutation();

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
    await downloadWorkbook(
      workbook,
      `antropometria_${safeFilenamePart(data.participantId)}_${dateStamp()}.xlsx`,
    );
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
    worksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Mão Dominante", key: "dominantHand", width: 15 },
      { header: "Perna Melhor", key: "bestLeg", width: 15 },
      { header: "Medida", key: "round", width: 10 },
      { header: "Lado Direito (kgf)", key: "right", width: 15 },
      { header: "Lado Esquerdo (kgf)", key: "left", width: 15 },
    ];

    data.rightMeasurements.forEach((right, index) => {
      worksheet.addRow({
        participantId: data.participantId,
        date: new Date(data.date).toLocaleDateString("pt-BR"),
        dominantHand: data.dominantHand,
        bestLeg: data.bestLeg,
        round: index + 1,
        right,
        left: data.leftMeasurements[index],
      });
    });

    styleHeader(worksheet, "FF70AD47");
    const stats = getFpmStats(data.rightMeasurements, data.leftMeasurements);
    const summaryWorksheet = workbook.addWorksheet("Resumo FPM");
    summaryWorksheet.columns = [
      { header: "Indicador", key: "metric", width: 34 },
      { header: "Valor (kgf)", key: "value", width: 16 },
    ];
    summaryWorksheet.addRows([
      { metric: "Força média - lado direito", value: stats.rightAverage },
      { metric: "Força média - lado esquerdo", value: stats.leftAverage },
      { metric: "Força máxima - lado direito", value: stats.rightMax },
      { metric: "Força máxima - lado esquerdo", value: stats.leftMax },
      { metric: "Força máxima total", value: stats.totalMax },
    ]);
    styleHeader(summaryWorksheet, "FF70AD47");

    await downloadWorkbook(workbook, `fpm_${safeFilenamePart(data.participantId)}_${dateStamp()}.xlsx`);
  };

  const generateLocalIsakExcel = async (data: {
    participantId: string;
    date: string;
    measurements: Record<string, number[]>;
  }) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ISAK");
    worksheet.columns = [
      { header: "ID Participante", key: "participantId", width: 15 },
      { header: "Data", key: "date", width: 15 },
      { header: "Rodada", key: "round", width: 10 },
      ...isakFields.map((field) => ({ header: field.label, key: field.key, width: 18 })),
    ];

    const numRounds = data.measurements[isakFields[0].key]?.length || 0;
    for (let round = 0; round < numRounds; round++) {
      const row: Record<string, unknown> = {
        participantId: data.participantId,
        date: new Date(data.date).toLocaleDateString("pt-BR"),
        round: round + 1,
      };

      isakFields.forEach((field) => {
        row[field.key] = data.measurements[field.key]?.[round];
      });

      worksheet.addRow(row);
    }

    styleHeader(worksheet, "FFC00000");
    await downloadWorkbook(workbook, `isak_${safeFilenamePart(data.participantId)}_${dateStamp()}.xlsx`);
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
    if (activeApp !== "home") {
      setHasUnsavedChanges(true);
    }
  }, [activeApp, participantId, date, antropoInputs, fpmInputs, fpmDominantHand, fpmBestLeg, isakInputs]);

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

  const isakEtmRows = isakFields.map((field) => {
    const values = isakData[field.key] ?? [];
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

  const antropoHasInvalidEtm = antropoEtmRows.some((row) => !row.isValid);
  const isakHasInvalidEtm = isakEtmRows.some((row) => !row.isValid);

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
  };

  const resetParticipant = () => {
    setParticipantId("");
    setDate(new Date().toISOString().split("T")[0]);
  };

  const resetAfterSave = (app: "antropo" | "fpm" | "isak") => {
    setHasUnsavedChanges(false);
    setActiveApp("home");
    if (app === "antropo") resetAntropo();
    if (app === "fpm") resetFpm();
    if (app === "isak") resetIsak();
    resetParticipant();
  };

  const updateAntropoReviewValue = (
    key: (typeof antropoFields)[number]["key"],
    index: number,
    value: string
  ) => {
    const nextValue = value === "" ? Number.NaN : Number(value);
    setAntropoEtmOverrideConfirmed(false);
    setAntropoData((current) => ({
      ...current,
      [key]: current[key].map((measurement, measurementIndex) =>
        measurementIndex === index ? nextValue : measurement
      ),
    }));
  };

  const updateIsakReviewValue = (key: string, index: number, value: string) => {
    const nextValue = value === "" ? Number.NaN : Number(value);
    setIsakEtmOverrideConfirmed(false);
    setIsakData((current) => ({
      ...current,
      [key]: (current[key] ?? []).map((measurement, measurementIndex) =>
        measurementIndex === index ? nextValue : measurement
      ),
    }));
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
        const result = await saveAntropoMutation.mutateAsync({
          ...payload,
          date: new Date(payload.date),
        });

        downloadExcel(result);
      }

      toast.success(isStaticPages ? "Antropometria salva em Excel!" : "Antropometria salva com Excel local e online!");
      resetAfterSave("antropo");
    } catch (error) {
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
    const payload = {
      participantId,
      date,
      measurements: data,
    };

    try {
      if (isStaticPages) {
        await generateLocalIsakExcel(payload);
      } else {
        const result = await saveIsakMutation.mutateAsync({
          ...payload,
          date: new Date(payload.date),
        });

        downloadExcel(result);
      }

      toast.success(isStaticPages ? "ISAK salva em Excel!" : "ISAK salva com Excel local e online!");
      resetAfterSave("isak");
    } catch (error) {
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
      // Salvar rodada intermediária
      setAntropoData(newData);
      setAntropoRound(antropoRound + 1);
      setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
      toast.success(`Rodada ${antropoRound} salva! Próxima rodada...`);
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

    const newData = {
      right: [...fpmData.right, right],
      left: [...fpmData.left, left],
    };

    if (fpmRound < 3) {
      // Salvar rodada intermediária
      setFpmData(newData);
      setFpmRound(fpmRound + 1);
      setFpmInputs({ right: "", left: "" });
      toast.success(`Medida ${fpmRound} salva! Próxima medida...`);
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
          const result = await saveFpmMutation.mutateAsync({
            ...payload,
            date: new Date(payload.date),
          });

          downloadExcel(result);
        }

        toast.success(isStaticPages ? "FPM salva em Excel!" : "FPM salva com Excel local e online!");
        setHasUnsavedChanges(false);
        setActiveApp("home");
        setFpmRound(1);
        setFpmData({ right: [], left: [] });
        setFpmInputs({ right: "", left: "" });
        setFpmDominantHand("");
        setFpmBestLeg("");
        setParticipantId("");
        setDate(new Date().toISOString().split("T")[0]);
      } catch (error) {
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
          setHasUnsavedChanges(false);
          setActiveApp("home");
          setFpmRound(1);
          setFpmData({ right: [], left: [] });
          setFpmInputs({ right: "", left: "" });
          setFpmDominantHand("");
          setFpmBestLeg("");
          setParticipantId("");
          setDate(new Date().toISOString().split("T")[0]);
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

    isakFields.forEach((field) => {
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
      // Salvar rodada intermediária
      setIsakData(newData);
      setIsakRound(isakRound + 1);
      setIsakInputs({});
      toast.success(`Rodada ${isakRound} salva! Próxima rodada...`);
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
        setAntropoRound(1);
        setAntropoData({ braco: [], cintura: [], panturrilha: [] });
        setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
        setAntropoReview(false);
        setAntropoEtmOverrideConfirmed(false);
        setFpmRound(1);
        setFpmData({ right: [], left: [] });
        setFpmInputs({ right: "", left: "" });
        setFpmDominantHand("");
        setFpmBestLeg("");
        setIsakRound(1);
        setIsakData({});
        setIsakInputs({});
        setIsakReview(false);
        setIsakEtmOverrideConfirmed(false);
        setParticipantId("");
        setDate(new Date().toISOString().split("T")[0]);
      }
    } else {
      setActiveApp("home");
    }
  };

  if (!canUseApp) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="text-center">
          <h1 className="text-3xl font-bold mb-4">Portal Saúde de Atleta</h1>
          <p className="text-slate-600 mb-6">Faça login para acessar a ferramenta</p>
          <Button onClick={() => window.location.href = "/api/oauth/login"}>
            Fazer Login
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-slate-900 text-white p-4 sticky top-0 z-30">
        <div className="max-w-4xl mx-auto flex justify-between items-center">
          <h1 className="text-xl font-bold">Portal Saúde de Atleta</h1>
          <div className="flex items-center gap-4">
            {hasUnsavedChanges && activeApp !== "home" && (
              <span className="text-yellow-300 text-sm font-semibold">● Alterações não salvas</span>
            )}
            <div className="text-sm">{user?.name ?? (isStaticPages ? "Modo público" : "")}</div>
          </div>
        </div>
      </header>

      <main className="max-w-4xl mx-auto p-4 md:p-6">
        {activeApp === "home" && (
          <div>
            <h2 className="text-2xl font-bold mb-6">Selecione o Protocolo</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <Button
                onClick={() => {
                  setActiveApp("antropo");
                  resetAntropo();
                  resetParticipant();
                }}
                className="h-32 text-lg"
              >
                Antropometria
              </Button>
              <Button
                onClick={() => {
                  setActiveApp("fpm");
                  resetFpm();
                  resetParticipant();
                }}
                className="h-32 text-lg"
              >
                Força de Preensão Manual
              </Button>
              <Button
                onClick={() => {
                  setActiveApp("isak");
                  resetIsak();
                  resetParticipant();
                }}
                className="h-32 text-lg"
              >
                Antropometria ISAK 1
              </Button>
            </div>
          </div>
        )}

        {activeApp === "antropo" && (
          <div className="bg-white p-6 rounded-lg border">
            <h2 className="text-xl font-bold mb-4">Antropometria</h2>
            <div className="space-y-4">
              <input
                type="text"
                placeholder="ID do Participante"
                value={participantId}
                onChange={(e) => setParticipantId(e.target.value)}
                className="w-full border rounded p-2"
              />
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full border rounded p-2"
              />
              {!antropoReview && antropoRound <= 3 && (
                <>
                  <h3 className="font-semibold">Rodada {antropoRound} de 3</h3>
                  <div className="bg-blue-50 p-3 rounded text-sm text-blue-800">
                    Dados salvos: {antropoData.braco.length > 0 && `Rodadas: ${antropoData.braco.length}`}
                  </div>
                  <input
                    type="number"
                    step="0.1"
                    placeholder="Braço Direito (cm)"
                    value={antropoInputs.braco}
                    onChange={(e) => setAntropoInputs({ ...antropoInputs, braco: e.target.value })}
                    className="w-full border rounded p-2"
                  />
                  <input
                    type="number"
                    step="0.1"
                    placeholder="Cintura (cm)"
                    value={antropoInputs.cintura}
                    onChange={(e) => setAntropoInputs({ ...antropoInputs, cintura: e.target.value })}
                    className="w-full border rounded p-2"
                  />
                  <input
                    type="number"
                    step="0.1"
                    placeholder="Panturrilha Direita (cm)"
                    value={antropoInputs.panturrilha}
                    onChange={(e) => setAntropoInputs({ ...antropoInputs, panturrilha: e.target.value })}
                    className="w-full border rounded p-2"
                  />
                  <div className="flex gap-4">
                    <Button variant="outline" onClick={handleCancel}>
                      Cancelar
                    </Button>
                    <Button onClick={handleAntropoRound} disabled={saveAntropoMutation.isPending}>
                      {antropoRound === 3 ? "Revisar ETM" : "Próxima Rodada"}
                    </Button>
                  </div>
                </>
              )}
              {antropoReview && (
                <>
                  <h3 className="font-semibold">Revisão do ETM</h3>
                  <div className="overflow-x-auto border rounded">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-100">
                        <tr>
                          <th className="text-left p-2">Medida</th>
                          <th className="text-left p-2">Rodada 1</th>
                          <th className="text-left p-2">Rodada 2</th>
                          <th className="text-left p-2">Rodada 3</th>
                          <th className="text-left p-2">ETM</th>
                          <th className="text-left p-2">Critério</th>
                          <th className="text-left p-2">Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {antropoEtmRows.map((row) => (
                          <tr key={row.key} className="border-t">
                            <td className="p-2 font-medium">{row.label}</td>
                            {row.values.map((value, index) => (
                              <td key={`${row.key}-${index}`} className="p-2 min-w-28">
                                <input
                                  type="number"
                                  step="0.1"
                                  value={Number.isFinite(value) ? value : ""}
                                  onChange={(event) => updateAntropoReviewValue(row.key, index, event.target.value)}
                                  className="w-24 border rounded p-2"
                                  aria-label={`${row.label} rodada ${index + 1}`}
                                />
                              </td>
                            ))}
                            <td className="p-2">{formatNumber(row.etmPercent)}%</td>
                            <td className="p-2">&lt; {row.limit}%</td>
                            <td className={`p-2 font-semibold ${row.isValid ? "text-green-700" : "text-red-700"}`}>
                              {row.isValid ? "OK" : "Ajustar"}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {antropoHasInvalidEtm && (
                    <label className="flex items-start gap-3 rounded border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900">
                      <input
                        type="checkbox"
                        checked={antropoEtmOverrideConfirmed}
                        onChange={(event) => setAntropoEtmOverrideConfirmed(event.target.checked)}
                        className="mt-1"
                      />
                      <span>
                        Confirmo que desejo salvar esta avaliação mesmo com ETM fora do alvo.
                      </span>
                    </label>
                  )}
                  <div className="flex gap-4">
                    <Button variant="outline" onClick={handleCancel}>
                      Cancelar
                    </Button>
                    <Button
                      onClick={handleSaveAntropoReview}
                      disabled={
                        saveAntropoMutation.isPending ||
                        (antropoHasInvalidEtm && !antropoEtmOverrideConfirmed)
                      }
                    >
                      Salvar avaliação
                    </Button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {activeApp === "fpm" && (
          <div className="bg-white p-6 rounded-lg border">
            <h2 className="text-xl font-bold mb-4">Força de Preensão Manual</h2>
            <div className="space-y-4">
              <input
                type="text"
                placeholder="ID do Participante"
                value={participantId}
                onChange={(e) => setParticipantId(e.target.value)}
                className="w-full border rounded p-2"
              />
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full border rounded p-2"
              />
              <select
                value={fpmDominantHand}
                onChange={(e) => setFpmDominantHand(e.target.value)}
                className="w-full border rounded p-2"
              >
                <option value="">Mão Dominante</option>
                <option value="Direita">Direita</option>
                <option value="Esquerda">Esquerda</option>
              </select>
              <select
                value={fpmBestLeg}
                onChange={(e) => setFpmBestLeg(e.target.value)}
                className="w-full border rounded p-2"
              >
                <option value="">Perna Melhor</option>
                <option value="Direita">Direita</option>
                <option value="Esquerda">Esquerda</option>
              </select>
              {fpmRound <= 3 && (
                <>
                  <h3 className="font-semibold">{fpmRound}ª Medida</h3>
                  <div className="bg-blue-50 p-3 rounded text-sm text-blue-800">
                    Dados salvos: {fpmData.right.length > 0 && `Medidas: ${fpmData.right.length}`}
                  </div>
                  <input
                    type="number"
                    step="0.1"
                    placeholder="Lado Direito (kgf)"
                    value={fpmInputs.right}
                    onChange={(e) => setFpmInputs({ ...fpmInputs, right: e.target.value })}
                    className="w-full border rounded p-2"
                  />
                  <input
                    type="number"
                    step="0.1"
                    placeholder="Lado Esquerdo (kgf)"
                    value={fpmInputs.left}
                    onChange={(e) => setFpmInputs({ ...fpmInputs, left: e.target.value })}
                    className="w-full border rounded p-2"
                  />
                  <div className="flex gap-4">
                    <Button variant="outline" onClick={handleCancel}>
                      Cancelar
                    </Button>
                    <Button onClick={handleFpmRound} disabled={saveFpmMutation.isPending}>
                      {fpmRound === 3 ? "Finalizar e Salvar" : "Próxima Medida"}
                    </Button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {activeApp === "isak" && (
          <div className="bg-white p-6 rounded-lg border">
            <h2 className="text-xl font-bold mb-4">Antropometria ISAK 1</h2>
            <div className="space-y-4">
              <input
                type="text"
                placeholder="ID do Participante"
                value={participantId}
                onChange={(e) => setParticipantId(e.target.value)}
                className="w-full border rounded p-2"
              />
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full border rounded p-2"
              />
              {!isakReview && isakRound <= 3 && (
                <>
                  <h3 className="font-semibold">Rodada {isakRound} de 3</h3>
                  <div className="bg-blue-50 p-3 rounded text-sm text-blue-800">
                    Dados salvos: {Object.keys(isakData).length > 0 && `Rodadas: ${isakData[Object.keys(isakData)[0]]?.length || 0}`}
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {isakFields.map((field) => (
                      <input
                        key={field.key}
                        type="number"
                        step="0.1"
                        placeholder={field.label}
                        value={isakInputs[field.key] || ""}
                        onChange={(e) => setIsakInputs({ ...isakInputs, [field.key]: e.target.value })}
                        className="border rounded p-2"
                      />
                    ))}
                  </div>
                  <div className="flex gap-4">
                    <Button variant="outline" onClick={handleCancel}>
                      Cancelar
                    </Button>
                    <Button onClick={handleIsakRound} disabled={saveIsakMutation.isPending}>
                      {isakRound === 3 ? "Revisar ETM" : "Próxima Rodada"}
                    </Button>
                  </div>
                </>
              )}
              {isakReview && (
                <>
                  <h3 className="font-semibold">Revisão do ETM</h3>
                  <div className="overflow-x-auto border rounded">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-100">
                        <tr>
                          <th className="text-left p-2">Medida</th>
                          <th className="text-left p-2">Rodada 1</th>
                          <th className="text-left p-2">Rodada 2</th>
                          <th className="text-left p-2">Rodada 3</th>
                          <th className="text-left p-2">ETM</th>
                          <th className="text-left p-2">Critério</th>
                          <th className="text-left p-2">Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {isakEtmRows.map((row) => (
                          <tr key={row.key} className="border-t">
                            <td className="p-2 font-medium min-w-56">{row.label}</td>
                            {row.values.map((value, index) => (
                              <td key={`${row.key}-${index}`} className="p-2 min-w-28">
                                <input
                                  type="number"
                                  step="0.1"
                                  value={Number.isFinite(value) ? value : ""}
                                  onChange={(event) => updateIsakReviewValue(row.key, index, event.target.value)}
                                  className="w-24 border rounded p-2"
                                  aria-label={`${row.label} rodada ${index + 1}`}
                                />
                              </td>
                            ))}
                            <td className="p-2">{formatNumber(row.etmPercent)}%</td>
                            <td className="p-2">&lt; {row.limit}%</td>
                            <td className={`p-2 font-semibold ${row.isValid ? "text-green-700" : "text-red-700"}`}>
                              {row.isValid ? "OK" : "Ajustar"}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {isakHasInvalidEtm && (
                    <label className="flex items-start gap-3 rounded border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900">
                      <input
                        type="checkbox"
                        checked={isakEtmOverrideConfirmed}
                        onChange={(event) => setIsakEtmOverrideConfirmed(event.target.checked)}
                        className="mt-1"
                      />
                      <span>
                        Confirmo que desejo salvar esta avaliação mesmo com ETM fora do alvo.
                      </span>
                    </label>
                  )}
                  <div className="flex gap-4">
                    <Button variant="outline" onClick={handleCancel}>
                      Cancelar
                    </Button>
                    <Button
                      onClick={handleSaveIsakReview}
                      disabled={
                        saveIsakMutation.isPending ||
                        (isakHasInvalidEtm && !isakEtmOverrideConfirmed)
                      }
                    >
                      Salvar avaliação
                    </Button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
