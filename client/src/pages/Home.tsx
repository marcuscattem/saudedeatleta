import { useAuth } from "@/_core/hooks/useAuth";
import { Button } from "@/components/ui/button";
import ExcelJS from "exceljs";
import { useState, useEffect } from "react";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";

const EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

  const isakFields = [
    // Dobras
    { key: "subscap", label: "Dobra subescapular (mm)" },
    { key: "triceps", label: "Dobra de tríceps (mm)" },
    { key: "biceps", label: "Dobra de bíceps (mm)" },
    { key: "iliaca", label: "Dobra de crista ilíaca (mm)" },
    { key: "supraesp", label: "Dobra supraespinhal (mm)" },
    { key: "abdom", label: "Dobra abdominal (mm)" },
    { key: "coxa", label: "Dobra de coxa anterior (mm)" },
    { key: "pant_dobra", label: "Dobra de panturrilha medial (mm)" },
    // Perímetros
    { key: "torax", label: "Perímetro de tórax (cm)" },
    { key: "braco_rel", label: "Perímetro de braço relaxado (cm)" },
    { key: "braco_flet", label: "Perímetro de braço contraído (cm)" },
    { key: "cintura", label: "Perímetro de cintura (cm)" },
    { key: "gluteo", label: "Perímetro de quadril (cm)" },
    { key: "coxa_media", label: "Perímetro de coxa média (cm)" },
    { key: "pant_perim", label: "Perímetro de panturrilha medial (cm)" },
  ];

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
      // Salvar no banco de dados
      try {
        const payload = {
          participantId,
          date,
          bracoMeasurements: newData.braco,
          cinturaMeasurements: newData.cintura,
          panturrilhaMeasurements: newData.panturrilha,
        };

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
        setHasUnsavedChanges(false);
        setActiveApp("home");
        setAntropoRound(1);
        setAntropoData({ braco: [], cintura: [], panturrilha: [] });
        setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
        setParticipantId("");
        setDate(new Date().toISOString().split("T")[0]);
      } catch (error) {
        try {
          await generateLocalAntropoExcel({
            participantId,
            date,
            bracoMeasurements: newData.braco,
            cinturaMeasurements: newData.cintura,
            panturrilhaMeasurements: newData.panturrilha,
          });
          toast.success("Antropometria salva em Excel!");
          setHasUnsavedChanges(false);
          setActiveApp("home");
          setAntropoRound(1);
          setAntropoData({ braco: [], cintura: [], panturrilha: [] });
          setAntropoInputs({ braco: "", cintura: "", panturrilha: "" });
          setParticipantId("");
          setDate(new Date().toISOString().split("T")[0]);
        } catch {
          toast.error("Erro ao gerar Excel de antropometria");
        }
      }
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
      try {
        const payload = {
          participantId,
          date,
          measurements: newData,
        };

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
        setHasUnsavedChanges(false);
        setActiveApp("home");
        setIsakRound(1);
        setIsakData({});
        setIsakInputs({});
        setParticipantId("");
        setDate(new Date().toISOString().split("T")[0]);
      } catch (error) {
        try {
          await generateLocalIsakExcel({
            participantId,
            date,
            measurements: newData,
          });
          toast.success("ISAK salva em Excel!");
          setHasUnsavedChanges(false);
          setActiveApp("home");
          setIsakRound(1);
          setIsakData({});
          setIsakInputs({});
          setParticipantId("");
          setDate(new Date().toISOString().split("T")[0]);
        } catch {
          toast.error("Erro ao gerar Excel ISAK");
        }
      }
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
        setFpmRound(1);
        setFpmData({ right: [], left: [] });
        setFpmInputs({ right: "", left: "" });
        setFpmDominantHand("");
        setFpmBestLeg("");
        setIsakRound(1);
        setIsakData({});
        setIsakInputs({});
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
                  setParticipantId("");
                  setDate(new Date().toISOString().split("T")[0]);
                }}
                className="h-32 text-lg"
              >
                Antropometria
              </Button>
              <Button
                onClick={() => {
                  setActiveApp("fpm");
                  setParticipantId("");
                  setDate(new Date().toISOString().split("T")[0]);
                }}
                className="h-32 text-lg"
              >
                Força de Preensão Manual
              </Button>
              <Button
                onClick={() => {
                  setActiveApp("isak");
                  setParticipantId("");
                  setDate(new Date().toISOString().split("T")[0]);
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
              {antropoRound <= 3 && (
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
                      {antropoRound === 3 ? "Finalizar e Salvar" : "Próxima Rodada"}
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
              {isakRound <= 3 && (
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
                      {isakRound === 3 ? "Finalizar e Salvar" : "Próxima Rodada"}
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
