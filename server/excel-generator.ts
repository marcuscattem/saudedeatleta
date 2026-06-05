import ExcelJS from "exceljs";
import { Antropometria, FpmEvaluation, IsakEvaluation } from "../drizzle/schema";

function calculateMean(values: number[]) {
  const validValues = values.filter(Number.isFinite);
  if (validValues.length === 0) return Number.NaN;
  return validValues.reduce((sum, value) => sum + value, 0) / validValues.length;
}

function calculateMax(values: number[]) {
  const validValues = values.filter(Number.isFinite);
  return validValues.length > 0 ? Math.max(...validValues) : Number.NaN;
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

function styleHeader(worksheet: ExcelJS.Worksheet, color: string) {
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };
}

export async function generateAntropometriaExcel(data: Antropometria): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Antropometria");

  // Headers
  worksheet.columns = [
    { header: "ID Participante", key: "participantId", width: 15 },
    { header: "Data", key: "date", width: 15 },
    { header: "Rodada", key: "round", width: 10 },
    { header: "Braço (cm)", key: "braco", width: 15 },
    { header: "Cintura (cm)", key: "cintura", width: 15 },
    { header: "Panturrilha (cm)", key: "panturrilha", width: 15 },
  ];

  const bracoMeasurements = JSON.parse(data.bracoMeasurements);
  const cinturaMeasurements = JSON.parse(data.cinturaMeasurements);
  const panturrilhaMeasurements = JSON.parse(data.panturrilhaMeasurements);

  // Add rows for each round
  for (let i = 0; i < bracoMeasurements.length; i++) {
    worksheet.addRow({
      participantId: data.participantId,
      date: new Date(data.date).toLocaleDateString("pt-BR"),
      round: i + 1,
      braco: bracoMeasurements[i],
      cintura: cinturaMeasurements[i],
      panturrilha: panturrilhaMeasurements[i],
    });
  }

  // Format header row
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export async function generateFpmExcel(data: FpmEvaluation): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("FPM");
  const rightMeasurements = JSON.parse(data.rightMeasurements);
  const leftMeasurements = JSON.parse(data.leftMeasurements);
  const stats = getFpmStats(rightMeasurements, leftMeasurements);

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
    right1: rightMeasurements[0],
    right2: rightMeasurements[1],
    right3: rightMeasurements[2],
    left1: leftMeasurements[0],
    left2: leftMeasurements[1],
    left3: leftMeasurements[2],
    rightAverage: stats.rightAverage,
    leftAverage: stats.leftAverage,
    generalAverage: stats.generalAverage,
    rightMax: stats.rightMax,
    leftMax: stats.leftMax,
    generalMax: stats.generalMax,
  });

  styleHeader(worksheet, "FF70AD47");

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export async function generateIsakExcel(data: IsakEvaluation): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("ISAK");

  const measurements = JSON.parse(data.measurements);

  // Define field labels in the new order
  const fields = [
    { key: "subscap", label: "Dobra subescapular (mm)" },
    { key: "triceps", label: "Dobra de tríceps (mm)" },
    { key: "biceps", label: "Dobra de bíceps (mm)" },
    { key: "iliaca", label: "Dobra de crista ilíaca (mm)" },
    { key: "supraesp", label: "Dobra supraespinhal (mm)" },
    { key: "abdom", label: "Dobra abdominal (mm)" },
    { key: "coxa", label: "Dobra de coxa anterior (mm)" },
    { key: "pant_dobra", label: "Dobra de panturrilha medial (mm)" },
    { key: "torax", label: "Perímetro de tórax (cm)" },
    { key: "braco_rel", label: "Perímetro de braço relaxado (cm)" },
    { key: "braco_flet", label: "Perímetro de braço contraído (cm)" },
    { key: "cintura", label: "Perímetro de cintura (cm)" },
    { key: "abdome_perim", label: "Perímetro de abdome (cm)" },
    { key: "gluteo", label: "Perímetro de quadril (cm)" },
    { key: "coxa_media", label: "Perímetro de coxa média (cm)" },
    { key: "pant_perim", label: "Perímetro de panturrilha medial (cm)" },
  ];

  // Create columns dynamically
  const columns = [
    { header: "ID Participante", key: "participantId", width: 15 },
    { header: "Data", key: "date", width: 15 },
    { header: "Rodada", key: "round", width: 10 },
    ...fields.map((f) => ({ header: f.label, key: f.key, width: 18 })),
  ];

  worksheet.columns = columns;

  // Get the number of rounds
  const firstField = measurements[fields[0].key];
  const numRounds = firstField ? firstField.length : 0;

  // Add rows for each round
  for (let round = 0; round < numRounds; round++) {
    const row: Record<string, unknown> = {
      participantId: data.participantId,
      date: new Date(data.date).toLocaleDateString("pt-BR"),
      round: round + 1,
    };

    fields.forEach((field) => {
      if (measurements[field.key] && measurements[field.key][round] !== undefined) {
        row[field.key] = measurements[field.key][round];
      }
    });

    worksheet.addRow(row);
  }

  // Format header row
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFC00000" },
  };
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
