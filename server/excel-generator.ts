import ExcelJS from "exceljs";
import { Antropometria, FpmEvaluation, IsakEvaluation } from "../drizzle/schema";

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

  worksheet.columns = [
    { header: "ID Participante", key: "participantId", width: 15 },
    { header: "Data", key: "date", width: 15 },
    { header: "Mão Dominante", key: "dominantHand", width: 15 },
    { header: "Perna Melhor", key: "bestLeg", width: 15 },
    { header: "Medida", key: "round", width: 10 },
    { header: "Lado Direito (kgf)", key: "right", width: 15 },
    { header: "Lado Esquerdo (kgf)", key: "left", width: 15 },
  ];

  const rightMeasurements = JSON.parse(data.rightMeasurements);
  const leftMeasurements = JSON.parse(data.leftMeasurements);

  for (let i = 0; i < rightMeasurements.length; i++) {
    worksheet.addRow({
      participantId: data.participantId,
      date: new Date(data.date).toLocaleDateString("pt-BR"),
      dominantHand: data.dominantHand,
      bestLeg: data.bestLeg,
      round: i + 1,
      right: rightMeasurements[i],
      left: leftMeasurements[i],
    });
  }

  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF70AD47" },
  };
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

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
