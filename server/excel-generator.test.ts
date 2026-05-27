import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { generateAntropometriaExcel, generateFpmExcel, generateIsakExcel } from "./excel-generator";
import type { Antropometria, FpmEvaluation, IsakEvaluation } from "../drizzle/schema";

describe("Excel Generator", () => {
  describe("generateAntropometriaExcel", () => {
    it("should generate valid Excel buffer for Antropometria", async () => {
      const mockData: Antropometria = {
        id: 1,
        userId: 1,
        participantId: "P001",
        date: new Date(),
        bracoMeasurements: JSON.stringify([25.5, 25.6, 25.4]),
        cinturaMeasurements: JSON.stringify([80.0, 80.1, 79.9]),
        panturrilhaMeasurements: JSON.stringify([35.0, 35.1, 34.9]),
        excelUrl: null,
        createdAt: new Date(),
        updatedAt: new Date(),
      };

      const buffer = await generateAntropometriaExcel(mockData);

      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
      // Check for XLSX magic bytes
      expect(buffer[0]).toBe(0x50); // 'P'
      expect(buffer[1]).toBe(0x4b); // 'K'
    });
  });

  describe("generateFpmExcel", () => {
    it("should generate valid Excel buffer for FPM", async () => {
      const mockData: FpmEvaluation = {
        id: 1,
        userId: 1,
        participantId: "P001",
        date: new Date(),
        dominantHand: "Direita",
        bestLeg: "Direita",
        rightMeasurements: JSON.stringify([45.0, 45.5, 44.8]),
        leftMeasurements: JSON.stringify([42.0, 42.5, 41.8]),
        excelUrl: null,
        createdAt: new Date(),
        updatedAt: new Date(),
      };

      const buffer = await generateFpmExcel(mockData);

      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
      // Check for XLSX magic bytes
      expect(buffer[0]).toBe(0x50); // 'P'
      expect(buffer[1]).toBe(0x4b); // 'K'
    });

    it("should include FPM summary statistics", async () => {
      const mockData: FpmEvaluation = {
        id: 1,
        userId: 1,
        participantId: "P001",
        date: new Date(),
        dominantHand: "Direita",
        bestLeg: "Direita",
        rightMeasurements: JSON.stringify([45.0, 45.5, 44.5]),
        leftMeasurements: JSON.stringify([42.0, 42.5, 41.5]),
        excelUrl: null,
        createdAt: new Date(),
        updatedAt: new Date(),
      };

      const buffer = await generateFpmExcel(mockData);
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.getWorksheet("Resumo FPM");

      expect(worksheet?.getCell("A2").value).toBe("Força média - lado direito");
      expect(worksheet?.getCell("B2").value).toBe(45);
      expect(worksheet?.getCell("A3").value).toBe("Força média - lado esquerdo");
      expect(worksheet?.getCell("B3").value).toBe(42);
      expect(worksheet?.getCell("A4").value).toBe("Força máxima - lado direito");
      expect(worksheet?.getCell("B4").value).toBe(45.5);
      expect(worksheet?.getCell("A5").value).toBe("Força máxima - lado esquerdo");
      expect(worksheet?.getCell("B5").value).toBe(42.5);
      expect(worksheet?.getCell("A6").value).toBe("Força máxima total");
      expect(worksheet?.getCell("B6").value).toBe(45.5);
    });
  });

  describe("generateIsakExcel", () => {
    it("should generate valid Excel buffer for ISAK with new field order", async () => {
      const mockData: IsakEvaluation = {
        id: 1,
        userId: 1,
        participantId: "P001",
        date: new Date(),
        measurements: JSON.stringify({
          subscap: [10.5, 10.6, 10.4],
          triceps: [12.0, 12.1, 11.9],
          biceps: [8.5, 8.6, 8.4],
          iliaca: [15.0, 15.1, 14.9],
          supraesp: [9.0, 9.1, 8.9],
          abdom: [20.0, 20.1, 19.9],
          coxa: [18.0, 18.1, 17.9],
          pant_dobra: [11.0, 11.1, 10.9],
          torax: [95.0, 95.1, 94.9],
          braco_rel: [28.0, 28.1, 27.9],
          braco_flet: [32.0, 32.1, 31.9],
          cintura: [80.0, 80.1, 79.9],
          abdome_perim: [82.0, 82.1, 81.9],
          gluteo: [95.0, 95.1, 94.9],
          coxa_media: [55.0, 55.1, 54.9],
          pant_perim: [38.0, 38.1, 37.9],
        }),
        excelUrl: null,
        createdAt: new Date(),
        updatedAt: new Date(),
      };

      const buffer = await generateIsakExcel(mockData);

      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
      // Check for XLSX magic bytes
      expect(buffer[0]).toBe(0x50); // 'P'
      expect(buffer[1]).toBe(0x4b); // 'K'
    });

    it("should include all 16 ISAK fields in correct order", async () => {
      const mockData: IsakEvaluation = {
        id: 1,
        userId: 1,
        participantId: "P001",
        date: new Date(),
        measurements: JSON.stringify({
          subscap: [10.5],
          triceps: [12.0],
          biceps: [8.5],
          iliaca: [15.0],
          supraesp: [9.0],
          abdom: [20.0],
          coxa: [18.0],
          pant_dobra: [11.0],
          torax: [95.0],
          braco_rel: [28.0],
          braco_flet: [32.0],
          cintura: [80.0],
          abdome_perim: [82.0],
          gluteo: [95.0],
          coxa_media: [55.0],
          pant_perim: [38.0],
        }),
        excelUrl: null,
        createdAt: new Date(),
        updatedAt: new Date(),
      };

      const buffer = await generateIsakExcel(mockData);

      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(1000);

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.getWorksheet("ISAK");
      const headers = worksheet?.getRow(1).values;

      expect(headers).toEqual([
        undefined,
        "ID Participante",
        "Data",
        "Rodada",
        "Dobra subescapular (mm)",
        "Dobra de tríceps (mm)",
        "Dobra de bíceps (mm)",
        "Dobra de crista ilíaca (mm)",
        "Dobra supraespinhal (mm)",
        "Dobra abdominal (mm)",
        "Dobra de coxa anterior (mm)",
        "Dobra de panturrilha medial (mm)",
        "Perímetro de tórax (cm)",
        "Perímetro de braço relaxado (cm)",
        "Perímetro de braço contraído (cm)",
        "Perímetro de cintura (cm)",
        "Perímetro de abdome (cm)",
        "Perímetro de quadril (cm)",
        "Perímetro de coxa média (cm)",
        "Perímetro de panturrilha medial (cm)",
      ]);
    });
  });
});
