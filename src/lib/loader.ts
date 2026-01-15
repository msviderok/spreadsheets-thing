import Excel from "exceljs";
import { initMatrix } from "@/lib/initMatrix";
import mockDb from "../mock.xlsx?url";
import workbookUrl from "../output_template.xlsx?url";

async function getDb() {
	const mockResponse = await fetch(mockDb);
	const mockArrayBuffer = await mockResponse.arrayBuffer();
	const mockWorkbook = new Excel.Workbook();
	await mockWorkbook.xlsx.load(mockArrayBuffer);
	return initMatrix(mockWorkbook.worksheets[0], "db");
}

async function getTemplates() {
	const response = await fetch(workbookUrl);
	const arrayBuffer = await response.arrayBuffer();
	const workbook = new Excel.Workbook();
	await workbook.xlsx.load(arrayBuffer);
	return {
		workbook,
		contents: initMatrix(workbook.worksheets[0], "contents"),
		extra: initMatrix(workbook.worksheets[1], "extra"),
		outputPage: initMatrix(workbook.worksheets[2], "output_page"),
	};
}

export const loader = async () => {
	const db = await getDb();
	const templates = await getTemplates();
	return { db, ...templates };
};
