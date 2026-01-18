import Excel from "exceljs";
import { copyWorksheet } from "@/lib/copy";
import { buildContentsSheet } from "./buildContentsSheet";
import { buildExtraSheet } from "./buildExtraSheed";
import { buildOutputPages } from "./buildOutputPages";

export async function generateOutputFile(data: Data, templateWorkbook: Excel.Workbook) {
	const newWorkbook = new Excel.Workbook();
	const contentsTemplate = templateWorkbook.getWorksheet("CONTENTS")!;
	const extraTemplate = templateWorkbook.getWorksheet("EXTRA")!;

	const contentsSheet = copyWorksheet({
		template: contentsTemplate,
		workbook: newWorkbook,
		newSheetName: "Зміст",
	});

	const extraSheet = copyWorksheet({
		template: extraTemplate,
		workbook: newWorkbook,
		newSheetName: "Додаток 1",
	});

	buildOutputPages({
		templateWorkbook,
		workbook: newWorkbook,
		data,
	});

	buildContentsSheet({ data, sheet: contentsSheet });
	buildExtraSheet({ data, sheet: extraSheet });

	const buffer = await newWorkbook.xlsx.writeBuffer();
	return buffer;
}
