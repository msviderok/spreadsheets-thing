import type Excel from "exceljs";
import { copyWorksheet } from "./copyWorksheet";
import { sanitizeStr } from "./sanitizeStr";

interface BuildOutputPagesParams {
	templateWorkbook: Excel.Workbook;
	workbook: Excel.Workbook;
	distinctNames: Set<string>;
}

/**
 * Builds output page sheets for each distinct name
 */
export function buildOutputPages({
	templateWorkbook,
	workbook,
	distinctNames,
}: BuildOutputPagesParams): void {
	const outputPageTemplate = templateWorkbook.getWorksheet("OUTPUT_PAGE")!;

	for (const name of distinctNames) {
		const outputPageSheet = copyWorksheet({
			template: outputPageTemplate,
			workbook,
			newSheetName: sanitizeStr(name),
		});

		outputPageSheet.getCell("B3").value = {
			text: "←Повернутись на зміст",
			hyperlink: `#'Зміст'!A1`,
		};
		outputPageSheet.getCell("F5").value = name;
	}
}
