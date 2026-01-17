import type Excel from "exceljs";
import { copyWorksheet } from "./copy";
import { sanitizeStr } from "./utils";

const HEADER_RANGES = ["A1:F12", "A18:F22"];

export function buildExtraSheet({
	templateWorkbook,
	workbook,
	distinctNames,
}: {
	templateWorkbook: Excel.Workbook;
	workbook: Excel.Workbook;
	distinctNames: Set<string>;
}): void {
	const extraTemplate = templateWorkbook.getWorksheet("EXTRA")!;
}
