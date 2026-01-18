import type Excel from "exceljs";
import { FIRST_DATA_ROW } from "@/lib/buildOutputPages";
import { copyCellRange } from "@/lib/copy";
import { sanitizeStr } from "@/lib/utils";

const FIRST_ROW = 14;

export function buildExtraSheet({ sheet, data }: { sheet: Excel.Worksheet; data: Data }): void {
	const rowCount = data.listArray.length;
	const lastRow = FIRST_ROW + rowCount - 1;

	sheet.duplicateRow(FIRST_ROW, data.listArray.length - 1, true);

	sheet.fillFormula(`A${FIRST_ROW}:A${lastRow}`, `ROW()-${FIRST_ROW - 1}`);

	for (let i = 0; i < data.listArray.length; i++) {
		const name = data.listArray[i];
		const item = data.list[name];
		const rowIdx = FIRST_ROW + i;
		const row = sheet.getRow(rowIdx);
		const statementsLastRow = FIRST_DATA_ROW + item.statements.length - 1;
		const sanitizedName = sanitizeStr(name);
		row.getCell("B").value = name;
		row.getCell("D").value = { formula: `=SUM('${sanitizedName}'!G${FIRST_DATA_ROW}:G${statementsLastRow})` };
		row.getCell("E").value = { formula: `=SUM('${sanitizedName}'!H${FIRST_DATA_ROW}:H${statementsLastRow})` };
		row.getCell("F").value = { formula: `='${sanitizedName}'!I${statementsLastRow}` };
	}

	copyCellRange({
		targetSheet: sheet,
		targetStartCell: `A${lastRow + 4}`,
		sourceRange: "G1:K5",
		onlyText: true,
	});

	sheet.spliceColumns(7, 5);
}
