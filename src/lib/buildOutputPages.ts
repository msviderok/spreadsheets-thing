import type Excel from "exceljs";
import { copyCellRange, copyWorksheet, numberToCol } from "./copy";
import { sanitizeStr } from "./utils";

/**
 * Builds output page sheets for each distinct name
 */
export function buildOutputPages({
	templateWorkbook,
	workbook,
	data,
}: {
	templateWorkbook: Excel.Workbook;
	workbook: Excel.Workbook;
	data: Data;
}): void {
	const outputPageTemplate = templateWorkbook.getWorksheet("OUTPUT_PAGE")!;

	for (const name in data.list) {
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

		// Copy O8-T12 range for each entity in repetitive chunks
		// Each entity gets 6 columns (O through T)
		const COLUMNS_PER_ENTITY = 6; // O, P, Q, R, S, T
		const SOURCE_RANGE = "O8:T14";

		const addresses: Record<string, string[]> = {};

		data.entitiesArray.forEach((entity, entityIndex) => {
			// Calculate target start column: O (15) + (entityIndex * 6)
			const targetStartCol = 15 + entityIndex * COLUMNS_PER_ENTITY;
			const targetStartCell = `${numberToCol(targetStartCol)}8`;

			// Copy the range from template to target position
			copyCellRange({
				sourceSheet: outputPageTemplate,
				targetSheet: outputPageSheet,
				sourceRange: SOURCE_RANGE,
				targetStartCell,
			});

			// Set entity name in O8-T9 range (relative to the chunk)
			// O8-T9 means rows 8-9, columns O through T (6 columns)
			for (let row = 8; row <= 9; row++) {
				for (let colOffset = 0; colOffset < COLUMNS_PER_ENTITY; colOffset++) {
					const col = targetStartCol + colOffset;
					const cell = outputPageSheet.getCell(row, col);
					cell.value = entity;
				}
			}

			// Store column addresses for this entity
			const entityColumns: string[] = [];
			for (let colOffset = 0; colOffset < COLUMNS_PER_ENTITY; colOffset++) {
				const col = targetStartCol + colOffset;
				entityColumns.push(numberToCol(col));
			}

			addresses[entity] = entityColumns;
		});

		const leftoversRow = outputPageSheet.getRow(13);
		const leftovers = data.list[name]!.leftovers;

		data.entitiesArray.forEach((entity) => {
			const entityColumns = addresses[entity];
			if (!entityColumns || entityColumns.length !== 6) return;

			const [totalCol, I, II, III, IV, V] = entityColumns;
			leftoversRow.getCell(totalCol).value = { formula: `=SUM(${I}13:${V}13)` };
			leftoversRow.getCell(I).value = leftovers.entities[entity]?.[0];
			leftoversRow.getCell(II).value = leftovers.entities[entity]?.[1];
			leftoversRow.getCell(III).value = leftovers.entities[entity]?.[2];
			leftoversRow.getCell(IV).value = leftovers.entities[entity]?.[3];
			leftoversRow.getCell(V).value = leftovers.entities[entity]?.[4];
		});
	}
}
