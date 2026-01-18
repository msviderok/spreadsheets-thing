import type Excel from "exceljs";
import { copyCellRange } from "./copy";
import { sanitizeStr } from "./utils";

const DATA_START_ROW = 4; // Data starts at row 4
const ROWS_PER_COLUMN_GROUP = 20; // Maximum rows per column group
const COLUMNS_PER_GROUP = 3; // Columns A-C (name, link, next)
const COLUMN_SPACING = 0; // Space between column groups (A-C, then E-G, etc.)
const NAME_COLUMN_OFFSET = 0; // Column A (relative to group start)
const LINK_COLUMN_OFFSET = 1; // Column B (relative to group start)

/**
 * Builds the Contents (Зміст) sheet with multi-column layout
 */
export function buildContentsSheet({ sheet, data }: { sheet: Excel.Worksheet; data: Data }): void {
	const distinctNamesArray = data.listArray;
	const totalEntries = data.listArray.length;
	const columnGroupsNeeded = Math.ceil(totalEntries / ROWS_PER_COLUMN_GROUP);

	for (let groupIndex = 1; groupIndex < columnGroupsNeeded; groupIndex++) {
		const targetStartCol = 1 + groupIndex * (COLUMNS_PER_GROUP + COLUMN_SPACING);
		const targetStartCell = `${sheet.getColumn(targetStartCol).letter}1`;
		copyCellRange({ targetSheet: sheet, sourceRange: "A1:C4", targetStartCell });
	}

	if (ROWS_PER_COLUMN_GROUP > 1) {
		for (let i = 1; i < ROWS_PER_COLUMN_GROUP; i++) {
			sheet.duplicateRow(DATA_START_ROW, 1, true);
		}
	}

	let entryIndex = 0;

	for (let groupIndex = 0; groupIndex < columnGroupsNeeded; groupIndex++) {
		const groupStartCol = 1 + groupIndex * (COLUMNS_PER_GROUP + COLUMN_SPACING);
		const entriesInThisGroup = Math.min(ROWS_PER_COLUMN_GROUP, totalEntries - entryIndex);

		for (let i = 0; i < entriesInThisGroup; i++) {
			const rowIndex = DATA_START_ROW + i;
			const name = distinctNamesArray[entryIndex];
			const sheetIndex = entryIndex + 1;

			sheet.getCell(rowIndex, groupStartCol + NAME_COLUMN_OFFSET).value = name;
			sheet.getCell(rowIndex, groupStartCol + LINK_COLUMN_OFFSET).value = {
				text: `Аркуш ${sheetIndex}`,
				hyperlink: `#'${sanitizeStr(name)}'!A1`,
			};

			entryIndex++;
		}

		if (sheet.model?.merges) {
			const mergesToRemove: string[] = [];
			for (let i = 0; i < sheet.model.merges.length; i++) {
				const mergeRange = sheet.model.merges[i];
				if (!mergeRange) return;
				const match = mergeRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
				if (!match) return;
				const [, , startRow, , endRow] = match;
				if (Number(startRow) === 1 && Number(endRow) === 1) {
					mergesToRemove.push(mergeRange);
				}
			}

			for (let i = 0; i < mergesToRemove.length; i++) {
				const range = mergesToRemove[i];
				try {
					sheet.unMergeCells(range);
				} catch {}
			}
		}

		sheet.mergeCells(1, 1, 1, columnGroupsNeeded * COLUMNS_PER_GROUP);
		const titleCell = sheet.getCell(1, 1);
		titleCell.value = "Зміст";
	}
}
