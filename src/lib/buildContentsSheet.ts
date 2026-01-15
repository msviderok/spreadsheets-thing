import type Excel from "exceljs";
import { copyCellProperties, copyWorksheet } from "./copyWorksheet";
import { sanitizeStr } from "./sanitizeStr";

interface BuildContentsSheetParams {
	templateWorkbook: Excel.Workbook;
	workbook: Excel.Workbook;
	distinctNames: Set<string>;
}

// Constants for Contents sheet layout
const HEADER_ROWS = 3; // Rows 1-3 are headers
const DATA_START_ROW = 4; // Data starts at row 4
const ROWS_PER_COLUMN_GROUP = 20; // Maximum rows per column group
const COLUMNS_PER_GROUP = 3; // Columns A-C (name, link, next)
const COLUMN_SPACING = 0; // Space between column groups (A-C, then E-G, etc.)
const NAME_COLUMN_OFFSET = 0; // Column A (relative to group start)
const LINK_COLUMN_OFFSET = 1; // Column B (relative to group start)
const NEXT_COLUMN_OFFSET = 2; // Column C (relative to group start)

/**
 * Builds the Contents (Зміст) sheet with multi-column layout
 */
export function buildContentsSheet({
	templateWorkbook,
	workbook,
	distinctNames,
}: BuildContentsSheetParams): void {
	const contentsTemplate = templateWorkbook.getWorksheet("CONTENTS")!;

	const contentsSheet = copyWorksheet({
		template: contentsTemplate,
		workbook,
		newSheetName: "Зміст",
	});

	// Get template cells from the original template for copying styles
	// We'll use these template cells to copy properties to data cells
	const templateDataRow = contentsTemplate.getRow(DATA_START_ROW);
	const templateBorders: Record<number, Partial<Excel.Borders>> = {};
	const templateCells: Record<number, Excel.Cell> = {};

	templateDataRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
		if (cell.border) {
			templateBorders[colNumber] = JSON.parse(JSON.stringify(cell.border));
		}
		// Save template cells for copying properties
		if (colNumber <= COLUMNS_PER_GROUP) {
			templateCells[colNumber] = cell;
		}
	});

	const totalEntries = distinctNames.size;
	const columnGroupsNeeded = Math.ceil(totalEntries / ROWS_PER_COLUMN_GROUP);
	const distinctNamesArray = Array.from(distinctNames);

	// First, expand rows for the first column group (duplicate row 4 to fill ROWS_PER_COLUMN_GROUP rows)
	if (ROWS_PER_COLUMN_GROUP > 1) {
		for (let i = 1; i < ROWS_PER_COLUMN_GROUP; i++) {
			contentsSheet.duplicateRow(DATA_START_ROW, 1, true);
		}
	}

	// Helper function to parse merged cell range
	const parseMergeRange = (range: string) => {
		const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
		if (!match) return null;
		const [, startCol, startRow, endCol, endRow] = match;

		const colToNumber = (col: string) => {
			let total = 0;
			for (const char of col) {
				total = total * 26 + (char.charCodeAt(0) - 64);
			}
			return total;
		};

		const numberToCol = (num: number) => {
			let result = "";
			let n = num;
			while (n > 0) {
				const remainder = (n - 1) % 26;
				result = String.fromCharCode(65 + remainder) + result;
				n = Math.floor((n - 1) / 26);
			}
			return result;
		};

		return {
			startRow: Number(startRow),
			endRow: Number(endRow),
			startCol: colToNumber(startCol),
			endCol: colToNumber(endCol),
			numberToCol,
		};
	};

	// Copy headers to each new column group
	for (let groupIndex = 1; groupIndex < columnGroupsNeeded; groupIndex++) {
		const sourceStartCol = 1;
		const sourceEndCol = COLUMNS_PER_GROUP;
		const targetStartCol =
			1 + groupIndex * (COLUMNS_PER_GROUP + COLUMN_SPACING);
		const colOffset = targetStartCol - sourceStartCol;

		// Copy all cells in header range (rows 1-3, columns 1-COLUMNS_PER_GROUP) from template
		for (let headerRow = 1; headerRow <= HEADER_ROWS; headerRow++) {
			for (let colIdx = 0; colIdx < COLUMNS_PER_GROUP; colIdx++) {
				const sourceCol = sourceStartCol + colIdx;
				const targetCol = targetStartCol + colIdx;
				const sourceCell = contentsTemplate.getCell(headerRow, sourceCol);
				const targetCell = contentsSheet.getCell(headerRow, targetCol);
				copyCellProperties(sourceCell, targetCell);
			}
		}

		// Copy column widths from template to new column group
		for (let colIdx = 0; colIdx < COLUMNS_PER_GROUP; colIdx++) {
			const sourceCol = sourceStartCol + colIdx;
			const targetCol = targetStartCol + colIdx;
			const sourceColumn = contentsTemplate.getColumn(sourceCol);
			const targetColumn = contentsSheet.getColumn(targetCol);
			if (sourceColumn.width !== undefined) {
				targetColumn.width = sourceColumn.width;
			}
		}

		// Copy merged cells that overlap with the header range from the template
		if (contentsTemplate.model?.merges) {
			contentsTemplate.model.merges.forEach((mergeRange) => {
				if (!mergeRange) return;

				const parsed = parseMergeRange(mergeRange);
				if (!parsed) return;

				// Check if merge is within header rows (1-3) and overlaps with source columns
				const isInHeaderRows =
					parsed.startRow >= 1 &&
					parsed.startRow <= HEADER_ROWS &&
					parsed.endRow >= 1 &&
					parsed.endRow <= HEADER_ROWS;

				const overlapsSourceColumns =
					parsed.startCol <= sourceEndCol && parsed.endCol >= sourceStartCol;

				if (isInHeaderRows && overlapsSourceColumns) {
					// Calculate the intersection with our source column range
					const intersectStartRow = Math.max(parsed.startRow, 1);
					const intersectEndRow = Math.min(parsed.endRow, HEADER_ROWS);
					const intersectStartCol = Math.max(parsed.startCol, sourceStartCol);
					const intersectEndCol = Math.min(parsed.endCol, sourceEndCol);

					// Only merge if we have more than one cell
					const hasMultipleCells =
						intersectEndRow > intersectStartRow ||
						intersectEndCol > intersectStartCol;

					if (hasMultipleCells) {
						// Translate to target column group
						const targetStartRow = intersectStartRow;
						const targetEndRow = intersectEndRow;
						const targetStartColNum = intersectStartCol + colOffset;
						const targetEndColNum = intersectEndCol + colOffset;

						// Create new merge range
						const newRange = `${parsed.numberToCol(targetStartColNum)}${targetStartRow}:${parsed.numberToCol(targetEndColNum)}${targetEndRow}`;

						// Check if cells are already merged - if so, unmerge first
						const startCell = contentsSheet.getCell(
							targetStartRow,
							targetStartColNum,
						);
						if (startCell.isMerged) {
							try {
								contentsSheet.unMergeCells(newRange);
							} catch {
								// Ignore unmerge errors
							}
						}

						try {
							contentsSheet.mergeCells(newRange);
						} catch {
							// Merge might already exist or be invalid
							// Try alternative merge method using coordinates
							try {
								contentsSheet.mergeCells(
									targetStartRow,
									targetStartColNum,
									targetEndRow,
									targetEndColNum,
								);
							} catch {
								// If all methods fail, skip this merge
							}
						}
					}
				}
			});
		}
	}

	// Process each column group - all groups use the same rows (DATA_START_ROW to DATA_START_ROW + ROWS_PER_COLUMN_GROUP - 1)
	let entryIndex = 0;

	for (let groupIndex = 0; groupIndex < columnGroupsNeeded; groupIndex++) {
		const groupStartCol = 1 + groupIndex * (COLUMNS_PER_GROUP + COLUMN_SPACING);
		const entriesInThisGroup = Math.min(
			ROWS_PER_COLUMN_GROUP,
			totalEntries - entryIndex,
		);

		// Update rows with values for this column group
		for (let i = 0; i < entriesInThisGroup; i++) {
			const rowIndex = DATA_START_ROW + i;
			const name = distinctNamesArray[entryIndex];
			const sanitizedName = sanitizeStr(name);
			const sheetName = sanitizedName;
			const sheetIndex = entryIndex + 1;

			// Column A (or corresponding column in group): name of the output_page
			const nameCell = contentsSheet.getCell(
				rowIndex,
				groupStartCol + NAME_COLUMN_OFFSET,
			);
			// Copy all properties from template cell first
			const nameTemplateCell = templateCells[NAME_COLUMN_OFFSET + 1];
			if (nameTemplateCell) {
				copyCellProperties(nameTemplateCell, nameCell);
			}
			// Then set the value (this preserves all copied styles)
			nameCell.value = name;

			// Column B (or corresponding column in group): hyperlink to the output_page sheet
			const linkCell = contentsSheet.getCell(
				rowIndex,
				groupStartCol + LINK_COLUMN_OFFSET,
			);
			// Copy all properties from template cell first
			const linkTemplateCell = templateCells[LINK_COLUMN_OFFSET + 1];
			if (linkTemplateCell) {
				copyCellProperties(linkTemplateCell, linkCell);
			}
			// Then set the hyperlink value (this preserves all copied styles)
			linkCell.value = {
				text: `Аркуш ${sheetIndex}`,
				hyperlink: `#'${sheetName}'!A1`,
			};

			// Column C: copy formatting even though it's empty
			const nextCell = contentsSheet.getCell(
				rowIndex,
				groupStartCol + NEXT_COLUMN_OFFSET,
			);
			// Copy all properties from template cell
			const nextTemplateCell = templateCells[NEXT_COLUMN_OFFSET + 1];
			if (nextTemplateCell) {
				copyCellProperties(nextTemplateCell, nextCell);
			}

			// Remove bottom border from all rows except the last row in this group
			// This preserves top, left, and right borders (and their colors) for all cells
			const isLastRowInGroup = i === entriesInThisGroup - 1;
			if (!isLastRowInGroup) {
				const row = contentsSheet.getRow(rowIndex);
				// Only remove bottom border from cells in this column group (A, B, C)
				for (let colOffset = 0; colOffset < COLUMNS_PER_GROUP; colOffset++) {
					const cell = row.getCell(groupStartCol + colOffset);
					// Preserve existing borders but remove bottom border
					// Use deep copy to preserve nested border objects with colors
					if (cell.border) {
						const borderCopy = JSON.parse(JSON.stringify(cell.border));
						delete borderCopy.bottom;
						cell.border = borderCopy;
					}
				}
			}

			entryIndex++;
		}

		// Ensure borders are correctly applied to the last row in this column group
		// Copy borders directly from template cells to preserve colors correctly
		const lastRowInGroupIndex = DATA_START_ROW + entriesInThisGroup - 1;
		const lastRow = contentsSheet.getRow(lastRowInGroupIndex);
		for (let colOffset = 0; colOffset < COLUMNS_PER_GROUP; colOffset++) {
			const sourceCol = colOffset + 1;
			const targetCol = groupStartCol + colOffset;
			const templateCell = templateCells[sourceCol];
			const cell = lastRow.getCell(targetCol);
			if (templateCell?.border) {
				// Copy border directly from template cell to preserve color
				cell.border = JSON.parse(JSON.stringify(templateCell.border));
			}
		}
	}
}
