import Excel from "exceljs";

export function copyCellProperties(sourceCell: Excel.Cell, targetCell: Excel.Cell): void {
	targetCell.value = sourceCell.value;
	targetCell.numFmt = sourceCell.numFmt;
	targetCell.protection = sourceCell.protection;
	targetCell.style = sourceCell.style;

	copyCellStyle(sourceCell, targetCell);
}

export function copyCellStyle(sourceCell: Excel.Cell, targetCell: Excel.Cell): void {
	targetCell.border = sourceCell.border;
	targetCell.fill = sourceCell.fill;
	targetCell.font = sourceCell.font;
	targetCell.alignment = sourceCell.alignment;
}

export function copyWorksheet({
	template,
	workbook,
	newSheetName,
}: {
	template: Excel.Worksheet;
	workbook: Excel.Workbook;
	newSheetName: string;
}): Excel.Worksheet {
	const newSheet = workbook.addWorksheet(newSheetName);

	// Copy all cells with their values, styles, and formulas
	template.eachRow({ includeEmpty: true }, (row, rowNumber) => {
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const newCell = newSheet.getCell(rowNumber, colNumber);
			copyCellProperties(cell, newCell);
		});
	});

	// Copy merged cells using the merges property
	if (template.model?.merges) {
		template.model.merges.forEach((mergeRange) => {
			if (mergeRange) {
				newSheet.mergeCells(mergeRange);
			}
		});
	}

	// Copy column widths
	template.columns.forEach((column, index) => {
		if (column.width !== undefined) {
			const newColumn = newSheet.getColumn(index + 1);
			newColumn.width = column.width;
		}
	});

	// Copy row heights and hidden state
	template.eachRow((row, rowNumber) => {
		const newRow = newSheet.getRow(rowNumber);
		if (row.height !== undefined) {
			newRow.height = row.height;
		}
		if (row.hidden !== undefined) {
			newRow.hidden = row.hidden;
		}
	});

	return newSheet;
}

/**
 * Helper function to convert column letter to number
 */
function colToNumber(col: string): number {
	let total = 0;
	for (const char of col) {
		total = total * 26 + (char.charCodeAt(0) - 64);
	}
	return total;
}

/**
 * Helper function to convert column number to letter
 */
export function numberToCol(num: number): string {
	let result = "";
	let n = num;
	while (n > 0) {
		const remainder = (n - 1) % 26;
		result = String.fromCharCode(65 + remainder) + result;
		n = Math.floor((n - 1) / 26);
	}
	return result;
}

/**
 * Helper function to parse merged cell range
 */
function parseMergeRange(range: string) {
	const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
	if (!match) return null;
	const [, startCol, startRow, endCol, endRow] = match;

	return {
		startRow: Number(startRow),
		endRow: Number(endRow),
		startCol: colToNumber(startCol),
		endCol: colToNumber(endCol),
		numberToCol,
	};
}

/**
 * Copies a header column group from template sheet to target sheet
 * This includes copying cells, column widths, and merged cells in the header range
 */
export function copyHeaderColumnGroup({
	templateSheet,
	targetSheet,
	sourceStartCol,
	sourceEndCol,
	targetStartCol,
	headerRows,
}: {
	templateSheet: Excel.Worksheet;
	targetSheet: Excel.Worksheet;
	sourceStartCol: number;
	sourceEndCol: number;
	targetStartCol: number;
	headerRows: number;
}): void {
	const colOffset = targetStartCol - sourceStartCol;
	const columnsPerGroup = sourceEndCol - sourceStartCol + 1;

	// Copy all cells in header range (rows 1-headerRows, columns sourceStartCol-sourceEndCol) from template
	for (let headerRow = 1; headerRow <= headerRows; headerRow++) {
		for (let colIdx = 0; colIdx < columnsPerGroup; colIdx++) {
			const sourceCol = sourceStartCol + colIdx;
			const targetCol = targetStartCol + colIdx;
			const sourceCell = templateSheet.getCell(headerRow, sourceCol);
			const targetCell = targetSheet.getCell(headerRow, targetCol);
			copyCellProperties(sourceCell, targetCell);
		}
	}

	// Copy column widths from template to new column group
	for (let colIdx = 0; colIdx < columnsPerGroup; colIdx++) {
		const sourceCol = sourceStartCol + colIdx;
		const targetCol = targetStartCol + colIdx;
		const sourceColumn = templateSheet.getColumn(sourceCol);
		const targetColumn = targetSheet.getColumn(targetCol);
		if (sourceColumn.width !== undefined) {
			targetColumn.width = sourceColumn.width;
		}
	}

	// Copy merged cells that overlap with the header range from the template
	if (templateSheet.model?.merges) {
		templateSheet.model.merges.forEach((mergeRange) => {
			if (!mergeRange) return;

			const parsed = parseMergeRange(mergeRange);
			if (!parsed) return;

			// Check if merge is within header rows (1-headerRows) and overlaps with source columns
			const isInHeaderRows =
				parsed.startRow >= 1 && parsed.startRow <= headerRows && parsed.endRow >= 1 && parsed.endRow <= headerRows;

			const overlapsSourceColumns = parsed.startCol <= sourceEndCol && parsed.endCol >= sourceStartCol;

			if (isInHeaderRows && overlapsSourceColumns) {
				// Calculate the intersection with our source column range
				const intersectStartRow = Math.max(parsed.startRow, 1);
				const intersectEndRow = Math.min(parsed.endRow, headerRows);
				const intersectStartCol = Math.max(parsed.startCol, sourceStartCol);
				const intersectEndCol = Math.min(parsed.endCol, sourceEndCol);

				// Only merge if we have more than one cell
				const hasMultipleCells = intersectEndRow > intersectStartRow || intersectEndCol > intersectStartCol;

				if (hasMultipleCells) {
					// Translate to target column group
					const targetStartRow = intersectStartRow;
					const targetEndRow = intersectEndRow;
					const targetStartColNum = intersectStartCol + colOffset;
					const targetEndColNum = intersectEndCol + colOffset;

					// Create new merge range
					const newRange = `${parsed.numberToCol(targetStartColNum)}${targetStartRow}:${parsed.numberToCol(targetEndColNum)}${targetEndRow}`;

					// Check if cells are already merged - if so, unmerge first
					const startCell = targetSheet.getCell(targetStartRow, targetStartColNum);
					if (startCell.isMerged) {
						try {
							targetSheet.unMergeCells(newRange);
						} catch {
							// Ignore unmerge errors
						}
					}

					try {
						targetSheet.mergeCells(newRange);
					} catch {
						// Merge might already exist or be invalid
						// Try alternative merge method using coordinates
						try {
							targetSheet.mergeCells(targetStartRow, targetStartColNum, targetEndRow, targetEndColNum);
						} catch {
							// If all methods fail, skip this merge
						}
					}
				}
			}
		});
	}
}

/**
 * Copies a cell range from source sheet to target sheet at a new position
 * @param sourceRange - Excel range string like "O8:T12"
 * @param targetStartCell - Excel cell string like "O8" where to start copying
 */
export function copyCellRange({
	sourceSheet,
	targetSheet,
	sourceRange,
	targetStartCell,
}: {
	sourceSheet: Excel.Worksheet;
	targetSheet: Excel.Worksheet;
	sourceRange: string;
	targetStartCell: string;
}): void {
	// Parse source range
	const sourceMatch = sourceRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
	if (!sourceMatch) throw new Error(`Invalid source range: ${sourceRange}`);
	const [, sourceStartCol, sourceStartRow, sourceEndCol, sourceEndRow] = sourceMatch;
	const sourceStartColNum = colToNumber(sourceStartCol);
	const sourceEndColNum = colToNumber(sourceEndCol);
	const sourceStartRowNum = Number(sourceStartRow);
	const sourceEndRowNum = Number(sourceEndRow);

	// Parse target start cell
	const targetMatch = targetStartCell.match(/^([A-Z]+)(\d+)$/);
	if (!targetMatch) throw new Error(`Invalid target cell: ${targetStartCell}`);
	const [, targetStartCol, targetStartRow] = targetMatch;
	const targetStartColNum = colToNumber(targetStartCol);
	const targetStartRowNum = Number(targetStartRow);

	// Calculate offsets
	const rowOffset = targetStartRowNum - sourceStartRowNum;
	const colOffset = targetStartColNum - sourceStartColNum;

	// Copy all cells in the range
	for (let row = sourceStartRowNum; row <= sourceEndRowNum; row++) {
		for (let col = sourceStartColNum; col <= sourceEndColNum; col++) {
			const sourceCell = sourceSheet.getCell(row, col);
			const targetCell = targetSheet.getCell(row + rowOffset, col + colOffset);
			copyCellProperties(sourceCell, targetCell);
		}
	}

	// Copy column widths
	for (let col = sourceStartColNum; col <= sourceEndColNum; col++) {
		const sourceColumn = sourceSheet.getColumn(col);
		const targetColumn = targetSheet.getColumn(col + colOffset);
		if (sourceColumn.width !== undefined) {
			targetColumn.width = sourceColumn.width;
		}
	}

	// Copy row heights, hidden state, and row-level styles
	for (let row = sourceStartRowNum; row <= sourceEndRowNum; row++) {
		const sourceRow = sourceSheet.getRow(row);
		const targetRow = targetSheet.getRow(row + rowOffset);
		if (sourceRow.height !== undefined) {
			targetRow.height = sourceRow.height;
		}
	}

	// Copy merged cells that overlap with the source range
	if (sourceSheet.model?.merges) {
		sourceSheet.model.merges.forEach((mergeRange) => {
			if (!mergeRange) return;

			const parsed = parseMergeRange(mergeRange);
			if (!parsed) return;

			// Check if merge overlaps with source range
			const overlaps =
				parsed.startRow <= sourceEndRowNum &&
				parsed.endRow >= sourceStartRowNum &&
				parsed.startCol <= sourceEndColNum &&
				parsed.endCol >= sourceStartColNum;

			if (overlaps) {
				// Calculate intersection with source range
				const intersectStartRow = Math.max(parsed.startRow, sourceStartRowNum);
				const intersectEndRow = Math.min(parsed.endRow, sourceEndRowNum);
				const intersectStartCol = Math.max(parsed.startCol, sourceStartColNum);
				const intersectEndCol = Math.min(parsed.endCol, sourceEndColNum);

				// Only merge if we have more than one cell
				const hasMultipleCells = intersectEndRow > intersectStartRow || intersectEndCol > intersectStartCol;

				if (hasMultipleCells) {
					// Translate to target position
					const targetStartRow = intersectStartRow + rowOffset;
					const targetEndRow = intersectEndRow + rowOffset;
					const targetStartColNum = intersectStartCol + colOffset;
					const targetEndColNum = intersectEndCol + colOffset;

					// Create new merge range
					const newRange = `${numberToCol(targetStartColNum)}${targetStartRow}:${numberToCol(targetEndColNum)}${targetEndRow}`;

					// Check if cells are already merged - if so, unmerge first
					const startCell = targetSheet.getCell(targetStartRow, targetStartColNum);
					if (startCell.isMerged) {
						try {
							targetSheet.unMergeCells(newRange);
						} catch {
							// Ignore unmerge errors
						}
					}

					try {
						targetSheet.mergeCells(newRange);
					} catch {
						// Merge might already exist or be invalid
						// Try alternative merge method using coordinates
						try {
							targetSheet.mergeCells(targetStartRow, targetStartColNum, targetEndRow, targetEndColNum);
						} catch {
							// If all methods fail, skip this merge
						}
					}
				}
			}
		});
	}
}
