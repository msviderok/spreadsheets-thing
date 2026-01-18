import type Excel from "exceljs";

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

	template.eachRow({ includeEmpty: true }, (row, rowNumber) => {
		const newRow = newSheet.getRow(rowNumber);
		newRow.height = row.height;
		newRow.hidden = row.hidden;
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const newCell = newSheet.getCell(rowNumber, colNumber);
			copyCellProperties(cell, newCell);
		});
	});

	if (template.model?.merges) {
		template.model.merges.forEach((mergeRange) => {
			newSheet.mergeCells(mergeRange);
		});
	}

	template.columns.forEach((column, index) => {
		newSheet.getColumn(index + 1).width = column.width;
	});

	return newSheet;
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
	sourceSheet?: Excel.Worksheet;
	targetSheet: Excel.Worksheet;
	sourceRange: string;
	targetStartCell: string;
}): void {
	const src = sourceSheet ?? targetSheet;
	const sourceMatch = sourceRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
	if (!sourceMatch) throw new Error(`Invalid source range: ${sourceRange}`);
	const [, sourceStartCol, sourceStartRow, sourceEndCol, sourceEndRow] = sourceMatch;
	const sourceStartColNum = src.getColumn(sourceStartCol).number;
	const sourceEndColNum = src.getColumn(sourceEndCol).number;
	const sourceStartRowNum = Number(sourceStartRow);
	const sourceEndRowNum = Number(sourceEndRow);

	const targetMatch = targetStartCell.match(/^([A-Z]+)(\d+)$/);
	if (!targetMatch) throw new Error(`Invalid target cell: ${targetStartCell}`);
	const [, targetStartCol, targetStartRow] = targetMatch;
	const targetStartColNum = targetSheet.getColumn(targetStartCol).number;
	const targetStartRowNum = Number(targetStartRow);

	const rowOffset = targetStartRowNum - sourceStartRowNum;
	const colOffset = targetStartColNum - sourceStartColNum;

	for (let row = sourceStartRowNum; row <= sourceEndRowNum; row++) {
		for (let col = sourceStartColNum; col <= sourceEndColNum; col++) {
			const sourceCell = src.getCell(row, col);
			const targetCell = targetSheet.getCell(row + rowOffset, col + colOffset);
			copyCellProperties(sourceCell, targetCell);
		}
	}

	for (let col = sourceStartColNum; col <= sourceEndColNum; col++) {
		targetSheet.getColumn(col + colOffset).width = src.getColumn(col).width;
	}

	for (let row = sourceStartRowNum; row <= sourceEndRowNum; row++) {
		targetSheet.getRow(row + rowOffset).height = src.getRow(row).height;
	}

	if (src.model?.merges) {
		src.model.merges.forEach((mergeRange) => {
			if (!mergeRange) return;

			const parsed = parseMergeRange(src, mergeRange);
			if (!parsed) return;

			const overlaps =
				parsed.startRow <= sourceEndRowNum &&
				parsed.endRow >= sourceStartRowNum &&
				parsed.startCol <= sourceEndColNum &&
				parsed.endCol >= sourceStartColNum;

			if (overlaps) {
				const intersectStartRow = Math.max(parsed.startRow, sourceStartRowNum);
				const intersectEndRow = Math.min(parsed.endRow, sourceEndRowNum);
				const intersectStartCol = Math.max(parsed.startCol, sourceStartColNum);
				const intersectEndCol = Math.min(parsed.endCol, sourceEndColNum);

				const hasMultipleCells = intersectEndRow > intersectStartRow || intersectEndCol > intersectStartCol;

				if (hasMultipleCells) {
					const targetStartRow = intersectStartRow + rowOffset;
					const targetEndRow = intersectEndRow + rowOffset;
					const targetStartColNum = intersectStartCol + colOffset;
					const targetEndColNum = intersectEndCol + colOffset;

					const newRange = `${targetSheet.getColumn(targetStartColNum).letter}${targetStartRow}:${targetSheet.getColumn(targetEndColNum).letter}${targetEndRow}`;

					const startCell = targetSheet.getCell(targetStartRow, targetStartColNum);
					if (startCell.isMerged) {
						try {
							targetSheet.unMergeCells(newRange);
						} catch {}
					}

					try {
						targetSheet.mergeCells(newRange);
					} catch {
						try {
							targetSheet.mergeCells(targetStartRow, targetStartColNum, targetEndRow, targetEndColNum);
						} catch {}
					}
				}
			}
		});
	}
}

function parseMergeRange(sheet: Excel.Worksheet, range: string) {
	const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
	if (!match) return null;
	const [, startCol, startRow, endCol, endRow] = match;

	return {
		startRow: Number(startRow),
		endRow: Number(endRow),
		startCol: sheet.getColumn(startCol).number,
		endCol: sheet.getColumn(endCol).number,
	};
}
