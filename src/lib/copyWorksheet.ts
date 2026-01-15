import Excel from "exceljs";

/**
 * Copies all properties from a source cell to a target cell
 */
export function copyCellProperties(
	sourceCell: Excel.Cell,
	targetCell: Excel.Cell,
): void {
	// Copy cell value
	if (sourceCell.type === Excel.ValueType.Formula) {
		const formulaValue = sourceCell.value as Excel.CellFormulaValue;
		targetCell.value = {
			formula: formulaValue.formula,
			result: formulaValue.result,
		};
	} else {
		targetCell.value = sourceCell.value;
	}

	// Copy all cell properties
	if (sourceCell.style) {
		targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
	}
	if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
	if (sourceCell.border) {
		targetCell.border = JSON.parse(JSON.stringify(sourceCell.border));
	}
	if (sourceCell.fill) {
		targetCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
	}
	if (sourceCell.font) {
		targetCell.font = JSON.parse(JSON.stringify(sourceCell.font));
	}
	if (sourceCell.alignment) {
		targetCell.alignment = JSON.parse(JSON.stringify(sourceCell.alignment));
	}
	if (sourceCell.protection) {
		targetCell.protection = JSON.parse(JSON.stringify(sourceCell.protection));
	}
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
