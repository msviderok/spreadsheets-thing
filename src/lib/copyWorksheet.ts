import Excel from "exceljs";

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

			// Copy cell value
			if (cell.type === Excel.ValueType.Formula) {
				const formulaValue = cell.value as Excel.CellFormulaValue;
				newCell.value = {
					formula: formulaValue.formula,
					result: formulaValue.result,
				};
			} else {
				newCell.value = cell.value;
			}

			if (cell.style) newCell.style = JSON.parse(JSON.stringify(cell.style));
			if (cell.numFmt) newCell.numFmt = cell.numFmt;
			if (cell.border) newCell.border = JSON.parse(JSON.stringify(cell.border));
			if (cell.fill) newCell.fill = JSON.parse(JSON.stringify(cell.fill));
			if (cell.font) newCell.font = JSON.parse(JSON.stringify(cell.font));
			if (cell.alignment)
				newCell.alignment = JSON.parse(JSON.stringify(cell.alignment));
			if (cell.protection)
				newCell.protection = JSON.parse(JSON.stringify(cell.protection));
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
