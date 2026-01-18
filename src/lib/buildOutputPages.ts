import type Excel from "exceljs";
import { copyCellRange, copyCellStyle, copyWorksheet } from "./copy";
import { sanitizeStr } from "./utils";

const FIRST_ROW = 13;

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
	const HEIGHT = outputPageTemplate.getRow(FIRST_ROW).height;

	for (const name in data.list) {
		const outputPageSheet = copyWorksheet({
			template: outputPageTemplate,
			workbook,
			newSheetName: sanitizeStr(name),
		});

		outputPageSheet.getCell("F5").value = name;
		outputPageSheet.getCell("B3").value = {
			text: "←Повернутись на зміст",
			hyperlink: `#'Зміст'!A1`,
		};

		const COLUMNS_PER_ENTITY = 6; // O, P, Q, R, S, T
		const SOURCE_RANGE = "O8:T13";
		const DATA_ROW = outputPageSheet.getRow(FIRST_ROW);
		const addresses: Record<string, string[]> = {};

		data.entitiesArray.forEach((entity, entityIndex) => {
			// Calculate target start column: O (15) + (entityIndex * 6)
			const targetStartCol = 15 + entityIndex * COLUMNS_PER_ENTITY;
			const targetStartCell = `${outputPageTemplate.getColumn(targetStartCol).letter}8`;

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
				const col = outputPageTemplate.getColumn(targetStartCol + colOffset).letter;
				entityColumns.push(col);
			}

			addresses[entity] = entityColumns;
		});

		const colCount = outputPageSheet.actualColumnCount;
		const lastCol = outputPageSheet.getColumn(colCount).letter;
		const leftovers = data.list[name]!.leftovers;
		const leftoversRow = outputPageSheet.getRow(FIRST_ROW);

		leftoversRow.getCell("A").value = leftovers.date;
		leftoversRow.getCell("B").value = "Перенесено з книги №10";

		data.entitiesArray.forEach((entity) => {
			const entityColumns = addresses[entity];
			if (!entityColumns || entityColumns.length !== 6) return;

			const [totalCol, I, II, III, IV, V] = entityColumns;
			leftoversRow.getCell(totalCol).value = { formula: `=SUM(${I}${FIRST_ROW}:${V}${FIRST_ROW})` };
			leftoversRow.getCell(I).value = leftovers.entities[entity]?.categories[0];
			leftoversRow.getCell(II).value = leftovers.entities[entity]?.categories[1];
			leftoversRow.getCell(III).value = leftovers.entities[entity]?.categories[2];
			leftoversRow.getCell(IV).value = leftovers.entities[entity]?.categories[3];
			leftoversRow.getCell(V).value = leftovers.entities[entity]?.categories[4];
		});

		const statements = data.list[name]!.statements;
		let row = FIRST_ROW;
		let prevRow = FIRST_ROW - 1;

		statements.forEach((statement) => {
			row++;
			prevRow++;
			const statementRow = outputPageSheet.getRow(row);
			statementRow.getCell("A").value = statement.date ?? "";
			statementRow.getCell("B").value = statement.docName ?? "";
			statementRow.getCell("C").value = statement.docNumber ?? "";
			statementRow.getCell("D").value = statement.docDate ?? "";
			statementRow.getCell("E").value = statement.from ?? "";
			statementRow.getCell("F").value = statement.to ?? "";
			statementRow.getCell("G").value = statement.quantityIn ?? "";
			statementRow.getCell("H").value = statement.quantityOut ?? "";

			const colToTakeValueFrom =
				statement.takeValueFrom === "in" ? "G" : statement.takeValueFrom === "out" ? "H" : undefined;

			data.entitiesArray.forEach((entityName) => {
				const entity = statement.entities[entityName];
				const entityColumns = addresses[entityName];
				if (!entityColumns || entityColumns.length !== 6) return;

				const [, I, , , , V] = entityColumns;
				for (let i = 0; i < entityColumns.length; i++) {
					const col = entityColumns[i];
					if (i === 0) {
						statementRow.getCell(col).value = { formula: `=SUM(${I}${row}:${V}${row})` };
						continue;
					}

					if (!colToTakeValueFrom) {
						throw new Error(`colToTakeValueFrom is not set for ${col}`);
					}

					if (entity?.categories[i - 1] !== undefined) {
						statementRow.getCell(col).value = {
							formula: `${col}${prevRow}+${colToTakeValueFrom}${row}*${entity.operation}`,
						};
					} else {
						statementRow.getCell(col).value = { formula: `${col}${prevRow}` };
						statementRow.getCell(col).numFmt = "0;-0;;";
					}
				}
			});
		});

		outputPageSheet.fillFormula(`A12:${lastCol}12`, "COLUMN()");
		outputPageSheet.fillFormula(
			`I${FIRST_ROW}:I${row}`,
			`SUMPRODUCT((MOD(COLUMN(O${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(O${FIRST_ROW}),6)=0)*O${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`J${FIRST_ROW}:J${row}`,
			`SUMPRODUCT((MOD(COLUMN(P${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(P${FIRST_ROW}),6)=0)*P${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`K${FIRST_ROW}:K${row}`,
			`SUMPRODUCT((MOD(COLUMN(Q${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(Q${FIRST_ROW}),6)=0)*Q${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`L${FIRST_ROW}:L${row}`,
			`SUMPRODUCT((MOD(COLUMN(R${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(R${FIRST_ROW}),6)=0)*R${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`M${FIRST_ROW}:M${row}`,
			`SUMPRODUCT((MOD(COLUMN(S${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(S${FIRST_ROW}),6)=0)*S${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`N${FIRST_ROW}:N${row}`,
			`SUMPRODUCT((MOD(COLUMN(T${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(T${FIRST_ROW}),6)=0)*T${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);

		outputPageSheet.eachRow((row, rowNumber) => {
			if (rowNumber >= FIRST_ROW) {
				row.height = HEIGHT;
				row.eachCell({ includeEmpty: true }, (cell) => {
					copyCellStyle(DATA_ROW.getCell(cell.col), cell);
				});
				row.commit();
			}
		});

		outputPageSheet.addConditionalFormatting({
			ref: `A${FIRST_ROW}:${lastCol}${FIRST_ROW}`,
			rules: [
				{
					type: "expression",
					formulae: ["COLUMN()"],
					style: { font: { color: { argb: "FFFF0000" } } },
					priority: 0,
				},
			],
		});

		outputPageSheet.addConditionalFormatting({
			ref: `I10:I${row}`,
			rules: [
				{
					type: "expression",
					formulae: ["OR(ROW()=10,ROW()>=13)"],
					style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFE2EFDA" } } },
					priority: 1,
				},
			],
		});
	}
}
