import type Excel from "exceljs";
import { copyCellProperties, copyCellRange, copyCellStyle, copyWorksheet, numberToCol } from "./copy";
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

		const colCount = outputPageSheet.actualColumnCount;

		const leftoversRow = outputPageSheet.getRow(13);
		const leftovers = data.list[name]!.leftovers;

		leftoversRow.getCell("A").value = leftovers.date;
		leftoversRow.getCell("B").value = "Перенесено з книги №10";

		data.entitiesArray.forEach((entity) => {
			const entityColumns = addresses[entity];
			if (!entityColumns || entityColumns.length !== 6) return;

			const [totalCol, I, II, III, IV, V] = entityColumns;
			leftoversRow.getCell(totalCol).value = { formula: `=SUM(${I}13:${V}13)` };
			leftoversRow.getCell(I).value = leftovers.entities[entity]?.categories[0];
			leftoversRow.getCell(II).value = leftovers.entities[entity]?.categories[1];
			leftoversRow.getCell(III).value = leftovers.entities[entity]?.categories[2];
			leftoversRow.getCell(IV).value = leftovers.entities[entity]?.categories[3];
			leftoversRow.getCell(V).value = leftovers.entities[entity]?.categories[4];
		});

		const dataCell = outputPageSheet.getCell(14, 1);
		const dataRowHeight = outputPageSheet.getRow(14).height;

		const statements = data.list[name]!.statements;
		let statementRowNum = 14;

		statements.forEach((statement) => {
			const statementRow = outputPageSheet.getRow(statementRowNum);
			statementRow.height = dataRowHeight;

			for (let col = 1; col <= colCount; col++) {
				const cell = statementRow.getCell(col);
				copyCellStyle(dataCell, cell);
			}

			statementRow.getCell("A").value = statement.date ?? "";
			statementRow.getCell("B").value = statement.docName ?? "";
			statementRow.getCell("C").value = statement.docNumber ?? "";
			statementRow.getCell("D").value = statement.docDate ?? "";
			statementRow.getCell("E").value = statement.from ?? "";
			statementRow.getCell("F").value = statement.to ?? "";
			statementRow.getCell("G").value = statement.quantityIn ?? "";
			statementRow.getCell("H").value = statement.quantityOut ?? "";
			statementRowNum++;
		});

		const lastCol = outputPageSheet.getColumn(colCount).letter;
		const lastRow = statementRowNum - 1;

		outputPageSheet.fillFormula(`A12:${lastCol}12`, "COLUMN()");
		outputPageSheet.fillFormula(
			`I${FIRST_ROW}:I${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(O${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(O${FIRST_ROW}),6)=0)*O${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`J${FIRST_ROW}:J${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(P${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(P${FIRST_ROW}),6)=0)*P${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`K${FIRST_ROW}:K${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(Q${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(Q${FIRST_ROW}),6)=0)*Q${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`L${FIRST_ROW}:L${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(R${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(R${FIRST_ROW}),6)=0)*R${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`M${FIRST_ROW}:M${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(S${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(S${FIRST_ROW}),6)=0)*S${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
		outputPageSheet.fillFormula(
			`N${FIRST_ROW}:N${lastRow}`,
			`SUMPRODUCT((MOD(COLUMN(T${FIRST_ROW}:${lastCol}${FIRST_ROW})-COLUMN(T${FIRST_ROW}),6)=0)*T${FIRST_ROW}:${lastCol}${FIRST_ROW})`,
		);
	}
}
