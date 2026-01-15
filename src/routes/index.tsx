import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import fileSaver from "file-saver";
import Spreadsheet from "react-spreadsheet";
import { Button } from "@/components/ui/button";
import { copyWorksheet } from "@/lib/copyWorksheet";
import { loader } from "@/lib/loader";
import { sanitizeStr } from "@/lib/sanitizeStr";

export const Route = createFileRoute("/")({
	ssr: false,
	component: App,
	loader,
});

function App() {
	const data = Route.useLoaderData();
	const distinctNames = data.db.reduce((names: Set<string>, row: Cell[]) => {
		const secondColumnCell = row[1];
		if (secondColumnCell?.value && secondColumnCell.isHeader === false) {
			const name = String(secondColumnCell.value).trim();
			if (name) {
				names.add(name);
			}
		}
		return names;
	}, new Set<string>());

	return (
		<div className="min-h-screen flex flex-col gap-4 p-4 text-xs">
			<div className="flex items-center justify-between gap-4">
				<h1 className="text-lg font-semibold">Spreadsheet Viewer</h1>
				<Button
					onClick={() => generateOutputFile(distinctNames, data.workbook)}
					variant="default"
				>
					Generate Output ({distinctNames.size} pages)
				</Button>
			</div>
			<Spreadsheet data={data.db} />
			{/* <Spreadsheet data={data.contents} /> */}
			{/* <Spreadsheet data={data.extra} /> */}
			{/* <Spreadsheet data={data.outputPage} /> */}
		</div>
	);
}

async function generateOutputFile(
	distinctNames: Set<string>,
	templateWorkbook: Excel.Workbook,
): Promise<Buffer | ArrayBuffer> {
	const newWorkbook = new Excel.Workbook();
	const outputPageTemplate = templateWorkbook.getWorksheet("OUTPUT_PAGE")!;
	const contentsTemplate = templateWorkbook.getWorksheet("CONTENTS")!;
	const extraTemplate = templateWorkbook.getWorksheet("EXTRA")!;

	for (const name of distinctNames) {
		const outputPageSheet = copyWorksheet({
			template: outputPageTemplate,
			workbook: newWorkbook,
			newSheetName: sanitizeStr(name),
		});

		outputPageSheet.getCell("B3").value = {
			text: "←Повернутись на зміст",
			hyperlink: `#'Зміст'!A1`,
		};
		outputPageSheet.getCell("F5").value = name;
	}

	const contentsSheet = copyWorksheet({
		template: contentsTemplate,
		workbook: newWorkbook,
		newSheetName: "Зміст",
	});

	// Save the border styles from row 4 (template row)
	const templateRow = contentsSheet.getRow(4);
	const templateBorders: Record<number, Partial<Excel.Borders>> = {};
	templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
		if (cell.border) {
			templateBorders[colNumber] = JSON.parse(JSON.stringify(cell.border));
		}
	});

	// Expand the section by duplicating row 4 for each additional entry
	const distinctNamesArray = Array.from(distinctNames);
	if (distinctNamesArray.length > 1) {
		// Duplicate row 4 for each additional entry (we already have row 4, so duplicate for the rest)
		for (let i = 1; i < distinctNamesArray.length; i++) {
			contentsSheet.duplicateRow(4, 1, true);
		}
	}

	// Update all rows with values and remove borders from non-last rows
	let rowIndex = 4;
	let sheetIndex = 1;
	const lastRowIndex = 4 + distinctNamesArray.length - 1;

	for (const name of distinctNamesArray) {
		const sanitizedName = sanitizeStr(name);
		const sheetName = sanitizedName;

		// Column A: name of the output_page
		contentsSheet.getCell(rowIndex, 1).value = name;

		// Column B: hyperlink to the output_page sheet
		contentsSheet.getCell(rowIndex, 2).value = {
			text: `Аркуш ${sheetIndex}`,
			hyperlink: `#'${sheetName}'!A1`,
		};

		// Remove borders from all rows except the last one
		if (rowIndex !== lastRowIndex) {
			const row = contentsSheet.getRow(rowIndex);
			row.eachCell({ includeEmpty: true }, (cell) => {
				cell.border = {};
			});
		}

		rowIndex++;
		sheetIndex++;
	}

	// Apply borders to the last row
	const lastRow = contentsSheet.getRow(lastRowIndex);
	Object.entries(templateBorders).forEach(([colNumber, border]) => {
		const cell = lastRow.getCell(Number(colNumber));
		cell.border = border as Partial<Excel.Borders>;
	});

	// console.log(newWorkbook.worksheets[0].getSheetValues());
	const buffer = await newWorkbook.xlsx.writeBuffer();
	fileSaver.saveAs(new Blob([buffer]), "output.xlsx");
	return buffer;
}
