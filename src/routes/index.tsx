import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import fileSaver from "file-saver";
import Spreadsheet from "react-spreadsheet";
import { Button } from "@/components/ui/button";
import { buildContentsSheet } from "@/lib/buildContentsSheet";
import { buildOutputPages } from "@/lib/buildOutputPages";
import { loader } from "@/lib/loader";

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
		</div>
	);
}

async function generateOutputFile(
	distinctNames: Set<string>,
	templateWorkbook: Excel.Workbook,
): Promise<Buffer | ArrayBuffer> {
	const newWorkbook = new Excel.Workbook();

	buildOutputPages({
		templateWorkbook,
		workbook: newWorkbook,
		distinctNames,
	});

	buildContentsSheet({
		templateWorkbook,
		workbook: newWorkbook,
		distinctNames,
	});

	const buffer = await newWorkbook.xlsx.writeBuffer();
	fileSaver.saveAs(new Blob([buffer]), "output.xlsx");
	return buffer;
}
