import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import fileSaver from "file-saver";
import { useEffect, useState } from "react";
import DataPreview from "@/components/DataPreview";
import { Input } from "@/components/input";
import { Label } from "@/components/label";
import { buildContentsSheet } from "@/lib/buildContentsSheet";
import { buildOutputPages } from "@/lib/buildOutputPages";
import { CATEGORY_MAP, formatDate } from "@/lib/utils";
import mockDb from "../mock.xlsx?url";
import workbookUrl from "../output_template.xlsx?url";

export const Route = createFileRoute("/")({
	ssr: false,
	component: App,
	async loader() {
		const response = await fetch(workbookUrl);
		const arrayBuffer = await response.arrayBuffer();
		const workbook = new Excel.Workbook();
		await workbook.xlsx.load(arrayBuffer);
		return workbook;
	},
});

function App() {
	const workbook = Route.useLoaderData();
	const [sheet, setSheet] = useState<Excel.Worksheet>();
	useEffect(() => void loadInputFile().then(setSheet), []);

	useEffect(() => {
		if (!sheet) return;

		const data = processData(sheet);
		if (data) {
			generateOutputFile(data, workbook);
		}
	}, [sheet]);

	return (
		<div className="min-h-screen flex flex-col gap-4 p-4 text-xs">
			<div className="grid w-full max-w-sm items-center gap-3">
				<Label htmlFor="picture">Upload Excel File</Label>
				<Input id="picture" type="file" />
			</div>

			{sheet && <DataPreview sheet={sheet} />}

			{/* 
			<Button onClick={() => generateOutputFile([], workbook)}>
				Generate Output
			</Button> */}
		</div>
	);
}

const INPUT_RANGE = ["A1:J1"];

declare global {
	interface Statement {
		date: string;
		entities: { [entity: string]: { [category: number]: number } };
	}
	interface Data {
		entities: Set<string>;
		entitiesArray: string[];
		list: {
			[name: string]: {
				leftovers: Statement;
				statements: Statement[];
			};
		};
	}
}

function processData(sheet: Excel.Worksheet | undefined) {
	if (!sheet) return null;

	const data: Data = {
		list: {},
		entities: new Set(),
		get entitiesArray() {
			return Array.from(this.entities)
				.filter((entity) => {
					if (entity === "-" || entity === undefined || entity === "") return false;

					const l = entity.toLocaleLowerCase();
					if ((l.startsWith("a") || l.startsWith("а")) && l !== "a1815" && l !== "а1815") return false;
					return true;
				})
				.sort((a, b) => a.localeCompare(b));
		},
	};
	for (const r of sheet.getSheetValues().slice(2)) {
		if (Array.isArray(r) === false) {
			continue;
		}

		const row = r.slice(1);
		const [date, name, docName, docNumber, docDate, from, to, price, romanCategory, quantity] = row as string[];
		data.entities.add(from);
		data.entities.add(to);
		const category = CATEGORY_MAP[romanCategory];

		if (!data.list[name]) {
			data.list[name] = { leftovers: { date: "", entities: {} }, statements: [] };
		}

		if (docName === "Перенесено з книги №10") {
			if (data.list[name]!.leftovers.date === "") {
				data.list[name]!.leftovers.date = formatDate(date);
			}

			if (!data.list[name]!.leftovers.entities[to]) {
				data.list[name]!.leftovers.entities[to] = {};
			}

			data.list[name]!.leftovers.entities[to][category] = Number(quantity);
		} else {
			// data.list[name]!.statements.push({
			// 	date: formatDate(date),
			// 	entities: { [from]: [Number(quantity)] },
			// });
		}
	}

	// console.log(data);
	return data;
}

async function generateOutputFile(data: Data, templateWorkbook: Excel.Workbook): Promise<Buffer | ArrayBuffer> {
	const newWorkbook = new Excel.Workbook();

	buildOutputPages({
		templateWorkbook,
		workbook: newWorkbook,
		data,
	});

	buildContentsSheet({
		templateWorkbook,
		workbook: newWorkbook,
		data,
	});

	const buffer = await newWorkbook.xlsx.writeBuffer();
	fileSaver.saveAs(new Blob([buffer]), "output.xlsx");
	return buffer;
}

async function loadInputFile(file?: File) {
	let arrayBuffer: ArrayBuffer | undefined;
	if (!file) {
		const mockResponse = await fetch(mockDb);
		const mockArrayBuffer = await mockResponse.arrayBuffer();
		arrayBuffer = mockArrayBuffer;
	} else {
		arrayBuffer = await file.arrayBuffer();
	}

	const inputDataWorkbook = new Excel.Workbook();
	await inputDataWorkbook.xlsx.load(arrayBuffer);
	return inputDataWorkbook.worksheets[0];
}
