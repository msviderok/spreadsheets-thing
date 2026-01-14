import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import { useEffect, useRef } from "react";
import Spreadsheet, {
	type CellBase,
	type DataViewerComponent,
	type Matrix,
} from "react-spreadsheet";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn, indexedColors } from "@/lib/utils";
import mockDb from "../mock.xlsx?url";
import workbookUrl from "../output_template.xlsx?url";

interface Cell extends CellBase {
	cell: Excel.Cell;
	isHeader?: boolean;
	shouldHide?: boolean;
	rowSpan?: number;
	colSpan?: number;
}

function getValue(data: Excel.Cell | undefined) {
	if (!data) return "";
	switch (data?.type) {
		case Excel.ValueType.Date:
			return Intl.DateTimeFormat("uk-UA", {
				day: "2-digit",
				month: "2-digit",
				year: "numeric",
			}).format(data.value as Date);
		case Excel.ValueType.Formula:
			if (data.value.formula === "CH1") {
				return data.value;
			}
			return `{${(data.value as Excel.CellFormulaValue)?.formula ?? ""}}`;
		case Excel.ValueType.RichText:
			return (data.value as Excel.CellRichTextValue)?.richText
				.map((text) => text.text)
				.join("");

		default:
			return data?.value as string;
	}
}

const ValueTypeCell: DataViewerComponent<Cell> = ({ cell }) => {
	const data = cell?.cell;
	const style = data?.style;
	const ref = useRef<HTMLSpanElement>(null);
	const isHiddenMergedCell = data ? isMergedHiddenCell(data) : false;
	const shouldHide = Boolean(cell?.shouldHide);

	const value = getValue(data);

	useEffect(() => {
		if (!cell?.isHeader) return;

		const parent = ref.current?.parentElement;
		if (!parent) return;

		if (cell?.shouldHide) {
			parent.style.display = "none";
		} else {
			if (cell?.rowSpan) {
				parent.setAttribute("rowspan", String(cell.rowSpan));
			}

			if (cell?.colSpan) {
				parent.setAttribute("colspan", String(cell.colSpan));
			}
		}
	}, [cell]);

	if (!data) return null;

	return (
		<span
			ref={ref}
			className={cn("w-full text-xs flex px-3", getAlignmentStyle(style), {
				"bg-red-500": shouldHide || isHiddenMergedCell,
			})}
		>
			{shouldHide || isHiddenMergedCell ? null : value}
		</span>
	);
};

export const Route = createFileRoute("/")({
	ssr: false,
	component: App,
	loader: async () => {
		const db = await getDb();
		const templates = await getTemplates();
		return { db, ...templates };
	},
});

function App() {
	const data = Route.useLoaderData();

	return (
		<div className="min-h-screen flex flex-col gap-4 p-4 text-xs">
			<Spreadsheet data={data.db} />
			<Spreadsheet data={data.contents} />
			<Spreadsheet data={data.extra} />
			<Spreadsheet data={data.outputPage} />

			{/* <Tabs>
				<TabsList>
					{matrix.map((sheet) => (
						<TabsTrigger key={sheet.sheetName} value={sheet.sheetName}>
							{sheet.sheetName}
						</TabsTrigger>
					))}
				</TabsList>
				{matrix.map((sheet) => (
					<TabsContent key={sheet.sheetName} value={sheet.sheetName}>
						<Spreadsheet data={sheet.rows} />
					</TabsContent>
				))}
			</Tabs> */}
		</div>
	);
}

function getAlignmentStyle(style: Partial<Excel.Style> | undefined) {
	return {
		"justify-end": style?.alignment?.horizontal === "right",
		"justify-start": style?.alignment?.horizontal === "left",
		"justify-center": style?.alignment?.horizontal === "center",
		"items-center": style?.alignment?.vertical === "middle",
		"items-start": style?.alignment?.vertical === "top",
		"items-end": style?.alignment?.vertical === "bottom",
	};
}

function isMergedHiddenCell(cell: Excel.Cell) {
	if (!cell.isMerged) return false;
	const master = cell.master ?? cell;
	return cell.address !== master.address;
}

async function getDb() {
	const mockResponse = await fetch(mockDb);
	const mockArrayBuffer = await mockResponse.arrayBuffer();
	const mockWorkbook = new Excel.Workbook();
	await mockWorkbook.xlsx.load(mockArrayBuffer);
	return initMatrix(mockWorkbook.worksheets[0]);
}

async function getTemplates() {
	const response = await fetch(workbookUrl);
	const arrayBuffer = await response.arrayBuffer();
	const workbook = new Excel.Workbook();
	await workbook.xlsx.load(arrayBuffer);
	return {
		contents: initMatrix(workbook.worksheets[0], "contents"),
		extra: initMatrix(workbook.worksheets[1], "extra"),
		outputPage: initMatrix(workbook.worksheets[2], "output_page"),
	};
}

type TemplateKey = "contents" | "extra" | "output_page";

const HEADER_RANGES: Record<TemplateKey, string[]> = {
	contents: ["A1:C3"],
	extra: ["A1:F12", "A18:F22"],
	output_page: ["A1:N12"],
};

function initMatrix(sheet: Excel.Worksheet, templateKey?: TemplateKey) {
	const matrix = Array.from({ length: sheet.actualRowCount }, () =>
		Array.from(
			{ length: sheet.actualColumnCount },
			(): Cell => ({
				value: "",
				readOnly: true,
				cell: null as unknown as Excel.Cell,
				DataViewer: ValueTypeCell as any,
				isHeader: false,
				shouldHide: false,
			}),
		),
	);

	const headerCells = templateKey
		? buildHeaderCellSet(HEADER_RANGES[templateKey])
		: new Set<string>();

	applyHeaderMap(matrix, headerCells);

	sheet.eachRow((row) => {
		row.eachCell((cell) => {
			const col = +cell.col - 1;
			const row = +cell.row - 1;

			if (cell.type === Excel.ValueType.Date) {
				cell.style.numFmt = "dd.mm.yyyy";
			}
			if (matrix?.[row]?.[col]) {
				matrix[row][col].cell = cell;
				matrix[row][col].value = cell.value;
				if (!matrix[row][col].isHeader) {
					matrix[row][col].isHeader = headerCells.has(cell.address);
				}
			}
		});
	});

	applyMergedHiddenMap(sheet, matrix, headerCells);

	return matrix;
}

function applyMergedHiddenMap(
	sheet: Excel.Worksheet,
	matrix: Cell[][],
	headerCells: Set<string>,
) {
	const processed = new Set<string>();

	const isSameMerged = (masterAddress: string, row: number, col: number) => {
		if (
			row < 1 ||
			col < 1 ||
			row > sheet.actualRowCount ||
			col > sheet.actualColumnCount
		) {
			return false;
		}
		const neighbor = sheet.getCell(row, col);
		if (!neighbor.isMerged) return false;
		const neighborMaster = neighbor.master ?? neighbor;
		return neighborMaster.address === masterAddress;
	};

	const getBounds = (masterAddress: string, row: number, col: number) => {
		let top = row;
		let bottom = row;
		let left = col;
		let right = col;

		while (isSameMerged(masterAddress, top - 1, col)) top -= 1;
		while (isSameMerged(masterAddress, bottom + 1, col)) bottom += 1;
		while (isSameMerged(masterAddress, row, left - 1)) left -= 1;
		while (isSameMerged(masterAddress, row, right + 1)) right += 1;

		return { top, bottom, left, right };
	};

	sheet.eachRow((row) => {
		row.eachCell((cell) => {
			if (!cell.isMerged) return;
			if (!headerCells.has(cell.address)) return;
			const master = cell.master ?? cell;
			const masterAddress = master.address;
			if (processed.has(masterAddress)) return;

			const rowNumber = Number(cell.row);
			const colNumber = Number(cell.col);
			const isStart =
				!isSameMerged(masterAddress, rowNumber - 1, colNumber) &&
				!isSameMerged(masterAddress, rowNumber, colNumber - 1);
			const isEnd =
				!isSameMerged(masterAddress, rowNumber + 1, colNumber) &&
				!isSameMerged(masterAddress, rowNumber, colNumber + 1);

			if (!isStart && !isEnd) return;

			const bounds = getBounds(masterAddress, rowNumber, colNumber);

			for (let r = bounds.top; r <= bounds.bottom; r += 1) {
				for (let c = bounds.left; c <= bounds.right; c += 1) {
					const isMaster = r === Number(master.row) && c === Number(master.col);
					const address = `${numberToColLetters(c)}${r}`;
					if (!headerCells.has(address)) continue;
					if (!isMaster && matrix?.[r - 1]?.[c - 1]) {
						matrix[r - 1][c - 1].shouldHide = true;
					}
				}
			}

			const masterRow = Number(master.row);
			const masterCol = Number(master.col);
			if (matrix?.[masterRow - 1]?.[masterCol - 1]) {
				matrix[masterRow - 1][masterCol - 1].rowSpan =
					bounds.bottom - bounds.top + 1;
				matrix[masterRow - 1][masterCol - 1].colSpan =
					bounds.right - bounds.left + 1;
			}

			processed.add(masterAddress);
		});
	});
}

function buildHeaderCellSet(ranges: string[]) {
	const headerCells = new Set<string>();
	ranges.forEach((range) => {
		expandRange(range).forEach((address) => {
			headerCells.add(address);
		});
	});
	return headerCells;
}

function expandRange(range: string) {
	const match = range.toUpperCase().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
	if (!match) return [];
	const [, startCol, startRow, endCol, endRow] = match;
	const startColNum = colLettersToNumber(startCol);
	const endColNum = colLettersToNumber(endCol);
	const startRowNum = Number(startRow);
	const endRowNum = Number(endRow);
	const addresses: string[] = [];

	for (let row = startRowNum; row <= endRowNum; row += 1) {
		for (let col = startColNum; col <= endColNum; col += 1) {
			addresses.push(`${numberToColLetters(col)}${row}`);
		}
	}

	return addresses;
}

function applyHeaderMap(matrix: Cell[][], headerCells: Set<string>) {
	headerCells.forEach((address) => {
		const parsed = addressToRowCol(address);
		if (!parsed) return;
		const rowIndex = parsed.row - 1;
		const colIndex = parsed.col - 1;
		if (!matrix?.[rowIndex]?.[colIndex]) return;
		matrix[rowIndex][colIndex].isHeader = true;
	});
}

function addressToRowCol(address: string) {
	const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/);
	if (!match) return null;
	const [, colLetters, rowNumber] = match;
	return {
		row: Number(rowNumber),
		col: colLettersToNumber(colLetters),
	};
}

function colLettersToNumber(letters: string) {
	let total = 0;
	for (const char of letters) {
		total = total * 26 + (char.charCodeAt(0) - 64);
	}
	return total;
}

function numberToColLetters(num: number) {
	let result = "";
	let n = num;
	while (n > 0) {
		const remainder = (n - 1) % 26;
		result = String.fromCharCode(65 + remainder) + result;
		n = Math.floor((n - 1) / 26);
	}
	return result;
}
