import Excel from "exceljs";
import { ValueTypeCell } from "@/components/ValueTypeCell";

export interface ExcelCellProperties {
	cell: Excel.Cell;
	isHeader?: boolean;
	shouldHide?: boolean;
	rowSpan?: number;
	colSpan?: number;
}

type TemplateKey = "contents" | "extra" | "output_page" | "db";

const HEADER_RANGES: Record<TemplateKey, string[]> = {
	db: ["A1:J1"],
	contents: ["A1:C3"],
	extra: ["A1:F12", "A18:F22"],
	output_page: ["A1:N12"],
};

export function initMatrix(sheet: Excel.Worksheet, templateKey?: TemplateKey) {
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
	matrix: ExcelCellProperties[][],
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

function applyHeaderMap(
	matrix: ExcelCellProperties[][],
	headerCells: Set<string>,
) {
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
