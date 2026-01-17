import Excel from "exceljs";
import { useEffect, useState } from "react";
import type { CellBase, DataViewerComponent } from "react-spreadsheet";
import Spreadsheet from "react-spreadsheet";
import { formatDate } from "@/lib/utils";

const ValueTypeCell: DataViewerComponent<CellBase> = ({ cell }) => {
	return <span className="w-full text-xs flex px-3">{cell?.value}</span>;
};

async function getDb(sheet: Excel.Worksheet) {
	const matrix: CellBase[][] = Array.from({ length: sheet.actualRowCount }, () =>
		Array.from({ length: sheet.actualColumnCount }).map(() => ({
			value: "",
			readOnly: true,
			DataViewer: ValueTypeCell as any,
		})),
	);

	sheet.eachRow((row) => {
		const rowIndex = +row.number - 1;
		row.eachCell((cell) => {
			const colIndex = +cell.col - 1;
			matrix[rowIndex][colIndex].value = cell.value;

			if (cell.type === Excel.ValueType.Date) {
				cell.style.numFmt = "dd.mm.yyyy";
				matrix[rowIndex][colIndex].value = formatDate(cell.value as any);
			}
		});
	});

	return matrix;
}

export default function DataPreview(props: { sheet: Excel.Worksheet }) {
	const [data, setData] = useState<CellBase[][]>([]);
	useEffect(() => void getDb(props.sheet).then(setData), [props.sheet]);
	return (
		<div className="flex flex-col gap-8">
			<Spreadsheet data={data} />
		</div>
	);
}
