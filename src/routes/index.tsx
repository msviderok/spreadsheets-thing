import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import Spreadsheet, {
	type CellBase,
	type DataViewerComponent,
	type Matrix,
} from "react-spreadsheet";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";
import mockDb from "../mock.xlsx?url";
import workbookUrl from "../output_template.xlsx?url";

interface Cell extends CellBase {
	cell: Excel.Cell;
}

const ValueTypeCell: DataViewerComponent<Cell> = ({ cell }) => {
	const data = cell?.cell;
	const value =
		data?.type === Excel.ValueType.Date
			? Intl.DateTimeFormat("uk-UA", {
					day: "2-digit",
					month: "2-digit",
					year: "numeric",
				}).format(data.value as Date)
			: (data?.value as string);

	return (
		<span
			className={cn("w-full text-xs flex px-2", {
				"justify-end": data?.style.alignment?.horizontal === "right",
				"justify-start": data?.style.alignment?.horizontal === "left",
				"justify-center": data?.style.alignment?.horizontal === "center",
				"items-center": data?.style.alignment?.vertical === "middle",
				"items-start": data?.style.alignment?.vertical === "top",
				"items-end": data?.style.alignment?.vertical === "bottom",
			})}
		>
			{value}
		</span>
	);
};

export const Route = createFileRoute("/")({
	component: App,
	ssr: false,
	loader: async () => {
		const response = await fetch(workbookUrl);
		const arrayBuffer = await response.arrayBuffer();
		const workbook = new Excel.Workbook();
		await workbook.xlsx.load(arrayBuffer);
		const CONTENTS_SHEET = workbook.worksheets[0];
		const EXTRA_SHEET = workbook.worksheets[1];
		const OUTPUT_PAGE_SHEET = workbook.worksheets[2];

		const mockResponse = await fetch(mockDb);
		const mockArrayBuffer = await mockResponse.arrayBuffer();
		const mockWorkbook = new Excel.Workbook();
		await mockWorkbook.xlsx.load(mockArrayBuffer);
		const sheet = mockWorkbook.worksheets[0];

		const matrix = Array.from({ length: sheet.actualRowCount }, () =>
			Array.from(
				{ length: sheet.actualColumnCount },
				(): Cell => ({
					value: "",
					readOnly: true,
					cell: null as unknown as Excel.Cell,
					DataViewer: ValueTypeCell as any,
				}),
			),
		);

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
				}
			});
		});

		return {
			inputData: {
				matrix,
			},
		};
	},
});

function App() {
	const data = Route.useLoaderData();

	return (
		<div className="min-h-screen flex flex-col gap-4 p-4 text-xs">
			<Spreadsheet data={data.inputData.matrix} />
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
