import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import { saveAs } from "file-saver";
import { FileUp, Loader2 } from "lucide-react";
import { useEffect, useState } from "react";
import { Badge } from "@/components/badge";
import { Button } from "@/components/button";
import { Collapsible, CollapsibleContent } from "@/components/collapsible";
import { Input } from "@/components/input";
import { Label } from "@/components/label";
import { generateOutputFile } from "@/lib/generate";
import { processData } from "@/lib/processData";
import mockDb from "../mock.xlsx?url";
import workbookUrl from "../output_template.xlsx?url";

const pluralize = (count: number, one: string, few: string, many: string) => {
	if (count === 1) return one;
	if (count % 10 === 1 && count % 100 !== 11) return one;
	if (count % 10 === 2 && count % 100 !== 12) return few;
	if (count % 10 === 3 && count % 100 !== 13) return few;
	if (count % 10 === 4 && count % 100 !== 14) return few;
	return many;
};

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
	const [data, setData] = useState<Data | null>(null);
	const [buffer, setBuffer] = useState<ArrayBuffer>();
	const [fileName, setFileName] = useState<string>();
	const [loading, setLoading] = useState(false);

	useEffect(() => {
		const handleBeforeUnload = (e: BeforeUnloadEvent) => {
			e.preventDefault();
			e.returnValue = "";
		};

		window.addEventListener("beforeunload", handleBeforeUnload);
		return () => window.removeEventListener("beforeunload", handleBeforeUnload);
	}, []);

	return (
		<div className="min-h-screen bg-background flex items-center justify-center p-6">
			<div className="max-w-2xl w-full">
				<div className="flex flex-col gap-4">
					<div className="mb-8 flex items-start justify-between">
						<div className="flex flex-col gap-2 w-full">
							<h1 className="text-3xl font-bold text-foreground mb-2 uppercase">Додаток 47 – на радість усім</h1>
							<ul className="list-disc list-inside">
								<li>Дані не завантажуються на сервер.</li>
								<li>Дані не зберігаються на сервері.</li>
								<li>Всі дані обробляються локально на Вашому пристрої.</li>
								<li>Файл для скачування генерується локально на Вашому пристрої.</li>
							</ul>
						</div>
					</div>

					<div className="flex items-center gap-4">
						<div>
							<Label htmlFor="file" className="contents">
								<div className="flex items-center justify-center gap-2 border w-max p-3 rounded-sm bg-card cursor-pointer hover:bg-secondary transition-colors">
									<FileUp className="size-5 text-foreground" />
									<span className="text-sm text-muted-foreground">
										{fileName ?? (
											<>
												Оберіть файл <strong>.xlsx</strong>
											</>
										)}
									</span>
								</div>
							</Label>
							<Input
								hidden
								id="file"
								type="file"
								accept=".xlsx"
								className="hidden"
								onChange={async (e) => {
									const file = e.target.files?.[0];
									if (!file) return;

									const inputData = await loadInputFile(file);
									const data = processData(inputData);
									setData(data);
									setFileName(file.name);
								}}
							/>
						</div>

						{data ? (
							<Button
								size="lg"
								className="h-[46px] rounded-sm px-4 text-sm"
								onClick={() => {
									if (!data) return;
									setLoading(true);
									generateOutputFile(data, workbook).then((buffer) => {
										setBuffer(buffer);
										setLoading(false);
										setData(null);
									});
								}}
							>
								Згенерувати файли
							</Button>
						) : buffer ? (
							<Button
								size="lg"
								className="h-[46px] rounded-sm px-4 text-sm bg-green-500 text-white hover:bg-green-600"
								onClick={() => {
									saveAs(new Blob([buffer]), "output.xlsx");
								}}
							>
								Скачати файли
							</Button>
						) : null}

						{loading && (
							<div className="flex items-center gap-2">
								<Loader2 className="size-7 text-foreground animate-spin" />
								<span className="text-sm text-muted-foreground">Файли генеруються...</span>
							</div>
						)}
					</div>

					{data && (
						<div className="text-md text-muted-foreground uppercase tracking-wide mt-2">
							Буде згенеровано{" "}
							<Badge variant="secondary" className="text-md p-3">
								{data!.listArray.length}
							</Badge>{" "}
							{pluralize(data!.listArray.length, "сторінка", "сторінки", "сторінок")} з даними для{" "}
							<Badge variant="secondary" className="text-md p-3">
								{data!.entitiesArray.length}
							</Badge>{" "}
							{pluralize(data!.entitiesArray.length, "підрозділу", "підрозділів", "підрозділів")}
						</div>
					)}
				</div>
			</div>
		</div>
	);
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
