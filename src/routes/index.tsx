import { createFileRoute } from "@tanstack/react-router";
import Excel from "exceljs";
import { saveAs } from "file-saver";
import { CheckCircle, Download, FileUp, Loader2, TriangleAlert } from "lucide-react";
import { useState } from "react";
import { Badge } from "@/components/badge";
import { Button } from "@/components/button";
import { Input } from "@/components/input";
import { Label } from "@/components/label";
import { copyWorksheet } from "@/lib/copy";
import { generateOutputFile } from "@/lib/generate";
import { processData } from "@/lib/processData";
import { cn, formatBytes } from "@/lib/utils";
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
	const [error, setError] = useState<string>();
	const fileSelected = !!data || !!buffer;

	return (
		<div className="min-h-screen bg-background flex items-center justify-center p-6">
			<div className="max-w-2xl w-full">
				<div className="flex flex-col gap-4">
					<div className="mb-8 flex items-start justify-between">
						<div className="flex flex-col gap-2 w-full items-start">
							<h1 className="text-3xl font-bold text-foreground mb-2 uppercase">Додаток 47 – на радість усім</h1>
							<ul className="list-disc list-inside">
								<li>Дані не завантажуються на сервер.</li>
								<li>Дані не зберігаються на сервері.</li>
								<li>Всі дані обробляються локально на Вашому пристрої.</li>
								<li>Файл для скачування генерується локально на Вашому пристрої.</li>
							</ul>
							<Button
								variant="link"
								className="text-blue-400/50 text-sm p-0"
								onClick={async () => {
									const newWorkbook = new Excel.Workbook();
									copyWorksheet({
										template: workbook.getWorksheet("INPUT")!,
										workbook: newWorkbook,
										newSheetName: "INPUT",
									});
									const buffer = await newWorkbook.xlsx.writeBuffer();
									saveAs(new Blob([buffer]), "Вхідні дані для заповнення Додатку 47.xlsx");
								}}
							>
								<Download className="size-4" />
								<span>Завантажити шаблон для заповнення даних</span>
							</Button>
						</div>
					</div>

					<div className="flex items-center gap-4 max-h-[46px] h-[46px] min-h-[46px]">
						<div className="flex gap-2 h-[46px]">
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

									try {
										const arrayBuffer = await file.arrayBuffer();
										const inputDataWorkbook = new Excel.Workbook();
										await inputDataWorkbook.xlsx.load(arrayBuffer);
										const inputData = inputDataWorkbook.worksheets[0];
										const data = processData(inputData);
										setData(data);
										setFileName(file.name);
										setError(undefined);
									} catch (_e) {
										setError("Некоректна структура файлу");
										setFileName(file.name);
										setData(null);
									} finally {
										setBuffer(undefined);
										setLoading(false);
									}
								}}
							/>

							{error && (
								<div className="flex items-center gap-2 bg-destructive/10 p-3 rounded-md">
									<TriangleAlert className="size-5 text-destructive" />
									<span className="text-sm text-destructive">{error}</span>
								</div>
							)}
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
								className="h-[46px] rounded-sm px-4 text-sm max-h-[46px]"
								onClick={() => saveAs(new Blob([buffer]), "Додаток 47.xlsx")}
							>
								<Download className="size-5" />
								<div className="flex flex-col justify-start items-start">
									<span>Завантажити</span>
									<span className="text-xs">{formatBytes(buffer.byteLength)}</span>
								</div>
							</Button>
						) : null}

						{loading && (
							<div className="flex items-center gap-2">
								<Loader2 className="size-7 text-foreground animate-spin" />
								<span className="text-sm text-muted-foreground">Файли генеруються...</span>
							</div>
						)}
					</div>

					{buffer && (
						<div className="flex items-center gap-2 bg-green-500/10 p-3 rounded-md h-[46px] text-green-500/50">
							<CheckCircle className="size-5" />
							<span className="text-sm">Файл успішно згенеровано!</span>
						</div>
					)}

					<div
						className={cn(
							"text-md text-muted-foreground uppercase tracking-wide mt-2 opacity-0 transition-opacity duration-200",
							data && "opacity-100",
						)}
					>
						Буде згенеровано{" "}
						<Badge variant="secondary" className="text-md p-3">
							{data?.listArray.length}
						</Badge>{" "}
						{pluralize(data?.listArray.length ?? 0, "сторінка", "сторінки", "сторінок")} з даними для{" "}
						<Badge variant="secondary" className="text-md p-3">
							{data?.entitiesArray.length ?? 0}
						</Badge>{" "}
						{pluralize(data?.entitiesArray.length ?? 0, "підрозділу", "підрозділів", "підрозділів")}
					</div>
				</div>
			</div>
		</div>
	);
}

async function loadInputFile(file: File) {
	const arrayBuffer = await file.arrayBuffer();
	const inputDataWorkbook = new Excel.Workbook();
	await inputDataWorkbook.xlsx.load(arrayBuffer);
	return inputDataWorkbook.worksheets[0];
}
