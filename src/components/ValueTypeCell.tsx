import Excel from "exceljs";
import { useEffect, useRef } from "react";
import type { DataViewerComponent } from "react-spreadsheet";
import { cn } from "@/lib/utils";

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
			return `{${(data.value as Excel.CellFormulaValue)?.formula ?? ""}}`;
		case Excel.ValueType.RichText:
			return (data.value as Excel.CellRichTextValue)?.richText
				.map((text) => text.text)
				.join("");

		default:
			return data?.value as string;
	}
}

export const ValueTypeCell: DataViewerComponent<Cell> = ({ cell }) => {
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
			className={cn("w-full text-xs flex px-3", getAlignmentStyle(style))}
		>
			{shouldHide || isHiddenMergedCell ? null : value}
		</span>
	);
};

function isMergedHiddenCell(cell: Excel.Cell) {
	if (!cell.isMerged) return false;
	const master = cell.master ?? cell;
	return cell.address !== master.address;
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
