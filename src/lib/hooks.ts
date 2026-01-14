import type Excel from "exceljs";
import { useEffect } from "react";

export function useWorkbookStyles(
	style: Partial<Excel.Style>,
	ref: React.RefObject<HTMLDivElement>,
) {
	useEffect(() => {
		if (!ref.current) return;

		const parentElement = ref.current.parentElement;
		if (!parentElement) return;

		const borderStyles = getBorderStyle(style.border);

		// Apply border styles to parent td/th
		if (borderStyles.borderTop) {
			parentElement.style.borderTop = borderStyles.borderTop;
		}
		if (borderStyles.borderRight) {
			parentElement.style.borderRight = borderStyles.borderRight;
		}
		if (borderStyles.borderBottom) {
			parentElement.style.borderBottom = borderStyles.borderBottom;
		}
		if (borderStyles.borderLeft) {
			parentElement.style.borderLeft = borderStyles.borderLeft;
		}
	}, [style?.border, ref]);
}

function getBorderStyle(borders?: Partial<Excel.Borders>) {
	if (!borders) {
		return {
			borderTop: undefined,
			borderRight: undefined,
			borderBottom: undefined,
			borderLeft: undefined,
		};
	}

	const getBorderWidth = (style?: Excel.BorderStyle): number => {
		switch (style) {
			case "thin":
			case "hair":
			case "dotted":
			case "dashed":
			case "dashDot":
			case "dashDotDot":
			case "slantDashDot":
				return 1;
			case "medium":
			case "mediumDashed":
			case "mediumDashDot":
			case "mediumDashDotDot":
				return 2;
			case "thick":
			case "double":
				return 3;
			default:
				return 0;
		}
	};

	const getBorderStyleType = (style?: Excel.BorderStyle): string => {
		switch (style) {
			case "dotted":
				return "dotted";
			case "dashed":
			case "mediumDashed":
				return "dashed";
			case "dashDot":
			case "mediumDashDot":
				return "dashed";
			case "dashDotDot":
			case "mediumDashDotDot":
			case "slantDashDot":
				return "dashed";
			case "double":
				return "double";
			case "thin":
			case "hair":
			case "medium":
			case "thick":
			default:
				return "solid";
		}
	};

	const getSideBorder = (side?: Partial<Excel.Border>) => {
		if (!side?.style) return undefined;
		const width = getBorderWidth(side.style);
		if (width === 0) return undefined;
		const styleType = getBorderStyleType(side.style);
		const color = side.color?.argb
			? `#${side.color.argb.slice(2)}` // Convert ARGB to hex (skip alpha)
			: "inherit";
		return `${width}px ${styleType} ${color}`;
	};

	return {
		borderTop: getSideBorder(borders.top),
		borderRight: getSideBorder(borders.right),
		borderBottom: getSideBorder(borders.bottom),
		borderLeft: getSideBorder(borders.left),
	};
}
