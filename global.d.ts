import type Excel from "exceljs";
import type { CellBase } from "react-spreadsheet";

declare global {
  interface Cell extends CellBase {
    cell: Excel.Cell;
    isHeader?: boolean;
    shouldHide?: boolean;
    rowSpan?: number;
    colSpan?: number;
  }

}

export {};