import type Excel from "exceljs";
import { CATEGORY_MAP, formatDate, sortByCaliber } from "@/lib/utils";

declare global {
	interface Statement {
		date: string;
		docName?: string;
		docNumber?: string;
		docDate?: string;
		from?: string;
		to?: string;
		quantityIn?: number;
		quantityOut?: number;
		takeValueFrom?: "in" | "out";
		entities: {
			[entity: string]: {
				operation: 1 | -1;
				categories: {
					[category: number]: number;
				};
			};
		};
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
		listArray: string[];
	}
}

export function processData(sheet: Excel.Worksheet | undefined) {
	if (!sheet) return null;

	const data: Data = {
		list: {},
		get listArray() {
			return Object.keys(this.list).sort(sortByCaliber);
		},

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
				data.list[name]!.leftovers.entities[to] = { operation: 1, categories: {} };
			}

			data.list[name]!.leftovers.entities[to].categories[category] = Number(quantity);
		} else {
			const knownFrom = data.entitiesArray.includes(from);
			const knownTo = data.entitiesArray.includes(to);

			const q = Number(quantity);
			let quantityIn: number | undefined;
			let quantityOut: number | undefined;
			let takeValueFrom: "in" | "out" | undefined;
			let entities: Statement["entities"] = {};

			if (knownFrom && knownTo) {
				quantityIn = q;
				quantityOut = undefined;
				takeValueFrom = "in";
				entities = {
					[from]: { operation: -1, categories: { [category]: q } },
					[to]: { operation: 1, categories: { [category]: q } },
				};
			} else if (knownFrom && (knownTo === false || to == null || to === "")) {
				quantityIn = undefined;
				quantityOut = q;
				takeValueFrom = "out";
				entities = {
					[from]: { operation: -1, categories: { [category]: q } },
				};
			} else if ((knownFrom === false || from == null || from === "") && knownTo) {
				quantityIn = q;
				quantityOut = undefined;
				takeValueFrom = "in";
				entities = {
					[to]: { operation: 1, categories: { [category]: q } },
				};
			}

			data.list[name]!.statements.push({
				date: formatDate(date),
				docName,
				docNumber,
				docDate: formatDate(docDate),
				from,
				to,
				quantityIn,
				quantityOut,
				takeValueFrom,
				entities,
			});
		}
	}

	return data;
}
