import { type ClassValue, clsx } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
	return twMerge(clsx(inputs));
}

export function sanitizeStr(name: string): string {
	let sanitized = name.replace(/[:\\/?*[\]]/g, "");

	sanitized = sanitized.replace(/^['\s]+|['\s]+$/g, "");

	if (sanitized.length > 31) {
		sanitized = sanitized.substring(0, 31);
	}

	if (!sanitized || sanitized.trim().length === 0) {
		sanitized = "Сторінка";
	}

	return sanitized;
}

export function formatDate(date: string): string {
	return Intl.DateTimeFormat("uk-UA", {
		day: "2-digit",
		month: "2-digit",
		year: "numeric",
	}).format(date as any);
}

export const CATEGORY_MAP: Record<any, number> = {
	I: 0,
	І: 0,
	1: 0,
	II: 1,
	ІІ: 1,
	2: 1,
	III: 2,
	ІІІ: 2,
	3: 2,
	IV: 3,
	ІV: 3,
	4: 3,
	V: 4,
	5: 4,
};

function extractCaliber(str: string) {
	const match = str.match(/(\d+(?:[.,]\d+)?)(?=\s*мм)?/i);
	if (!match) return null;

	return parseFloat(match[1].replace(",", "."));
}

export function sortByCaliber(a: string, b: string) {
	const calA = extractCaliber(a);
	const calB = extractCaliber(b);

	// If both have calibers → numeric compare
	if (calA !== null && calB !== null) {
		if (calA !== calB) return calA - calB;
	}

	// If only one has caliber → push missing ones to the end
	if (calA === null) return 1;
	if (calB === null) return -1;

	// Same caliber → locale-aware text sort (Ukrainian)
	return a.localeCompare(b, "uk", { sensitivity: "base" });
}
