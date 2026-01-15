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
