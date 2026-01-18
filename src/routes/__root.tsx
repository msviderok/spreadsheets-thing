import { createRootRoute, HeadContent, Scripts } from "@tanstack/react-router";
import appCss from "../styles.css?url";

export const Route = createRootRoute({
	shellComponent: RootDocument,
	head: () => ({
		meta: [
			{ charSet: "utf-8" },
			{ name: "viewport", content: "width=device-width, initial-scale=1" },
			{ title: "TanStack Start Starter" },
		],
		links: [{ rel: "stylesheet", href: appCss }],
	}),
});

function RootDocument({ children }: { children: React.ReactNode }) {
	return (
		<html lang="en" className="dark min-h-screen min-w-screen w-full overflow-auto">
			<head>
				<HeadContent />
			</head>
			<body className="min-h-screen min-w-screen w-full overflow-auto">
				{children}
				<Scripts />
			</body>
		</html>
	);
}
