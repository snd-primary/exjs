import { chromium } from "playwright";
import type { Page, Browser, BrowserContext } from "playwright";
import { readFile, writeFile } from "node:fs/promises";

import path from "path";
import { join } from "path";
import { dirname } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export const createBrowserSession = async () => {
	const browser = await chromium.launch({ headless: false, slowMo: 500 });
	const context = await browser.newContext();

	return { browser, context };
};

type DOMElement = Element & {
	id: string;
	className: string;
	textContent: string | null;
	getBoundingClientRect: () => DOMRect;
	attributes: NamedNodeMap;
};

interface BrowserSession {
	browser: Browser;
	context: BrowserContext;
}

const analyzeElements = async (page: Page) => {
	return await page.evaluate(() => {
		const getElementInfo = (el: DOMElement) => ({
			tagName: el.tagName.toLowerCase(),
			id: el.id || "",
			className: el.className || "",
			text: el.textContent?.trim() || "",
			attributes: Array.from(el.attributes).map((attr) => ({
				name: attr.name,
				value: attr.value,
			})),
			position: (() => {
				const rect = el.getBoundingClientRect();
				return {
					x: Math.round(rect.x),
					y: Math.round(rect.y),
					width: Math.round(rect.width),
					height: Math.round(rect.height),
				};
			})(),
		});

		return Array.from(document.querySelectorAll("*"))
			.filter((el) => !["HTML", "HEAD", "BODY"].includes(el.tagName))
			.map(getElementInfo);
	});
};

const captureScreenshot = async (page: Page, outputPath: string) => {
	await page.screenshot({ path: outputPath, fullPage: true });
	return outputPath;
};

const saveAnalysisResult = async (result: any, outputDir: string) => {
	const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
	const filePath = path.join(outputDir, `analysis-${timestamp}.json`);

	await writeFile(filePath, JSON.stringify(result, null, 2));
	return filePath;
};

export const analyzePage = async (
	session: BrowserSession,
	htmlPath: string,
	outputDir
) => {
	const page = await session.context.newPage();
	await page.goto(`http://localhost:5050/${htmlPath}`);

	const [elements, screenshotPath] = await Promise.all([
		analyzeElements(page),
		captureScreenshot(page, "screenshot.png"),
	]);

	const result = { elements, screenshotPath };
	const jsonPath = await saveAnalysisResult(result, outputDir);

	await page.close();
	return { elements, jsonPath };
};

export const cleanupSession = async (session: BrowserSession) => {
	await session.browser.close();
};

/* const analyze = async (htmlPath: string) => {
	const session = await createBrowserSession();
	try {
		return await analyzePage(session, htmlPath, );
	} finally {
		await cleanupSession(session);
	}
} */

const runAnalysis = async () => {
	// HTMLファイルのパス
	const targetPath = "index.html";
	const outputDir = join(__dirname, "output");
	let session;

	try {
		const session = await createBrowserSession();
		const result = await analyzePage(session, targetPath, outputDir);

		console.log("Analysis Results:", result);
		return result;
	} catch (error) {
		console.error("Analysis failed:", error);
		throw error;
	} finally {
		if (session) {
			await cleanupSession(session);
		}
	}
};

// スクリプト実行
runAnalysis().catch(console.error);
