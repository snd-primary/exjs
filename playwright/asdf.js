import { chromium } from "playwright";
import ExcelJS from "exceljs";
import { promises as fs } from "fs";
import path from "path";
import { fileURLToPath } from "url";

// Get current file path
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function mapWebElementsToExcel(url) {
	console.log(`Starting analysis of ${url}`);

	// Launch browser and navigate to page
	const browser = await chromium.launch();
	const context = await browser.newContext();
	const page = await context.newPage();

	try {
		// Navigate to page and wait for load
		await page.goto(url, { waitUntil: "networkidle" });
		console.log("Page loaded successfully");

		// Get page dimensions first
		const pageWidth = await page.evaluate(
			() => document.documentElement.scrollWidth
		);
		const pageHeight = await page.evaluate(
			() => document.documentElement.scrollHeight
		);

		// Set viewport to match full page size
		await page.setViewportSize({
			width: pageWidth,
			height: pageHeight,
		});
		console.log(`Viewport set to ${pageWidth}x${pageHeight}`);

		// Get all elements and their positions
		const elements = await page.evaluate(() => {
			const allElements = document.querySelectorAll("*");
			return Array.from(allElements).map((el) => {
				const rect = el.getBoundingClientRect();
				return {
					tag: el.tagName.toLowerCase(),
					x: rect.x,
					y: rect.y,
					width: rect.width,
					height: rect.height,
				};
			});
		});
		console.log(`Found ${elements.length} elements`);

		// Take screenshot
		const screenshot = await page.screenshot({
			fullPage: true,
		});
		console.log("Screenshot captured");

		// Create Excel workbook
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet("Web Elements");

		// Add screenshot to Excel
		const imageId = workbook.addImage({
			buffer: screenshot,
			extension: "png",
		});

		// Excel's column width is approximately 8 pixels
		const colWidth = pageWidth / 8;
		const rowHeight = pageHeight / 20; // Excel's default row height is about 20 pixels

		// Resize columns and rows to match screenshot dimensions
		worksheet.getColumn(1).width = colWidth;
		worksheet.getRow(1).height = pageHeight;

		// Add screenshot to worksheet
		worksheet.addImage(imageId, {
			tl: { col: 0, row: 0 },
			br: { col: colWidth, row: pageHeight / rowHeight },
			editAs: "oneCell",
		});

		// Add element borders
		let addedShapes = 0;
		elements.forEach((element, index) => {
			try {
				// Convert pixel positions to Excel cell positions
				const startCol = element.x / 8;
				const startRow = element.y / 20;
				const endCol = (element.x + element.width) / 8;
				const endRow = (element.y + element.height) / 20;

				// Only add shapes for visible elements
				if (element.width > 0 && element.height > 0) {
					worksheet.addShape({
						type: "rect",
						text: element.tag,
						textColor: "#FF0000",
						fill: {
							type: "none",
						},
						line: {
							color: "#FF0000",
							width: 1,
						},
						position: {
							tl: { col: startCol, row: startRow },
							br: { col: endCol, row: endRow },
						},
					});
					addedShapes++;
				}
			} catch (error) {
				console.warn(`Failed to add shape for element ${element.tag}:`, error);
			}
		});
		console.log(`Added ${addedShapes} element borders`);

		// Save workbook
		const outputPath = path.join(process.cwd(), "web-elements-mapping.xlsx");
		await workbook.xlsx.writeFile(outputPath);
		console.log(`Excel file saved to: ${outputPath}`);
	} catch (error) {
		console.error("Error during processing:", error);
		throw error;
	} finally {
		// Ensure browser is closed even if an error occurs
		await browser.close();
		console.log("Browser closed");
	}
}

// Export the function for use in other modules
export { mapWebElementsToExcel };

// Check if file is being run directly
if (process.argv[1] === fileURLToPath(import.meta.url)) {
	const targetUrl = process.argv[2] || "http://localhost:5050";
	console.log("Running script directly");
	mapWebElementsToExcel(targetUrl).catch((error) => {
		console.error("Script failed:", error);
		process.exit(1);
	});
}
