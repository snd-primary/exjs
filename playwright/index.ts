import { chromium } from "playwright";
import type { Page, Browser } from "playwright";

type HtmlNodeContent = HtmlNode | string;

interface HtmlNode {
	tag: string;
	attributes: { [key: string]: string };
	children: HtmlNodeContent[];
}

(async () => {
	const browser: Browser = await chromium.launch({
		headless: false,
		slowMo: 500,
	});
	const page: Page = await browser.newPage();

	await page.goto("http://localhost:5050");

	const pageLinkLocator = page.locator("text=mock");
	await pageLinkLocator.click();

	const bodyContent = await page.evaluate(() => {
		function extractNode(node: Node): HtmlNodeContent | null {
			if (node.nodeType === Node.TEXT_NODE) {
				return node.textContent?.trim() || null;
			}

			if (node.nodeType === Node.ELEMENT_NODE) {
				const element = node as Element;

				const result = {
					tag: element.tagName.toLowerCase(),
					attributes: {},
					children: [],
				};

				for (const attr of element.attributes) {
					result.attributes[attr.name] = attr.value;
				}

				for (const childNode of element.childNodes) {
					const childResult = extractNode(childNode);
					if (childResult === null) continue;
					result.children.push(childResult);
				}
				return result;
			}

			return null;
		}

		const body = document.body;
		return extractNode(body);
	});

	await page.close();

	console.log(JSON.stringify(bodyContent, null, 2));
})();
