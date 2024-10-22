import ExcelJS from "exceljs";

const screenDataList = [
	{
		node: [
			{
				serial: 1,
				tagName: "div",
				id: "screen1Id",
				class: ["screen1Class1"],
				style: {
					screen1Class1: ".screen1Class1 {display: block; position: absolute}",
				},
				size: { x: 500, y: 500 },
				position: { x: 0, y: 0 },
			},
		],
	},
	{
		node: [
			{
				serial: 1,
				tagName: "div",
				id: "screen2Id",
				class: ["screen2Class1"],
				style: {
					screen2Class1: ".screen2Class1 {display: flex; position: relative}",
				},
				size: { x: 600, y: 400 },
				position: { x: 10, y: 20 },
			},
		],
	},
	{
		node: [
			{
				serial: 1,
				tagName: "div",
				id: "screen3Id",
				class: ["screen2Class1"],
				style: {
					screen2Class1: ".screen2Class1 {display: flex; position: relative}",
				},
				size: { x: 600, y: 400 },
				position: { x: 10, y: 20 },
			},
		],
	},
];

function extractCSSProperties(cssString) {
	// ブラケット内のコンテンツのみを抽出
	const matches = cssString.match(/{([^}]+)}/);
	if (!matches) return [];

	// プロパティを分割して整形
	return matches[1]
		.split(";")
		.map((prop) => prop.trim())
		.filter((prop) => prop) // 空文字を除去
		.map((prop) => {
			const [property, value] = prop.split(":").map((p) => p.trim());
			return `${property}: ${value}`;
		});
}

function formatDataForExcel(data) {
	const headers = [
		"No.",
		"タグ名",
		"ID",
		"クラス名",
		"サイズ(幅)",
		"サイズ(高さ)",
		"位置(X)",
		"位置(Y)",
		"スタイル適用クラス",
		"スタイル定義",
	];

	const rows = [];

	for (const node of data.node) {
		let isFirstRow = true;
		const allStyles = [];

		for (const [className, style] of Object.entries(node.style)) {
			const styleRules = style
				.split("}")
				.filter((rule) => rule.trim())
				.map((rule) => `${rule.trim()}}`);

			for (const rule of styleRules) {
				const properties = extractCSSProperties(rule);
				for (const prop of properties) {
					allStyles.push({
						className,
						property: prop,
					});
				}
			}
		}

		for (const style of allStyles) {
			const row = {
				styleClass: style.className,
				styleDefinition: style.property,
			};

			if (isFirstRow) {
				row.serial = node.serial;
				row.tagName = node.tagName;
				row.id = node.id;
				row.className = node.class.join(", ");
				row.width = node.size.x;
				row.height = node.size.y;
				row.positionX = node.position.x;
				row.positionY = node.position.y;
				isFirstRow = false;
			} else {
				row.serial = "";
				row.tagName = "";
				row.id = "";
				row.className = "";
				row.width = "";
				row.height = "";
				row.positionX = "";
				row.positionY = "";
			}

			rows.push(row);
		}
	}

	return {
		headers,
		rows,
	};
}

async function writeMultipleScreensToExcel(
	templatePath,
	outputPath,
	screenDataList,
	sheetNames = []
) {
	try {
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(templatePath);

		for (const [index, screenData] of screenDataList.entries()) {
			// シート名の決定
			const sheetName = sheetNames[index] || `Screen${index + 1}`;

			// シートの取得または作成
			let worksheet = workbook.getWorksheet(sheetName);
			if (!worksheet) {
				worksheet = workbook.addWorksheet(sheetName);
			}

			// 既存データのクリーンアップ
			if (worksheet.rowCount > 0) {
				worksheet.spliceRows(1, worksheet.rowCount);
			}

			// データをExcel形式に変換
			const excelData = formatDataForExcel(screenData);

			// ヘッダー行の設定
			const headerRow = worksheet.addRow(excelData.headers);
			headerRow.height = 30;

			for (const [colIndex, header] of excelData.headers.entries()) {
				const cell = headerRow.getCell(colIndex + 1);
				cell.fill = {
					type: "pattern",
					pattern: "solid",
					fgColor: { argb: "FFE0E0E0" },
				};
				cell.border = {
					top: { style: "thin" },
					left: { style: "thin" },
					bottom: { style: "thin" },
					right: { style: "thin" },
				};
				cell.font = { bold: true };
				cell.alignment = {
					vertical: "middle",
					horizontal: "center",
					wrapText: true,
				};
			}

			// データ行の追加
			let currentSerial = null;
			let rowColor = "FFFFFF";

			for (const [rowIndex, row] of excelData.rows.entries()) {
				if (row.serial && row.serial !== currentSerial) {
					currentSerial = row.serial;
					rowColor = rowColor === "FFFFFF" ? "F5F5F5" : "FFFFFF";
				}

				const excelRow = worksheet.addRow([
					row.serial,
					row.tagName,
					row.id,
					row.className,
					row.width,
					row.height,
					row.positionX,
					row.positionY,
					row.styleClass,
					row.styleDefinition,
				]);

				excelRow.height = 25;

				for (const cell of excelRow._cells) {
					if (!cell) continue;

					cell.border = {
						top: { style: "thin" },
						left: { style: "thin" },
						bottom: { style: "thin" },
						right: { style: "thin" },
					};
					cell.fill = {
						type: "pattern",
						pattern: "solid",
						fgColor: { argb: rowColor },
					};

					const columnName = excelData.headers[cell.col - 1];
					if (["width", "height", "positionX", "positionY"].includes(columnName)) {
						cell.alignment = { horizontal: "right", vertical: "middle" };
					} else {
						cell.alignment = {
							horizontal: "left",
							vertical: "middle",
							wrapText: true,
						};
					}
				}
			}

			// 列幅の自動調整
			for (const column of worksheet.columns) {
				let maxLength = 0;
				column.eachCell({ includeEmpty: true }, (cell) => {
					const columnLength = cell.value ? cell.value.toString().length : 10;
					maxLength = Math.max(maxLength, columnLength);
				});
				column.width = Math.min(maxLength + 2, 50);
			}
		}

		await workbook.xlsx.writeFile(outputPath);
		console.log("Excel file has been written successfully");
	} catch (error) {
		console.error("Error writing Excel file:", error);
		throw error;
	}
}

// const asdf = screenDataList.map((node) => console.log(node));

function getSheetNames(list) {
	return list.map((item) => item.node[0].id);
}

const sheetNames = getSheetNames(screenDataList);

// シート名を指定して実行
writeMultipleScreensToExcel(
	"./sample.xlsx",
	"./output.xlsx",
	screenDataList,
	sheetNames
)
	.then(() => console.log("すべての画面データの書き込みが完了しました"))
	.catch((error) => console.error("エラーが発生しました:", error));
