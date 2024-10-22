import ExcelJS from "exceljs";

const data = {
	node: [
		{
			serial: 1,
			tagName: "div",
			id: "hogeId",
			class: ["hogeClass1", "hogeClass2"],
			style: {
				hogeClass1:
					".hogeClass1 {display: block; position: absolute} .hogeClass1:hover { background: red;}",
				hogeClass2:
					".hogeClass2 {display: block; position:relative} .hogeClass2::before { position: aboslute; background: red;}",
			},
			size: {
				x: 500,
				y: 500,
			},
			position: {
				x: 0,
				y: 0,
			},
		},
		{
			serial: 2,
			tagName: "span",
			id: "hogeId",
			class: ["hogeClass1", "hogeClass2"],
			style: {
				hogeClass1:
					".hogeClass1 {display: block; position: absolute} .hogeClass1:hover { background: red;}",
				hogeClass2:
					".hogeClass2 {display: block; position:relative padding: 500px;} .hogeClass2::before { position: aboslute; background: red;}",
			},
			size: {
				x: 500,
				y: 4400,
			},
			position: {
				x: 8,
				y: 0,
			},
		},
	],
};

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

async function writeToExcel(templatePath, outputPath, data) {
	try {
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(templatePath);

		const worksheet = workbook.getWorksheet("Sheet1");

		// 既存データのクリーンアップを改善
		if (worksheet.rowCount > 1) {
			// 下から順番に行を削除（ヘッダー以外）
			for (let i = worksheet.rowCount; i > 1; i--) {
				worksheet.spliceRows(i, 1);
			}

			// ワークシートのプロパティをリセット
			worksheet.properties.outlineLevelRow = 0;
			worksheet.properties.outlineLevelCol = 0;
		}

		// ヘッダー行のスタイルを再設定
		const headerRow = worksheet.getRow(1);
		headerRow.height = 30; // ヘッダーの高さを固定

		for (const [index, header] of data.headers.entries()) {
			const cell = headerRow.getCell(index + 1);
			cell.value = header;
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

		let currentSerial = null;
		let rowColor = "FFFFFF";

		// データ行の追加
		for (const [rowIndex, row] of data.rows.entries()) {
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

			// 行の高さを固定
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

				const columnName = data.headers[cell.col - 1];
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

		await workbook.xlsx.writeFile(outputPath);
		console.log("Excel file has been written successfully");
	} catch (error) {
		console.error("Error writing Excel file:", error);
		throw error;
	}
}

// 使用例
const templatePath = "./sample.xlsx"; // 既存のExcelファイルパス
const outputPath = "./output.xlsx"; // 出力先のファイルパス

// 先ほどの整形関数で作成したデータを使用
const excelData = formatDataForExcel(data);

// Excelファイルに書き込み
writeToExcel(templatePath, outputPath, excelData)
	.then(() => console.log("処理が完了しました"))
	.catch((error) => console.error("エラーが発生しました:", error));
