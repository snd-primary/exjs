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
	],
};

function formatCSSProperties(cssString) {
	// セレクタとプロパティ部分を分離
	const parts = cssString.split("{");
	if (parts.length !== 2) return cssString;

	const selector = parts[0].trim();
	const properties = parts[1].replace("}", "").trim();

	// プロパティを分割
	const formattedProperties = properties
		.split(";")
		.filter((prop) => prop.trim()) // 空の要素を除去
		.map((prop) => {
			// プロパティと値を取得して整形
			const [property, value] = prop.split(":").map((p) => p.trim());
			return `${property}: ${value};`;
		})
		.join("\n");

	return `${selector} {\n${formattedProperties}\n}`;
}

function formatDataForExcel(data) {
	// ヘッダー行の定義
	const headers = [
		"No.",
		"タグ名",
		"ID",
		"クラス名",
		"サイズ(幅)",
		"サイズ(高さ)",
		"位置(X)",
		"位置(Y)",
		"スタイル定義",
	];

	// データ行の作成
	const rows = data.node.map((node) => {
		// クラス名を文字列に結合
		const classNames = node.class.join(", ");

		// スタイル定義を整形
		const styleDefinitions = Object.entries(node.style)
			.map(([className, style]) => {
				// 複数のCSSルールを分割（スペースで区切られている場合に対応）
				const cssRules = style
					.split("}")
					.filter((rule) => rule.trim())
					.map((rule) => rule.trim() + "}");

				// 各ルールを整形
				const formattedRules = cssRules
					.map((rule) => formatCSSProperties(rule))
					.join("\n\n");

				return `/* ${className} */\n${formattedRules}`;
			})
			.join("\n\n");

		// 1行分のデータを作成
		return {
			serial: node.serial,
			tagName: node.tagName,
			id: node.id,
			className: classNames,
			width: node.size.x,
			height: node.size.y,
			positionX: node.position.x,
			positionY: node.position.y,
			styles: styleDefinitions,
		};
	});

	// ヘッダーとデータを結合
	return {
		headers,
		rows,
	};
}

/**
 * 既存のExcelファイルにデータを書き込む関数
 * @param {string} templatePath - 既存のExcelファイルのパス
 * @param {string} outputPath - 出力するExcelファイルのパス
 * @param {Object} data - 書き込むデータ
 */
async function writeToExcel(templatePath, outputPath, data) {
	try {
		// ワークブックを読み込む
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(templatePath);

		// Sheet1を取得
		const worksheet = workbook.getWorksheet("Sheet1");

		// 前回のデータをクリア（2行目以降をクリア）
		const lastRow = worksheet.lastRow;
		if (lastRow && lastRow.number > 1) {
			worksheet.spliceRows(2, lastRow.number - 1);
		}

		// ヘッダー行のスタイルを設定
		const headerRow = worksheet.getRow(1);
		data.headers.forEach((header, index) => {
			const cell = headerRow.getCell(index + 1);
			cell.value = header;

			cell.border = {
				top: { style: "thin" },
				left: { style: "thin" },
				bottom: { style: "thin" },
				right: { style: "thin" },
			};
			cell.font = { bold: true };
		});

		// データ行を追加
		data.rows.forEach((row, rowIndex) => {
			const excelRow = worksheet.addRow([
				row.serial,
				row.tagName,
				row.id,
				row.className,
				row.width,
				row.height,
				row.positionX,
				row.positionY,
				row.styles,
			]);

			// データ行のスタイルを設定
			excelRow.eachCell((cell) => {
				cell.border = {
					top: { style: "thin" },
					left: { style: "thin" },
					bottom: { style: "thin" },
					right: { style: "thin" },
				};

				// スタイル定義列は折り返して表示
				if (cell.col === 9) {
					// スタイル定義列
					cell.alignment = {
						wrapText: true,
						vertical: "top",
					};
				}
			});

			// 行の高さを自動調整
			excelRow.height = 20;
		});

		// 列幅の自動調整
		worksheet.columns.forEach((column) => {
			let maxLength = 0;
			column.eachCell({ includeEmpty: true }, (cell) => {
				const columnLength = cell.value ? cell.value.toString().length : 10;
				if (columnLength > maxLength) {
					maxLength = columnLength;
				}
			});
			column.width = Math.min(maxLength + 2, 50); // 最大幅を50文字に制限
		});

		// ファイルを保存
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
