# フォルダ構成草案

```
webpage-extractor/
├── src/
│ ├── main/ # Electron メインプロセス
│ │ ├── index.ts # メインプロセスのエントリーポイント
│ │ └── ipc/ # IPC 通信ハンドラー
│ │ └── handlers.ts
│ │
│ ├── renderer/ # Electron レンダラープロセス
│ │ ├── App.tsx # メインの React コンポーネント
│ │ ├── index.html
│ │ └── styles/
│ │
│ ├── core/ # 共通のビジネスロジック
│ │ ├── types/ # 型定義
│ │ │ ├── analysis.ts # 解析結果の型定義
│ │ │ └── config.ts # 設定の型定義
│ │ │
│ │ ├── services/ # ビジネスロジック
│ │ │ ├── analyzer.ts # Playwright 解析ロジック
│ │ │ └── excel.ts # Excel 生成ロジック
│ │ │
│ │ └── utils/ # ユーティリティ関数
│ │ ├── file.ts # ファイル操作
│ │ └── path.ts # パス操作
│ │
│ └── config/ # 設定ファイル
│ └── default.ts # デフォルト設定
│
├── dist/ # ビルド成果物
├── dist-electron/ # Electron ビルド成果物
├── output/ # 解析結果の出力先
│ ├── json/ # JSON 中間ファイル
│ ├── excel/ # Excel 出力ファイル
│ └── screenshots/ # スクリーンショット
│
├── test/ # テストファイル
│ ├── analyzer.test.ts
│ └── excel.test.ts
│
├── scripts/ # ビルドスクリプトなど
├── .gitignore
├── electron-builder.json5 # Electron Builder 設定
├── package.json
├── tsconfig.json
└── vite.config.ts
```
