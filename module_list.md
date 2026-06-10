# Module List

ファイルの追加・削除・役割変更時はこのファイルを必ず更新すること。

---

## ExcelChatAddin

### エントリーポイント・リボン

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/ThisAddIn.cs` | アドイン起動・終了の中核。右クリックメニュー（Cell CommandBar）の追加/削除、ホットキー（Ctrl+Alt+M）登録、各ダイアログ起動メソッドを集約 |
| `ExcelChatAddin/ThisAddIn.Designer.cs` | VSTO自動生成コード（編集不要） |
| `ExcelChatAddin/ChatRibbon.cs` | `IRibbonExtensibility` 実装。リボンタブ「Secure Chat」のボタン定義と各 `ThisAddIn` メソッドへの橋渡し |

### チャットUI

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/TaskPaneHost.cs` | カスタムタスクペインの WinForms ホスト。`ElementHost` で WPF `ChatView` を埋め込み、Excel セル操作のブリッジ |
| `ExcelChatAddin/TaskPaneHost.Designer.cs` | `TaskPaneHost` 自動生成コード（編集不要） |
| `ExcelChatAddin/ChatView.xaml` | チャット画面レイアウト（WPF） |
| `ExcelChatAddin/ChatView.xaml.cs` | チャットUIの全制御。`@range`/`@table` トークン解析、送信ペイロード構築（マスキング有無を切替・スキーマ同梱）、Gemini/ローカルLLM への送信分岐、応答のシート反映、検証ループ呼び出し。モデル選択コンボ（Gemini静的＋ローカル動的）、Local由来履歴がある状態でのGemini送信ブロック、マスキングOFFバッジ表示 |

### AI連携

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/GeminiClient.cs` | Gemini API への HTTP 通信（シングルトン）。system instruction 生成、ストリーミング非対応の GenerateAsync |
| `ExcelChatAddin/GeminiDtos.cs` | Gemini API リクエスト/レスポンス用 DTO |
| `ExcelChatAddin/GeminiResponseWindow.xaml.cs` | Gemini 応答のポップアップ表示ウィンドウ（WPF） |
| `ExcelChatAddin/OllamaClient.cs` | ローカルLLM（ollama）への HTTP 通信ラッパ。`/api/tags` でモデル一覧取得、`/api/chat` で送信。接続不可/モデル未pull/タイムアウト時はフォールバックせず明確な例外を投げる。パース/URL組み立ては `OllamaProtocol` に委譲 |
| `ExcelChatAddin/ContentValidator.cs` | LLM 生成 JSON の定義準拠バリデーション。検証プロンプト構築・ループ制御（最大3回）・date型誤指摘フィルタ |

### マスキング

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/DictionaryManager.cs` | マスキング辞書の一覧・検索・編集・削除・新規登録 UI（WinForms） |
| `ExcelChatAddin/RegisterDialog.cs` | 文字列のマスキング登録 UI（WinForms）。カテゴリ選択/既存タグ選択/カテゴリ履歴管理 |
| `ExcelChatAddin/MaskPreviewWindow.xaml` | マスク後テキストのプレビューレイアウト（WPF） |
| `ExcelChatAddin/MaskPreviewWindow.xaml.cs` | マスク後テキストの表示・プレビュー内選択語の辞書登録 |
| `ExcelChatAddin/DebugMaskingLogger.cs` | `IMaskingLogger` の `Debug.WriteLine` 実装。起動時に `MaskingEngine.SetLogger` へセット |

### テーブルスキーマ

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/IssueSchemaConfig.cs` | `IssueSchemaColumn` / `IssueSchemaConfig` / `TableSchemaStore` モデル定義 + `IssueSchemaManager`（`table_schema.json` の CRUD） |
| `ExcelChatAddin/IssueSchemaSettingsDialog.cs` | テーブル定義の追加・削除・列定義編集 UI。テンプレート保存/挿入・定義 Excel 出力ボタンを含む |
| `ExcelChatAddin/SchemaTemplateManager.cs` | スキーマテンプレートの CRUD（`schema_templates.json`）。`SchemaTemplateEntry` モデルを定義 |
| `ExcelChatAddin/SchemaTemplateEditDialog.cs` | テンプレートのフル編集ダイアログ（名前・説明・行設定・列定義グリッド） |
| `ExcelChatAddin/SchemaTemplateListDialog.cs` | テンプレート一覧・選択・管理ダイアログ（通常選択 / `manageOnly` モード兼用） |
| `ExcelChatAddin/SchemaTemplateSaveDialog.cs` | テンプレート名・説明入力ダイアログ |
| `ExcelChatAddin/SchemaTemplateTableNameDialog.cs` | テンプレート挿入時のテーブル名入力ダイアログ |
| `ExcelChatAddin/ColumnDetailDialog.cs` | 列定義の詳細編集ダイアログ |

### テーブル関係

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/TableRelationConfig.cs` | `RelationTypeMasterItem` / `TableRelationRule` / `TableRecordRelation` モデル + `TableRelationManager`（`table_relations.json` の CRUD・バックアップ） + `TableRelationSheetStore`（Excelシート「関係データ」の読み書き） |
| `ExcelChatAddin/TableRelationSettingsDialog.cs` | 関係種別マスタ・テーブル間ルール・レコード間関係の3タブ編集 UI（TSV一括貼付対応） |
| `ExcelChatAddin/TableRelationMatrixForm.cs` | レコード間関係のマトリクス UI。VirtualMode DataGridView で大量データに対応。詳細パネルで関係種別・意味・疎結合フラグを編集 |

### 設定・永続化

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/Paths.cs` | 永続データ保存先の統一管理。共通パスは `MaskingPaths` へ委譲し、Excel固有パス（`table_schema.json` / `table_relations.json` / テンプレート等）を追加定義 |
| `ExcelChatAddin/AddinConfig.cs` | `config.json` の読み書き。ローカルLLM のエンドポイント（`ollamaBaseUrl`）と最後に選択したモデル（`lastModel`）を永続化。未知のキーは保持 |

### テンプレート

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/TemplateManager.cs` | チャットプロンプトテンプレートの CRUD（`diagram_templates.json`）。旧形式キー（`Name`/`Prompt`）からの互換読み込みを含む |
| `ExcelChatAddin/TemplateDialog.cs` | テンプレート一覧・選択 UI（WinForms） |
| `ExcelChatAddin/TemplateEditDialog.cs` | テンプレート編集 UI（WinForms） |

### ユーティリティ

| ファイル | 役割 |
|---|---|
| `ExcelChatAddin/HotKeyWindow.cs` | `WM_HOTKEY` 受信用メッセージ専用ウィンドウ（Win32 API `RegisterHotKey` と対応） |
| `ExcelChatAddin/Win32Window.cs` | Win32 ウィンドウハンドルを `IWin32Window` として扱うラッパー（ダイアログのオーナー指定用） |
| `ExcelChatAddin/DebugLogger.cs` | `Debug.WriteLine` ラッパー |
| `ExcelChatAddin/DiffPreviewDialog.cs` | 差分プレビューダイアログ |

---

## OfficeMasking.Core

マスキング共通ライブラリ（.NET Framework 4.8 SDK スタイル）。Excel/PowerPoint 両アドインで共用。

| ファイル | 役割 |
|---|---|
| `OfficeMasking.Core/MaskingEngine.cs` | マスキングのシングルトンエンジン。`Mask` / `Unmask` / `AddRule` / `AddRuleWithPlaceholder` / `OverrideRules` を提供。ロード失敗時は `IsAvailable=false` で書き込みを保護 |
| `OfficeMasking.Core/MaskingRulesStore.cs` | `rules.json` の読み書きと50世代バックアップローテーション。保存のたびにバックアップを実行 |
| `OfficeMasking.Core/MaskingPaths.cs` | データ保存先の決定ロジック。優先順：環境変数 `OFFICE_MASKING_DATA_DIR` > `AppData\OfficeChatMasking`。旧フォルダ（`PowerPointMasking`/DLL直下）からの自動移行も管理 |
| `OfficeMasking.Core/DictionaryManagerLogic.cs` | フィルタ判定・削除候補抽出・バリデーション（`ValidateNewEntry`）・プレースホルダー生成（`GeneratePlaceholder`）のロジック層。UI非依存 |
| `OfficeMasking.Core/IMaskingLogger.cs` | ログ出力インターフェース定義 |
| `OfficeMasking.Core/NullMaskingLogger.cs` | Null オブジェクトパターンの `IMaskingLogger` 実装（テスト・デフォルト用） |
| `OfficeMasking.Core/OllamaProtocol.cs` | ローカルLLM（ollama）API の純粋ロジック。URL正規化、`/api/tags` のモデル一覧パース、`/api/chat` リクエスト生成・レスポンスパース。HTTP通信を持たず単体テスト可能 |

---

## OfficeMasking.Core.Tests

MSTest テストプロジェクト。`OfficeMasking.Core` のみを対象とする。

| ファイル | 役割 |
|---|---|
| `OfficeMasking.Core.Tests/MaskingEngineTests.cs` | `MaskingEngine` の単体テスト（Mask/Unmask/AddRule/ロード失敗保護等） |
| `OfficeMasking.Core.Tests/MaskingRulesStoreTests.cs` | `MaskingRulesStore` の単体テスト（Load/Save/50世代バックアップ/復元/旧形式エラー） |
| `OfficeMasking.Core.Tests/MaskingPathsTests.cs` | `MaskingPaths` の単体テスト（DataDir解決/IsDataDirEnvironmentConfigured等） |
| `OfficeMasking.Core.Tests/DictionaryManagerLogicTests.cs` | `DictionaryManagerLogic` の単体テスト（バリデーション/プレースホルダー生成/削除候補抽出等） |
| `OfficeMasking.Core.Tests/OllamaProtocolTests.cs` | `OllamaProtocol` の単体テスト（URL正規化/モデル一覧パース/チャットリクエスト生成/レスポンスパース） |
