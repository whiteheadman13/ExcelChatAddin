# Module List

このファイルは修正開始前に必ず参照する前提の一覧です。
以後の変更では、対象機能に関連するモジュールをこの一覧で先に確認します。

## Session Update (複数テーブル定義対応+@table記法+更新対象ComboBox)
- 事象:
  1. table_schema.json が1テーブル分しか保持できない
  2. テーブル参照が @range 形式で分かりにくい
  3. 更新対象テーブルを明示的に指定する手段がない
  4. 更新対象テーブル未選択時もJSON強制されてしまう
- 修正内容:
  1. `TableSchemaStore` クラス追加。`table_schema.json` を配列形式（`{ "Tables": [...] }`）に変更
  2. 旧単体JSON / `issue_schema.json` からの自動移行ロジック
  3. `IssueSchemaManager` に `LoadStore` / `SaveStore` / `FindByTableName` / `Upsert` を追加（旧 `LoadOrCreate` / `Save` は互換ラッパー化）
  4. `IssueSchemaSettingsDialog` をテーブル切替対応に改修（ComboBox選択で定義切替 + 定義削除ボタン追加）
  5. `@table("テーブル名")` Regex (`TableTagRegex`) を追加
  6. テーブル一覧ダブルクリック時の挿入形式を `@range(Sheet,Addr)` → `@table("テーブル名")` に変更
  7. `ResolveTableToRangeKey` ヘルパーで `@table` → `Sheet!Addr` 参照キーへ解決
  8. `BuildMaskedPayload` で `@table` タグを `@range_ref` に変換して参照データに含める
  9. ChatView に「更新対象」ComboBox (`cmbUpdateTarget`) を追加
  10. `RefreshSheetList` で更新対象ComboBoxを連動更新（複数テーブル定義のHasSchema判定も配列対応）
  11. `BuildSchemaSection` を全面改修: 全参照テーブルの定義を同梱 + `_selectedUpdateTable` 選択時のみJSON強制指示を付加
  12. `RenderPreview` で `@table` タグもクリック可能なリンクとして表示
- 対象モジュール:
  - `ExcelChatAddin/IssueSchemaConfig.cs`
  - `ExcelChatAddin/IssueSchemaSettingsDialog.cs`
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (JSON operations形式でのLLM応答と反映)
- 事象:
  1. LLMがCSV/テキスト形式で返すため反映時に全データが1列に入ってしまう
  2. 項目定義を送っているのにJSON出力指示がない
- 修正内容:
  1. `BuildSchemaSection` にJSON出力フォーマット指示を追加（operations/key/fields/errors）
  2. `TryParseJsonOperations` を追加: ```json``` ブロックまたは `{` で始まるJSONを抽出・パース
  3. `ApplyJsonOperations` を追加: スキーマの列名→列インデックスでマッピングし、キー列で既存行検索→upsert/insert/update
  4. `ApplyResponseToSheet` をJSON優先→Markdown/TSVフォールバックの2段構えに変更
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (テーブル一覧UIへの変更と項目定義のLLM送信)
- 事象:
  1. シート一覧ではなくテーブル一覧を表示したい
  2. テーブル名に一致する定義がある場合のみ、その定義をLLMに送信したい
  3. 表設定画面も「対象シート」→「対象テーブル名」に変更したい
- 修正内容:
  1. ChatView UIを「テーブル一覧」に変更（ListObject.Name + シート名 + 範囲を表示）
  2. 定義ありテーブルに「★定義あり」を表示
  3. ダブルクリックで `@range(シート名,テーブル範囲)` を入力欄に挿入
  4. `BuildMaskedPayload` に `BuildSchemaSection` を追加: 参照範囲内のテーブル名と定義が一致する場合のみ項目定義をペイロードに同梱
  5. `IssueSchemaConfig` に `TableName` プロパティを追加（旧 `SheetName` と後方互換）
  6. `IssueSchemaSettingsDialog` のラベルを「対象テーブル名」に変更
  7. `EnsureTableIfMissing` をテーブル名ベースの検索に変更
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`
  - `ExcelChatAddin/IssueSchemaConfig.cs`
  - `ExcelChatAddin/IssueSchemaSettingsDialog.cs`

## Session Update (シート指定時に全表範囲を送付し、反映を回答単位に変更)
- 事象:
  1. シート一覧ダブルクリックで `@range(Sheet,A1)` ではなく表全体を送りたい
  2. 反映対象はチャット履歴全体ではなく、各Gemini回答ごとに指定したい
- 修正内容:
  1. シート一覧ダブルクリック時、`ListObject.Range` 優先 / 無ければ `UsedRange` の `@range(Sheet,Address)` を入力欄へ挿入
  2. ヘッダ右上の共通「反映」ボタンを削除
  3. 各Gemini回答の「コピー」横に「反映」ボタンを表示（回答単位）
  4. 反映処理は回答ごとの関連入力（@range）を使って開始セルを解決
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (反映ボタンの配置変更と反映データ抽出の改善)
- 事象:
  1. 反映ボタンをチャット履歴ヘッダ（クリア履歴/再解析の横）に置きたい
  2. Gemini回答が曖昧な場合に反映できない
- 修正内容:
  1. 反映ボタンを入力欄側からヘッダ右上（クリア履歴/再解析の横）へ移動
  2. 反映データ抽出を強化（Markdown/TSVに加えて `A-001: 田中 / ...` のアクション行を解析）
  3. 回答が曖昧な場合は直近入力本文からもアクション行を救済抽出
  4. 開始セルがヘッダー値で先頭データがID形式の場合は1行下へ書き込み（ヘッダー上書き回避）
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (Gemini回答をシートへ反映する機能を追加)
- 事象: チャット回答が表示されるだけで、シートへ自動反映されない。
- 修正内容:
  1. `ChatView.xaml` に「反映」ボタンを追加
  2. 直近送信入力と直近Gemini回答を保持
  3. 反映時に `@range(シート,開始セル)` を優先、なければ選択セル/アクティブセルへ書き込み
  4. Markdown表/TSV(および簡易1列)を抽出して `Range.Value2` へ反映
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (チャット欄に現在シート一覧の分かりやすいUIを追加)
- 事象: チャット欄で現在ブックのシート一覧を見やすく表示し、手動更新できるUIがほしい。
- 修正内容:
  1. `ChatView.xaml` に「シート一覧」パネル（一覧を更新ボタン + ListBox）を追加
  2. 画面ロード時にシート一覧を初期表示
  3. 「一覧を更新」クリックで現在ブックのシート一覧を再取得
  4. シート名をダブルクリックすると `@range(シート名,A1)` を入力欄に追加
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml`
  - `ExcelChatAddin/ChatView.xaml.cs`

## Session Update (表設定保存時に未作成テーブルを自動作成)
- 事象: 表設定の保存時、対象シートに表が未作成なら自動で作成したい。
- 修正内容:
  1. 保存後に `EnsureTableIfMissing` を実行
  2. 対象シートが無ければ新規作成
  3. ヘッダー未作成時のみ列名ヘッダーを書き込み
  4. ListObject が無い場合に Excel テーブルを自動作成
- 対象モジュール:
  - `ExcelChatAddin/IssueSchemaSettingsDialog.cs`

## Session Update (課題設定を汎用の表設定へ名称・保存先を変更)
- 事象: 表更新機能は課題管理表に限定しないため、UI文言と設定保存先を汎用化したい。
- 修正内容:
  1. リボンボタン文言を「課題設定」から「表設定」へ変更
  2. `ThisAddIn` に `ShowTableSchemaSettings()` を追加し、既存 `ShowIssueSchemaSettings()` は互換ラッパー化
  3. 設定画面タイトルを「表スキーマ設定」に変更
  4. 保存先を `table_schema.json` 優先に変更し、既存 `issue_schema.json` を互換読み込み・移行保存
- 対象モジュール:
  - `ExcelChatAddin/ChatRibbon.cs`
  - `ExcelChatAddin/ThisAddIn.cs`
  - `ExcelChatAddin/IssueSchemaSettingsDialog.cs`
  - `ExcelChatAddin/IssueSchemaConfig.cs`
  - `ExcelChatAddin/Paths.cs`

## Session Update (リボンから課題設定画面を起動し、issue_schema.jsonを作成)
- 事象: 課題管理表スキーマをユーザーが設定できる画面をリボンから起動したい。
- 修正内容:
  1. `ChatRibbon` に「課題設定」ボタンを追加し、設定画面を起動
  2. `ThisAddIn` に `ShowIssueSchemaSettings()` を追加
  3. `IssueSchemaSettingsDialog` を追加（列名/列位置/キー列/値候補/記載例を編集） 
  4. `IssueSchemaManager` を追加し `issue_schema.json` を作成・保存
  5. `Paths` に `IssueSchemaPath` を追加
- 対象モジュール:
  - `ExcelChatAddin/ChatRibbon.cs`
  - `ExcelChatAddin/ThisAddIn.cs`
  - `ExcelChatAddin/IssueSchemaSettingsDialog.cs`
  - `ExcelChatAddin/IssueSchemaConfig.cs`
  - `ExcelChatAddin/Paths.cs`
  - `ExcelChatAddin/ExcelChatAddin.csproj`

## Session Update (用語登録画面の分類登録をチェックボックス判断へ変更)
- 事象: 用語登録画面で、分類履歴への保存有無をチェックボックスで判断したい。
- 参照元:
  - `C:\Users\Masay\source\repos\powerpoint_masking2\docs\メソッド一覧.md`
  - `C:\Users\Masay\source\repos\powerpoint_masking2\powerpoint_masking2\RegisterDialog.cs`
- 修正内容:
  1. `RegisterDialog` に「履歴に保存」チェックボックスを追加
  2. 新規カテゴリ登録時、チェックON時のみ `categories.txt` へ保存する
  3. カテゴリ履歴削除ボタンを追加し、履歴ファイルへ反映
  4. 新規/既存切替時の有効状態を更新（チェックボックス/削除ボタン含む）
- 対象モジュール:
  - `ExcelChatAddin/RegisterDialog.cs`

## Session Update (チャットプレビュー画面からの用語登録)
- 事象: チャットのマスキング確認プレビュー画面から、選択語を直接用語登録したい。
- 参照元:
  - `C:\Users\Masay\source\repos\powerpoint_masking2\docs\メソッド一覧.md`
- 修正内容:
  1. `MaskPreviewWindow` に「選択語を登録」ボタンを追加
  2. プレビュー内選択テキストで `RegisterDialog` を開いて辞書登録できるようにする
  3. 登録後、プレビュー中の選択箇所をプレースホルダへ置換する
- 対象モジュール:
  - `ExcelChatAddin/MaskPreviewWindow.xaml`
  - `ExcelChatAddin/MaskPreviewWindow.xaml.cs`

## Session Update (リボンに辞書表示機能を追加)
- 事象: 右クリックメニューだけでなく、リボンからも辞書管理画面を開きたい。
- 修正内容:
  1. `ChatRibbon` に「辞書表示」ボタンを追加
  2. `ThisAddIn` に `ShowDictionaryManager()` を追加して右クリックとリボンで共通利用
- 対象モジュール:
  - `ExcelChatAddin/ChatRibbon.cs`
  - `ExcelChatAddin/ThisAddIn.cs`

## Session Update (コミット運用ルール追記)
- 事象: コミット運用ルールを指示ファイルへ明示したい。
- 修正内容:
  1. `.github/copilot-instructions.md` に「コミットは常にタスク終了後に実施する」を追記する
- 対象モジュール:
  - `.github/copilot-instructions.md`

## Session Update (辞書管理更新処理の PowerPoint 実装反映)
- 事象: Excel 側の辞書管理画面は、PowerPoint 側で更新済みの安全な保存ロジックが未反映だった。
- 参照元:
  - `C:\Users\Masay\source\repos\powerpoint_masking2\docs\メソッド一覧.md`
  - `C:\Users\Masay\source\repos\powerpoint_masking2\powerpoint_masking2\DictionaryManager.cs`
  - `C:\Users\Masay\source\repos\powerpoint_masking2\powerpoint_masking2\DictionaryManagerLogic.cs`
- 修正内容:
  1. フィルタ中でも表示中行の編集内容を `_originalData` へマージする
  2. フィルタ非表示の行は保存時に削除しない
  3. 非フィルタ時のみ、グリッドから消えた行を削除扱いにする
  4. 補助ロジックを `DictionaryManagerLogic.cs` として分離する
- 対象モジュール:
  - `ExcelChatAddin/DictionaryManager.cs`
  - `ExcelChatAddin/DictionaryManagerLogic.cs`
  - `ExcelChatAddin/ExcelChatAddin.csproj`

## Session Update (rules.json 読込失敗時の保護強化)
- 事象: `rules.json` 読込失敗時にデータ消失へつながる恐れがある。
- 修正内容:
  1. 起動時に `rules.json` を毎回バックアップし、2世代保持する
  2. `rules.json` を読めない場合は理由を保持し、利用時に明示する
  3. 読込失敗時は Secure Chat を開けないようにする
  4. 読込失敗時は辞書管理/マスキング登録の書き込みも止める
- 対象モジュール:
  - `ExcelChatAddin/MaskingEngine.cs`
  - `ExcelChatAddin/ThisAddIn.cs`
  - `ExcelChatAddin/ChatRibbon.cs`

## Session Update (マスキング・辞書管理の修正)
- 事象: マスキング機能と辞書管理機能が正常に機能しない。ソースファイルのフォーマット崩れあり。
- 修正内容:
  1. `MaskingEngine.cs` — `AddRule` メソッドのインデント不整合を修正（4sp→8sp）、不適切なコメント削除
  2. `RegisterDialog.cs` — クラス外の孤立コメント `// Paths.cs` と余分な空行を削除
  3. `TaskPaneHost.cs` — `RangeToPlainText` のセル区切りを `" | "` → `\t` に修正。`TsvToMarkdownTable` がTSVを前提としておりセル単位マスキングが破壊されていた
  4. `ThisAddIn.cs` — `UnregisterHotKeys` のHWNDを `this.Application.Hwnd` → `_hotKeyWindow.WindowHandle` に修正（登録先と解除先の不一致）
- 対象モジュール:
  - `ExcelChatAddin/MaskingEngine.cs`
  - `ExcelChatAddin/RegisterDialog.cs`
  - `ExcelChatAddin/TaskPaneHost.cs`
  - `ExcelChatAddin/ThisAddIn.cs`

## Session Update (チャット履歴の表の折りたたみ対応)
- 事象: チャット履歴に大きい表が表示されると、履歴欄を圧迫して可読性が下がる。
- 要望: 表部分だけを折りたたみ可能にする。
- 対象モジュール:
  - `ExcelChatAddin/ChatView.xaml.cs`（表レンダリング部・再解析後の表表示部）

## Session Update (テンプレート不具合対応)
- 事象: テンプレートファイル `diagram_templates.json` が旧形式キー（`Name`/`Prompt`）の場合、現行読み込みが `Title`/`Body` 前提のため一覧表示が崩れる。
- 対象モジュール:
  - `ExcelChatAddin/TemplateManager.cs`（テンプレート読込互換・正規化）
  - `ExcelChatAddin/TemplateDialog.cs`（既存利用、読み込み結果の表示側確認）

## Project Overview
- Project: `ExcelChatAddin/ExcelChatAddin.csproj`
- Type: Excel VSTO Add-in (.NET Framework 4.8)
- Main responsibilities: Excel連携、チャットUI、Gemini API連携、マスキング辞書管理、テンプレート管理

## Modules

| 物理名 | 種別 | 役割 | VBA対応/備考 |
|---|---|---|---|
| `ExcelChatAddin/ThisAddIn.cs` | C# | アドイン起動・終了、右クリックメニュー追加、ホットキー登録、各UI起動の中核 | Excelアドイン本体のエントリーポイント |
| `ExcelChatAddin/ThisAddIn.Designer.cs` | C# Designer | VSTO生成コード、アドインイベント配線 | 自動生成コード |
| `ExcelChatAddin/ChatRibbon.cs` | C# | リボン拡張（`IRibbonExtensibility`実装）、「チャット表示」ボタン定義、クリック時に`ShowChat()`呼び出し | Excelリボンタブ「Secure Chat」 |
| `ExcelChatAddin/TaskPaneHost.cs` | C# WinForms | カスタムタスクペインのホスト、WPF `ChatView` 埋め込み、Excel選択範囲操作 | Excel画面との橋渡し |
| `ExcelChatAddin/TaskPaneHost.Designer.cs` | C# Designer | `TaskPaneHost` の自動生成UI | 自動生成コード |
| `ExcelChatAddin/ChatView.xaml` | XAML | チャット画面レイアウト | WPF UI |
| `ExcelChatAddin/ChatView.xaml.cs` | C# WPF | チャットUI制御、範囲参照解析、表形式テキスト処理、送受信補助 | チャット体験の中心 |
| `ExcelChatAddin/ChatModels.cs` | C# | チャットセッション/メッセージのデータモデル | モデル定義 |
| `ExcelChatAddin/GeminiClient.cs` | C# | Gemini API呼び出し、system instruction生成、HTTP送受信、応答取得 | 外部AI連携 |
| `ExcelChatAddin/GeminiDtos.cs` | C# | Gemini APIリクエスト用DTO | `GeminiClient` 用データ契約 |
| `ExcelChatAddin/GeminiResponseWindow.xaml.cs` | C# WPF | Gemini応答表示ウィンドウのコードビハインド | 表示用補助UI |
| `ExcelChatAddin/MaskingEngine.cs` | C# | マスキング辞書の保持、保存/読込、マスク/復元、プレースホルダ生成 | 既存マスキング機能の中核 |
| `ExcelChatAddin/DictionaryManager.cs` | C# WinForms | マスキング辞書の一覧・検索・編集・削除UI | 辞書管理画面 |
| `ExcelChatAddin/RegisterDialog.cs` | C# WinForms | 選択文字列のマスキング登録UI、カテゴリ/既存タグ選択 | マスキング登録画面 |
| `ExcelChatAddin/MaskPreviewWindow.xaml` | XAML | マスク後文字列のプレビュー画面レイアウト | WPF UI |
| `ExcelChatAddin/MaskPreviewWindow.xaml.cs` | C# WPF | マスク後文字列の表示 | 表示用補助UI |
| `ExcelChatAddin/TemplateManager.cs` | C# | テンプレート保存/読込、ID採番 | テンプレート永続化 |
| `ExcelChatAddin/TemplateDialog.cs` | C# WinForms | テンプレート一覧、選択、新規、編集UI | テンプレート選択画面 |
| `ExcelChatAddin/TemplateEditDialog.cs` | C# WinForms | テンプレート編集UI | テンプレート編集画面 |
| `ExcelChatAddin/Paths.cs` | C# | AppData/環境変数ベースの保存先統一、旧データ移行 | 永続データ保存先の基盤 |
| `ExcelChatAddin/DebugLogger.cs` | C# | ローカルログ出力 | デバッグ補助 |
| `ExcelChatAddin/HotKeyWindow.cs` | C# | WM_HOTKEY受信用のメッセージ専用ウィンドウ | グローバルホットキー補助 |
| `ExcelChatAddin/Win32Window.cs` | C# | Win32ハンドルを `IWin32Window` として扱うラッパー | ダイアログオーナー設定補助 |
| `ExcelChatAddin/Properties/AssemblyInfo.cs` | C# | アセンブリ属性定義 | メタ情報 |
| `ExcelChatAddin/Properties/Resources.Designer.cs` | C# Designer | 埋め込みリソースの自動生成アクセサ | 自動生成コード |
| `ExcelChatAddin/Properties/Resources.resx` | Resource | リソース定義 | リソース |
| `ExcelChatAddin/Properties/Settings.settings` | Settings | 設定定義 | 設定元ファイル |

## Related Data Files

| 物理名 | 役割 |
|---|---|
| `AppData/OfficeChatMasking/rules.json` | マスキング辞書 |
| `AppData/OfficeChatMasking/categories.txt` | カテゴリ履歴 |
| `AppData/OfficeChatMasking/diagram_templates.json` | テンプレート保存 |
| `AppData/OfficeChatMasking/config.json` | 将来/既存設定用 |

## Change Procedure
1. 修正対象の機能を確認する。
2. この `module_list.md` で関連モジュールを先に確認する。
3. 影響範囲を絞って最小変更で実装する。
4. 修正後はビルド/必要テストで確認する。
