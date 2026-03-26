# Module List

このファイルは修正開始前に必ず参照する前提の一覧です。
以後の変更では、対象機能に関連するモジュールをこの一覧で先に確認します。

## Project Overview
- Project: `ExcelChatAddin/ExcelChatAddin.csproj`
- Type: Excel VSTO Add-in (.NET Framework 4.8)
- Main responsibilities: Excel連携、チャットUI、Gemini API連携、マスキング辞書管理、テンプレート管理

## Modules

| 物理名 | 種別 | 役割 | VBA対応/備考 |
|---|---|---|---|
| `ExcelChatAddin/ThisAddIn.cs` | C# | アドイン起動・終了、右クリックメニュー追加、ホットキー登録、各UI起動の中核 | Excelアドイン本体のエントリーポイント |
| `ExcelChatAddin/ThisAddIn.Designer.cs` | C# Designer | VSTO生成コード、アドインイベント配線 | 自動生成コード |
| `ExcelChatAddin/ChatRibbon.cs` | C# | リボン拡張用ファイル | 現状は実装が薄い/要確認 |
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
