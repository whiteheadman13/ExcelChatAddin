# Module List

このファイルは修正開始前に必ず参照する前提の一覧です。
以後の変更では、対象機能に関連するモジュールをこの一覧で先に確認します。

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
