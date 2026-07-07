# ロードマップ / TODO（ExcelChatAddin）

このファイルはプロジェクトの**永続 TODO リスト**です。ユーザーが「TODO」と言った場合はこのファイルを指します。
今後発生した TODO は原則すべてここに追記すること。着手したら状態を `進行中`、完了したら `完了` に更新すること。

データは `powerpoint_masking2` とマスキング辞書（`AppData\OfficeChatMasking\rules.json` ほか）を**共有**しているが、
**マスキング機能の実装自体は共有されておらず**、Excel 側は PowerPoint 側より大幅に機能が少ない。
下表は `powerpoint_masking2`（先行実装）との差分から起票した TODO。

---

## 優先度：高（機密漏洩リスクに直結）

| # | 状態 | 種別 | 項目 | 内容・PowerPoint 側の参照 |
|---|---|---|---|---|
| H-1 | ✅ 完了 | 機能追加 | **送信前セーフティネット（MaskingSendGuard）を導入** | `OfficeMasking.Core/MaskingSendGuard.cs` を新規追加。`GeminiClient.SendAsync` 冒頭で `EnsureSafeForExternalSend(maskedText)` を呼び、平文残存を検出したら `ThisAddIn` の警告ダイアログで確認→中止可。**Ollama は生データ送信が仕様のため対象外**。参照: `powerpoint_masking2/MaskingSendGuard.cs` |
| H-2 | ✅ 完了 | 不具合修正 | **読込失敗時の Mask() 素通しを是正（フェイルセーフ）** | `MaskingEngine.Mask()` を読込失敗時に `InvalidOperationException` で停止する `EnsureAvailableForMask()` を追加。正常だが辞書空の場合は従来どおり素通し。参照: `OfficeMasking.Core/MaskingEngine.cs` |
| H-3 | ✅ 完了 | 機能追加 | **未復元プレースホルダー検出（FindUnresolvedPlaceholders / 警告付加）** | `MaskingEngine` に `FindUnresolvedPlaceholders` / `AppendUnresolvedPlaceholderWarning(ForDisplay)` を追加。**カテゴリが日本語のプレースホルダー（`__人名_1__`）に対応した正規表現**。`ChatView` の Gemini 応答表示で警告を付加。参照: `powerpoint_masking2/MaskingEngine.cs:457,480` |

## 優先度：中（マスク精度・実用性）

| # | 状態 | 種別 | 項目 | 内容・PowerPoint 側の参照 |
|---|---|---|---|---|
| M-1 | ✅ 完了 | 機能追加 | **rules.json v2（MaskingRule）へ移行** | ⚠**共有 rules.json が PowerPoint により v2 化され、Excel が読込失敗でクラッシュしたため緊急対応**。`MaskingRule`/`MaskingRuleFile` を Core へ移植し、`MaskingRulesStore`/`MaskingEngine` を v2（エントリ）ベースへ全面改修。エイリアス・意味・有効フラグ・大小文字区別に対応。互換 Dictionary API は v2 メタデータ・無効エントリを保全して更新（共有相手のデータを壊さない）。実データ102件の読込を確認。参照: `powerpoint_masking2/MaskingRule.cs`, `MaskingEngine.cs` |
| M-2 | 未着手 | 機能追加 | **意味ヒント送信（BuildMeaningHintBlock / AppendMeaningHintsForSend）** | マスク語の「意味」を機密を含めずに LLM へ渡し理解を助ける。設定で ON/OFF。Excel は未対応（v2 前提）。参照: `powerpoint_masking2/MaskingEngine.cs:389,508` |
| M-3 | 未着手 | 機能追加 | **@トークン保護マスキング（MaskExcludingAtTokens 相当）** | Excel は `@range(...)` を独自処理しているが、PowerPoint 型の「@token を退避→マスク→復元」の汎用ガードは無い。@range 解決経路の取りこぼし確認と統一。参照: `powerpoint_masking2/MaskingEngine.cs:325`, `ExcelChatAddin/ChatView.xaml.cs:67,479` |
| M-4 | ✅ 完了 | 機能追加 | **Unmask の大小文字非区別フォールバック** | M-1 の v2 化と同時に `Unmask()` へ大小文字非区別の再置換フォールバックを実装済み。参照: `OfficeMasking.Core/MaskingEngine.cs` |

## 優先度：低（LLM プロバイダ拡充）

| # | 状態 | 種別 | 項目 | 内容・PowerPoint 側の参照 |
|---|---|---|---|---|
| L-1 | 未着手 | 機能追加 | **LLM プロバイダのルーター化（LlmClientRouter / LlmProvider 相当）** | PowerPoint は Gemini / Ollama / LM Studio / Claude CLI を1つのルーターで切替。Excel は Gemini + Ollama を `ChatView` から直接呼び分けており、送信前ガードの一元化も兼ねてルーター導入が望ましい。参照: `powerpoint_masking2/LlmClientRouter.cs`, `LlmProvider.cs` |
| L-2 | 未着手 | 機能追加 | **LM Studio 対応** | Excel は未対応。参照: `powerpoint_masking2/LmStudioClient`（`LlmClientRouter.cs`） |
| L-3 | 未着手 | 機能追加 | **Claude Code CLI 対応** | Excel は未対応。hook 設定・作業ディレクトリ隔離含む。参照: `powerpoint_masking2/ClaudeCliClient`, `ClaudeHookSettingsWriter.cs`, `Paths.cs:85,91` |

## 開発環境・運用

| # | 状態 | 種別 | 項目 | 内容 |
|---|---|---|---|---|
| E-1 | 完了 | 環境整備 | **`.claude/settings.local.json` を追加** | PowerPoint 側 `.claude` を踏襲し、MSBuild ビルド・`dotnet test` の権限許可リストを Excel 用に追加。 |
| E-2 | 未着手 | ドキュメント | **docs/ 一式の整備検討** | PowerPoint は `docs/`（システム機能一覧・MethodOverview・MaskingSpecification・プロンプト一覧）を持つ。Excel にも必要になれば追加。 |
| E-3 | ✅ 完了 | デバッグ機能 | **マスキング診断フォームを追加** | リボンに「デバッグ」グループを新設し「マスキング診断」ボタンを追加。元テキスト→マスク→（Gemini送信）→アンマスクの往復を1画面で段階表示し、往復一致・送信前ガード(H-1)・未復元プレースホルダー(H-3)・Mask停止(H-2)を目視確認できる。`ExcelChatAddin/MaskingDebugForm.cs` |
