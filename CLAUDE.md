# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## module_list.md の維持

`module_list.md` はプロジェクト全体のモジュール一覧ファイル。以下のいずれかを行ったときは必ず更新すること。

- ファイルを**新規追加**した → 適切なセクションに行を追加
- ファイルを**削除**した → 該当行を削除
- ファイルの**役割が変わった**（責務追加・移管・リネーム等） → 役割の説明を更新

コミット前に `module_list.md` の更新が漏れていないか確認すること。

---

## テスト駆動開発（TDD）

`OfficeMasking.Core` に対する変更は TDD で進めること。

1. **Red** — 失敗するテストを先に書く（`OfficeMasking.Core.Tests/` に追加）
2. **Green** — テストが通る最小限の実装をする
3. **Refactor** — テストを壊さずコードを整理する

```powershell
# 変更しながらテストを回す
dotnet test OfficeMasking.Core.Tests\OfficeMasking.Core.Tests.csproj --logger "console;verbosity=normal"
```

`ExcelChatAddin` 本体（VSTO側）は Excel 依存のため自動テストは書けない。ロジックを `OfficeMasking.Core` に切り出してテストする設計を優先すること。

---

## ビルド

このプロジェクトは VSTO (Visual Studio Tools for Office) アドインのため、**`dotnet` CLI ではビルドできない**。Visual Studio 2022 Insiders の MSBuild を使うこと。

```powershell
# ExcelChatAddin（メインプロジェクト）
& "C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\MSBuild.exe" `
  "ExcelChatAddin\ExcelChatAddin.csproj" /p:Configuration=Debug /nologo /v:minimal

# OfficeMasking.Core（共通ライブラリ）は dotnet でビルド可
dotnet build OfficeMasking.Core\OfficeMasking.Core.csproj
```

## テスト

```powershell
# 全テスト実行
dotnet test OfficeMasking.Core.Tests\OfficeMasking.Core.Tests.csproj

# 単一テストクラスを指定
dotnet test OfficeMasking.Core.Tests\OfficeMasking.Core.Tests.csproj --filter "ClassName=MaskingEngineTests"

# 単一テストメソッドを指定
dotnet test OfficeMasking.Core.Tests\OfficeMasking.Core.Tests.csproj --filter "FullyQualifiedName~MaskingEngineTests.Mask_"
```

テストは MSTest (v3.1.1)。テスト対象は `OfficeMasking.Core` のみ。`ExcelChatAddin` プロジェクト自体に自動テストはない。

## 署名証明書の再生成

`ExcelChatAddin_TemporaryKey.pfx` は `.gitignore` で除外されており、環境ごとに再生成が必要。ビルドエラー `MSB3482` が出たら以下を実行：

```powershell
$thumb = (New-SelfSignedCertificate -Subject "CN=ExcelChatAddin_TemporaryKey" `
  -CertStoreLocation "Cert:\CurrentUser\My" -KeySpec Signature `
  -NotAfter (Get-Date).AddYears(10)).Thumbprint

$pwd = New-Object System.Security.SecureString
Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$thumb" `
  -FilePath "ExcelChatAddin\ExcelChatAddin_TemporaryKey.pfx" -Password $pwd -Force
```

その後 `ExcelChatAddin.csproj` の `<ManifestCertificateThumbprint>` を新しいサムプリントに更新する。

## アーキテクチャ概要

### プロジェクト構成

```
ExcelChatAddin/          ← VSTO Excel アドイン (.NET Framework 4.8)
OfficeMasking.Core/      ← マスキング共通ライブラリ (net48, SDK スタイル)
OfficeMasking.Core.Tests/← MSTest テストプロジェクト (net48)
```

`ExcelChatAddin` は `OfficeMasking.Core` を参照する。逆方向の依存はない。

### ExcelChatAddin の主要な責務分担

**エントリーポイント**  
`ThisAddIn.cs` が Startup/Shutdown を管理し、右クリックメニュー（CommandBar "Cell"）の追加・削除、ホットキー（Ctrl+Alt+M）の登録、各ダイアログの起動を一手に担う。`ChatRibbon.cs` はリボンUIの定義。

**チャット機能**  
`ChatCoordinator` がマスキング → Gemini送信 → 履歴保存のパイプラインを調整する。`GeminiClient`（シングルトン）が実際のAPI通信を担当。`ChatView.xaml` + `TaskPaneHost` でカスタムタスクペインとして表示される。`@range(シート名, A1形式)` トークンを介してセル範囲をチャットに挿入できる。

**テーブルスキーマ機能**  
`IssueSchemaConfig` / `IssueSchemaManager` がテーブル定義（列名・型・主キー・更新モード等）を管理し `table_schema.json` に永続化。`IssueSchemaSettingsDialog` で編集。`SchemaTemplateManager` でテンプレートとして保存・再利用できる。

**テーブル関係機能**  
`TableRelationConfig.cs` に `RelationTypeMasterItem`（関係種別マスタ）、`TableRelationRule`（テーブル間許可ルール）、`TableRecordRelation`（レコード間の実関係）の3階層のモデルを定義。`TableRelationManager` が `table_relations.json` への読み書きを担当。`TableRelationMatrixForm` がマトリクスUIを提供（VirtualMode の DataGridView で大量データに対応）。

**マスキング機能（OfficeMasking.Core）**  
`MaskingEngine`（シングルトン）が辞書に基づく文字列置換を実施。プレースホルダー形式は `__カテゴリ_連番__`。`MaskingRulesStore` が `rules.json` の読み書き（50世代バックアップ付き）を担う。`MaskingPaths` がデータ保存先を決定する（優先順：環境変数 `OFFICE_MASKING_DATA_DIR` > `AppData\OfficeChatMasking`）。

### 永続データの保存場所

すべてのデータは `MaskingPaths.DataDir` 配下に JSON で保存される。

| ファイル | 内容 |
|---|---|
| `rules.json` | マスキング辞書（原文→プレースホルダーのマップ） |
| `categories.txt` | マスキングカテゴリ一覧 |
| `config.json` | アドイン設定（Geminiモデル名等） |
| `table_schema.json` | テーブルスキーマ定義（複数テーブル） |
| `table_relations.json` | テーブル間・レコード間の関係定義 |
| `diagram_templates.json` | ダイアグラムテンプレート |
| `schema_templates.json` | スキーマテンプレート |

### UI技術の混在

`ExcelChatAddin` は WinForms と WPF が混在している。`TaskPaneHost`（WinForms UserControl）内に `ElementHost` を使って WPF の `ChatView.xaml` を埋め込む構造。ダイアログは WinForms が多いが、`MaskPreviewWindow`・`GeminiResponseWindow` は WPF。
