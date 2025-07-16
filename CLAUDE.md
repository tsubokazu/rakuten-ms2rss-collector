# Stock AI Trade - Rakuten MS2RSS株式データ収集システム

## プロジェクト概要

楽天証券MarketSpeed2のRSS API機能を使用した株式データ収集システム。VBAベースのExcelアドインとして実装されており、株式チャートデータをCSV形式で出力する。

**現在の状態**: テストモード - MarketSpeed2なしでサンプルデータを生成

## 開発コマンド

```bash
# Python環境（PDF処理ツール用）
uv sync                    # 依存関係インストール
uv run python -m tools.pdf_reader  # PDF処理実行

# プロジェクト管理
ls vba/src-sjis/          # VBAソースコード確認
ls output/csv/            # CSV出力ファイル確認
ls output/logs/           # ログファイル確認
```

## アーキテクチャ

### ディレクトリ構造
```
stock-ai-trade/
├── vba/
│   ├── src-sjis/         # VBA源码（推奨・修正版）
│   ├── src-sjis-fixed/   # エラー回避版
│   └── COMPILE_ERROR_FIX.md
├── docs/ms2rss/          # RSS API文档
├── tools/pdf-reader/     # Python PDF処理
├── output/
│   ├── csv/             # データ出力
│   └── logs/            # ログファイル
└── pyproject.toml       # Python設定
```

### VBAモジュール構成

| モジュール | 機能 | 重要な関数 |
|-----------|------|-----------|
| **DataCollector.bas** | データ収集エンジン | `CollectStockData()`, `CollectMultipleStocks()` |
| **MainModule.bas** | メインエントリ | `ShowMainForm()`, `QuickTest()` |
| **WorksheetMacros.bas** | ワークシート用マクロ | `StartDataCollection()` |
| **CSVExporter.bas** | CSV出力機能 | `ExportStockDataToCSV()` |
| **Utils.bas** | ユーティリティ | `LogMessage()`, `ValidateStockCode()` |

### データフロー
1. **入力**: 銘柄コード, 時間軸, 期間指定
2. **検証**: `ValidateStockCode()` で銘柄コード形式チェック
3. **データ生成**: `CreateSampleCSVFileWithDateRange()` でテストデータ作成
4. **出力**: `output/csv/` フォルダにCSVファイル保存

### RSS API統合アーキテクチャ

**本番環境での使用**: MarketSpeed2とRSS Chart関数（`RssChart`, `RssChartPast`）
**現在**: サンプルデータ生成でテスト実行

対応時間軸:
- **分足**: 1M, 5M, 15M, 30M, 60M
- **日足**: D
- **対応市場**: T(東証), JAX, JNX, CHJ

## 重要な設計決定

### 1. コンパイルエラー解決
- `Attribute VB_Name` 行を全削除（VBAエディタが自動生成）
- UserFormとクラスモジュールを除去してInputBox方式に変更
- エラー回避版（`src-sjis-fixed/`）を提供

### 2. テストモード実装
- MarketSpeed2なしでの動作確認を可能に
- リアルな日付範囲でのサンプルデータ生成
- 本番環境への切り替えは最小限の変更で対応

### 3. 文字エンコーディング対応
- Shift_JIS形式のVBAソースコード
- UTF-8/Shift_JIS変換問題を解決

## 主要機能

### データ収集関数
```vba
' 単一銘柄データ収集
CollectStockData(stockCode, timeFrame, startDate, endDate)

' 複数銘柄一括処理
CollectMultipleStocks("7203,6758,9984", "5M", Date-7, Date)
```

### 設定可能パラメータ
- **銘柄コード**: "7203", "7203.T"形式
- **時間軸**: "1M", "5M", "15M", "30M", "60M", "D"
- **期間**: 開始日〜終了日の日付範囲

## Excel設定手順

### VBAインポート順序
```
1. Utils.bas          # ユーティリティ（最初）
2. CSVExporter.bas    # 出力機能
3. DataCollector.bas  # データ収集
4. MainModule.bas     # メイン機能
5. WorksheetMacros.bas # ボタン用マクロ
```

### 参照設定
- Microsoft Office Object Library
- Microsoft Forms 2.0 Object Library

## テスト・デバッグ

### 基本テスト手順
```vba
' 1. 基本機能テスト
Call TestBasic

' 2. メインフォーム表示
Call ShowMainForm

' 3. クイックテスト
Call QuickTest

' 4. 直接データ収集テスト
result = CollectStockData("7203", "5M", Date-7, Date)
```

### デバッグ方法
- F8: ステップ実行
- F9: ブレークポイント設定
- Ctrl+G: イミディエイトウィンドウ
- ログファイル: `output/logs/` で確認

## 本番環境への移行

### 必要な準備
1. MarketSpeed2のインストールと設定
2. RSS Chart Add-inの有効化
3. テストモードからライブモードへの切り替え
4. 十分なテスト実行

### 制限事項
- チャート関数は最大3000本のデータ取得制限
- 立会時間外では一部機能制限
- VBA関数での引数数制限

## Python連携

### PDF処理ツール
```bash
cd tools/pdf-reader
uv run python main.py
```

**用途**: MarketSpeed2関連PDFドキュメントの解析・処理

## 実装時の注意事項

### 1. VBAコード作成・修正時の必須事項
- **絶対に`Attribute VB_Name`行を含めない** - VBAエディタが自動生成するため、手動追加するとコンパイルエラーが発生
- **UserFormや複雑なクラスモジュールは避ける** - 構文エラーの原因となるため、InputBox方式を使用
- **文字エンコーディングはShift_JISで統一** - UTF-8で保存するとVBAエディタで文字化けが発生

### 2. ディレクトリ構造の維持
- **出力ファイルは必ず`output/`以下に配置** - CSV: `output/csv/`, ログ: `output/logs/`
- **VBAソースコードの格納場所を変更しない** - `vba/src-sjis/`が基本、エラー時は`vba/src-sjis-fixed/`を使用

### 3. テストモードの重要性
- **MarketSpeed2なしでの動作確認を必ず実施** - 本番環境設定前に基本機能をテスト
- **サンプルデータ生成機能を維持** - `CreateSampleCSVFileWithDateRange()`でリアルな日付範囲データを生成
- **テストモードと本番モードの切り替えは最小限の変更で対応** - 大幅な構造変更は避ける

### 4. エラーハンドリングとデバッグ
- **コンパイルエラーが発生した場合は`src-sjis-fixed/`を使用** - 修正版で基本動作確認後、段階的に機能追加
- **モジュールインポートは必ず指定順序で実行** - Utils.bas → CSVExporter.bas → DataCollector.bas → MainModule.bas → WorksheetMacros.bas
- **Debug.Print文を活用** - ログファイルと併用してデバッグ情報を出力

### 5. 本番環境移行時の注意
- **MarketSpeed2のRSS Chart機能確認** - `RssChart`, `RssChartPast`関数の動作テストを実施
- **データ制限を考慮した設計** - 最大3000本の制限に対応したバッチ処理を実装
- **エラー処理の充実** - 接続エラー、データ取得エラーに対する適切なハンドリング

### 6. 過去の修正事項（重要）
- **文字化けの解決**: README.mdの一部で文字エンコーディング問題が発生済み - 新規作成時は注意
- **構文エラーの解決**: Attribute行除去により解決済み - 再発防止のため手動追加禁止
- **UI簡素化**: 複雑なUserFormを削除してInputBox方式に変更済み - 元に戻さない

## 重要なファイル

- `vba/src-sjis/README.md`: VBAインポート手順とトラブルシューティング
- `vba/COMPILE_ERROR_FIX.md`: コンパイルエラー対処法詳細
- `docs/ms2rss/function-reference.md`: RSS API関数リファレンス
- `vba/src-sjis/modules/DataCollector.bas`: コアデータ収集ロジック