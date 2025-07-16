# Excelテンプレートファイル説明

## ファイル概要

**StockDataCollector.xlsm** - 楽天MS2RSS株価データコレクターのメインExcelファイル

このファイルは、楽天証券MarketSpeed2のRSS APIを使用して株価データを取得し、CSV形式で出力するVBAアプリケーションです。

## ⚠️ 重要：最新のVBAファイル構成

**VBAファイルは `vba/src-sjis/` フォルダから最新版をインポートしてください！**

```
✅ vba/src-sjis/ - 最新版（コンパイルエラー修正済み）⭐
❌ vba/src/     - 旧版（問題あり）
❌ vba/src-sjis-fixed/ - 中間版（使用非推奨）
```

**🔧 修正済み問題**：
- Attribute VB_Name エラー解決
- UserForm削除（InputBox方式に変更）
- 全文英語化でコンパイルエラー解決
- テストモード実装（MarketSpeed2不要でテスト可能）

## セットアップ手順

### 1. 前提条件確認

- [ ] Microsoft Excel 2016以降がインストール済み
- [ ] 楽天証券口座を保有し、MarketSpeed2の利用契約済み
- [ ] MarketSpeed2がインストール済み
- [ ] VBAマクロの実行が許可されている

### 2. ファイルの配置

1. このREADMEと同じフォルダに`StockDataCollector.xlsm`を作成
2. 以下のディレクトリ構造を確認：

```
vba/
├── templates/
│   ├── StockDataCollector.xlsm    # ← メインファイル
│   └── README.md                  # ← このファイル
├── src/                           # VBAソースコード（UTF-8版）
├── src-sjis/                      # VBAソースコード（Shift_JIS版）⭐
├── tests/                         # テスト用データ
└── ENCODING_GUIDE.md              # エンコーディングガイド
```

### 3. VBAモジュールの組み込み（重要：文字化け対策）

**必ず `src-sjis/` フォルダからインポートしてください！**

`StockDataCollector.xlsm`を作成し、以下のVBAコンポーネントを組み込みます：

#### ⭐ インポート元フォルダ
```
vba/src-sjis/  ← このフォルダからインポート（文字化けしない）
```

#### モジュール（.bas）
- **Utils.bas** - ユーティリティ・ログ機能
- **CSVExporter.bas** - CSV出力機能（簡易版）
- **DataCollector.bas** - データ取得エンジン（テストモード対応）
- **MainModule.bas** - ShowMainForm()実装（InputBox方式）
- **WorksheetMacros.bas** - ワークシートボタン用マクロ
- **SimpleTest.bas** - 基本テスト機能

#### ユーザーフォーム（.frm）
~~削除済み~~ - コンパイルエラー回避のため削除（InputBox方式に変更）

#### クラスモジュール（.cls）
~~削除済み~~ - Attribute エラー回避のため削除

### 4. 参照設定の追加

VBAエディタ（Alt+F11）で「ツール」→「参照設定」を開き、以下を有効化：

- [ ] Microsoft Office 16.0 Object Library
- [ ] Microsoft Forms 2.0 Object Library  
- [ ] Microsoft Windows Common Controls 6.0 (SP6)
- [ ] Microsoft Windows Common Controls-2 6.0 (SP6)

### 5. 詳細なインポート手順

#### Step 1: VBAエディタを開く
1. Excelで新しいブックを作成
2. `StockDataCollector.xlsm` として保存（マクロ有効ブック）
3. `Alt + F11` でVBAエディタを開く

#### Step 2: 参照設定
1. 「ツール」→「参照設定」
2. 上記の4つの参照ライブラリを有効化

#### Step 3: モジュールインポート（重要）
1. プロジェクトエクスプローラーで右クリック
2. 「ファイルのインポート」を選択
3. **`vba/src-sjis/modules/`** から以下を順番にインポート：
   ```
   Utils.bas              ⭐ 基本ユーティリティ
   CSVExporter.bas        ⭐ CSV出力機能
   DataCollector.bas      ⭐ データ取得（テストモード）
   MainModule.bas         ⭐ ShowMainForm()含む（InputBox方式）
   WorksheetMacros.bas    ⭐ ボタン用マクロ
   SimpleTest.bas         ⭐ 基本テスト機能
   ```

#### Step 4: コンパイル確認
1. VBAエディタで「デバッグ」→「VBAProjectのコンパイル」
2. エラーが表示されないことを確認

#### Step 5: 基本テスト実行
```vba
Sub Test()
    Call TestBasic
End Sub
```

**✅ 期待される結果**：
- メッセージボックス「VBA is working correctly!」が表示
- エラーが発生しない

## ワークシート構成

### Sheet1: メイン画面

```
A1: 楽天MS2RSS株価データコレクター v1.0

A3: [データ収集開始] ← StartDataCollection マクロ
A5: [クイックテスト] ← RunQuickTest マクロ  
A7: [接続テスト] ← TestConnection マクロ
A9: [設定表示] ← DisplaySettings マクロ
A11: [ヘルプ] ← ShowHelp マクロ

C3: [出力フォルダを開く] ← OpenOutputFolder マクロ
C5: [ログフォルダを開く] ← OpenLogFolder マクロ
C7: [バージョン情報] ← AboutApp マクロ
C9: [マクロ一覧] ← ShowMacroList マクロ
```

### ボタンとマクロの対応表

| ボタン名 | マクロ名 | 機能 |
|----------|----------|------|
| **データ収集開始** | `StartDataCollection` | メインGUIを表示 |
| **クイックテスト** | `RunQuickTest` | 接続・データ取得テスト |
| **接続テスト** | `TestConnection` | MarketSpeed2接続確認 |
| **設定表示** | `DisplaySettings` | 現在の設定を表示 |
| **ヘルプ** | `ShowHelp` | 使用方法を表示 |
| **出力フォルダを開く** | `OpenOutputFolder` | CSVファイル保存場所 |
| **ログフォルダを開く** | `OpenLogFolder` | ログファイル保存場所 |
| **バージョン情報** | `AboutApp` | アプリ情報表示 |

### ボタンの作成方法

1. 「開発」タブ→「挿入」→「ボタン (フォーム コントロール)」
2. ワークシート上でボタンを描画
3. 「マクロの登録」ダイアログで対応するマクロを選択
4. ボタンのテキストを設定

## 基本的な使用方法

### 1. 基本動作テスト
```vba
' VBA動作確認
Sub Test_Basic()
    Call TestBasic
End Sub
```

### 2. データ収集（InputBox方式）
```vba
' メインインターフェース表示
Sub Test_ShowMainForm()
    Call ShowMainForm
End Sub
```

### 3. クイックテスト実行
```vba
' 接続とデータ取得のテスト
Sub Test_QuickTest()
    Call QuickTest
End Sub
```

### 4. プログラムからの直接実行
```vba
Sub Test_DirectCall()
    Dim result As Boolean
    
    ' トヨタ自動車の5分足データを1週間分取得
    result = CollectStockData("7203", "5M", Date-7, Date)
    
    If result Then
        MsgBox "データ取得成功"
    Else
        MsgBox "データ取得失敗"
    End If
End Sub
```

### 5. 複数銘柄の一括取得
```vba
Sub Test_MultiplStocks()
    Dim result As Boolean
    
    ' 複数銘柄を一括取得
    result = CollectMultipleStocks("7203,6758,9984", "1M", Date-1, Date)
    
    If result Then
        MsgBox "全銘柄取得成功"
    End If
End Sub
```

## データ取得の流れ

### 1. 手動実行（推奨）

1. `StockDataCollector.xlsm`を開く
2. 「データ収集開始」ボタンをクリック
3. MainFormで設定を入力：
   - **銘柄コード**：7203,6758,9984 など
   - **取得期間**：開始日〜終了日
   - **足種**：1M, 5M, 15M, 30M, 60M, D
   - **出力先フォルダ**：CSVファイル保存場所
4. 「実行」ボタンでデータ取得開始
5. 進捗を確認し、完了まで待機
6. 指定フォルダにCSVファイルが出力される

### 2. プログラム自動実行

```vba
Sub AutoCollectData()
    ' 複数銘柄を自動取得
    Dim stocks As String
    Dim result As Boolean
    
    stocks = "7203,6758,9984,8306,9434"
    result = CollectMultipleStocks(stocks, "5M", Date-30, Date)
    
    If result Then
        Call LogMessage("INFO", "自動データ取得完了")
    End If
End Sub
```

## 主要関数リファレンス

### ShowMainForm()
メインGUIフォームを表示してデータ収集を開始

### CollectStockData(stockCode, timeFrame, startDate, endDate)
- **stockCode**: 銘柄コード（"7203", "7203.T" など）
- **timeFrame**: 足種（"1M", "5M", "15M", "30M", "60M", "D"）
- **startDate**: 開始日
- **endDate**: 終了日
- **戻り値**: Boolean（成功時True）

### CollectMultipleStocks(stockCodes, timeFrame, startDate, endDate)
複数銘柄の一括データ取得
- **stockCodes**: カンマ区切りの銘柄コード（"7203,6758,9984"）

## エラー対処法

### よくあるエラー

1. **「文字化けしています」**
   - **原因**: UTF-8版ファイルをインポートした
   - **対処**: `vba/src-sjis/` フォルダからインポートし直す

2. **「MarketSpeed2に接続できません」**
   - MarketSpeed2を起動
   - RSS機能を有効化
   - 接続状態を確認

3. **「ShowMainFormが見つかりません」**
   - `MainModule.bas` がインポートされているか確認
   - 参照設定が正しいか確認

4. **「ファイルが保存できません」**
   - 出力フォルダの書き込み権限を確認
   - ファイルが開かれていないか確認

5. **「無効な銘柄コードです」**
   - 4桁または5桁の数字で入力
   - 市場コード（.T, .JAX等）を確認

### デバッグ方法

1. **Ctrl+G** でイミディエイトウィンドウを開く
2. **F8** でステップ実行
3. **F9** でブレークポイント設定
4. ログファイル（`output/logs/`）を確認

### 接続テスト

```vba
Sub TestMS2Connection()
    Call TestConnection()
End Sub
```

## 出力データ形式

### CSV形式

```csv
DateTime,Open,High,Low,Close,Volume
2025-01-16 09:00:00,2500,2520,2495,2510,150000
2025-01-16 09:01:00,2510,2525,2505,2520,120000
```

### ファイル名形式

```
{銘柄コード}_{足種}_{開始日}-{終了日}.csv

例：
7203_5M_20250101-20250131.csv
6758_1M_20250115-20250116.csv
```

## カスタマイズ

### UIの変更

MainForm.frmを修正してインターフェースをカスタマイズ：

- ボタンの追加
- 入力項目の変更
- 表示項目の拡張

### 出力形式の変更

CSVExporter.basを修正して出力形式を変更：

- ヘッダー項目の追加
- 数値フォーマットの変更
- ファイル名規則の変更

### 新機能の追加

- データベース連携
- グラフ表示機能
- 自動実行スケジュール
- メール通知機能

## セキュリティ設定

### マクロセキュリティ

1. Excelの「ファイル」→「オプション」→「セキュリティセンター」
2. 「セキュリティセンターの設定」→「マクロの設定」
3. 「警告を表示してすべてのマクロを無効にする」または「デジタル署名されたマクロを除き、すべてのマクロを無効にする」を選択

### 信頼できる場所

1. セキュリティセンターで「信頼できる場所」を設定
2. プロジェクトフォルダを信頼できる場所に追加

## バックアップとバージョン管理

### 定期バックアップ

- 毎週ファイルをバックアップ
- 設定ファイル（config/settings.json）も保存
- 出力CSVファイルのアーカイブ

### バージョン管理

```vba
' ファイルバージョン管理
Const APP_VERSION As String = "1.0.0"
Const BUILD_DATE As String = "2025-01-16"

Sub ShowVersion()
    MsgBox "楽天MS2RSS株価データコレクター" & vbCrLf & _
           "バージョン: " & APP_VERSION & vbCrLf & _
           "ビルド日: " & BUILD_DATE
End Sub
```

## サポート・問い合わせ

### ドキュメント

- **VBA詳細ガイド**: `docs/vba-guide.md`
- **エンコーディングガイド**: `vba/ENCODING_GUIDE.md`
- **API仕様**: `docs/ms2rss/function-reference.md`

### 問い合わせ

技術的な問題やご質問は、GitHubのIssuesページまでお願いします：

**https://github.com/tsubokazu/rakuten-ms2rss-collector/issues**

### よくある質問

1. **Q: 文字化けします**
   - A: `vba/src-sjis/` フォルダからインポートしてください

2. **Q: ShowMainFormが動きません**
   - A: `MainModule.bas` をインポートしてください

3. **Q: データが取得できません**
   - A: MarketSpeed2のRSS機能を確認してください

## ライセンス

MIT License - 詳細はプロジェクトルートのLICENSEファイルを参照

## 免責事項

このソフトウェアは教育・研究目的で提供されています。投資判断や取引結果について、開発者は一切の責任を負いません。ご自身の責任でご利用ください。