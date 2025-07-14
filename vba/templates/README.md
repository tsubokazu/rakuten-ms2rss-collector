# Excelテンプレートファイル説明

## ファイル概要

**StockDataCollector.xlsm** - 楽天MS2RSS株価データコレクターのメインExcelファイル

このファイルは、楽天証券MarketSpeed2のRSS APIを使用して株価データを取得し、CSV形式で出力するVBAアプリケーションです。

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
├── src/                           # VBAソースコード
└── tests/                         # テスト用データ
```

### 3. VBAモジュールの組み込み

`StockDataCollector.xlsm`を作成し、以下のVBAコンポーネントを組み込みます：

#### モジュール（.bas）
- **DataCollector.bas** - データ取得エンジン
- **CSVExporter.bas** - CSV出力機能
- **Utils.bas** - ユーティリティ・ログ機能

#### ユーザーフォーム（.frm）
- **MainForm.frm** - メインユーザーインターフェース

#### クラスモジュール（.cls）
- **StockData.cls** - 株価データ構造
- **Configuration.cls** - 設定管理

### 4. 参照設定の追加

VBAエディタ（Alt+F11）で「ツール」→「参照設定」を開き、以下を有効化：

- [ ] Microsoft Office 16.0 Object Library
- [ ] Microsoft Forms 2.0 Object Library  
- [ ] Microsoft Scripting Runtime

## ワークシート構成

### Sheet1: メイン画面

- **A1**: タイトル "楽天MS2RSS株価データコレクター"
- **A3**: "データ収集開始" ボタン（MainFormを表示）
- **A5-A10**: 簡単な使用方法説明
- **A12-A20**: 最近の実行ログ表示エリア

### Sheet2: 設定画面

- **A1-B10**: 基本設定項目
  - B1: デフォルト足種
  - B2: デフォルト出力パス
  - B3: ログレベル
  - B4: 最大取得本数
  - B5: CSV小数点桁数

### Sheet3: テスト用

- **A1-F100**: テスト用データ表示エリア
- RSS関数のテスト実行用

## ボタンとマクロ

### メインボタン

1. **データ収集開始** (`ShowMainForm`)
   - MainFormを表示してGUIによるデータ収集を開始

2. **設定画面** (`ShowConfigForm`)
   - 設定画面を表示（将来実装予定）

3. **ログ表示** (`ShowLogViewer`)
   - ログファイルの内容を表示

4. **テスト実行** (`RunTest`)
   - 接続テストとサンプルデータ取得

### サンプルマクロ

```vba
' メインフォーム表示
Sub ShowMainForm()
    MainForm.Show
End Sub

' 簡単テスト
Sub QuickTest()
    Dim result As Boolean
    result = CollectStockData("7203", "5M", Date-1, Date)
    If result Then
        MsgBox "テスト成功"
    Else
        MsgBox "テスト失敗"
    End If
End Sub

' 設定表示
Sub ShowCurrentConfig()
    Dim config As Configuration
    Set config = New Configuration
    config.LoadFromFile
    MsgBox config.ToString()
    Set config = Nothing
End Sub
```

## データ取得の流れ

### 1. 手動実行

1. `StockDataCollector.xlsm`を開く
2. 「データ収集開始」ボタンをクリック
3. MainFormで設定を入力：
   - 銘柄コード（例：7203,6758,9984）
   - 取得期間（開始日〜終了日）
   - 足種（1M, 5M, 15M, 30M, 60M, D）
   - 出力先フォルダ
4. 「実行」ボタンでデータ取得開始
5. 進捗を確認し、完了まで待機
6. 指定フォルダにCSVファイルが出力される

### 2. プログラム実行

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

## エラー対処法

### よくあるエラー

1. **「MarketSpeed2に接続できません」**
   - MarketSpeed2を起動
   - RSS機能を有効化
   - 接続状態を確認

2. **「ファイルが保存できません」**
   - 出力フォルダの書き込み権限を確認
   - ファイルが開かれていないか確認

3. **「無効な銘柄コードです」**
   - 4桁または5桁の数字で入力
   - 市場コード（.T, .JAX等）を確認

### デバッグ方法

1. **Ctrl+G** でイミディエイトウィンドウを開く
2. **F8** でステップ実行
3. **F9** でブレークポイント設定
4. ログファイル（`output/logs/`）を確認

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
Const BUILD_DATE As String = "2025-01-14"

Sub ShowVersion()
    MsgBox "楽天MS2RSS株価データコレクター" & vbCrLf & _
           "バージョン: " & APP_VERSION & vbCrLf & _
           "ビルド日: " & BUILD_DATE
End Sub
```

## サポート・問い合わせ

技術的な問題やご質問は、GitHubのIssuesページまでお願いします：

https://github.com/tsubokazu/rakuten-ms2rss-collector/issues

## ライセンス

MIT License - 詳細はプロジェクトルートのLICENSEファイルを参照