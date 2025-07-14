# VBA使用方法ガイド

## 概要

楽天MS2RSS株価データコレクターのVBA実装に関する詳細ガイドです。このドキュメントでは、VBAコードの構造、使用方法、カスタマイズ方法について説明します。

## 前提条件

### 必要環境

- **Microsoft Excel 2016以降**（VBA対応版）
- **楽天証券口座**（MarketSpeed2利用契約済み）
- **MarketSpeed2**（RSS機能有効）
- **Windows OS**（Mac版Excelでは一部機能制限あり）

### MarketSpeed2設定

1. MarketSpeed2を起動
2. 「設定」→「RSS設定」を開く
3. 「RSS機能を有効にする」をチェック
4. 接続テストを実行して正常性を確認

## インストール手順

### 1. ファイルの配置

```
rakuten-ms2rss-collector/
├── vba/
│   ├── templates/
│   │   └── StockDataCollector.xlsm    # メインExcelファイル
│   └── src/                           # VBAソースコード
├── output/                            # 出力ディレクトリ
└── config/                            # 設定ディレクトリ
```

### 2. VBAプロジェクトのインポート

1. `StockDataCollector.xlsm`を開く
2. `Alt + F11`でVBAエディタを開く
3. 「ファイル」→「ファイルのインポート」で各モジュールをインポート：
   - `DataCollector.bas`
   - `CSVExporter.bas`  
   - `Utils.bas`
   - `MainForm.frm`
   - `StockData.cls`
   - `Configuration.cls`

### 3. 参照設定

VBAエディタで「ツール」→「参照設定」を開き、以下を有効化：

- Microsoft Office 16.0 Object Library
- Microsoft Forms 2.0 Object Library
- Microsoft Scripting Runtime

## 基本的な使用方法

### 1. メインフォームの起動

```vba
Sub ShowMainForm()
    MainForm.Show
End Sub
```

### 2. プログラムからの直接実行

```vba
Sub CollectSingleStock()
    Dim result As Boolean
    
    ' 単一銘柄のデータ取得
    result = CollectStockData("7203", "5M", #1/1/2025#, #1/31/2025#)
    
    If result Then
        MsgBox "データ取得完了"
    Else
        MsgBox "データ取得失敗"
    End If
End Sub
```

### 3. 複数銘柄の一括処理

```vba
Sub CollectMultipleStocks()
    Dim result As Boolean
    
    ' 複数銘柄のデータ取得
    result = CollectMultipleStocks("7203,6758,9984", "1M", #1/1/2025#, #1/31/2025#)
    
    If result Then
        MsgBox "全銘柄のデータ取得完了"
    Else
        MsgBox "一部または全ての銘柄でデータ取得失敗"
    End If
End Sub
```

## 高度な使用方法

### 1. StockDataクラスの活用

```vba
Sub UseStockDataClass()
    Dim stockData As StockData
    Set stockData = New StockData
    
    ' 設定
    stockData.StockCode = "7203.T"
    stockData.StockName = "トヨタ自動車"
    stockData.TimeFrame = "5M"
    stockData.StartDate = #1/1/2025#
    stockData.EndDate = #1/31/2025#
    
    ' データの妥当性チェック
    If stockData.IsValid() Then
        ' データ取得処理
        Debug.Print stockData.GetStatistics()
    End If
    
    Set stockData = Nothing
End Sub
```

### 2. カスタム設定の利用

```vba
Sub CustomConfiguration()
    Dim config As Configuration
    Set config = New Configuration
    
    ' 設定の変更
    config.MaxBarsPerRequest = 2000
    config.DefaultTimeFrame = "1M"
    config.DecimalPlaces = 3
    
    ' 設定の保存
    config.SaveToFile
    
    ' 設定の表示
    Debug.Print config.ToString()
    
    Set config = Nothing
End Sub
```

### 3. CSV出力のカスタマイズ

```vba
Sub CustomCSVExport()
    Dim data(1 To 3, 1 To 6) As Variant
    
    ' サンプルデータ
    data(1, 1) = "2025-01-14 09:00:00": data(1, 2) = 2500: data(1, 3) = 2520
    data(1, 4) = 2495: data(1, 5) = 2510: data(1, 6) = 150000
    
    ' CSV設定のカスタマイズ
    Call SetCSVConfig(True, "YYYY-MM-DD HH:MM:SS", 3, ",")
    
    ' CSV出力
    If ExportStockDataToCSV(data, "C:\temp\custom_data.csv") Then
        MsgBox "カスタムCSV出力完了"
    End If
End Sub
```

## エラーハンドリング

### 1. 一般的なエラーパターン

```vba
Sub ErrorHandlingExample()
    On Error GoTo ErrorHandler
    
    ' データ取得処理
    Call CollectStockData("INVALID", "5M", Date, Date)
    
    Exit Sub
    
ErrorHandler:
    ' 詳細エラーログ
    Call LogDetailedError("ErrorHandlingExample", Err.Description, "銘柄コード: INVALID")
    
    ' ユーザー向けエラーメッセージ
    MsgBox "データ取得でエラーが発生しました。ログを確認してください。", vbCritical
End Sub
```

### 2. MarketSpeed2接続エラー

```vba
Sub CheckMS2Connection()
    On Error GoTo ConnectionError
    
    ' 簡易接続テスト
    Dim testResult As Variant
    testResult = Application.WorksheetFunction.RssMarket("7203", "現在値")
    
    If IsError(testResult) Then
        MsgBox "MarketSpeed2に接続できません。RSS機能を確認してください。", vbCritical
    Else
        MsgBox "MarketSpeed2接続正常: " & testResult, vbInformation
    End If
    
    Exit Sub
    
ConnectionError:
    MsgBox "MarketSpeed2接続エラー: " & Err.Description, vbCritical
End Sub
```

## カスタマイズ

### 1. 対応足種の追加

`Utils.bas`の`ValidateTimeFrame`関数を修正：

```vba
Public Function ValidateTimeFrame(timeFrame As String) As Boolean
    Dim validFrames As Variant
    
    ' 新しい足種を追加
    validFrames = Array("T", "1M", "2M", "3M", "4M", "5M", "10M", "15M", "30M", "60M", _
                       "2H", "4H", "8H", "12H", "D", "W", "M")  ' 12H追加
    
    ' 以下同じ...
End Function
```

### 2. 出力フォーマットの変更

`CSVExporter.bas`の`GenerateCSVHeader`を修正：

```vba
Private Function GenerateCSVHeader() As String
    ' カスタムヘッダー
    GenerateCSVHeader = "Date,Time,Open,High,Low,Close,Volume,VWAP"
End Function
```

### 3. 新しい市場の追加

`DataCollector.bas`の`ValidateStockCode`を修正：

```vba
Select Case UCase(marketPart)
    Case "T", "JAX", "JNX", "CHJ", "NEW_MARKET"  ' 新市場追加
        ' 有効な市場コード
    Case Else
        ValidateStockCode = False
        Exit Function
End Select
```

## パフォーマンス最適化

### 1. 大量データ処理

```vba
Sub OptimizedBatchProcessing()
    ' アプリケーション設定の最適化
    Call OptimizeApplicationSettings()
    
    Try
        ' バッチ処理
        Call CollectMultipleStocks("大量の銘柄リスト", "1M", startDate, endDate)
    Finally
        ' 設定復元
        Call RestoreApplicationSettings()
    End Try
End Sub
```

### 2. メモリ使用量の監視

```vba
Sub MonitorMemoryUsage()
    Debug.Print "開始時: " & GetMemoryUsage()
    
    ' データ処理
    Call CollectStockData("7203", "1M", Date - 30, Date)
    
    Debug.Print "終了時: " & GetMemoryUsage()
End Sub
```

## トラブルシューティング

### よくある問題と解決方法

| 問題 | 原因 | 解決方法 |
|------|------|----------|
| RSS関数エラー | MarketSpeed2未接続 | MS2を起動し、RSS機能を有効化 |
| ファイル保存エラー | フォルダが存在しない | `EnsureDirectoryExists`で作成 |
| メモリ不足 | 大量データ処理 | バッチサイズを小さくする |
| 日付形式エラー | 地域設定の違い | 明示的な日付フォーマット使用 |

### デバッグ方法

1. **ログ確認**: `output/logs/`フォルダのログファイル
2. **イミディエイトウィンドウ**: `Ctrl + G`でデバッグ出力確認
3. **ブレークポイント**: VBAエディタでの段階実行
4. **エラー詳細**: `LogDetailedError`関数の活用

## サンプルコード集

### 完全な実行例

```vba
Sub CompleteExample()
    ' 設定読み込み
    Dim config As Configuration
    Set config = New Configuration
    config.LoadFromFile
    
    ' データオブジェクト作成
    Dim stockData As StockData
    Set stockData = New StockData
    stockData.StockCode = "7203"
    stockData.TimeFrame = "5M"
    stockData.StartDate = DateAdd("d", -7, Date)
    stockData.EndDate = Date
    
    ' データ取得
    If stockData.IsValid() Then
        If CollectStockData(stockData.GetFullStockCode(), stockData.TimeFrame, _
                          stockData.StartDate, stockData.EndDate) Then
            
            ' 統計表示
            Debug.Print stockData.GetStatistics()
            
            ' JSON出力
            Dim jsonFile As String
            jsonFile = config.DefaultOutputPath & "stats.json"
            
            Open jsonFile For Output As #1
            Print #1, stockData.ToJSON()
            Close #1
            
            MsgBox "完了: " & jsonFile
        End If
    End If
    
    ' クリーンアップ
    Set stockData = Nothing
    Set config = Nothing
End Sub
```

## API リファレンス

### 主要関数

| 関数名 | 説明 | 戻り値 |
|--------|------|--------|
| `CollectStockData` | 単一銘柄データ取得 | Boolean |
| `CollectMultipleStocks` | 複数銘柄データ取得 | Boolean |
| `ExportStockDataToCSV` | CSV出力 | Boolean |
| `ValidateStockCode` | 銘柄コード検証 | Boolean |
| `LogMessage` | ログ出力 | なし |

詳細な仕様は各モジュールのコメントを参照してください。