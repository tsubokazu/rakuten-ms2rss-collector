# VBAソースコード - インポート手順

## 概要

このフォルダには、楽天MS2RSS株価データコレクターのすべてのVBAソースコードが含まれています。

## ファイル構成

### 📁 modules/ - VBAモジュール
| ファイル名 | 説明 | 主要関数 |
|------------|------|----------|
| **MainModule.bas** | メインエントリーポイント | `ShowMainForm()`, `QuickTest()` |
| **WorksheetMacros.bas** | ワークシートボタン用マクロ | `StartDataCollection()` など |
| **DataCollector.bas** | データ取得エンジン | `CollectStockData()` |
| **CSVExporter.bas** | CSV出力機能 | `ExportStockDataToCSV()` |
| **Utils.bas** | ユーティリティ・ログ | `LogMessage()`, `ValidateTimeFrame()` |

### 📁 forms/ - ユーザーフォーム
| ファイル名 | 説明 |
|------------|------|
| **MainForm.frm** | メインGUIフォーム |

### 📁 classes/ - クラスモジュール
| ファイル名 | 説明 |
|------------|------|
| **StockData.cls** | 株価データ構造クラス |
| **Configuration.cls** | 設定管理クラス |

## Excelへのインポート手順

### 1. 新しいExcelファイルを作成
1. Microsoft Excelを起動
2. 新しいブックを作成
3. ファイル名を`StockDataCollector.xlsm`として保存（マクロ有効ブック形式）

### 2. VBAエディタを開く
1. `Alt + F11`を押してVBAエディタを開く
2. プロジェクトエクスプローラーでVBAProjectを確認

### 3. 参照設定を追加
1. VBAエディタで「ツール」→「参照設定」を選択
2. 以下の項目にチェックを入れる：
   - ✅ Microsoft Office 16.0 Object Library
   - ✅ Microsoft Forms 2.0 Object Library
   - ✅ Microsoft Windows Common Controls 6.0 (SP6)
   - ✅ Microsoft Windows Common Controls-2 6.0 (SP6)

### 4. モジュールをインポート

#### 標準モジュール (.bas)
1. プロジェクトエクスプローラーで右クリック
2. 「ファイルのインポート」を選択
3. 以下のファイルを順番にインポート：
   ```
   modules/MainModule.bas
   modules/WorksheetMacros.bas
   modules/DataCollector.bas
   modules/CSVExporter.bas
   modules/Utils.bas
   ```

#### ユーザーフォーム (.frm)
1. プロジェクトエクスプローラーで右クリック
2. 「ファイルのインポート」を選択
3. `forms/MainForm.frm`をインポート

#### クラスモジュール (.cls)
1. プロジェクトエクスプローラーで右クリック
2. 「ファイルのインポート」を選択
3. 以下のファイルをインポート：
   ```
   classes/StockData.cls
   classes/Configuration.cls
   ```

### 5. ワークシートの設定

#### Sheet1の設定
1. Sheet1を選択し、以下のように設定：

```
A1: 楽天MS2RSS株価データコレクター v1.0
A3: [データ収集開始] (ボタン)
A5: [クイックテスト] (ボタン)
A7: [接続テスト] (ボタン)
A9: [設定表示] (ボタン)
A11: [ヘルプ] (ボタン)

C3: [出力フォルダを開く] (ボタン)
C5: [ログフォルダを開く] (ボタン)
C7: [バージョン情報] (ボタン)
```

#### ボタンのマクロ割り当て
各ボタンに以下のマクロを割り当て：

| ボタン名 | マクロ名 |
|----------|----------|
| データ収集開始 | `StartDataCollection` |
| クイックテスト | `RunQuickTest` |
| 接続テスト | `TestConnection` |
| 設定表示 | `DisplaySettings` |
| ヘルプ | `ShowHelp` |
| 出力フォルダを開く | `OpenOutputFolder` |
| ログフォルダを開く | `OpenLogFolder` |
| バージョン情報 | `AboutApp` |

## 基本的な使用方法

### 1. アプリケーション起動
```vba
' メインフォームを表示
Sub Test_ShowMainForm()
    Call ShowMainForm
End Sub
```

### 2. クイックテスト実行
```vba
' 接続とデータ取得のテスト
Sub Test_QuickTest()
    Call QuickTest
End Sub
```

### 3. プログラムからの直接実行
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

## トラブルシューティング

### よくあるエラー

1. **「プロシージャが見つかりません」**
   - モジュールが正しくインポートされているか確認
   - 参照設定が正しく設定されているか確認

2. **「RSS関数がエラーを返します」**
   - MarketSpeed2が起動しているか確認
   - RSS機能が有効になっているか確認

3. **「ファイルが保存できません」**
   - 出力フォルダが存在するか確認
   - フォルダの書き込み権限を確認

### デバッグ方法

1. **ステップ実行**: F8キーで行単位実行
2. **ブレークポイント**: F9キーで設定
3. **イミディエイトウィンドウ**: Ctrl+Gで表示
4. **ログ確認**: `output/logs/`フォルダのログファイル

## 注意事項

- マクロのセキュリティ設定で、マクロの実行を許可してください
- MarketSpeed2のRSS機能が有効になっている必要があります
- 大量データ取得時は処理時間がかかる場合があります
- 本番環境での使用前に十分なテストを実施してください

## カスタマイズ

### 新しい足種の追加
`Utils.bas`の`ValidateTimeFrame`関数を修正

### 新しい市場の追加
`DataCollector.bas`の`ValidateStockCode`関数を修正

### UI表示項目の変更
`MainForm.frm`のデザインを修正

詳細なカスタマイズ方法は、`docs/vba-guide.md`を参照してください。