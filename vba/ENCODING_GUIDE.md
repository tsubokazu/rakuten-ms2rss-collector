# VBAファイル文字エンコーディング ガイド

## 文字化け対策

VBAファイルをExcelにインポートする際に文字化けが発生する場合は、**Shift_JIS版のファイル**を使用してください。

## ファイル構成

### 📁 src/ - UTF-8版（開発・閲覧用）
- GitHub上での表示・編集に適している
- 現代的なエディタでの編集に適している
- 日本語コメントが正しく表示される

### 📁 src-sjis/ - Shift_JIS版（Excelインポート用）
- **VBAエディタへのインポート専用**
- 文字化けしない日本語表示
- Excelが期待する文字エンコーディング

## 使い分け

| 用途 | 使用フォルダ | 理由 |
|------|-------------|------|
| **Excelへのインポート** | `src-sjis/` | 文字化け防止 |
| GitHub上での閲覧 | `src/` | UTF-8表示 |
| エディタでの編集 | `src/` | 現代的なエンコーディング |
| 配布・共有 | `src-sjis/` | 受信者の環境に依存しない |

## インポート手順（文字化け対策版）

### 1. 準備
```
rakuten-ms2rss-collector/
├── vba/
│   ├── src/           ← UTF-8版（GitHub表示用）
│   └── src-sjis/      ← Shift_JIS版（インポート用）⭐
```

### 2. VBAエディタでのインポート

1. **Alt+F11** でVBAエディタを開く
2. **「ファイル」→「ファイルのインポート」**
3. **`vba/src-sjis/`フォルダから** 以下のファイルを選択：

#### 標準モジュール (.bas)
```
src-sjis/modules/MainModule.bas         ⭐ ShowMainForm()
src-sjis/modules/WorksheetMacros.bas    ⭐ ボタン用マクロ
src-sjis/modules/DataCollector.bas
src-sjis/modules/CSVExporter.bas
src-sjis/modules/Utils.bas
```

#### ユーザーフォーム (.frm)
```
src-sjis/forms/MainForm.frm
```

#### クラスモジュール (.cls)
```
src-sjis/classes/StockData.cls
src-sjis/classes/Configuration.cls
```

### 3. 文字化け確認

インポート後、以下を確認：

- ✅ 日本語コメントが正しく表示される
- ✅ 関数名・変数名が正しく表示される
- ✅ 文字列リテラルが正しく表示される

### 4. 文字化けした場合

1. **すべてのモジュールを削除**
2. **Excelを再起動**
3. **`src-sjis/`フォルダから再インポート**

## エンコーディング変換

新しくファイルを編集した場合の変換方法：

### Pythonスクリプトで変換
```python
# tools/convert_encoding.py を実行
python tools/convert_encoding.py
```

### 手動変換（テキストエディタ）
1. `src/`のファイルを開く
2. 「名前を付けて保存」→「エンコーディング: Shift_JIS」
3. `src-sjis/`フォルダに保存

### iconv コマンド（Unix系）
```bash
# 例：MainModule.basを変換
iconv -f UTF-8 -t SHIFT_JIS src/modules/MainModule.bas > src-sjis/modules/MainModule.bas
```

## 文字化けトラブルシューティング

### 症状別対処法

| 症状 | 原因 | 対処法 |
|------|------|--------|
| 日本語コメントが文字化け | UTF-8版をインポート | `src-sjis/`版を使用 |
| 一部文字だけ化ける | 特定文字の変換エラー | ファイルを再生成 |
| インポート時エラー | ファイル破損 | 元ファイルから再変換 |

### 確認方法

インポート後、VBAエディタで以下のコードを実行：

```vba
Sub TestEncoding()
    Debug.Print "文字化けテスト: 日本語表示確認"
    MsgBox "楽天MS2RSS株価データコレクター"
End Sub
```

正しく表示されればエンコーディングは適切です。

## 開発者向け情報

### エンコーディング自動変換

開発時にUTF-8で編集後、自動でShift_JIS版を生成：

```python
# 開発用スクリプト例
def auto_convert():
    # src/ の更新を監視
    # 変更があれば src-sjis/ を自動更新
    pass
```

### Git管理

```gitignore
# UTF-8版のみをGit管理
vba/src/          # ✅ Gitで管理
vba/src-sjis/     # ❌ Gitで管理しない（生成物）
```

## まとめ

**重要**: Excelへのインポートには必ず `vba/src-sjis/` フォルダのファイルを使用してください。

- 📂 `src/` → GitHub閲覧・開発用（UTF-8）
- 📂 `src-sjis/` → Excelインポート用（Shift_JIS）⭐

これで文字化けなくVBAシステムを利用できます！