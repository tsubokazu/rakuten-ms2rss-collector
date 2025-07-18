# 📈 Japanese Stock Data Collector - Yahoo Finance API統合システム

## プロジェクト概要

Yahoo Finance APIを使用した日本株式データ収集システム。Streamlit Webアプリケーションとして実装されており、ブラウザから簡単に株式チャートデータを取得・分析できる。

**現在の状態**: 完全に動作する本番環境 - インターネット接続のみで利用可能

## 🚀 クイックスタート

```bash
# 1. ディレクトリに移動
cd tools/yahoo-finance-client

# 2. 依存関係をインストール
uv sync

# 3. Webアプリケーションを起動
uv run streamlit run yahoo_finance_client/streamlit_app.py --server.port 8501

# 4. ブラウザでアクセス
# http://localhost:8501
```

## 🎯 主要機能

### Streamlit Webアプリケーション
- **直感的なUI**: サイドバーでの簡単設定
- **人気銘柄プリセット**: トヨタ、ソニー、ソフトバンク等
- **柔軟な期間指定**: 日付範囲または「過去N日間」選択
- **複数時間軸**: 1分、5分、15分、30分、60分、日足
- **リアルタイム収集**: 進捗バー付きデータ収集
- **インタラクティブチャート**: Plotlyによるローソク足チャート
- **CSV出力**: ワンクリックダウンロード機能

### データ収集機能
- **日本株式対応**: 自動.T拡張子付加
- **複数銘柄**: 一度に複数銘柄のデータ収集
- **エラーハンドリング**: 包括的なエラー処理
- **データ検証**: 入力値の自動検証

## 📊 アーキテクチャ

### ディレクトリ構造
```
stock-ai-trade/
├── tools/
│   ├── yahoo-finance-client/      # メインアプリケーション
│   │   ├── yahoo_finance_client/
│   │   │   ├── streamlit_app.py   # Webアプリケーション
│   │   │   ├── client.py          # API クライアント
│   │   │   ├── main.py           # CLI インターフェース
│   │   │   └── vba_bridge.py     # レガシーブリッジ
│   │   ├── run_app.py            # アプリ起動スクリプト
│   │   └── README.md             # 使用方法
│   ├── csv-visualizer/           # データ可視化
│   └── pdf-reader/               # PDF処理
├── output/
│   ├── csv/                      # データ出力
│   └── logs/                     # ログファイル
├── CLAUDE.md                     # このファイル
├── README.md                     # プロジェクト説明
└── pyproject.toml               # Python設定
```

### 技術スタック
- **Frontend**: Streamlit (Web UI)
- **Backend**: Python + Yahoo Finance API
- **データ処理**: pandas, yfinance
- **可視化**: Plotly (ローソク足チャート)
- **パッケージ管理**: uv

## 🎨 チャート機能

### 専用設計
- **日本式カラーリング**: 上昇=赤、下降=青
- **連続表示**: 取引時間のみ表示（空白時間除去）
- **出来高連動**: 株価動向と連動した出来高バーの色分け
- **プロフェッショナル**: 適切なグリッドとスタイリング

### 統計情報
- **リアルタイム価格**: 最新価格、期間変化
- **価格帯情報**: 最高値、最安値
- **データサマリー**: レコード数、銘柄数、期間

## 💻 使用方法

### 1. Webアプリケーション（推奨）
```bash
# アプリケーション起動
uv run python tools/yahoo-finance-client/run_app.py

# ブラウザでアクセス
# http://localhost:8501
```

### 2. コマンドライン
```bash
# 基本的な使用
uv run python -m yahoo_finance_client.main 7203.T 5m 30

# 期間指定
uv run python -m yahoo_finance_client.main 7203.T 5m 30 --start-date 2025-07-01 --end-date 2025-07-18

# 複数銘柄
uv run python -m yahoo_finance_client.main 7203.T,6758.T,9984.T 5m 30
```

## 📈 対応銘柄

### 人気銘柄（プリセット）
- **7203.T**: トヨタ自動車
- **6758.T**: ソニーグループ
- **9984.T**: ソフトバンクグループ
- **9432.T**: 日本電信電話
- **6861.T**: キーエンス
- **7974.T**: 任天堂
- **9983.T**: ファーストリテイリング
- **6098.T**: リクルートホールディングス

### 対応市場
- **東証**: 全銘柄（.T拡張子自動付加）
- **その他**: Yahoo Finance対応の全日本株

## ⚠️ 制限事項

### Yahoo Finance API制限
- **1分足**: 最大7日間
- **5分足**: 最大60日間
- **その他**: Yahoo Finance APIの標準制限に依存

### システム要件
- **Python**: 3.11以上
- **インターネット接続**: 必須
- **ブラウザ**: モダンブラウザ（Chrome, Firefox, Safari, Edge）

## 🔧 技術仕様

### 主要ライブラリ
```toml
dependencies = [
    "yfinance>=0.2.65",    # Yahoo Finance API
    "pandas>=2.0.3",       # データ処理
    "streamlit>=1.47.0",   # Web UI
    "plotly>=6.2.0",       # チャート描画
    "click>=8.0.0",        # CLI
]
```

### API設計
```python
class YahooFinanceClient:
    def get_stock_data(symbol, interval, start_date, end_date)
    def get_multiple_stocks_data(symbols, interval, start_date, end_date)
    def save_to_csv(data, filename)
```

### データ形式
```csv
Symbol,Datetime,Open,High,Low,Close,Volume,Dividends,Stock Splits
7203.T,2025-07-18 09:05:00+09:00,2518.0,2525.0,2515.0,2520.0,125000,0.0,0.0
```

## 🚀 デプロイ・共有

### ローカル実行
```bash
# 開発用
uv run streamlit run yahoo_finance_client/streamlit_app.py --server.port 8501

# 本番用（外部アクセス許可）
uv run streamlit run yahoo_finance_client/streamlit_app.py --server.port 8501 --server.address 0.0.0.0
```

### Docker化（将来）
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install uv && uv sync
EXPOSE 8501
CMD ["uv", "run", "streamlit", "run", "yahoo_finance_client/streamlit_app.py", "--server.port", "8501", "--server.address", "0.0.0.0"]
```

## 🎉 主な成果

### 完全な機能実装
- ✅ **Web UI**: 直感的なStreamlitアプリケーション
- ✅ **データ収集**: Yahoo Finance APIによる確実なデータ取得
- ✅ **チャート表示**: プロフェッショナルなローソク足チャート
- ✅ **CSV出力**: 分析用データのエクスポート
- ✅ **エラーハンドリング**: 包括的なエラー処理
- ✅ **日本株対応**: 東証銘柄の完全サポート

### 開発・運用の改善
- ✅ **VBA廃止**: 複雑なExcel/VBA依存を排除
- ✅ **クロスプラットフォーム**: Windows, Mac, Linux対応
- ✅ **セットアップ簡素化**: 依存関係の最小化
- ✅ **保守性向上**: モジュール化された設計

## 📝 今後の拡張可能性

### 機能拡張
- **銘柄スクリーニング**: 条件検索機能
- **テクニカル指標**: 移動平均線、RSI、MACD等
- **アラート機能**: 価格変動通知
- **ポートフォリオ管理**: 複数銘柄の一括管理

### 技術的改善
- **キャッシュ機能**: データ取得の高速化
- **バックグラウンド処理**: 大量データの非同期処理
- **データベース統合**: 履歴データの永続化
- **API Rate Limiting**: レート制限の自動調整

## 🛠️ 開発・デバッグ

### 開発コマンド
```bash
# 依存関係の更新
uv sync

# テスト実行
uv run python -m yahoo_finance_client.main 7203.T 5m 5 --verbose

# Streamlitアプリの開発モード
uv run streamlit run yahoo_finance_client/streamlit_app.py --server.runOnSave true
```

### ログ確認
```bash
# アプリケーションログ
tail -f output/logs/stock_data_collector_*.log

# データ出力確認
ls -la output/csv/
```

## 🤝 コントリビューション

### コード品質
- **型ヒント**: 全関数に型注釈
- **ドキュメント**: 包括的なコメント
- **エラーハンドリング**: 堅牢なエラー処理
- **テスト**: 主要機能のテスト

### Git フロー
```bash
# 機能開発
git checkout -b feature/new-feature
git commit -m "Add new feature"
git push origin feature/new-feature

# リリース
git checkout main
git merge feature/new-feature
git push origin main
```

---

**Japanese Stock Data Collector** - 日本株式データ収集の決定版

Created with ❤️ using Streamlit, Yahoo Finance API, and Python