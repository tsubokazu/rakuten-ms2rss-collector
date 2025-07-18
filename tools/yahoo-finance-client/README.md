# 📈 Japanese Stock Data Collector

日本株式の分足データを取得するためのYahoo Finance APIクライアント

## 🎯 機能

- **Webアプリケーション**: 使いやすいStreamlit GUI
- **日本株式対応**: 1分、5分、15分、30分、60分、日足データ
- **期間指定**: 柔軟な日付範囲指定
- **チャート表示**: インタラクティブなローソク足チャート
- **CSV出力**: データのダウンロード機能
- **複数銘柄対応**: 一度に複数銘柄のデータ収集

## 🚀 使用方法

### 1. Webアプリケーション（推奨）

```bash
# Streamlit Webアプリを起動
uv run python run_app.py

# または直接起動
uv run streamlit run yahoo_finance_client/streamlit_app.py
```

ブラウザで http://localhost:8501 にアクセス

### 2. コマンドライン

```bash
# 基本的な使用例
uv run python -m yahoo_finance_client.main 7203.T 5m 30

# 期間指定
uv run python -m yahoo_finance_client.main 7203.T 5m 30 --start-date 2025-07-01 --end-date 2025-07-18

# 複数銘柄
uv run python -m yahoo_finance_client.main 7203.T,6758.T,9984.T 5m 30
```

## 📊 対応銘柄

- **トヨタ自動車**: 7203.T
- **ソニーグループ**: 6758.T
- **ソフトバンクグループ**: 9984.T
- **日本電信電話**: 9432.T
- **キーエンス**: 6861.T
- **任天堂**: 7974.T
- **ファーストリテイリング**: 9983.T
- **リクルートホールディングス**: 6098.T
- その他すべての日本株（.T拡張子自動付加）

## ⚠️ 制限事項

- **1分足**: 最大7日間
- **5分足**: 最大60日間  
- **その他**: Yahoo Finance APIの制限に依存
- インターネット接続が必要

## 🛠️ 技術仕様

- **Python**: 3.11+
- **主要ライブラリ**: yfinance, streamlit, plotly, pandas
- **データソース**: Yahoo Finance API
- **出力形式**: CSV (Symbol, Datetime, Open, High, Low, Close, Volume)