# Yahoo Finance Client

日本株式の分足データを取得するためのYahoo Finance APIクライアント

## 機能

- 日本株式の分足データ取得（1分、5分、15分、30分、60分）
- 期間指定でのデータ取得
- CSV形式での出力
- VBAからの呼び出し対応

## 使用方法

```bash
# 基本的な使用例
uv run yahoo-finance-client 7203.T 5m 30 --output output.csv

# 期間指定
uv run yahoo-finance-client 7203.T 5m 30 --start-date 2025-07-01 --end-date 2025-07-18

# 複数銘柄
uv run yahoo-finance-client 7203.T,6758.T,9984.T 5m 30
```

## 制限事項

- 1分足: 最大7日間
- 5分足: 最大60日間
- その他の時間軸: Yahoo Finance APIの制限に依存