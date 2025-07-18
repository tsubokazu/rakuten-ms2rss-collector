# CSV Candlestick Visualizer

このツールは、自動生成された株価 CSV をローソク足チャートとして可視化し、HTML ファイルとして出力します。

## 使い方

```bash
python tools/csv-visualizer/main.py -i output/csv/7203_5M_20250701-20250715.csv -o output/charts/7203_ローソク足.html
```

- `-i, --input` : 入力 CSV ファイルパス（必須）
- `-o, --output`: 出力 HTML ファイルパス（省略時は同ディレクトリに `candlestick.html` が生成されます）

## 依存関係

- Python 3.11 以上
- pandas
- plotly
- kaleido (静的イメージ書き出し用)

各依存は `pyproject.toml` に記載されています。`uv pip install -r` などでインストールしてください。
