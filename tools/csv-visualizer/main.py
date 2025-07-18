import argparse
import os
from pathlib import Path
from typing import List, Optional

import pandas as pd
import plotly.graph_objects as go

# デフォルトの列名 (ヘッダはSJISで壊れているため固定で付与)
DEFAULT_COLUMNS: List[str] = [
    "symbol",  # 銘柄コード
    "market",  # 市場区分
    "timeframe",  # 5M など
    "date",  # YYYY/MM/DD
    "time",  # HH:MM
    "open",
    "high",
    "low",
    "close",
    "volume",
]


def read_csv(path: Path, encoding_candidates: Optional[List[str]] = None) -> pd.DataFrame:
    """CSV を読み込み、DataFrame を返す。

    ヘッダー行は破損しているためスキップし、固定ヘッダーを使用する。
    入力CSVのエンコーディングは Shift-JIS のことが多いが、失敗時には可読エンコーディングを総当たりで試す。
    """
    encodings = encoding_candidates or ["shift_jis", "cp932", "utf-8"]
    last_error: Exception | None = None
    for enc in encodings:
        try:
            df = pd.read_csv(
                path,
                header=None,
                names=DEFAULT_COLUMNS,
                skiprows=1,  # 先頭行はヘッダーなのでスキップ
                encoding=enc,
                dtype={
                    "symbol": str,
                    "market": str,
                    "timeframe": str,
                    "date": str,
                    "time": str,
                    "open": float,
                    "high": float,
                    "low": float,
                    "close": float,
                    "volume": int,
                },
            )
            return df
        except Exception as e:  # pragma: no cover
            last_error = e
            continue

    raise RuntimeError(f"CSV の読み込みに失敗しました. 最後のエラー: {last_error}")


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """日付と時刻を結合し、ローソク足描画用に整形する。"""
    # pandas 2.0: format="mixed" で警告が出る場合があるため errors='coerce' で安全に変換
    df["datetime"] = pd.to_datetime(df["date"] + " " + df["time"], errors="coerce")
    # NaT を含む行は落とす
    df = df.dropna(subset=["datetime"])
    df = df.sort_values("datetime").reset_index(drop=True)
    return df


def create_candlestick_chart(df: pd.DataFrame) -> go.Figure:
    """Plotly でローソク足チャートを生成する。"""
    fig = go.Figure(
        data=[
            go.Candlestick(
                x=df["datetime"],
                open=df["open"],
                high=df["high"],
                low=df["low"],
                close=df["close"],
                name="Price",
            )
        ]
    )

    fig.update_layout(
        xaxis_title="DateTime",
        yaxis_title="Price (JPY)",
        title="Candlestick Chart",
        xaxis_rangeslider_visible=False,
        template="plotly_white",
    )
    return fig


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Candlestick chart generator from CSV.")
    parser.add_argument("-i", "--input", required=True, help="入力CSVファイル")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="出力HTMLファイル (省略時は input と同じディレクトリに candlestick.html)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    csv_path = Path(args.input)
    if not csv_path.exists():
        raise FileNotFoundError(f"入力CSVが見つかりません: {csv_path}")

    df = read_csv(csv_path)
    df = preprocess(df)

    fig = create_candlestick_chart(df)

    output_path = (
        Path(args.output)
        if args.output
        else csv_path.with_suffix("").with_name("candlestick.html")
    )
    # 出力ディレクトリが存在しない場合は作成
    os.makedirs(output_path.parent, exist_ok=True)
    fig.write_html(str(output_path))

    print(f"✅ チャートを生成しました: {output_path}")


if __name__ == "__main__":
    main() 