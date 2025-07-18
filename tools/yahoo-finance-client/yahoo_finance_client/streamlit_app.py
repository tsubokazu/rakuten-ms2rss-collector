"""
Streamlit Web Application for Japanese Stock Data Collection
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import yfinance as yf
from datetime import datetime, timedelta, date
import io
import time
from typing import List, Dict, Any
import logging

from .client import YahooFinanceClient

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="Japanese Stock Data Collector",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Title and description
st.title("📈 Japanese Stock Data Collector")
st.markdown("**Yahoo Finance API を使用した日本株データ収集ツール**")
st.markdown("---")

# Sidebar for input parameters
st.sidebar.header("⚙️ データ収集設定")

# Stock symbols input
st.sidebar.subheader("📊 銘柄選択")
stock_input_method = st.sidebar.radio(
    "入力方法を選択:",
    ["手動入力", "人気銘柄から選択"]
)

if stock_input_method == "手動入力":
    stock_symbols = st.sidebar.text_input(
        "銘柄コード (カンマ区切り)",
        value="7203,6758,9984",
        help="例: 7203,6758,9984 または 7203.T,6758.T,9984.T"
    )
else:
    # Popular Japanese stocks
    popular_stocks = {
        "トヨタ自動車 (7203)": "7203.T",
        "ソニーグループ (6758)": "6758.T", 
        "ソフトバンクグループ (9984)": "9984.T",
        "日本電信電話 (9432)": "9432.T",
        "キーエンス (6861)": "6861.T",
        "任天堂 (7974)": "7974.T",
        "ファーストリテイリング (9983)": "9983.T",
        "リクルートホールディングス (6098)": "6098.T"
    }
    
    selected_stocks = st.sidebar.multiselect(
        "人気銘柄から選択:",
        options=list(popular_stocks.keys()),
        default=["トヨタ自動車 (7203)", "ソニーグループ (6758)", "ソフトバンクグループ (9984)"]
    )
    
    if selected_stocks:
        stock_symbols = ",".join([popular_stocks[stock] for stock in selected_stocks])
    else:
        stock_symbols = "7203.T,6758.T,9984.T"

# Date range selection
st.sidebar.subheader("📅 期間選択")
date_method = st.sidebar.radio(
    "期間指定方法:",
    ["日付範囲指定", "過去N日間"]
)

if date_method == "日付範囲指定":
    start_date = st.sidebar.date_input(
        "開始日",
        value=date.today() - timedelta(days=7),
        max_value=date.today()
    )
    end_date = st.sidebar.date_input(
        "終了日", 
        value=date.today(),
        max_value=date.today()
    )
else:
    days_back = st.sidebar.selectbox(
        "過去何日間のデータ:",
        options=[1, 3, 7, 14, 30, 60],
        index=2
    )
    end_date = date.today()
    start_date = end_date - timedelta(days=days_back)

# Time interval selection
st.sidebar.subheader("⏰ 時間軸選択")
interval_options = {
    "1分足": "1m",
    "5分足": "5m", 
    "15分足": "15m",
    "30分足": "30m",
    "1時間足": "60m",
    "日足": "1d"
}

selected_interval = st.sidebar.selectbox(
    "時間軸:",
    options=list(interval_options.keys()),
    index=1  # Default to 5分足
)
interval = interval_options[selected_interval]

# Display current settings
st.sidebar.markdown("---")
st.sidebar.subheader("📋 現在の設定")
st.sidebar.write(f"**銘柄**: {stock_symbols}")
st.sidebar.write(f"**期間**: {start_date} ～ {end_date}")
st.sidebar.write(f"**時間軸**: {selected_interval}")

# Data collection button
if st.sidebar.button("📥 データ収集開始", type="primary"):
    # Validate inputs
    if not stock_symbols.strip():
        st.error("銘柄コードを入力してください")
        st.stop()
    
    if start_date >= end_date:
        st.error("開始日は終了日より前である必要があります")
        st.stop()
    
    # Parse stock symbols
    symbols = [s.strip() for s in stock_symbols.split(",")]
    if len(symbols) == 0:
        st.error("有効な銘柄コードを入力してください")
        st.stop()
    
    # Auto-add .T suffix for Japanese stocks
    processed_symbols = []
    for symbol in symbols:
        if not "." in symbol:
            processed_symbols.append(f"{symbol}.T")
        else:
            processed_symbols.append(symbol)
    
    # Initialize client
    client = YahooFinanceClient()
    
    # Create progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Collect data
    all_data = []
    total_symbols = len(processed_symbols)
    
    for i, symbol in enumerate(processed_symbols):
        status_text.text(f"データ収集中: {symbol} ({i+1}/{total_symbols})")
        
        try:
            # Get data for this symbol
            data = client.get_stock_data(
                symbol=symbol,
                interval=interval,
                start_date=start_date.strftime("%Y-%m-%d"),
                end_date=end_date.strftime("%Y-%m-%d")
            )
            
            if not data.empty:
                all_data.append(data)
                st.success(f"✅ {symbol}: {len(data)} 件のデータを取得")
            else:
                st.warning(f"⚠️ {symbol}: データが取得できませんでした")
                
        except Exception as e:
            st.error(f"❌ {symbol}: エラー - {str(e)}")
            
        # Update progress
        progress_bar.progress((i + 1) / total_symbols)
    
    # Combine all data
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        combined_data = combined_data.sort_values(['Symbol', 'Datetime'])
        
        # Store in session state
        st.session_state.stock_data = combined_data
        st.session_state.collection_success = True
        
        status_text.text("✅ データ収集完了!")
        st.success(f"🎉 合計 {len(combined_data)} 件のデータを収集しました")
        
    else:
        st.error("❌ データを取得できませんでした")
        st.session_state.collection_success = False

# Display results if data is available
if 'stock_data' in st.session_state and st.session_state.collection_success:
    data = st.session_state.stock_data
    
    st.markdown("---")
    st.header("📊 データ表示・分析")
    
    # Create tabs
    tab1, tab2, tab3 = st.tabs(["📈 チャート", "📋 データテーブル", "📥 ダウンロード"])
    
    with tab1:
        st.subheader("株価チャート")
        
        # Symbol selection for chart
        available_symbols = data['Symbol'].unique()
        selected_symbol = st.selectbox(
            "表示する銘柄を選択:",
            options=available_symbols,
            index=0
        )
        
        # Filter data for selected symbol
        symbol_data = data[data['Symbol'] == selected_symbol].copy()
        
        if not symbol_data.empty:
            # Create candlestick chart
            fig = make_subplots(
                rows=2, cols=1,
                shared_xaxes=True,
                vertical_spacing=0.03,
                subplot_titles=(f'{selected_symbol} 株価チャート', '出来高'),
                row_width=[0.7, 0.3]
            )
            
            # Add candlestick chart
            fig.add_trace(
                go.Candlestick(
                    x=symbol_data['Datetime'],
                    open=symbol_data['Open'],
                    high=symbol_data['High'],
                    low=symbol_data['Low'],
                    close=symbol_data['Close'],
                    name=selected_symbol
                ),
                row=1, col=1
            )
            
            # Add volume chart
            fig.add_trace(
                go.Bar(
                    x=symbol_data['Datetime'],
                    y=symbol_data['Volume'],
                    name='出来高',
                    marker_color='rgba(158,202,225,0.8)'
                ),
                row=2, col=1
            )
            
            # Update layout
            fig.update_layout(
                title=f'{selected_symbol} - {selected_interval}',
                xaxis_rangeslider_visible=False,
                height=600,
                showlegend=False
            )
            
            # Update axes
            fig.update_yaxes(title_text="価格 (円)", row=1, col=1)
            fig.update_yaxes(title_text="出来高", row=2, col=1)
            fig.update_xaxes(title_text="時間", row=2, col=1)
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display summary statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("最新価格", f"¥{symbol_data['Close'].iloc[-1]:,.0f}")
            with col2:
                price_change = symbol_data['Close'].iloc[-1] - symbol_data['Close'].iloc[0]
                st.metric("期間変化", f"¥{price_change:,.0f}")
            with col3:
                st.metric("最高値", f"¥{symbol_data['High'].max():,.0f}")
            with col4:
                st.metric("最安値", f"¥{symbol_data['Low'].min():,.0f}")
    
    with tab2:
        st.subheader("データテーブル")
        
        # Display data table
        st.dataframe(
            data,
            use_container_width=True,
            height=400
        )
        
        # Data summary
        st.subheader("データサマリー")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**基本統計**")
            st.write(f"• 総レコード数: {len(data):,}")
            st.write(f"• 銘柄数: {data['Symbol'].nunique()}")
            st.write(f"• 期間: {data['Datetime'].min().strftime('%Y-%m-%d %H:%M')} ～ {data['Datetime'].max().strftime('%Y-%m-%d %H:%M')}")
        
        with col2:
            st.write("**銘柄別データ数**")
            symbol_counts = data['Symbol'].value_counts()
            for symbol, count in symbol_counts.items():
                st.write(f"• {symbol}: {count:,} 件")
    
    with tab3:
        st.subheader("データダウンロード")
        
        # Generate filename
        symbols_str = "_".join([s.replace('.T', '') for s in data['Symbol'].unique()])
        start_str = start_date.strftime("%Y%m%d")
        end_str = end_date.strftime("%Y%m%d")
        filename = f"{symbols_str}_{interval}_{start_str}-{end_str}.csv"
        
        # Convert to CSV
        csv_buffer = io.StringIO()
        data.to_csv(csv_buffer, index=False)
        csv_string = csv_buffer.getvalue()
        
        st.download_button(
            label="📥 CSVファイルをダウンロード",
            data=csv_string,
            file_name=filename,
            mime="text/csv"
        )
        
        st.info(f"ファイル名: {filename}")
        st.success(f"💾 {len(data)} 件のデータを CSV として保存可能です")

# Footer
st.markdown("---")
st.markdown("**Japanese Stock Data Collector** - Yahoo Finance API を使用")
st.markdown("Created with ❤️ using Streamlit")

# Instructions for first-time users
if 'collection_success' not in st.session_state:
    st.markdown("---")
    st.info("""
    ### 📖 使用方法
    
    1. **左側のサイドバー**で銘柄・期間・時間軸を設定
    2. **「データ収集開始」**ボタンをクリック
    3. データが取得されたら、**チャート・テーブル・ダウンロード**タブで結果を確認
    
    ### 📝 注意事項
    
    - **1分足**: 最大7日間のデータ
    - **5分足**: 最大60日間のデータ  
    - **その他**: Yahoo Finance の制限に依存
    - 日本株は自動で「.T」拡張子が付加されます
    """)