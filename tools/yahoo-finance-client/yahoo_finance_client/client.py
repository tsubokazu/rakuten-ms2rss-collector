"""Yahoo Finance API client implementation."""

import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class YahooFinanceClient:
    """Yahoo Finance API client for Japanese stock data."""
    
    # Yahoo Finance interval limits
    INTERVAL_LIMITS = {
        '1m': 7,    # 7 days max
        '5m': 60,   # 60 days max
        '15m': 60,  # 60 days max
        '30m': 60,  # 60 days max
        '60m': 730, # 730 days max
        '1d': 3650, # 10 years max
    }
    
    def __init__(self):
        """Initialize the Yahoo Finance client."""
        pass
    
    def get_stock_data(
        self,
        symbol: str,
        interval: str = '5m',
        period_days: int = 30,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Get stock data from Yahoo Finance.
        
        Args:
            symbol: Stock symbol (e.g., '7203.T' for Toyota)
            interval: Data interval ('1m', '5m', '15m', '30m', '60m', '1d')
            period_days: Number of days to retrieve (if start_date/end_date not specified)
            start_date: Start date in YYYY-MM-DD format
            end_date: End date in YYYY-MM-DD format
            
        Returns:
            DataFrame with OHLCV data
        """
        try:
            # Validate interval
            if interval not in self.INTERVAL_LIMITS:
                raise ValueError(f"Unsupported interval: {interval}. Supported: {list(self.INTERVAL_LIMITS.keys())}")
            
            # Apply interval-specific limits
            max_days = self.INTERVAL_LIMITS[interval]
            if period_days > max_days:
                logger.warning(f"Period {period_days} days exceeds {interval} limit of {max_days} days. Using {max_days} days.")
                period_days = max_days
            
            # Create ticker object
            ticker = yf.Ticker(symbol)
            
            # Get data based on date specification
            if start_date and end_date:
                # Use specific date range
                data = ticker.history(start=start_date, end=end_date, interval=interval)
            elif start_date:
                # Use start date with period
                start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                end_dt = start_dt + timedelta(days=period_days)
                data = ticker.history(start=start_dt, end=end_dt, interval=interval)
            else:
                # Use period only
                data = ticker.history(period=f"{period_days}d", interval=interval)
            
            if data.empty:
                logger.warning(f"No data returned for {symbol} with interval {interval}")
                return pd.DataFrame()
            
            # Reset index to make Datetime a column
            data = data.reset_index()
            
            # Add symbol column
            data['Symbol'] = symbol
            
            # Reorder columns
            columns = ['Symbol', 'Datetime', 'Open', 'High', 'Low', 'Close', 'Volume', 'Dividends', 'Stock Splits']
            data = data.reindex(columns=columns)
            
            logger.info(f"Retrieved {len(data)} records for {symbol} ({interval})")
            return data
            
        except Exception as e:
            logger.error(f"Error retrieving data for {symbol}: {e}")
            raise
    
    def get_multiple_stocks_data(
        self,
        symbols: List[str],
        interval: str = '5m',
        period_days: int = 30,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Get data for multiple stocks.
        
        Args:
            symbols: List of stock symbols
            interval: Data interval
            period_days: Number of days to retrieve
            start_date: Start date in YYYY-MM-DD format
            end_date: End date in YYYY-MM-DD format
            
        Returns:
            Combined DataFrame with all stock data
        """
        all_data = []
        
        for symbol in symbols:
            try:
                data = self.get_stock_data(symbol, interval, period_days, start_date, end_date)
                if not data.empty:
                    all_data.append(data)
            except Exception as e:
                logger.error(f"Failed to get data for {symbol}: {e}")
                continue
        
        if not all_data:
            return pd.DataFrame()
        
        # Combine all data
        combined_data = pd.concat(all_data, ignore_index=True)
        
        # Sort by symbol and datetime
        combined_data = combined_data.sort_values(['Symbol', 'Datetime'])
        
        return combined_data
    
    def save_to_csv(self, data: pd.DataFrame, filename: str) -> None:
        """
        Save data to CSV file.
        
        Args:
            data: DataFrame to save
            filename: Output filename
        """
        try:
            data.to_csv(filename, index=False)
            logger.info(f"Data saved to {filename}")
        except Exception as e:
            logger.error(f"Error saving to CSV: {e}")
            raise
    
    def get_stock_info(self, symbol: str) -> Dict[str, Any]:
        """
        Get stock information.
        
        Args:
            symbol: Stock symbol
            
        Returns:
            Dictionary with stock info
        """
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info
            return {
                'symbol': symbol,
                'name': info.get('longName', 'N/A'),
                'currency': info.get('currency', 'N/A'),
                'exchange': info.get('exchange', 'N/A'),
                'marketCap': info.get('marketCap', 'N/A'),
                'sector': info.get('sector', 'N/A'),
                'industry': info.get('industry', 'N/A'),
            }
        except Exception as e:
            logger.error(f"Error getting info for {symbol}: {e}")
            return {'symbol': symbol, 'error': str(e)}