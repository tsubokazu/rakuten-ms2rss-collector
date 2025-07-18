"""VBA bridge for Yahoo Finance client."""

import sys
import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from .client import YahooFinanceClient

# Setup logging for VBA bridge
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class VBABridge:
    """Bridge between VBA and Yahoo Finance API."""
    
    def __init__(self, output_dir: str = "output/csv"):
        """Initialize the VBA bridge."""
        self.client = YahooFinanceClient()
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def collect_stock_data(
        self,
        symbol: str,
        interval: str,
        start_date: str,
        end_date: str,
        output_file: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Collect stock data for VBA integration.
        
        Args:
            symbol: Stock symbol (e.g., '7203.T')
            interval: Data interval ('1m', '5m', '15m', '30m', '60m', '1d')
            start_date: Start date in YYYY-MM-DD format
            end_date: End date in YYYY-MM-DD format
            output_file: Optional output file path
            
        Returns:
            Dictionary with result information
        """
        try:
            logger.info(f"Collecting data for {symbol} ({interval}) from {start_date} to {end_date}")
            
            # Get stock data
            data = self.client.get_stock_data(
                symbol=symbol,
                interval=interval,
                start_date=start_date,
                end_date=end_date
            )
            
            if data.empty:
                return {
                    'success': False,
                    'error': 'No data retrieved',
                    'record_count': 0,
                    'output_file': None
                }
            
            # Generate output filename if not provided
            if not output_file:
                symbol_clean = symbol.replace('.T', '')
                start_clean = start_date.replace('-', '')
                end_clean = end_date.replace('-', '')
                output_file = self.output_dir / f"{symbol_clean}_{interval}_{start_clean}-{end_clean}.csv"
            else:
                output_file = Path(output_file)
                output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Save data
            self.client.save_to_csv(data, str(output_file))
            
            return {
                'success': True,
                'record_count': len(data),
                'output_file': str(output_file),
                'date_range': {
                    'start': str(data['Datetime'].min()),
                    'end': str(data['Datetime'].max())
                },
                'symbol': symbol,
                'interval': interval
            }
            
        except Exception as e:
            logger.error(f"Error collecting data: {e}")
            return {
                'success': False,
                'error': str(e),
                'record_count': 0,
                'output_file': None
            }
    
    def collect_multiple_stocks(
        self,
        symbols: str,
        interval: str,
        start_date: str,
        end_date: str,
        output_file: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Collect data for multiple stocks.
        
        Args:
            symbols: Comma-separated stock symbols
            interval: Data interval
            start_date: Start date in YYYY-MM-DD format
            end_date: End date in YYYY-MM-DD format
            output_file: Optional output file path
            
        Returns:
            Dictionary with result information
        """
        try:
            symbol_list = [s.strip() for s in symbols.split(',')]
            logger.info(f"Collecting data for {len(symbol_list)} symbols")
            
            # Get data for multiple stocks
            data = self.client.get_multiple_stocks_data(
                symbols=symbol_list,
                interval=interval,
                start_date=start_date,
                end_date=end_date
            )
            
            if data.empty:
                return {
                    'success': False,
                    'error': 'No data retrieved',
                    'record_count': 0,
                    'output_file': None
                }
            
            # Generate output filename if not provided
            if not output_file:
                symbols_clean = '_'.join([s.replace('.T', '') for s in symbol_list])
                start_clean = start_date.replace('-', '')
                end_clean = end_date.replace('-', '')
                output_file = self.output_dir / f"{symbols_clean}_{interval}_{start_clean}-{end_clean}.csv"
            else:
                output_file = Path(output_file)
                output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Save data
            self.client.save_to_csv(data, str(output_file))
            
            return {
                'success': True,
                'record_count': len(data),
                'output_file': str(output_file),
                'date_range': {
                    'start': str(data['Datetime'].min()),
                    'end': str(data['Datetime'].max())
                },
                'symbols': symbol_list,
                'interval': interval
            }
            
        except Exception as e:
            logger.error(f"Error collecting multiple stocks data: {e}")
            return {
                'success': False,
                'error': str(e),
                'record_count': 0,
                'output_file': None
            }


def main():
    """Main entry point for VBA bridge."""
    if len(sys.argv) < 6:
        print("Usage: python -m yahoo_finance_client.vba_bridge <symbol> <interval> <start_date> <end_date> <output_file>")
        sys.exit(1)
    
    symbol = sys.argv[1]
    interval = sys.argv[2]
    start_date = sys.argv[3]
    end_date = sys.argv[4]
    output_file = sys.argv[5] if len(sys.argv) > 5 else None
    
    bridge = VBABridge()
    result = bridge.collect_stock_data(symbol, interval, start_date, end_date, output_file)
    
    # Print result as JSON for VBA parsing
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()