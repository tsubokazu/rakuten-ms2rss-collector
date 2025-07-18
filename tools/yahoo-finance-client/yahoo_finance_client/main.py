"""Command line interface for Yahoo Finance client."""

import click
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path
from .client import YahooFinanceClient

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@click.command()
@click.argument('symbols', required=True)
@click.argument('interval', required=True)
@click.argument('period_days', type=int, default=30)
@click.option('--start-date', '-s', help='Start date (YYYY-MM-DD)')
@click.option('--end-date', '-e', help='End date (YYYY-MM-DD)')
@click.option('--output', '-o', help='Output CSV file path')
@click.option('--output-dir', '-d', default='output/csv', help='Output directory')
@click.option('--verbose', '-v', is_flag=True, help='Verbose logging')
def main(symbols, interval, period_days, start_date, end_date, output, output_dir, verbose):
    """
    Yahoo Finance API client for Japanese stock data.
    
    SYMBOLS: Comma-separated stock symbols (e.g., 7203.T,6758.T,9984.T)
    INTERVAL: Data interval (1m, 5m, 15m, 30m, 60m, 1d)
    PERIOD_DAYS: Number of days to retrieve (default: 30)
    
    Examples:
        yahoo-finance-client 7203.T 5m 30
        yahoo-finance-client 7203.T,6758.T 5m 30 --start-date 2025-07-01 --end-date 2025-07-18
        yahoo-finance-client 7203.T 1m 7 --output toyota_1m.csv
    """
    
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        # Parse symbols
        symbol_list = [s.strip() for s in symbols.split(',')]
        logger.info(f"Retrieving data for symbols: {symbol_list}")
        logger.info(f"Interval: {interval}, Period: {period_days} days")
        
        # Validate dates
        if start_date and end_date:
            try:
                start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                end_dt = datetime.strptime(end_date, '%Y-%m-%d')
                if start_dt >= end_dt:
                    raise ValueError("Start date must be before end date")
                logger.info(f"Date range: {start_date} to {end_date}")
            except ValueError as e:
                logger.error(f"Invalid date format: {e}")
                sys.exit(1)
        
        # Create client
        client = YahooFinanceClient()
        
        # Get data
        if len(symbol_list) == 1:
            data = client.get_stock_data(
                symbol_list[0], interval, period_days, start_date, end_date
            )
        else:
            data = client.get_multiple_stocks_data(
                symbol_list, interval, period_days, start_date, end_date
            )
        
        if data.empty:
            logger.warning("No data retrieved")
            sys.exit(1)
        
        logger.info(f"Retrieved {len(data)} records")
        
        # Generate output filename if not specified
        if not output:
            # Create output directory
            output_dir_path = Path(output_dir)
            output_dir_path.mkdir(parents=True, exist_ok=True)
            
            # Generate filename
            symbols_str = '_'.join(symbol_list).replace('.T', '')
            if start_date and end_date:
                date_str = f"{start_date.replace('-', '')}-{end_date.replace('-', '')}"
            else:
                end_date_str = datetime.now().strftime('%Y%m%d')
                start_date_str = (datetime.now() - timedelta(days=period_days)).strftime('%Y%m%d')
                date_str = f"{start_date_str}-{end_date_str}"
            
            output = output_dir_path / f"{symbols_str}_{interval}_{date_str}.csv"
        
        # Save to CSV
        client.save_to_csv(data, str(output))
        
        # Print summary
        print(f"✅ Data successfully retrieved and saved to: {output}")
        print(f"   Records: {len(data)}")
        print(f"   Symbols: {', '.join(symbol_list)}")
        print(f"   Interval: {interval}")
        if not data.empty:
            print(f"   Date range: {data['Datetime'].min()} to {data['Datetime'].max()}")
        
    except Exception as e:
        logger.error(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()