#!/usr/bin/env python3
"""
Japanese Stock Data Collector - Streamlit App Launcher
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Launch the Streamlit app"""
    
    # Get the directory where this script is located
    script_dir = Path(__file__).parent
    app_path = script_dir / "yahoo_finance_client" / "streamlit_app.py"
    
    # Check if app file exists
    if not app_path.exists():
        print(f"❌ Error: App file not found at {app_path}")
        sys.exit(1)
    
    print("🚀 Starting Japanese Stock Data Collector...")
    print("📈 Yahoo Finance API を使用した日本株データ収集ツール")
    print("-" * 50)
    
    # Configure Streamlit
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    
    # Launch Streamlit
    try:
        cmd = [
            sys.executable, "-m", "streamlit", "run", str(app_path),
            "--server.port", "8501",
            "--server.headless", "true",
            "--browser.gatherUsageStats", "false"
        ]
        
        print(f"📊 Streamlit アプリケーションを起動中...")
        print(f"🌐 URL: http://localhost:8501")
        print(f"⏹️  終了するには Ctrl+C を押してください")
        print("-" * 50)
        
        subprocess.run(cmd, cwd=script_dir)
        
    except KeyboardInterrupt:
        print("\n👋 アプリケーションを終了します...")
    except Exception as e:
        print(f"❌ エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()