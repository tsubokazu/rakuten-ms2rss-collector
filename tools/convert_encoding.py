#!/usr/bin/env python3
"""
VBAファイルをUTF-8からShift_JISに変換するスクリプト
"""

import os
import shutil
from pathlib import Path

def convert_vba_files():
    """VBAファイルをShift_JISに変換"""
    
    # ソースディレクトリとターゲットディレクトリ
    src_dir = Path("vba/src")
    target_dir = Path("vba/src-sjis")
    
    # ターゲットディレクトリの構造を作成
    for subdir in ["modules", "forms", "classes"]:
        (target_dir / subdir).mkdir(parents=True, exist_ok=True)
    
    # 変換対象ファイル
    vba_files = []
    for pattern in ["**/*.bas", "**/*.frm", "**/*.cls"]:
        vba_files.extend(src_dir.glob(pattern))
    
    converted_count = 0
    error_count = 0
    
    print("VBAファイルをShift_JISに変換中...")
    print("=" * 50)
    
    for file_path in vba_files:
        try:
            # UTF-8で読み取り
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 相対パスを取得
            rel_path = file_path.relative_to(src_dir)
            target_path = target_dir / rel_path
            
            # Shift_JISで書き込み
            with open(target_path, 'w', encoding='shift_jis', errors='replace') as f:
                f.write(content)
            
            print(f"✓ {rel_path}")
            converted_count += 1
            
        except Exception as e:
            print(f"✗ {file_path.name}: {e}")
            error_count += 1
    
    print("=" * 50)
    print(f"変換完了: {converted_count}件")
    if error_count > 0:
        print(f"エラー: {error_count}件")
    
    # READMEもコピー
    try:
        readme_src = src_dir / "README.md"
        readme_target = target_dir / "README.md"
        
        with open(readme_src, 'r', encoding='utf-8') as f:
            readme_content = f.read()
        
        with open(readme_target, 'w', encoding='shift_jis', errors='replace') as f:
            f.write(readme_content)
        
        print(f"✓ README.md もコピーしました")
        
    except Exception as e:
        print(f"✗ README.md のコピーに失敗: {e}")

def verify_conversion():
    """変換結果を確認"""
    
    target_dir = Path("vba/src-sjis")
    
    print("\n変換されたファイル一覧:")
    print("=" * 50)
    
    for file_path in sorted(target_dir.rglob("*")):
        if file_path.is_file():
            size = file_path.stat().st_size
            print(f"{file_path.relative_to(target_dir)} ({size:,} bytes)")
    
    print("=" * 50)

if __name__ == "__main__":
    try:
        convert_vba_files()
        verify_conversion()
        print("\n✅ Shift_JIS変換が完了しました！")
        print("📁 変換されたファイルは vba/src-sjis/ フォルダにあります")
        
    except Exception as e:
        print(f"❌ 変換でエラーが発生しました: {e}")