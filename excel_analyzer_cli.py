#!/usr/bin/env python3
# /// script
# requires-python = ">=3.7.16"
# dependencies = [
#     "openpyxl>=3.1.0",
# ]
# ///
"""
Excel檔案分析器 - 使用uv script的獨立執行檔

這個腳本可以分析Excel檔案並修復常見的尺寸問題
使用方法: uv run excel_analyzer_cli.py [檔案路徑] [選項]
"""

import sys
import os
from pathlib import Path
import argparse
import shutil
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter

def analyze_sheet_size(sheet):
    """分析工作表的尺寸問題"""
    # 獲取實際使用的範圍
    actual_max_row = 0
    actual_max_col = 0
    cell_count = 0
    non_empty_cells = 0
    
    # 掃描前1000行來找實際內容（避免掃描時間過長）
    scan_limit = min(1000, sheet.max_row)
    
    for row_idx in range(1, scan_limit + 1):
        for col_idx in range(1, min(sheet.max_column + 1, 100)):  # 限制列數掃描
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell_count += 1
            
            if cell.value is not None and str(cell.value).strip():
                non_empty_cells += 1
                actual_max_row = max(actual_max_row, row_idx)
                actual_max_col = max(actual_max_col, col_idx)
    
    return {
        'reported_rows': sheet.max_row,
        'reported_cols': sheet.max_column,
        'actual_rows': actual_max_row,
        'actual_cols': actual_max_col,
        'scanned_cells': cell_count,
        'non_empty_cells': non_empty_cells,
        'has_size_issue': sheet.max_row > actual_max_row * 10  # 如果報告的行數比實際多10倍以上
    }

def backup_file(original_path):
    """建立備份檔案"""
    backup_path = original_path.with_suffix(f'.backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    shutil.copy2(original_path, backup_path)
    return backup_path

def fix_sheet_by_copy(workbook, sheet_name, actual_rows, actual_cols):
    """透過複製資料修復工作表尺寸問題"""
    old_sheet = workbook[sheet_name]
    
    # 建立新工作表
    new_sheet = workbook.create_sheet(f"{sheet_name}_fixed")
    
    # 複製實際有內容的資料
    for row_idx in range(1, actual_rows + 1):
        for col_idx in range(1, actual_cols + 1):
            old_cell = old_sheet.cell(row=row_idx, column=col_idx)
            new_cell = new_sheet.cell(row=row_idx, column=col_idx)
            
            # 複製值
            new_cell.value = old_cell.value
            
            # 複製基本格式
            if old_cell.font.bold:
                new_cell.font = openpyxl.styles.Font(bold=True)
            if old_cell.alignment.horizontal:
                new_cell.alignment = openpyxl.styles.Alignment(horizontal=old_cell.alignment.horizontal)
    
    # 獲取原工作表位置
    old_index = workbook.sheetnames.index(sheet_name)
    
    # 刪除原工作表並重新命名
    workbook.remove(old_sheet)
    new_sheet.title = sheet_name
    workbook.move_sheet(new_sheet, old_index)
    
    return True

def analyze_excel(file_path, fix_issues=False):
    """分析Excel檔案"""
    excel_path = Path(file_path)
    
    if not excel_path.exists():
        print(f"❌ 錯誤: 檔案 {excel_path} 不存在")
        return False
    
    print(f"📊 正在分析Excel檔案: {excel_path.name}")
    print(f"📁 檔案位置: {excel_path}")
    print(f"💾 檔案大小: {excel_path.stat().st_size / 1024 / 1024:.2f} MB")
    
    try:
        workbook = openpyxl.load_workbook(excel_path)
        
        print(f"\n📋 工作表列表 ({len(workbook.sheetnames)} 個):")
        
        problem_sheets = []
        total_issues = 0
        
        for i, sheet_name in enumerate(workbook.sheetnames, 1):
            sheet = workbook[sheet_name]
            analysis = analyze_sheet_size(sheet)
            
            status_icon = "🔴" if analysis['has_size_issue'] else "✅"
            print(f"  {i:2d}. {status_icon} {sheet_name:<20} - {analysis['reported_rows']:>8,} x {analysis['reported_cols']:>3} 列")
            
            if analysis['has_size_issue']:
                problem_sheets.append((sheet_name, analysis))
                total_issues += 1
                print(f"      ⚠️  實際內容: {analysis['actual_rows']} x {analysis['actual_cols']}")
                print(f"      📊 有效資料: {analysis['non_empty_cells']}/{analysis['scanned_cells']} 個儲存格")
        
        if problem_sheets:
            print(f"\n🚨 發現 {total_issues} 個工作表有尺寸問題:")
            for sheet_name, analysis in problem_sheets:
                wastage = analysis['reported_rows'] - analysis['actual_rows']
                print(f"  • {sheet_name}: 多了 {wastage:,} 個空白行")
        
        if fix_issues and problem_sheets:
            print(f"\n🔧 開始修復問題...")
            
            # 建立備份
            backup_path = backup_file(excel_path)
            print(f"💾 已建立備份: {backup_path.name}")
            
            # 修復問題工作表
            for sheet_name, analysis in problem_sheets:
                print(f"  🔨 修復 {sheet_name}...")
                fix_sheet_by_copy(workbook, sheet_name, analysis['actual_rows'], analysis['actual_cols'])
            
            # 儲存修復後的檔案
            fixed_path = excel_path.with_suffix('.fixed.xlsx')
            workbook.save(fixed_path)
            
            print(f"\n✅ 修復完成!")
            print(f"📁 修復後檔案: {fixed_path}")
            print(f"💾 檔案大小: {fixed_path.stat().st_size / 1024 / 1024:.2f} MB")
            print(f"💰 節省空間: {(excel_path.stat().st_size - fixed_path.stat().st_size) / 1024 / 1024:.2f} MB")
            
        elif not problem_sheets:
            print(f"\n✅ 所有工作表尺寸都正常，無需修復")
        
        workbook.close()
        return True
        
    except Exception as e:
        print(f"❌ 分析過程中發生錯誤: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description='Excel檔案分析器 - 檢測並修復工作表尺寸問題',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例:
  uv run excel_analyzer_cli.py file.xlsx              # 分析檔案
  uv run excel_analyzer_cli.py file.xlsx --fix       # 分析並修復問題
  
常見問題:
  - product工作表顯示100萬行但實際只有幾百行
  - 這通常是因為格式化或意外操作導致Excel認為有大量使用中的行
  - 本工具會複製實際內容到新工作表來解決此問題
        """
    )
    
    parser.add_argument('excel_file', help='Excel檔案路徑')
    parser.add_argument('--fix', action='store_true', help='自動修復發現的問題')
    parser.add_argument('--version', action='version', version='Excel Analyzer v1.0')
    
    if len(sys.argv) == 1:
        parser.print_help()
        return
    
    args = parser.parse_args()
    
    print("🔍 Excel 檔案分析器")
    print("=" * 50)
    
    success = analyze_excel(args.excel_file, args.fix)
    
    if not success:
        sys.exit(1)
    
    print("\n" + "=" * 50)
    print("分析完成! 如有問題請檢查上方報告。")

if __name__ == "__main__":
    main()
