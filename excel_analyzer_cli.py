#!/usr/bin/env python3
# /// script
# requires-python = ">=3.7.16"
# dependencies = [
#     "openpyxl>=3.1.0",
#     "loguru>=0.6.0",
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
from loguru import logger
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
        'has_size_issue': (sheet.max_row > actual_max_row * 5 and sheet.max_row > 100) or (sheet.max_column > actual_max_col * 5 and sheet.max_column > 50)  # 檢測行和列的異常
    }

def backup_file(original_path):
    """建立備份檔案"""
    backup_path = original_path.with_suffix(f'.backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    shutil.copy2(original_path, backup_path)
    return backup_path

def fix_sheet_by_copy(workbook, sheet_name, actual_rows, actual_cols):
    """透過複製資料修復工作表尺寸問題"""
    old_sheet = workbook[sheet_name]
    
    # 確保至少複製基本行列數
    safe_rows = max(actual_rows, 10) if actual_rows > 0 else 10
    safe_cols = max(actual_cols, 10) if actual_cols > 0 else 10
    
    # 建立新工作表
    new_sheet = workbook.create_sheet(f"{sheet_name}_fixed")
    
    # 複製實際有內容的資料
    for row_idx in range(1, safe_rows + 1):
        for col_idx in range(1, safe_cols + 1):
            try:
                old_cell = old_sheet.cell(row=row_idx, column=col_idx)
                new_cell = new_sheet.cell(row=row_idx, column=col_idx)
                
                # 複製值
                new_cell.value = old_cell.value
                
                # 複製基本格式
                if hasattr(old_cell, 'font') and old_cell.font and old_cell.font.bold:
                    new_cell.font = openpyxl.styles.Font(bold=True)
                if hasattr(old_cell, 'alignment') and old_cell.alignment and old_cell.alignment.horizontal:
                    new_cell.alignment = openpyxl.styles.Alignment(horizontal=old_cell.alignment.horizontal)
            except Exception as e:
                # 如果某個儲存格複製失敗，跳過並繼續
                continue
    
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
        logger.error(f"檔案 {excel_path} 不存在")
        return None
    
    logger.info(f"正在分析Excel檔案: {excel_path.name}")
    logger.info(f"檔案位置: {excel_path}")
    logger.info(f"檔案大小: {excel_path.stat().st_size / 1024 / 1024:.2f} MB")
    
    try:
        workbook = openpyxl.load_workbook(excel_path)
        
        logger.info(f"工作表列表 ({len(workbook.sheetnames)} 個):")
        
        problem_sheets = []
        total_issues = 0
        
        for i, sheet_name in enumerate(workbook.sheetnames, 1):
            sheet = workbook[sheet_name]
            analysis = analyze_sheet_size(sheet)
            
            status = "問題" if analysis['has_size_issue'] else "正常"
            logger.info(f"  {i:2d}. {status} {sheet_name:<20} - {analysis['reported_rows']:>8,} x {analysis['reported_cols']:>3} 列")
            
            if analysis['has_size_issue']:
                problem_sheets.append((sheet_name, analysis))
                total_issues += 1
                logger.debug(f"      實際內容: {analysis['actual_rows']} x {analysis['actual_cols']}")
                logger.debug(f"      有效資料: {analysis['non_empty_cells']}/{analysis['scanned_cells']} 個儲存格")
                
                # 詳細說明問題類型
                row_issue = analysis['reported_rows'] > analysis['actual_rows'] * 5 and analysis['reported_rows'] > 100
                col_issue = analysis['reported_cols'] > analysis['actual_cols'] * 5 and analysis['reported_cols'] > 50
                if row_issue and col_issue:
                    logger.debug(f"      行列都有問題: 多了 {analysis['reported_rows'] - analysis['actual_rows']:,} 行, {analysis['reported_cols'] - analysis['actual_cols']} 列")
                elif row_issue:
                    logger.debug(f"      空白行問題: 多了 {analysis['reported_rows'] - analysis['actual_rows']:,} 行")
                elif col_issue:
                    logger.debug(f"      空白列問題: 多了 {analysis['reported_cols'] - analysis['actual_cols']} 列")
        
        if problem_sheets:
            logger.info(f"發現 {total_issues} 個工作表有尺寸問題:")
            for sheet_name, analysis in problem_sheets:
                wastage = analysis['reported_rows'] - analysis['actual_rows']
                logger.info(f"  • {sheet_name}: 多了 {wastage:,} 個空白行")
        
        if fix_issues and problem_sheets:
            logger.info("開始修復問題...")
            
            # 建立備份
            backup_path = backup_file(excel_path)
            logger.info(f"已建立備份: {backup_path.name}")
            
            # 修復問題工作表
            for sheet_name, analysis in problem_sheets:
                logger.info(f"修復 {sheet_name}...")
                fix_sheet_by_copy(workbook, sheet_name, analysis['actual_rows'], analysis['actual_cols'])
            
            # 儲存修復後的檔案
            fixed_path = excel_path.with_suffix('.fixed.xlsx')
            workbook.save(fixed_path)
            
            logger.info("修復完成!")
            logger.debug(f"修復後檔案: {fixed_path}")
            logger.debug(f"檔案大小: {fixed_path.stat().st_size / 1024 / 1024:.2f} MB")
            logger.debug(f"節省空間: {(excel_path.stat().st_size - fixed_path.stat().st_size) / 1024 / 1024:.2f} MB")
            
            workbook.close()
            return str(fixed_path.resolve())
            
        elif not problem_sheets:
            logger.info("所有工作表尺寸都正常，無需修復")
        
        workbook.close()
        return str(excel_path.resolve())
        
    except Exception as e:
        logger.error(f"分析過程中發生錯誤: {e}")
        return None

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
    parser.add_argument('--debug', action='store_true', help='啟用詳細除錯訊息')
    parser.add_argument('--version', action='version', version='Excel Analyzer v1.0')
    
    if len(sys.argv) == 1:
        parser.print_help()
        return
    
    args = parser.parse_args()
    
    # 設置日誌配置
    if not args.debug:
        logger.remove()
        logger.add(sys.stderr, level="WARNING")
    
    result = analyze_excel(args.excel_file, args.fix)
    
    if result is None:
        sys.exit(1)
    
    # 在標準終端輸出最終路徑
    print(result)
    

if __name__ == "__main__":
    main()
