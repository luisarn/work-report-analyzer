import os
import sys
import logging
import pandas as pd
from datetime import datetime, timedelta
from openai import OpenAI
import json
import re
from pathlib import Path
from dotenv import load_dotenv

# Logging setup
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# ==================== CONFIGURATION ====================
class Config:
    # Folder containing Excel files
    REPORT_FOLDER = os.getenv("REPORT_FOLDER", "./work_reports")
    
    # OpenAI Compatible API Settings
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
    OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
    MODEL_NAME = os.getenv("MODEL_NAME")

# ==================== ANALYZER CLASS ====================
class WorkReportAnalyzer:
    def __init__(self):
        try:
            self.client = OpenAI(
                api_key=Config.OPENAI_API_KEY,
                base_url=Config.OPENAI_BASE_URL
            )
        except Exception as e:
            logger.exception("Failed to initialize OpenAI client")
            raise
        self.results = []
        
    def get_check_info(self):
        """Determine next check date and which period/month to check."""
        today = datetime.now()
        
        if today.day < 15:
            next_check = datetime(today.year, today.month, 15)
            period = "上半月"
            sheet_month = today.month
        else:
            if today.month == 12:
                next_check = datetime(today.year + 1, 1, 1)
                sheet_month = 12
            else:
                next_check = datetime(today.year, today.month + 1, 1)
                sheet_month = today.month
            period = "下半月"
        
        # Calculate next Monday for the deadline
        days_until_monday = (7 - today.weekday()) % 7
        if days_until_monday == 0:
            days_until_monday = 7
        next_monday = today + timedelta(days=days_until_monday)
        
        return {
            'next_check': next_check,
            'period': period,
            'sheet_month': sheet_month,
            'deadline_monday': next_monday
        }

    def parse_filename(self, filename):
        """Extract team and staff name from filename."""
        name = Path(filename).stem
        team, staff = "Unknown", name
        
        for sep in ['_', '-', ' ']:
            if sep in name:
                parts = name.split(sep, 1)
                if len(parts) >= 2:
                    team = parts[0]
                    staff = parts[1]
                    break
        return team, staff

    def read_excel_data(self, file_path, sheet_month, current_period):
        """
        Read specific sheet and extract all relevant data including both periods.
        """
        try:
            logger.debug(f"讀取Excel文件: {file_path}")
            xl = pd.ExcelFile(file_path)
            
            logger.debug(f"可用的工作表: {xl.sheet_names}")
            
            # Find correct sheet - improved logic
            possible_names = [
                f"{sheet_month}月",
                f"Month {sheet_month}",
                f"{sheet_month}",
                f"第{sheet_month}月",
                f"{sheet_month}月份"
            ]
            
            # Also check for sheets that contain the month number
            sheet_name = None
            for name in possible_names:
                if name in xl.sheet_names:
                    sheet_name = name
                    logger.debug(f"找到匹配的工作表名稱: {name}")
                    break
            
            # If still not found, look for sheets that contain the month number
            if not sheet_name:
                for sheet in xl.sheet_names:
                    # Check if sheet name contains the month number
                    if str(sheet_month) in str(sheet):
                        sheet_name = sheet
                        logger.debug(f"找到包含月份的工作表: {sheet}")
                        break
            
            # Fallback to index-based approach
            if not sheet_name and len(xl.sheet_names) >= sheet_month:
                sheet_name = xl.sheet_names[sheet_month - 1]
                logger.debug(f"使用索引位置 {sheet_month-1} 的工作表: {sheet_name}")
            
            # Final fallback - use the first sheet
            if not sheet_name and len(xl.sheet_names) > 0:
                sheet_name = xl.sheet_names[0]
                logger.warning(f"使用第一個工作表 '{sheet_name}' 作為默認，因為未找到匹配的月份工作表")
            
            if not sheet_name:
                return None, f"Sheet for month {sheet_month} not found"
            
            logger.info(f"讀取工作表: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Find header rows
            logger.debug(f"數據框形狀: {df.shape}")
            header_row = None
            for idx, row in df.iterrows():
                row_values = [str(v) for v in row.values if pd.notna(v)]
                logger.debug(f"檢查行 {idx}: {row_values[:5]}")  # Log first 5 values
                if any('頂層項目' in v for v in row_values):
                    header_row = idx
                    logger.debug(f"找到標題行: {header_row}")
                    break
            
            if header_row is None:
                logger.warning("未找到標題行")
                return None, "Header row not found"
            
            # Log the type and value of header_row for debugging
            logger.debug(f"header_row 類型: {type(header_row)}, 值: {header_row}")
            
            # Extract data rows (skip 2 header rows)
            # header_row is guaranteed to be not None here
            # Use a fallback approach
            try:
                # Convert to string first, then to int
                header_row_index = int(str(header_row))
                data_df = df.iloc[header_row_index + 2:].copy()
            except (ValueError, TypeError) as e:
                logger.error(f"無法轉換標題行索引: {header_row}, 錯誤: {e}")
                # Fallback: use all data
                data_df = df.copy()
            
            # Determine column indices (handle variable formats)
            # Standard: 0=頂層項目, 1=項目, 2=目標, 3=上半月, 4=下半月
            # Sometimes merged: 3-4=上半月, 5-6=下半月
            n_cols = df.shape[1]
            
            tasks = []
            for _, row in data_df.iterrows():
                try:
                    top_item = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                    work_item = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                    objective = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                    
                    # Try to get both periods
                    first_half = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
                    second_half = str(row.iloc[4]) if n_cols > 4 and pd.notna(row.iloc[4]) else ""
                    
                    # Skip empty rows
                    if not any([top_item.strip(), work_item.strip(), objective.strip(), 
                               first_half.strip(), second_half.strip()]):
                        continue
                    
                    tasks.append({
                        'top_category': top_item.strip(),
                        'work_item': work_item.strip(),
                        'objective': objective.strip(),
                        'first_half_execution': first_half.strip(),
                        'second_half_execution': second_half.strip(),
                        'period_being_checked': current_period
                    })
                except Exception as e:
                    logger.debug(f"Skipping malformed row in {file_path}: {e}")
                    continue
            
            return tasks, None
            
        except Exception as e:
            logger.exception(f"Error reading Excel file {file_path}")
            return None, str(e)

    def analyze_compliance(self, tasks, staff_name, team, period):
        """
        Use LLM to check compliance with eased criteria.
        """
        if not tasks:
            return {
                "compliant": False,
                "summary": "未能找到此時期的工作項目，請確認是否已填寫工作內容。",
                "missing_elements": ["工作項目資料"],
                "violations": [],
                "score": "0",
                "all_others_warning": False
            }
        
        # Additional check for empty tasks (when tasks list exists but all items are empty)
        has_content = any(
            task.get('objective') or task.get('work_item') or task.get('top_category')
            for task in tasks
        )
        
        if not has_content:
            return {
                "compliant": False,
                "summary": "工作項目已建立但未填寫具體內容，請確認是否已填寫工作目標與執行情況。",
                "missing_elements": ["工作內容"],
                "violations": [],
                "score": "0",
                "all_others_warning": False
            }
        
        # Check if all tasks are "其他"
        all_others = all(
            task['top_category'] == '其他' or task['top_category'] == '' 
            for task in tasks if task['objective'] or task['work_item']
        )
        
        # Check if any real tasks exist (not just empty rows)
        has_real_tasks = any(
            task['objective'] or task['work_item'] 
            for task in tasks
        )
        
        tasks_json = json.dumps(tasks, ensure_ascii=False, indent=2)
        
        system_prompt = """你是一位友善且專業的部門主管，正在協助下屬改善工作報告的品質。
請以鼓勵和建設性的方式進行檢查，重點在於確保基本資訊完整。
絕對不能輸出「N/A」，必須提供具體、實用、有建設性的評語或建議。"""

        user_prompt = f"""請分析以下員工的工作報告：

【員工資料】
姓名: {staff_name}
團隊: {team}
檢查時期: {period}

【檢查標準 - 合理彈性原則】
1. 一般工作項目 (頂層項目不是"其他"):
   - 需要有基本的工作目標描述 (知道要做什麼)
   - 需要有目標日期或時間框架 (何時完成)
   - 不需要嚴格的SMART格式，只要有目標和時間即可

2. "其他"工作項目 (頂層項目是"其他"):
   - 只需有基本的工作描述即可，標準放寬
   - 不需要具體的完成日期或詳細產出定義

3. 執行工作欄位邏輯:
   - 檢查上半月時：如果「上半月已執行的工作」是空白，但「下半月已執行的工作」有內容，這是合理的（表示任務安排在下半月進行）
   - 只有當兩個時期欄位都空白時，才視為缺漏

4. 特殊警示條件:
   - 如果員工的所有工作項目都歸類為"其他"（沒有主要工作類別），這需要提醒

【工作內容資料】
{tasks_json}

【特別提醒】
- 所有項目都歸類為"其他": {"是" if all_others and has_real_tasks else "否"}

【輸出格式】
請輸出JSON格式，**絕對不能出現"N/A"**：
{{
  "compliant": true/false,
  "summary": "整體評語 (必須是建設性的建議或鼓勵，50-100字，絕對不能寫N/A)",
  "missing_elements": ["缺少的要素清單，如果沒有則留空陣列"],
  "violations": [
    {{"item": "工作項目名稱", "problem": "問題描述", "suggestion": "具體改進建議 (不能寫N/A)"}}
  ],
  "score": "1-10分",
  "all_others_warning": true/false (是否所有項目都是"其他"),
  "constructive_advice": "給主管的具體輔導建議，如何協助這位員工改進 (絕對不能寫N/A)"
}}

【評語範例】
- 合規時: "工作項目規劃完整，目標與時間明確。建議可在'其他'項目中補充更多背景說明，讓主管更了解工作脈絡。"
- 不合規時: "建議將'完成日常工作'改寫為具體事項，例如'完成每月報表整理並於5日前提交'，這樣更能展現工作價值。"
- 全是其他時: "目前所有工作都歸在'其他'類別，建議與員工確認主要職責範圍，將核心工作獨立分類，使報告結構更清晰。"""

        try:
            response = self.client.chat.completions.create(
                model=Config.MODEL_NAME,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,  # Slightly higher for more natural language
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            
            # Post-processing to ensure no N/A
            def clean_na(text):
                if not text or text.upper() in ['N/A', 'NA', '無', 'NONE', 'NULL']:
                    return "建議持續關注工作進展，並在下次回顧時確認具體產出。"
                return text
            
            result['summary'] = clean_na(result.get('summary', ''))
            result['constructive_advice'] = clean_na(result.get('constructive_advice', ''))
            
            # Ensure all_others_warning is set correctly
            if all_others and has_real_tasks:
                result['all_others_warning'] = True
                if result.get('compliant'):  # Force non-compliant if all others
                    result['compliant'] = False
                    result['summary'] = "雖然填寫內容完整，但所有工作都歸類為'其他'。建議區分主要工作職責與臨時事項，讓工作重點更明確。"
            
            return result
            
        except Exception as e:
            logger.exception("LLM analysis failed")
            return {
                "compliant": False,
                "summary": f"分析過程遇到技術問題，建議手動檢查此份報告。錯誤訊息：{str(e)}",
                "missing_elements": ["系統檢查"],
                "violations": [],
                "score": "0",
                "all_others_warning": all_others,
                "constructive_advice": "請直接與員工確認工作內容，確認基本資訊（目標與時間）是否已填寫。"
            }

    def generate_digest(self, check_info):
        """Generate the final report."""
        non_compliant = [r for r in self.results if not r['analysis'].get('compliant', False)]
        all_others_cases = [r for r in self.results if r['analysis'].get('all_others_warning', False)]
        
        deadline_str = check_info['deadline_monday'].strftime('%m/%d')
        period_str = f"{check_info['sheet_month']}月{check_info['period']}"
        
        report_lines = [
            "=" * 70,
            "工作報告檢查摘要 Work Report Review Digest",
            "=" * 70,
            f"生成時間 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"檢查時期 Review Period: {period_str}",
            f"截止日期 Deadline: 下週一({deadline_str}) 中午前上載",
            "-" * 70,
            f"總檔案數 Total: {len(self.results)} | 需改進 Need Improvement: {len(non_compliant)}",
            "=" * 70,
            ""
        ]
        
        if not non_compliant:
            report_lines.append("✓ 所有工作報告基本符合要求，請參閱個別建議進行優化。")
        else:
            report_lines.append("【需關注名單 Attention Required】")
            report_lines.append("")
            
            for item in non_compliant:
                analysis = item['analysis']
                report_lines.append(f"► {item['staff']} ({item['team']}) - {item['filename']}")
                
                # Special warning for all others
                if analysis.get('all_others_warning'):
                    report_lines.append("  ⚠️ 特別提醒：所有工作項目均為「其他」類別")
                
                report_lines.append(f"  評語: {analysis.get('summary', '')}")
                
                if analysis.get('violations') and len(analysis['violations']) > 0:
                    report_lines.append("  具體建議:")
                    for v in analysis['violations'][:2]:
                        if v.get('suggestion'):
                            report_lines.append(f"    • {v.get('item', '工作項目')}: {v['suggestion']}")
                
                if analysis.get('constructive_advice'):
                    report_lines.append(f"  輔導建議: {analysis['constructive_advice']}")
                
                report_lines.append("")
            
            # Action required section
            report_lines.append("-" * 70)
            report_lines.append("【行動要求 Action Required】")
            report_lines.append("")
            report_lines.append("請進行同事工作項目的適當分配後，要求同事填寫(檢視)清晰的工作目標，")
            report_lines.append("特別提醒目標應包括完成時間及基本產出說明即可，無需過於嚴格格式。")
            
            if all_others_cases:
                report_lines.append("")
                report_lines.append("【分類提醒】以下員工所有工作均為「其他」類別，請協助區分主要職責：")
                for item in all_others_cases:
                    report_lines.append(f"  - {item['staff']} ({item['team']})")
            
            report_lines.append("")
            report_lines.append(f"下週一({deadline_str}) 中午前完成並上載")
            report_lines.append("-" * 70)
        
        return "\n".join(report_lines)

    def run(self):
        """Main execution flow."""
        logger.info("開始分析工作報告 Starting analysis...")
        
        check_info = self.get_check_info()
        logger.info(f"檢查時期: {check_info['sheet_month']}月 - {check_info['period']}")
        logger.info("-" * 50)
        
        excel_files = [f for f in os.listdir(Config.REPORT_FOLDER) 
                      if f.endswith(('.xlsx', '.xls', '.xlsm'))]
        
        if not excel_files:
            logger.warning(f"未找到Excel檔案於 {Config.REPORT_FOLDER}")
            return
        
        for filename in excel_files:
            logger.info(f"處理中: {filename}...")
            
            team, staff = self.parse_filename(filename)
            # Use Path for better cross-platform compatibility and encoding handling
            file_path = Path(Config.REPORT_FOLDER) / filename
            
            tasks, error = self.read_excel_data(
                file_path, 
                check_info['sheet_month'], 
                check_info['period']
            )
            
            if error:
                logger.error(f"  讀取錯誤: {error}")
                self.results.append({
                    'filename': filename,
                    'staff': staff,
                    'team': team,
                    'analysis': {
                        'compliant': False,
                        'summary': f'無法讀取檔案：{error}，請確認Excel格式正確。',
                        'all_others_warning': False,
                        'constructive_advice': '請檢查檔案是否損壞，或嘗試重新儲存為標準Excel格式(.xlsx)。'
                    }
                })
                continue
            
            analysis = self.analyze_compliance(tasks, staff, team, check_info['period'])
            
            self.results.append({
                'filename': filename,
                'staff': staff,
                'team': team,
                'tasks_count': len(tasks) if tasks else 0,
                'analysis': analysis
            })
            
            status = "✓ OK" if analysis.get('compliant') else "✗ 需改進"
            warning = " [全其他]" if analysis.get('all_others_warning') else ""
            logger.info(f"  結果: {status}{warning}")
        
        digest = self.generate_digest(check_info)
        
        report_filename = f"digest_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        # Use Path for better cross-platform compatibility and encoding handling
        report_path = Path(Config.REPORT_FOLDER) / report_filename
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(digest)
        
        logger.info("\n" + "=" * 50)
        logger.info(f"分析完成！報告已儲存至: {report_path}")
        logger.info("=" * 50)
        logger.info("報告內容:\n%s", digest)

if __name__ == "__main__":
    if not os.path.exists(Config.REPORT_FOLDER):
        os.makedirs(Config.REPORT_FOLDER)
        logger.info(f"已創建資料夾: {Config.REPORT_FOLDER}")
        logger.info("請將Excel檔案放入此資料夾")
    else:
        # Validate required environment variables
        if not Config.OPENAI_API_KEY or not Config.MODEL_NAME:
            logger.error("OPENAI_API_KEY and MODEL_NAME must be set in the environment. Exiting.")
            logger.info("請設置環境變量或在 .env 文件中配置")
            sys.exit(1)
        
        analyzer = WorkReportAnalyzer()
        analyzer.run()