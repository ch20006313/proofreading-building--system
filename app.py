from flask import Flask, request, render_template_string, jsonify
import openpyxl
import time
import uuid
import math
import re
import gc
import ast
import operator
import datetime
import threading
import os
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 允許上傳最大 100 MB 檔案

# 資料庫 (全域變數)
history_db = {}
tasks_db = {}

# ==============================================================================
# 🏗️ 【建物調查估價表 - 專屬檢核參數設定區】
# ==============================================================================
BUILDING_RULES = {
    "default_max_row": 85,          # 若未設定列印範圍，預設最多檢核列數
    "data_start_row": 11,           # 明細資料開始列數 (第 11 列起)
    # 視覺化圖面顯示欄位：僅保留 B~L 與 N~Q (隱藏 A=1, M=13, R=18, S=19, T=20)
    "visible_cols": [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17],
    "enable_style_check": False,    # 若未來需要字體/框線檢核可改為 True
}

# ==============================================================================
# 📚 【需求 5：建物項目單價與單位資料庫】
# 格式：(C欄「構造/規格」, D欄「用途」): {"unit": "K欄單位", "price": H欄單價}
# 後續您可直接在此往下新增條目 (比對時系統會自動去除換行符號與前後空白)
# ==============================================================================
BUILDING_DB = {
    ("水泥地面", ""): {"unit": "㎡", "price": 480},
    # 以下為範例格式，您可自行取消註解或新增：
    # ("砌塊石圍牆(厚度35cm)", "擋土牆"): {"unit": "㎡", "price": 572},
    # ("鋼筋混凝土", "擋土牆"): {"unit": "㎡", "price": 6427},
    # ("深水井管徑10吋", ""): {"unit": "M", "price": 2310},
    # ("手動大門遷移費", ""): {"unit": "㎡", "price": 4100},
}

def custom_round(value, ndigits=0):
    """標準四捨五入 (避免 Python 銀行家捨入誤差)"""
    multiplier = 10 ** ndigits
    # 加入極小值 1e-12 避免二進位浮點數誤差 (例如 2.5 存成 2.4999999999999996)
    if value >= 0:
        rounded_val = math.floor(value * multiplier + 0.5 + 1e-12) / multiplier
    else:
        rounded_val = math.ceil(value * multiplier - 0.5 - 1e-12) / multiplier
    if ndigits == 0:
        return int(rounded_val)
    return rounded_val

def get_cell_value_safe(ws, row, col, merged_ranges):
    """安全讀取儲存格內容（若位於合併儲存格內，自動抓取左上角主格的值）"""
    for mr in merged_ranges:
        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
            return ws.cell(row=min_row, column=min_col).value
    return ws.cell(row=row, column=col).value

def is_pure_number(val):
    """判斷是否為純數值 (排除 None、bool、datetime、文字)"""
    if val is None or isinstance(val, (bool, datetime.date, datetime.datetime)):
        return False
    if isinstance(val, (int, float)):
        return not math.isnan(val)
    try:
        s = str(val).replace(',', '').strip()
        if s == "":
            return False
        float(s)
        return True
    except:
        return False

def to_float(val):
    """將數值或千分位字串安全轉為 float"""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = str(val).replace(',', '').strip()
    return float(s)

def parse_rate_value(val):
    """解析 L 欄發放比例 (支援 0.7、1、1.0 或 '70%'、'100%')"""
    if val is None or isinstance(val, (bool, datetime.date, datetime.datetime)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "":
        return None
    if s.endswith('%'):
        try:
            return float(s[:-1].strip()) / 100.0
        except:
            return None
    try:
        return float(s.replace(',', ''))
    except:
        return None

# 安全算術 AST 求值器 (絕不使用 eval)
_ALLOWED_BINOPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
}
_ALLOWED_UNARYOPS = {
    ast.UAdd: operator.pos,
    ast.USub: operator.neg,
}

def _safe_eval_node(node):
    if isinstance(node, ast.Expression):
        return _safe_eval_node(node.body)
    elif isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return float(node.value)
    elif isinstance(node, ast.BinOp) and type(node.op) in _ALLOWED_BINOPS:
        left = _safe_eval_node(node.left)
        right = _safe_eval_node(node.right)
        if isinstance(node.op, ast.Div) and abs(right) < 1e-12:
            raise ZeroDivisionError("除數不可為 0")
        return _ALLOWED_BINOPS[type(node.op)](left, right)
    elif isinstance(node, ast.UnaryOp) and type(node.op) in _ALLOWED_UNARYOPS:
        return _ALLOWED_UNARYOPS[type(node.op)](_safe_eval_node(node.operand))
    else:
        raise ValueError("包含不合法的運算結構")

def validate_and_eval_col_i(raw_val):
    """
    需求 2：檢核 I 欄計算式是否合理並計算結果。
    回傳 (is_valid: bool, result_or_None: float, error_msg: str)
    """
    if raw_val is None:
        return True, None, ""
    if isinstance(raw_val, (int, float)) and not isinstance(raw_val, bool):
        return True, float(raw_val), ""

    expr = str(raw_val).strip()
    if expr == "":
        return True, None, ""

    # 若開頭有 '='，移除後再檢核 (若含有儲存格參照字母則跳過純算式檢核)
    if expr.startswith('='):
        expr_body = expr[1:].strip()
        if re.search(r'[A-Za-z]', expr_body):
            return True, None, ""
        expr = expr_body

    # 移除空白
    compact = re.sub(r'\s+', '', expr)
    if compact == "":
        return True, None, ""

    # 1. 檢查全形運算符號或非合法字元
    fullwidth_match = re.findall(r'[×÷＋－＊／（）]', compact)
    if fullwidth_match:
        return False, None, f"計算式包含全形符號 ({''.join(set(fullwidth_match))})，請改為半形"

    invalid_chars = re.findall(r'[^0-9\.\+\-\*\/\(\)]', compact)
    if invalid_chars:
        return False, None, f"計算式包含非合法字元 ({''.join(set(invalid_chars))})"

    # 2. 檢查連續運算符號 (例如 **, +*, //, ++, *-, /+)
    dup_op = re.search(r'[\+\-\*\/]{2,}', compact)
    if dup_op:
        return False, None, f"計算式不合理：出現連續運算符號 ({dup_op.group()})"

    # 3. 檢查開頭或結尾為運算符號 (加減乘除後方無數字等)
    if re.match(r'^[\+\*\/]', compact):
        return False, None, f"計算式不合理：開頭不可為運算符號 ({compact[0]})"
    if re.search(r'[\+\-\*\/]$', compact):
        return False, None, f"計算式不合理：運算符號 ({compact[-1]}) 後方缺少數字"

    # 4. 檢查括號成對與空括號、括號前後缺少運算符
    depth = 0
    for ch in compact:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth < 0:
                return False, None, "計算式不合理：括號順序錯誤或不成對"
    if depth != 0:
        return False, None, "計算式不合理：左右括號數量不成對"

    if '()' in compact:
        return False, None, "計算式不合理：出現空括號 ()"
    if re.search(r'\([\+\*\/]', compact):
        return False, None, "計算式不合理：左括號後方不可直接接運算符號"
    if re.search(r'[\+\-\*\/]\)', compact):
        return False, None, "計算式不合理：右括號前方缺少數字"
    if re.search(r'\d\(', compact) or re.search(r'\)\d', compact) or ')(' in compact:
        return False, None, "計算式不合理：括號與數字之間缺少乘號 (*)"

    # 5. 檢查不合法小數點 (例如 12..5、1.2.3、孤立小數點)
    if '..' in compact or re.search(r'\d+\.\d+\.', compact):
        return False, None, "計算式不合理：小數點格式錯誤"
    if re.search(r'(?<!\d)\.(?!\d)', compact):
        return False, None, "計算式不合理：出現孤立的小數點"

    # 6. AST 安全解析與求值
    try:
        tree = ast.parse(compact, mode='eval')
        val = _safe_eval_node(tree)
        return True, val, ""
    except ZeroDivisionError:
        return False, None, "計算式不合理：除數不可為 0"
    except Exception:
        return False, None, "計算式語法無法解析，請檢查公式格式"

def format_cell_display(val, col_idx, row_idx):
    """美化視覺化網格中的儲存格文字顯示"""
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, float):
        # L 欄發放比例若為 0~1 小數，以百分比或簡潔數字顯示
        if col_idx == 12 and row_idx >= 11 and 0 <= val <= 2:
            pct = custom_round(val * 100, 1)
            return f"{int(pct)}%" if pct == int(pct) else f"{pct}%"
        if val == int(val):
            s = f"{int(val):,}"
        else:
            s = f"{round(val, 4)}"
    elif isinstance(val, int) and not isinstance(val, bool):
        if col_idx == 12 and row_idx >= 11 and val in (0, 1):
            return f"{val * 100}%"
        s = f"{val:,}" if abs(val) >= 1000 and col_idx in (8, 14, 17) else str(val)
    else:
        s = str(val).replace('\n', ' ').strip()
    return s[:9] + '..' if len(s) > 9 else s

def extract_max_row_from_print_area(print_area):
    """解析單一或多段列印範圍 (如 '$B$1:$Q$27,$B$28:$Q$44') 取得最大列數"""
    if not print_area:
        return None
    if isinstance(print_area, (list, tuple)):
        areas = [str(a) for a in print_area]
    elif isinstance(print_area, str):
        areas = [a.strip() for a in print_area.split(',') if a.strip()]
    else:
        return None

    max_r = None
    for area in areas:
        part = area.split('!')[-1]
        end_coord = part.split(':')[1] if ':' in part else part
        m = re.search(r'\d+', end_coord)
        if m:
            r_num = int(m.group())
            if max_r is None or r_num > max_r:
                max_r = r_num
    return max_r

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>建築改良物調查估價表 - 智能視覺化校對系統</title>
    <style>
        body { font-family: '微軟正黑體', sans-serif; background-color: #f4f7f6; margin: 0; padding: 14px 16px; color: #333; scroll-behavior: smooth; }
        .header-title { text-align: center; color: #2c3e50; margin: 4px 0 16px 0; font-size: 22px; }
        /* 需求 1：放大整體版面寬度並縮小右側欄寬 */
        .layout-container { display: flex; gap: 16px; max-width: 98%; width: 100%; margin: auto; align-items: flex-start; }
        .main-content { flex: 1; background: #fff; padding: 22px 24px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: relative; min-width: 0; }
        .upload-box { border: 2px dashed #3498db; padding: 22px; text-align: center; border-radius: 10px; margin-bottom: 18px; background: #f8fbff; transition: 0.3s; }
        
        /* 載入動畫 */
        #loadingOverlay { display: none; text-align: center; padding: 40px 20px; }
        .spinner { border: 6px solid #f3f3f3; border-top: 6px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .progress-bar-container { width: 80%; background-color: #e0e0e0; border-radius: 8px; margin: 15px auto; height: 20px; overflow: hidden; }
        .progress-bar { height: 100%; background-color: #2ecc71; width: 0%; transition: width 0.4s ease; }
        
        .success { color: #27ae60; font-weight: bold; background: #e9f7ef; padding: 10px 14px; border-radius: 5px; }
        .special-warning { background-color: #f3e5f5 !important; color: #6a1b9a !important; border: 2px solid #9c27b0 !important; font-weight: bold; cursor: help; }
        .correction-panel { background: #fff3cd; border-left: 6px solid #ffc107; padding: 14px 18px; border-radius: 6px; margin-bottom: 20px; }
        .correction-panel h4 { margin-top: 0; margin-bottom: 10px; color: #856404; font-size: 15px; }
        .correction-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 4px; overflow: hidden; }
        .correction-table th, .correction-table td { padding: 8px 10px; border: 1px solid #ffeeba; text-align: center; font-size: 13px; }
        .correction-table th { background: #ffe8a1; color: #856404; }
        
        /* 結構探測器面板 */
        .inspector-box { background: #eef6ff; border: 1px solid #b8daff; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; }
        .inspector-toggle { background: #2980b9; color: #fff; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold; }
        .inspector-toggle:hover { background: #1f6391; }
        .inspector-content { display: none; margin-top: 10px; max-height: 300px; overflow-y: auto; background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; font-size: 12px; }
        .inspector-table { width: 100%; border-collapse: collapse; }
        .inspector-table th, .inspector-table td { border: 1px solid #ddd; padding: 5px 8px; text-align: left; }
        .inspector-table th { background: #f1f5f9; position: sticky; top: 0; }

        /* 需求 1：縮小右側瀏覽紀錄與按鈕寬度 (由 320px 縮減至 215px) */
        .history-panel { width: 215px; flex-shrink: 0; background: #fff; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: sticky; top: 16px; overflow: hidden; }
        .panel-actions { display: flex; gap: 6px; padding: 8px; background: #ecf0f1; border-bottom: 1px solid #dcdde1; }
        .action-btn { flex: 1; padding: 6px 4px; border: none; border-radius: 5px; font-size: 12px; font-weight: bold; cursor: pointer; transition: 0.2s; font-family: inherit; white-space: nowrap; }
        .btn-top { background: #2ecc71; color: #fff; }
        .btn-top:hover { background: #27ae60; }
        .btn-clear { background: #e74c3c; color: #fff; }
        .btn-clear:hover { background: #c0392b; }
        
        .history-header { background: #2c3e50; color: #fff; padding: 11px 12px; font-size: 14px; font-weight: bold; display: flex; justify-content: space-between; cursor: pointer; }
        .history-content { max-height: 75vh; overflow-y: auto; background: #fdfdfd; }
        .history-item { padding: 10px 10px; border-bottom: 1px solid #eee; display: flex; flex-direction: column; transition: background 0.2s; font-size: 13px; word-break: break-all; }
        .history-item:hover { background: #f0f8ff; }
        .history-item.active { background: #e1f0fa; border-left: 4px solid #2980b9; }
        .history-time { font-size: 11px; color: #7f8c8d; }
        
        /* 視覺化 Excel 網格：僅保留 B~L, N~Q，每欄更寬敞 */
        .map-container { overflow-x: auto; border-radius: 8px; border: 1px solid #ddd; margin-top: 12px; }
        .excel-map { border-collapse: collapse; width: 100%; table-layout: fixed; background: #fff; }
        .excel-map th { background-color: #ecf0f1; border: 1px solid #bdc3c7; font-size: 12px; padding: 6px 4px; position: sticky; top: 0; z-index: 2; }
        .excel-map th.row-num { width: 34px; }
        .excel-map td { border: 1px solid #e5e8e8; height: 32px; padding: 2px 4px; text-align: center; position: relative; font-size: 12px; overflow: visible; word-break: break-word; }
        .cell-valid { background-color: #eafaf1; color: #27ae60; }
        .cell-empty { background-color: #f9f9f9; color: #ccc; }
        .cell-error { background-color: #fadbd8; color: #c0392b; border: 2px solid #e74c3c !important; font-weight: bold; cursor: help; }
        .tooltip { position: absolute; background-color: rgba(0,0,0,0.9); color: #fff; padding: 8px 12px; border-radius: 6px; font-size: 12px; font-weight: normal; bottom: 118%; left: 50%; transform: translateX(-50%); white-space: nowrap; opacity: 0; visibility: hidden; z-index: 9999; pointer-events: none; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
        .excel-map tr:nth-child(-n+6) .tooltip { bottom: auto; top: 118%; }
        .excel-map td:nth-child(-n+3) .tooltip { left: 0; transform: none; }
        .excel-map td:nth-last-child(-n+3) .tooltip { left: auto; right: 0; transform: none; }
        .cell-error:hover .tooltip, .special-warning:hover .tooltip { opacity: 1; visibility: visible; }
        .sheet-container { border: 2px solid #e0e0e0; border-radius: 8px; padding: 18px; margin-bottom: 30px; background: #fafbfc; }
        .sheet-title { margin-top: 0; color: #34495e; border-bottom: 2px solid #bdc3c7; padding: 10px; border-radius: 4px; transition: 0.2s; scroll-margin-top: 20px; cursor: pointer; display: flex; justify-content: space-between; font-size: 16px; }
        .sheet-title:hover { background-color: #e2e8f0; }
        .meta-badges { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; font-size: 12px; }
        .badge { background: #eef2f7; color: #34495e; padding: 4px 10px; border-radius: 12px; border: 1px solid #d5dde5; }
    </style>
</head>
<body id="pageTop">
    <h2 class="header-title">🏢 建築改良物調查估價表 - 智能視覺化校對系統</h2>
    <div class="layout-container">
        <div class="main-content">
            
            <!-- 上傳區塊 -->
            <div class="upload-box" id="uploadBox">
                <form id="uploadForm">
                    <input type="file" name="excel_file" id="excelFile" accept=".xlsx, .xlsm" required>
                    <button type="submit" style="padding: 8px 20px; font-weight: bold; cursor: pointer; background: #2980b9; color: #fff; border: none; border-radius: 5px;">上傳建物估價表 (.xlsx / .xlsm) 並開始校對</button>
                </form>
            </div>

            <!-- 載入動畫區塊 -->
            <div id="loadingOverlay">
                <div class="spinner"></div>
                <h3 style="color: #2980b9;">建物估價表正在背景火速解析與驗算中，請稍候... (<span id="progressText">0%</span>)</h3>
                <div class="progress-bar-container"><div class="progress-bar" id="progressBar"></div></div>
                <p style="color: #e74c3c; font-weight: bold; font-size: 11pt;">⚠️ 若工作表數量龐大，可能需要數分鐘處理，請勿關閉或重新整理網頁！</p>
                <p id="currentSheetText" style="color: #7f8c8d; font-size: 13px;"></p>
            </div>

            <script>
                document.getElementById('uploadForm').onsubmit = async function(e) {
                    e.preventDefault();
                    let formData = new FormData(this);
                    
                    document.getElementById('uploadBox').style.display = 'none';
                    document.getElementById('loadingOverlay').style.display = 'block';
                    if(document.getElementById('reportContainer')) {
                        document.getElementById('reportContainer').style.display = 'none';
                    }

                    try {
                        let response = await fetch('/upload', { method: 'POST', body: formData });
                        let data = await response.json();
                        
                        if (data.task_id) {
                            checkTaskStatus(data.task_id);
                        } else {
                            alert("上傳失敗: " + (data.error || "未知錯誤"));
                            location.reload();
                        }
                    } catch (err) {
                        alert("網路連線錯誤，請重試。");
                        location.reload();
                    }
                };

                async function checkTaskStatus(taskId) {
                    try {
                        let response = await fetch('/status/' + taskId);
                        let data = await response.json();
                        
                        if (data.status === 'processing') {
                            document.getElementById('progressText').innerText = data.progress + "%";
                            document.getElementById('progressBar').style.width = data.progress + "%";
                            if (data.current_sheet) {
                                document.getElementById('currentSheetText').innerText = "目前正在處理: " + data.current_sheet;
                            }
                            setTimeout(() => checkTaskStatus(taskId), 2000);
                        } else if (data.status === 'completed') {
                            document.getElementById('progressText').innerText = "100%";
                            document.getElementById('progressBar').style.width = "100%";
                            window.location.href = '/?id=' + data.record_id;
                        } else if (data.status === 'error') {
                            alert("背景運算發生錯誤: " + data.message);
                            location.reload();
                        }
                    } catch (err) {
                        setTimeout(() => checkTaskStatus(taskId), 2000);
                    }
                }

                function toggleHistory() {
                    let content = document.getElementById('historyContent');
                    content.style.display = (content.style.display === 'none') ? 'block' : 'none';
                }

                function toggleSheet(sheetId) {
                    let content = document.getElementById('content-' + sheetId);
                    let icon = document.getElementById('icon-' + sheetId);
                    if (content.style.display === 'none') {
                        content.style.display = 'block';
                        icon.innerHTML = '▼';
                    } else {
                        content.style.display = 'none';
                        icon.innerHTML = '▲';
                    }
                }

                function toggleInspector(sheetId) {
                    let box = document.getElementById('inspector-' + sheetId);
                    box.style.display = (box.style.display === 'block') ? 'none' : 'block';
                }

                function scrollToTop() {
                    const uploadBox = document.getElementById('uploadBox');
                    if (uploadBox) {
                        uploadBox.style.display = 'block';
                    }
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }

                async function confirmClearHistory() {
                    if (confirm("⚠️ 確定要清除所有瀏覽紀錄與暫存資料嗎？\\n\\n清除後可釋放伺服器記憶體避免卡頓，但已校對的歷史紀錄將無法復原。")) {
                        try {
                            let response = await fetch('/clear_history', { method: 'POST' });
                            let data = await response.json();
                            if (data.status === 'cleared') {
                                window.location.href = '/';
                            }
                        } catch (err) {
                            alert("清除紀錄時發生錯誤，請重試。");
                        }
                    }
                }
            </script>

            <!-- 渲染報表區塊 -->
            {% if record %}
            <div id="reportContainer">
                <div class="file-info" style="font-size:17px; margin-bottom: 14px; font-weight: bold;">📄 本次校對完成：{{ record.filename }}</div>
                {% for sheet_name, s_data in record.sheets.items() %}
                {% set s_idx = loop.index %}
                <div class="sheet-container">
                    <h3 class="sheet-title" id="sheet-{{ s_idx }}" onclick="toggleSheet('{{ s_idx }}')">
                        <span>🏷️ 工作表：{{ sheet_name }} (列印範圍：{{ s_data.print_area_str }})</span>
                        <span id="icon-{{ s_idx }}">▼</span>
                    </h3>
                    <div id="content-{{ s_idx }}">
                        
                        <!-- 關鍵欄位判讀狀態摘要 -->
                        <div class="meta-badges">
                            <span class="badge">H1：{{ s_data.meta.h1_text or '(空)' }}</span>
                            <span class="badge">S24：{{ s_data.meta.s24_text or '(空)' }}</span>
                            <span class="badge">Q5 面積：{{ s_data.meta.q5_text or '(空)' }}</span>
                            <span class="badge">S/T 合計：{{ s_data.meta.st_total_text or '未偵測' }}</span>
                        </div>

                        <!-- 結構探測器 -->
                        <div class="inspector-box">
                            <button type="button" class="inspector-toggle" onclick="toggleInspector('{{ s_idx }}')">
                                🔍 展開/收合「{{ sheet_name }}」座標、字體與公式結構表 (共 {{ s_data.structure_info|length }} 格有值)
                            </button>
                            <div class="inspector-content" id="inspector-{{ s_idx }}">
                                <table class="inspector-table">
                                    <tr>
                                        <th>座標</th><th>顯示數值</th><th>原始公式 / 內容</th><th>字體</th><th>合併範圍</th>
                                    </tr>
                                    {% for item in s_data.structure_info %}
                                    <tr>
                                        <td><strong>{{ item.coord }}</strong></td>
                                        <td>{{ item.val }}</td>
                                        <td style="color:#2980b9; font-family:monospace;">{{ item.formula }}</td>
                                        <td>{{ item.font_size }}pt</td>
                                        <td>{{ item.merge_info }}</td>
                                    </tr>
                                    {% endfor %}
                                </table>
                            </div>
                        </div>

                        {% if s_data.excess_rows %}
                            <p style="color: #c0392b; font-weight: bold; background: #fadbd8; padding: 10px; border-radius: 5px; font-size: 11pt;">⚠️ 注意事項：本工作表未設定列印範圍且資料超過 {{ s_data.max_row }} 列，系統僅校對前 {{ s_data.max_row }} 列。</p>
                        {% endif %}

                        {% if s_data.correction_list %}
                        <div class="correction-panel">
                            <h4>⚠️ 發現「小計 / 合計 / S-T欄面積加總」與系統運算不符：</h4>
                            <table class="correction-table">
                                <tr><th>位置</th><th>檢核項目</th><th>上傳檔案數值</th><th>系統校驗正確數值</th></tr>
                                {% for c in s_data.correction_list %}
                                <tr>
                                    <td>第 {{ c.row }} 列 ({{ c.coord }})</td>
                                    <td>{{ c.type }}</td>
                                    <td style="color:#c0392b; text-decoration: line-through;">{{ c.original }}</td>
                                    <td style="color:#27ae60; font-weight:bold; font-size:15px;">{{ c.correct }}</td>
                                </tr>
                                {% endfor %}
                            </table>
                        </div>
                        {% endif %}
                        
                        {% if s_data.error_list|length == 0 and not s_data.correction_list %}
                            <p class="success">✅ 完美！本工作表所有算式、發放比例、單價單位、小計、合計與面積加總皆完全符合規定。</p>
                        {% else %}
                            <p style="margin-bottom:8px;">💡 滑鼠移至<span style="color:#c0392b; font-weight:bold;">紅色/紫色格子</span>可查看錯誤原因（共發現 {{ s_data.error_list|length }} 處異常；已隱藏 A、M、R、S 欄以放大視野）。</p>
                        {% endif %}

                        <!-- 視覺化圖面 (僅保留 B~L, N~Q) -->
                        <div class="map-container">
                            <table class="excel-map">
                                <tr>
                                    <th class="row-num"></th>
                                    {% for c in visible_cols %}
                                        <th>{{ get_col_letter(c) }}</th>
                                    {% endfor %}
                                </tr>
                                {% for r in range(1, s_data.max_row + 1) %}
                                <tr>
                                    <th class="row-num">{{ r }}</th>
                                    {% for c in visible_cols %}
                                        {% set cell = s_data.map_data.get((r, c), {'status': 'empty', 'text': '', 'error': ''}) %}
                                        <td class="{% if cell.status == 'special' %}special-warning{% else %}cell-{{ cell.status }}{% endif %}">
                                            {{ cell.text }}
                                            {% if cell.status == 'error' or cell.status == 'special' %}
                                                <div class="tooltip">{{ cell.error }}</div>
                                            {% endif %}
                                        </td>
                                    {% endfor %}
                                </tr>
                                {% endfor %}
                            </table>
                        </div>

                    </div>
                </div>
                {% endfor %}
            </div>
            {% endif %}
        </div>

        <div class="history-panel">
            <div class="panel-actions">
                <button type="button" class="action-btn btn-top" onclick="scrollToTop()">⬆️ 置頂</button>
                <button type="button" class="action-btn btn-clear" onclick="confirmClearHistory()">🗑️ 清除紀錄</button>
            </div>
            <div class="history-header" onclick="toggleHistory()">
                <span>📂 瀏覽紀錄</span><span>▼</span>
            </div>
            <div class="history-content" id="historyContent">
                {% for uid, hist in history.items()|reverse %}
                    <div class="history-item {% if uid == current_id %}active{% endif %}">
                        <div onclick="window.location.href='/?id={{ uid }}'" style="cursor: pointer; font-weight:bold; color: #2980b9;">
                            {{ hist.filename }} <br>
                            <span class="history-time">{{ hist.time }}</span>
                        </div>
                        <ul style="margin: 6px 0 0 0; padding-left: 16px; font-size: 12px; color: #555;">
                            {% for s_name, s_data in hist.sheets.items() %}
                                <li onclick="window.location.href='/?id={{ uid }}#sheet-{{ loop.index }}'" style="cursor: pointer; padding: 2px 0; text-decoration: underline;">
                                    {{ s_name }} 
                                    {% if s_data.error_list|length == 0 and not s_data.correction_list %}
                                        <span style="color: #27ae60; font-weight: bold;">(✅)</span>
                                    {% else %}
                                        <span style="color: #c0392b; font-weight: bold;">(❌ {{ s_data.error_list|length }} 錯誤)</span>
                                    {% endif %}
                                </li>
                            {% endfor %}
                        </ul>
                    </div>
                {% else %}
                    <div style="padding: 14px; text-align: center; color: #999; font-size: 12px;">尚無紀錄</div>
                {% endfor %}
            </div>
        </div>
    </div>
</body>
</html>
"""

# 背景處理工作緒
def background_process_excel(task_id, filepath, original_filename):
    try:
        gc.enable()

        wb_data = openpyxl.load_workbook(filepath, data_only=True)
        wb_formula = openpyxl.load_workbook(filepath, data_only=False)

        sheets_data = {}
        sheet_names = wb_data.sheetnames
        total_sheets = len(sheet_names)

        default_max_row = BUILDING_RULES["default_max_row"]
        data_start_row = BUILDING_RULES["data_start_row"]
        visible_cols = BUILDING_RULES["visible_cols"]

        # 預先建立去換行/空白版本的 BUILDING_DB 索引
        normalized_db = {}
        for (c_k, d_k), info in BUILDING_DB.items():
            norm_key = (
                str(c_k).replace('\n', '').replace('\r', '').strip(),
                str(d_k).replace('\n', '').replace('\r', '').strip()
            )
            normalized_db[norm_key] = info

        for idx, sheet_name in enumerate(sheet_names):
            progress = int((idx / total_sheets) * 100)
            tasks_db[task_id]['progress'] = progress
            tasks_db[task_id]['current_sheet'] = sheet_name

            ws_data = wb_data[sheet_name]
            ws_formula = wb_formula[sheet_name]

            error_list = []
            map_data = {}
            correction_list = []
            structure_info = []

            excess_rows = False
            print_area = ws_data.print_area
            print_area_str = str(print_area) if print_area else "未設定"

            parsed_max_row = extract_max_row_from_print_area(print_area)
            if parsed_max_row:
                max_row = parsed_max_row
            else:
                max_row = min(ws_data.max_row, default_max_row)
                excess_rows = ws_data.max_row > default_max_row

            merged_ranges = list(ws_data.merged_cells.ranges)

            # 準備全域列錯誤容器 (涵蓋到第 20 欄 T 欄)
            all_calc_errors = {r: {c: [] for c in range(1, 21)} for r in range(1, max_row + 1)}
            all_special_flags = {r: {c: False for c in range(1, 21)} for r in range(1, max_row + 1)}

            # ==================================================================
            # 需求 3：讀取 H1 (1, 8) 與 S24 (24, 19)
            # ==================================================================
            h1_raw = get_cell_value_safe(ws_data, 1, 8, merged_ranges)
            h1_str = str(h1_raw or "").strip()
            is_public_land = ("公有地" in h1_str)

            s24_raw = get_cell_value_safe(ws_data, 24, 19, merged_ranges)
            s24_str = str(s24_raw if s24_raw is not None else "").strip()
            s24_upper = s24_str.upper()

            # S24 包含「承租」或「FALSE」 -> 原則 1.0；不包含「承租」或「TRUE」 -> 0.5
            s24_is_leased_or_false = (
                (s24_raw is False)
                or (s24_upper == "FALSE")
                or ("承租" in s24_str)
            )

            # ==================================================================
            # 需求 7：驗算 S 欄「合計」之 T 欄加總 (T2:合計上一列) 並比對 Q5
            # ==================================================================
            q5_raw = get_cell_value_safe(ws_data, 5, 17, merged_ranges)
            q5_num = to_float(q5_raw) if is_pure_number(q5_raw) else None
            st_total_text = "未偵測"

            scan_st_max = min(max(max_row, ws_data.max_row), 200)
            for r_s in range(2, scan_st_max + 1):
                s_val = str(ws_data.cell(row=r_s, column=19).value or "").strip()
                if "合計" in s_val:
                    # 計算 T2 : T{r_s - 1} 的加總
                    calc_t_sum = 0.0
                    for r_t in range(2, r_s):
                        t_item = ws_data.cell(row=r_t, column=20).value
                        if is_pure_number(t_item):
                            calc_t_sum += to_float(t_item)
                    calc_t_sum = custom_round(calc_t_sum, 2)

                    t_total_raw = ws_data.cell(row=r_s, column=20).value
                    t_total_num = to_float(t_total_raw) if is_pure_number(t_total_raw) else None
                    st_total_text = f"T{r_s}={t_total_num:,.2f} (重算={calc_t_sum:,.2f})" if t_total_num is not None else f"T{r_s}=空白 (重算={calc_t_sum:,.2f})"

                    # (1) 檢核 T{r_s} 是否等於 T2:T{r_s-1} 加總
                    if t_total_num is None or abs(t_total_num - calc_t_sum) > 0.01:
                        orig_disp = f"{t_total_num:,.2f}" if t_total_num is not None else "(空白/非數字)"
                        correction_list.append({
                            "row": r_s,
                            "coord": f"T{r_s}",
                            "type": "S欄合計 (T2至上一列加總)",
                            "original": orig_disp,
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                        error_list.append(f"❌ T{r_s}: S欄合計對應之T欄數值({orig_disp})與 T2:T{r_s-1} 加總({calc_t_sum:,.2f})不符")

                    # (2) 比對 Q5 數值是否與 T 欄加總相符
                    if q5_num is None:
                        if 5 <= max_row:
                            all_calc_errors[5][17].append(f"Q5面積為空白或非數字，應與 T{r_s} 加總 ({calc_t_sum:,.2f}) 相符")
                        correction_list.append({
                            "row": 5,
                            "coord": "Q5",
                            "type": "Q5面積與T欄合計比對",
                            "original": "(空白/非數字)",
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                    elif abs(q5_num - calc_t_sum) > 0.01 or (t_total_num is not None and abs(q5_num - t_total_num) > 0.01):
                        if 5 <= max_row:
                            all_calc_errors[5][17].append(
                                f"Q5數值({q5_num:,.2f})與T欄合計(T{r_s}={t_total_num if t_total_num is not None else '空'} / 重算={calc_t_sum:,.2f})不符"
                            )
                        correction_list.append({
                            "row": 5,
                            "coord": "Q5",
                            "type": "Q5面積與T欄合計比對",
                            "original": f"{q5_num:,.2f}",
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                    break  # 找到第一個 S 欄合計即結束

            # ==================================================================
            # 需求 6 累計變數 (小計與合計追蹤)
            # ==================================================================
            current_block_n_sum = 0.0
            subtotal_list = []
            has_subtotal_seen = False

            # ==================================================================
            # 逐列進行規則檢核 (第 1 列 ~ max_row)
            # ==================================================================
            for row in range(1, max_row + 1):
                if row < data_start_row:
                    continue

                col_B_val = str(ws_data.cell(row=row, column=2).value or "").strip()
                col_C_val = str(ws_data.cell(row=row, column=3).value or "").strip()
                col_D_val = str(ws_data.cell(row=row, column=4).value or "").strip()
                col_O_val = str(get_cell_value_safe(ws_data, row, 15, merged_ranges) or "").strip()

                # 判斷該列是否為頁尾/注意事項/日期列/以下空白列
                row_text_sample = "".join(
                    str(ws_data.cell(row=row, column=c).value or "") for c in [2, 3, 12, 14]
                )
                is_footer_row = any(kw in row_text_sample for kw in ["注意事項", "調查單位", "調查日期", "以下空白"])

                is_subtotal_row = ("小計" in col_B_val)
                is_total_row = ("合計" in col_B_val) and not is_subtotal_row

                # ---------------- 【需求 6：小計與合計驗算】 ----------------
                if is_subtotal_row:
                    n_raw = ws_data.cell(row=row, column=14).value
                    n_val_num = to_float(n_raw) if is_pure_number(n_raw) else 0.0
                    expected_sub = custom_round(current_block_n_sum, 0)

                    if abs(n_val_num - expected_sub) > 0.5:
                        correction_list.append({
                            "row": row,
                            "coord": f"N{row}",
                            "type": f"B{row} {col_B_val}",
                            "original": f"{n_val_num:,.0f}",
                            "correct": f"{expected_sub:,.0f}"
                        })
                        all_calc_errors[row][14].append(f"小計金額錯誤，上方區段加總應為 {expected_sub:,}")

                    subtotal_list.append(expected_sub)
                    has_subtotal_seen = True
                    current_block_n_sum = 0.0  # 歸零，下一區段不重複累加
                    continue

                if is_total_row:
                    n_raw = ws_data.cell(row=row, column=14).value
                    n_val_num = to_float(n_raw) if is_pure_number(n_raw) else 0.0

                    if has_subtotal_seen:
                        expected_tot = custom_round(sum(subtotal_list), 0)
                        tot_desc = "各小計加總"
                    else:
                        expected_tot = custom_round(current_block_n_sum, 0)
                        tot_desc = f"N11:N{row - 1} 加總"

                    if abs(n_val_num - expected_tot) > 0.5:
                        correction_list.append({
                            "row": row,
                            "coord": f"N{row}",
                            "type": f"B{row} {col_B_val} ({tot_desc})",
                            "original": f"{n_val_num:,.0f}",
                            "correct": f"{expected_tot:,.0f}"
                        })
                        all_calc_errors[row][14].append(f"合計金額錯誤，{tot_desc}應為 {expected_tot:,}")

                    # 結算後重置
                    current_block_n_sum = 0.0
                    subtotal_list = []
                    has_subtotal_seen = False
                    continue

                # 若為注意事項/調查日期/以下空白列，忽略不納入加總與明細檢核
                if is_footer_row:
                    continue

                # 累加本區段 N 欄數值 (供下一筆小計或合計驗算使用)
                n_cell_val = ws_data.cell(row=row, column=14).value
                if is_pure_number(n_cell_val):
                    current_block_n_sum += to_float(n_cell_val)

                # 判斷 B 欄是否有編號
                has_col_B_id = (
                    col_B_val != ""
                    and not any(kw in col_B_val for kw in ["編號", "項目", "小計", "合計", "總計", "注意", "調查", "以下空白", "備註"])
                )

                # ---------------- 【需求 5：C 欄 & D 欄 資料庫檢核 H 欄單價、K 欄單位】 ----------------
                norm_key = (
                    col_C_val.replace('\n', '').replace('\r', '').strip(),
                    col_D_val.replace('\n', '').replace('\r', '').strip()
                )
                if norm_key in normalized_db:
                    db_item = normalized_db[norm_key]
                    h_raw = ws_data.cell(row=row, column=8).value
                    k_raw = str(ws_data.cell(row=row, column=11).value or "").strip()

                    if not is_pure_number(h_raw):
                        all_calc_errors[row][8].append(f"單價空白或非數字，資料庫標準應為 {db_item['price']}")
                    elif abs(to_float(h_raw) - db_item["price"]) > 0.001:
                        all_calc_errors[row][8].append(f"單價與資料庫不符，應為 {db_item['price']}")

                    if k_raw != db_item["unit"]:
                        all_calc_errors[row][11].append(f"單位與資料庫不符，應為「{db_item['unit']}」")

                # ---------------- 【需求 2：檢核 I 欄計算式合理性，並驗算 J 欄為 I 欄四捨五入結果】 ----------------
                i_formula_raw = ws_formula.cell(row=row, column=9).value
                i_data_raw = ws_data.cell(row=row, column=9).value
                i_check_target = i_formula_raw if i_formula_raw is not None else i_data_raw

                if i_check_target is not None and str(i_check_target).strip() != "":
                    i_valid, i_eval_val, i_err_msg = validate_and_eval_col_i(i_check_target)
                    if not i_valid:
                        all_calc_errors[row][9].append(i_err_msg)
                    elif i_eval_val is not None:
                        expected_j = custom_round(i_eval_val, ndigits=2)
                        j_raw = ws_data.cell(row=row, column=10).value
                        if not is_pure_number(j_raw):
                            all_calc_errors[row][10].append(f"數量空白或非數字，依 I 欄計算應為 {expected_j}")
                        else:
                            j_num = to_float(j_raw)
                            if abs(j_num - expected_j) > 0.005:
                                all_calc_errors[row][10].append(f"數量驗算錯誤，I 欄四捨五入後應為 {expected_j}")

                # ---------------- 【需求 3：H1 含「公有地」與 S24 連動檢核 L 欄】 ----------------
                if is_public_land and has_col_B_id:
                    has_illegal_or_no_water = ("無合法" in col_O_val) or ("無水權" in col_O_val)
                    if s24_is_leased_or_false:
                        if has_illegal_or_no_water:
                            expected_l = 0.7
                            l_reason = "O欄含無合法/無水權，L欄應為 0.7 (70%)"
                        else:
                            expected_l = 1.0
                            l_reason = "H1含公有地且S24含承租/FALSE，L欄應為 1.0 (100%)"
                    else:
                        expected_l = 0.5
                        l_reason = "H1含公有地且S24不含承租/TRUE，L欄應為 0.5 (50%)"

                    l_raw = ws_data.cell(row=row, column=12).value
                    l_rate_parsed = parse_rate_value(l_raw)
                    if l_rate_parsed is None:
                        all_calc_errors[row][12].append(f"發放比例空白或非數字 ({l_reason})")
                    elif abs(l_rate_parsed - expected_l) > 0.001:
                        all_calc_errors[row][12].append(f"發放比例錯誤 ({l_reason})")

                # ---------------- 【需求 4：驗算 N 欄 = round(round(H * J, 0) * L, 0)】 ----------------
                h_raw = ws_data.cell(row=row, column=8).value
                j_raw = ws_data.cell(row=row, column=10).value
                l_raw = ws_data.cell(row=row, column=12).value

                # 僅針對有填寫 H 欄與 J 欄的項目驗算 (跳過房屋評點列如第 11 列)
                if is_pure_number(h_raw) and is_pure_number(j_raw):
                    h_num = to_float(h_raw)
                    j_num = to_float(j_raw)
                    l_num = parse_rate_value(l_raw)
                    if l_num is None:
                        all_calc_errors[row][14].append("無法驗算金額：L 欄發放比例空白或非數字")
                    else:
                        step1 = custom_round(h_num * j_num, ndigits=0)
                        expected_n = custom_round(step1 * l_num, ndigits=0)
                        if not is_pure_number(n_cell_val):
                            all_calc_errors[row][14].append(f"金額空白或非數字，應為 {expected_n:,}")
                        else:
                            actual_n = to_float(n_cell_val)
                            if abs(actual_n - expected_n) > 0.1:
                                all_calc_errors[row][14].append(
                                    f"金額驗算錯誤，應為 {expected_n:,} (H×J四捨五入={step1:,}，再×L四捨五入)"
                                )

            # ==================================================================
            # 組合視覺化網格 (僅渲染 visible_cols: B~L, N~Q) 與結構探測表
            # ==================================================================
            for row in range(1, max_row + 1):
                for col in visible_cols:
                    cell_formula = ws_formula.cell(row=row, column=col)
                    cell_data = ws_data.cell(row=row, column=col)
                    coord = cell_formula.coordinate

                    is_phantom = False
                    merge_range_str = "無"
                    for mr in merged_ranges:
                        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
                        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
                            if row == min_row and col == min_col:
                                merge_range_str = str(mr)
                            else:
                                is_phantom = True
                            break

                    errors = all_calc_errors[row][col].copy()
                    is_special = all_special_flags[row][col]

                    if is_phantom:
                        if errors:
                            err_msg = "、".join(errors)
                            map_data[(row, col)] = {'status': 'error', 'text': '(合併)', 'error': err_msg}
                            error_list.append(f"❌ {coord}: {err_msg}")
                        else:
                            map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    if cell_data.value is None and cell_formula.value is None:
                        if errors:
                            err_msg = "、".join(errors)
                            map_data[(row, col)] = {'status': 'error', 'text': '(空)', 'error': err_msg}
                            error_list.append(f"❌ {coord}: {err_msg}")
                        else:
                            map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    val_display = format_cell_display(
                        cell_data.value if cell_data.value is not None else cell_formula.value,
                        col,
                        row
                    )
                    formula_str = str(cell_formula.value if cell_formula.value is not None else "").strip()
                    font_size = cell_formula.font.size if cell_formula.font else None

                    if len(structure_info) < 250:
                        structure_info.append({
                            "coord": coord,
                            "val": str(cell_data.value if cell_data.value is not None else "")[:25],
                            "formula": formula_str[:35],
                            "font_size": font_size,
                            "merge_info": merge_range_str,
                        })

                    if errors:
                        err_msg = "、".join(errors)
                        status = 'special' if is_special else 'error'
                        map_data[(row, col)] = {'status': status, 'text': val_display, 'error': err_msg}
                        error_list.append(f"{'🟪' if is_special else '❌'} {coord}: {err_msg}")
                    else:
                        map_data[(row, col)] = {'status': 'valid', 'text': val_display}

            sheets_data[sheet_name] = {
                'map_data': map_data,
                'error_list': error_list,
                'correction_list': correction_list,
                'structure_info': structure_info,
                'print_area_str': print_area_str,
                'max_row': max_row,
                'excess_rows': excess_rows,
                'meta': {
                    'h1_text': h1_str[:22],
                    's24_text': s24_str if s24_str else '(空)',
                    'q5_text': f"{q5_num:,.2f}" if q5_num is not None else str(q5_raw or '(空)'),
                    'st_total_text': st_total_text,
                }
            }

            del ws_data
            del ws_formula
            gc.collect()

        wb_data.close()
        wb_formula.close()
        del wb_data
        del wb_formula
        gc.collect()

        record_id = str(uuid.uuid4())[:8]
        history_db[record_id] = {
            'filename': original_filename,
            'time': time.strftime("%H:%M:%S"),
            'sheets': sheets_data
        }

        tasks_db[task_id]['status'] = 'completed'
        tasks_db[task_id]['record_id'] = record_id

    except Exception as e:
        tasks_db[task_id]['status'] = 'error'
        tasks_db[task_id]['message'] = str(e)
    finally:
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass


# --- 路由設定 ---

@app.route("/", methods=["GET"])
def index():
    record_id = request.args.get('id')
    visible_cols = BUILDING_RULES["visible_cols"]
    if record_id and record_id in history_db:
        return render_template_string(
            HTML_TEMPLATE,
            record=history_db[record_id],
            history=history_db,
            current_id=record_id,
            visible_cols=visible_cols,
            get_col_letter=openpyxl.utils.get_column_letter
        )
    return render_template_string(
        HTML_TEMPLATE,
        record=None,
        history=history_db,
        current_id=None,
        visible_cols=visible_cols,
        get_col_letter=openpyxl.utils.get_column_letter
    )

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("excel_file")
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xlsm')):
        task_id = str(uuid.uuid4())

        temp_dir = tempfile.gettempdir()
        filepath = os.path.join(temp_dir, f"{task_id}_{file.filename}")
        file.save(filepath)

        tasks_db[task_id] = {'status': 'processing', 'progress': 0, 'current_sheet': ''}

        thread = threading.Thread(target=background_process_excel, args=(task_id, filepath, file.filename))
        thread.start()

        return jsonify({"task_id": task_id})
    return jsonify({"error": "不支援的檔案格式，請上傳 .xlsx 或 .xlsm 檔案"})

@app.route("/status/<task_id>", methods=["GET"])
def task_status(task_id):
    task = tasks_db.get(task_id)
    if task:
        return jsonify(task)
    return jsonify({"status": "error", "message": "找不到該任務"})

@app.route("/clear_history", methods=["POST"])
def clear_history():
    history_db.clear()
    finished_tasks = [tid for tid, t in tasks_db.items() if t.get('status') != 'processing']
    for tid in finished_tasks:
        del tasks_db[tid]
    gc.collect()
    return jsonify({"status": "cleared"})

if __name__ == "__main__":
    app.run(debug=True, port=5000)