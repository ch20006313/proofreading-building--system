from flask import Flask, request, render_template_string, jsonify
import openpyxl
import time
import uuid
import math
import re
import gc
import threading
import os
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 允許上傳最大 100 MB 檔案

# 資料庫 (全域變數)
history_db = {}
tasks_db = {}  # 記錄背景運算進度的新資料庫

# ==============================================================================
# 🏗️ 【建物調查估價表 - 專屬檢核參數設定區】
# 您可依據「建物1118.xlsm」上傳後在網頁「🔍 結構探測器」看到的欄位座標，隨時微調下方設定：
# ==============================================================================
BUILDING_RULES = {
    "max_col": 19,                  # 預設檢核最寬欄位 (19 = S欄，可依建物表寬度調整)
    "default_max_row": 85,          # 若未設定列印範圍，預設最多檢核列數
    "data_start_row": 10,           # 明細資料開始列數 (通常前 1~9 列為表頭與基本資料區)
    
    # 字體大小規範 (單位: pt)
    "title_rows": (1, 2),           # 大標題列範圍 (第1~2列)
    "title_max_col": 17,            # 大標題欄範圍 (A~Q欄)
    "title_font_size": 15,          # 大標題字體大小
    "default_font_size": 12,        # 一般儲存格預設字體大小
    "small_font_size": 10,          # 附註/簽章區小字體大小
    
    # 合計與金額驗算欄位設定 (欄位編號: A=1, B=2, ..., J=10, K=11, M=13, N=14, O=15, P=16, R=18)
    "total_keyword_col": 2,         # 出現「合計/總計/複估」關鍵字的欄位 (預設 B欄=2，亦會自動掃描 A~F 欄)
    "total_keywords": ["合計", "總計", "小計", "複估增加", "複估新增", "補償費合計"],
    "total_amount_col": 16,         # 合計列存放總金額的欄位 (預設 P欄=16)
    
    # 明細列公式與單價驗算設定 (若設為 True，將自動核對明細列的 數量/面積 * 單價 * 成數 = 金額)
    "enable_row_Formula_check": True,
    "qty_area_col": 11,             # 面積或數量欄 (預設 K欄=11)
    "unit_price_col": 13,           # 評點或單價欄 (預設 M欄=13)
    "subtotal_col": 14,             # 金額欄 (預設 N欄=14)
    "rate_col": 15,                 # 折舊率/補償成數欄 (預設 O欄=15)
    "comp_amount_col": 16,          # 補償金額欄 (預設 P欄=16)
    "note_col": 18                  # 備註欄 (預設 R欄=18)
}

# 內建建物構造/項目單價參考庫 (可選填：若 (構造類別, 規格項目) 吻合，即自動檢核單位與單價)
BUILDING_DB = {
    # 範例格式：("鋼筋混凝土造", "一樓"): {"unit": "㎡", "price": 12500},
}

def custom_round(value, ndigits=0):
    multiplier = 10 ** ndigits
    rounded_val = math.floor(value * multiplier + 0.5) / multiplier
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

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>建築改良物調查估價表 - 智能視覺化校對系統</title>
    <style>
        body { font-family: '微軟正黑體', sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; scroll-behavior: smooth; }
        .header-title { text-align: center; color: #2c3e50; margin-bottom: 20px; }
        .layout-container { display: flex; gap: 20px; max-width: 1450px; margin: auto; align-items: flex-start; }
        .main-content { flex: 1; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: relative; min-width: 0; }
        .upload-box { border: 2px dashed #3498db; padding: 30px; text-align: center; border-radius: 10px; margin-bottom: 20px; background: #f8fbff; transition: 0.3s; }
        
        /* 載入動畫 (Loading Spinner) 樣式 */
        #loadingOverlay { display: none; text-align: center; padding: 40px 20px; }
        .spinner { border: 6px solid #f3f3f3; border-top: 6px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .progress-bar-container { width: 80%; background-color: #e0e0e0; border-radius: 8px; margin: 15px auto; height: 20px; overflow: hidden; }
        .progress-bar { height: 100%; background-color: #2ecc71; width: 0%; transition: width 0.4s ease; }
        
        .success { color: #27ae60; font-weight: bold; background: #e9f7ef; padding: 10px; border-radius: 5px; }
        .special-warning { background-color: #f3e5f5 !important; color: #6a1b9a !important; border: 2px solid #9c27b0 !important; font-weight: bold; cursor: help; }
        .correction-panel { background: #fff3cd; border-left: 6px solid #ffc107; padding: 15px 20px; border-radius: 6px; margin-bottom: 25px; }
        .correction-panel h4 { margin-top: 0; color: #856404; font-size: 16px; }
        .correction-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 4px; overflow: hidden; }
        .correction-table th, .correction-table td { padding: 10px; border: 1px solid #ffeeba; text-align: center; font-size: 14px; }
        .correction-table th { background: #ffe8a1; color: #856404; }
        
        /* 結構探測器面板 */
        .inspector-box { background: #eef6ff; border: 1px solid #b8daff; border-radius: 6px; padding: 12px 15px; margin-bottom: 20px; }
        .inspector-toggle { background: #2980b9; color: #fff; border: none; padding: 6px 14px; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold; }
        .inspector-toggle:hover { background: #1f6391; }
        .inspector-content { display: none; margin-top: 12px; max-height: 320px; overflow-y: auto; background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; font-size: 12px; }
        .inspector-table { width: 100%; border-collapse: collapse; }
        .inspector-table th, .inspector-table td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; }
        .inspector-table th { background: #f1f5f9; position: sticky; top: 0; }

        /* 瀏覽紀錄面板與置頂/清除按鈕列 */
        .history-panel { width: 320px; background: #fff; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: sticky; top: 20px; overflow: hidden; }
        .panel-actions { display: flex; gap: 8px; padding: 10px 12px; background: #ecf0f1; border-bottom: 1px solid #dcdde1; }
        .action-btn { flex: 1; padding: 8px 10px; border: none; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; transition: 0.2s; font-family: inherit; }
        .btn-top { background: #2ecc71; color: #fff; }
        .btn-top:hover { background: #27ae60; }
        .btn-clear { background: #e74c3c; color: #fff; }
        .btn-clear:hover { background: #c0392b; }
        
        .history-header { background: #2c3e50; color: #fff; padding: 15px; font-size: 16px; font-weight: bold; display: flex; justify-content: space-between; cursor: pointer; }
        .history-content { max-height: 70vh; overflow-y: auto; background: #fdfdfd; }
        .history-item { padding: 12px 15px; border-bottom: 1px solid #eee; display: flex; flex-direction: column; transition: background 0.2s; }
        .history-item:hover { background: #f0f8ff; }
        .history-item.active { background: #e1f0fa; border-left: 4px solid #2980b9; }
        .history-time { font-size: 12px; color: #7f8c8d; }
        
        .map-container { overflow-x: auto; border-radius: 8px; border: 1px solid #ddd; margin-top: 15px; }
        .excel-map { border-collapse: collapse; width: 100%; background: #fff; }
        .excel-map th { background-color: #ecf0f1; border: 1px solid #bdc3c7; font-size: 12px; padding: 6px; position: sticky; top: 0; z-index: 2; }
        .excel-map td { border: 1px solid #ecf0f1; height: 30px; min-width: 52px; text-align: center; position: relative; font-size: 12px; }
        .cell-valid { background-color: #eafaf1; color: #27ae60; }
        .cell-empty { background-color: #f9f9f9; color: #ccc; }
        .cell-error { background-color: #fadbd8; color: #c0392b; border: 2px solid #e74c3c !important; font-weight: bold; cursor: help; }
        .tooltip { position: absolute; background-color: rgba(0,0,0,0.88); color: #fff; padding: 8px 12px; border-radius: 6px; font-size: 12px; bottom: 120%; left: 50%; transform: translateX(-50%); white-space: nowrap; opacity: 0; visibility: hidden; z-index: 9999; pointer-events: none; }
        .cell-error:hover .tooltip, .special-warning:hover .tooltip { opacity: 1; visibility: visible; }
        .sheet-container { border: 2px solid #e0e0e0; border-radius: 8px; padding: 20px; margin-bottom: 40px; background: #fafbfc; }
        .sheet-title { margin-top: 0; color: #34495e; border-bottom: 2px solid #bdc3c7; padding: 10px; border-radius: 4px; transition: 0.2s; scroll-margin-top: 20px; cursor: pointer; display: flex; justify-content: space-between; }
        .sheet-title:hover { background-color: #e2e8f0; }
    </style>
</head>
<body id="pageTop">
    <h2 class="header-title">🏢 建築改良物調查估價表 - 智能視覺化校對系統</h2>
    <div class="layout-container">
        <div class="main-content">
            
            <!-- 上傳區塊與 AJAX 非同步上傳腳本 -->
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

            <!-- 渲染報表區塊 (若已完成處理) -->
            {% if record %}
            <div id="reportContainer">
                <div class="file-info" style="font-size:18px; margin-bottom: 15px; font-weight: bold;">📄 本次校對完成：{{ record.filename }}</div>
                {% for sheet_name, s_data in record.sheets.items() %}
                {% set s_idx = loop.index %}
                <div class="sheet-container">
                    <h3 class="sheet-title" id="sheet-{{ s_idx }}" onclick="toggleSheet('{{ s_idx }}')">
                        <span>🏷️ 工作表：{{ sheet_name }} (列印範圍：{{ s_data.print_area_str }})</span>
                        <span id="icon-{{ s_idx }}">▼</span>
                    </h3>
                    <div id="content-{{ s_idx }}">
                        
                        <!-- 專為建物1118.xlsm設計的結構探測器 -->
                        <div class="inspector-box">
                            <button type="button" class="inspector-toggle" onclick="toggleInspector('{{ s_idx }}')">
                                🔍 展開/收合「{{ sheet_name }}」座標、字體、合併儲存格與公式結構表 (共 {{ s_data.structure_info|length }} 格有值)
                            </button>
                            <div class="inspector-content" id="inspector-{{ s_idx }}">
                                <p style="margin-top:0; color:#555;">💡 以下為系統自動從此工作表提取之儲存格規格（若需微調建物專屬欄位規則，可參考此表座標）：</p>
                                <table class="inspector-table">
                                    <tr>
                                        <th>座標</th>
                                        <th>顯示數值</th>
                                        <th>原始公式 / 內容</th>
                                        <th>字體大小</th>
                                        <th>粗體</th>
                                        <th>合併範圍</th>
                                        <th>底線狀態</th>
                                    </tr>
                                    {% for item in s_data.structure_info %}
                                    <tr>
                                        <td><strong>{{ item.coord }}</strong></td>
                                        <td>{{ item.val }}</td>
                                        <td style="color:#2980b9; font-family:monospace;">{{ item.formula }}</td>
                                        <td>{{ item.font_size }}pt</td>
                                        <td>{{ '是' if item.bold else '否' }}</td>
                                        <td>{{ item.merge_info }}</td>
                                        <td>{{ item.bottom_border }}</td>
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
                            <h4>⚠️ 發現「合計/總計/新增」金額與明細加總或公式運算不符：</h4>
                            <table class="correction-table">
                                <tr><th>列數 / 座標</th><th>類別項目</th><th>上傳檔案金額</th><th>系統校驗正確金額</th></tr>
                                {% for c in s_data.correction_list %}
                                <tr>
                                    <td>第 {{ c.row }} 列 ({{ c.coord }})</td><td>{{ c.type }}</td>
                                    <td style="color:#c0392b; text-decoration: line-through;">{{ c.original }}</td>
                                    <td style="color:#27ae60; font-weight:bold; font-size:16px;">{{ c.correct }}</td>
                                </tr>
                                {% endfor %}
                            </table>
                        </div>
                        {% endif %}
                        
                        {% if s_data.error_list|length == 0 and not s_data.correction_list %}
                            <p class="success">✅ 完美！本工作表格式、字體、格線與公式運算皆完全符合規定。</p>
                        {% else %}
                            <p>💡 滑鼠移至<span style="color:#c0392b; font-weight:bold;">紅色/紫色格子</span>可查看錯誤原因（共發現 {{ s_data.error_list|length }} 處異常）。</p>
                            <div class="map-container">
                                <table class="excel-map">
                                    <tr>
                                        <th></th>
                                        {% for c in range(1, s_data.max_col + 1) %}<th>{{ get_col_letter(c) }}</th>{% endfor %}
                                    </tr>
                                    {% for r in range(1, s_data.max_row + 1) %}
                                    <tr>
                                        <th>{{ r }}</th>
                                        {% for c in range(1, s_data.max_col + 1) %}
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
                        {% endif %}
                    </div>
                </div>
                {% endfor %}
            </div>
            {% endif %}
        </div>

        <div class="history-panel">
            <div class="panel-actions">
                <button type="button" class="action-btn btn-top" onclick="scrollToTop()">⬆️ 置頂 (上傳)</button>
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
                        <ul style="margin: 8px 0 0 0; padding-left: 15px; font-size: 13px; color: #555;">
                            {% for s_name, s_data in hist.sheets.items() %}
                                <li onclick="window.location.href='/?id={{ uid }}#sheet-{{ loop.index }}'" style="cursor: pointer; padding: 3px 0; text-decoration: underline;">
                                    {{ s_name }} 
                                    {% if s_data.error_list|length == 0 and not s_data.correction_list %}
                                        <span style="color: #27ae60; font-weight: bold;">(✅)</span>
                                    {% else %}
                                        <span style="color: #c0392b; font-weight: bold;">(❌ {{ s_data.error_list|length + s_data.correction_list|length }} 錯誤)</span>
                                    {% endif %}
                                </li>
                            {% endfor %}
                        </ul>
                    </div>
                {% else %}
                    <div style="padding: 15px; text-align: center; color: #999; font-size: 13px;">尚無紀錄</div>
                {% endfor %}
            </div>
        </div>
    </div>
</body>
</html>
"""

# 解析儲存格內建 SUM 公式並自動驗算區間加總
def verify_sum_formula(ws_data, formula_str):
    """若儲存格為 =SUM(X1:X2) 格式，自動計算實際加總值以供比對"""
    match = re.search(r'=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', str(formula_str).upper().replace('$', ''))
    if match:
        col_start, row_start, col_end, row_end = match.groups()
        if col_start == col_end:
            col_idx = openpyxl.utils.column_index_from_string(col_start)
            total = 0.0
            for r in range(int(row_start), int(row_end) + 1):
                v = ws_data.cell(row=r, column=col_idx).value
                try:
                    if v is not None:
                        total += float(v)
                except:
                    pass
            return total
    return None

# 背景處理工作緒
def background_process_excel(task_id, filepath, original_filename):
    try:
        gc.enable()
        
        wb_data = openpyxl.load_workbook(filepath, data_only=True)
        wb_formula = openpyxl.load_workbook(filepath, data_only=False)
        
        sheets_data = {}
        sheet_names = wb_data.sheetnames
        total_sheets = len(sheet_names)
        
        max_col_cfg = BUILDING_RULES["max_col"]
        default_max_row = BUILDING_RULES["default_max_row"]
        data_start_row = BUILDING_RULES["data_start_row"]
        
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
            running_correct_p = 0
            
            excess_rows = False
            print_area = ws_data.print_area
            print_area_str = str(print_area) if print_area else "未設定"
            
            # 動態判斷列印範圍與最大列數、欄數
            max_col = max_col_cfg
            if print_area and isinstance(print_area, str):
                first_area = print_area.split(',')[0]
                try:
                    end_coord = first_area.split(':')[1] if ':' in first_area else first_area
                    col_match = re.search(r'[A-Za-z]+', end_coord.replace('$', ''))
                    row_match = re.search(r'\d+', end_coord)
                    if col_match:
                        parsed_col = openpyxl.utils.column_index_from_string(col_match.group())
                        max_col = min(max(parsed_col, max_col_cfg), 26)
                    if row_match:
                        max_row = int(row_match.group())
                    else:
                        max_row = min(ws_data.max_row, default_max_row)
                        excess_rows = ws_data.max_row > default_max_row
                except:
                    max_row = min(ws_data.max_row, default_max_row)
                    excess_rows = ws_data.max_row > default_max_row
            else:
                max_row = min(ws_data.max_row, default_max_row)
                excess_rows = ws_data.max_row > default_max_row
            
            merged_ranges = ws_data.merged_cells.ranges

            for row in range(1, max_row + 1):
                row_calc_errors = {c: [] for c in range(1, max_col + 1)}
                row_special_flags = {c: False for c in range(1, max_col + 1)}
                
                # 掃描該列是否有「合計 / 總計 / 複估」關鍵字
                row_total_kw = None
                for check_c in range(1, 7):
                    c_val = str(ws_data.cell(row=row, column=check_c).value or "").strip()
                    if any(kw in c_val for kw in BUILDING_RULES["total_keywords"]):
                        row_total_kw = c_val
                        break
                
                # 若為合計列，自動檢核 SUM 公式與加總數值
                if row_total_kw:
                    for col_idx in range(1, max_col + 1):
                        f_val = str(ws_formula.cell(row=row, column=col_idx).value or "").strip()
                        d_val = ws_data.cell(row=row, column=col_idx).value
                        if f_val.upper().startswith("=SUM"):
                            calc_sum = verify_sum_formula(ws_data, f_val)
                            try:
                                d_float = float(d_val) if d_val is not None else 0.0
                            except:
                                d_float = 0.0
                            if calc_sum is not None and abs(d_float - calc_sum) > 0.5:
                                coord_str = ws_data.cell(row=row, column=col_idx).coordinate
                                correction_list.append({
                                    "row": row,
                                    "coord": coord_str,
                                    "type": row_total_kw,
                                    "original": f"{d_float:,.0f}",
                                    "correct": f"{calc_sum:,.0f}"
                                })
                                row_calc_errors[col_idx].append(f"合計金額與區間加總({calc_sum:,.0f})不符")
                    running_correct_p = 0

                # 明細列：若具備公式運算，自動驗證「面積/數量 * 單價」或「金額 * 成數」與快取結果是否一致
                if row >= data_start_row and not row_total_kw and BUILDING_RULES["enable_row_Formula_check"]:
                    for col_idx in range(1, max_col + 1):
                        f_str = str(ws_formula.cell(row=row, column=col_idx).value or "").strip()
                        d_val = ws_data.cell(row=row, column=col_idx).value
                        # 偵測簡單相乘公式 (例如 =K11*M11 或 =ROUND(N11*O11,0))
                        mult_match = re.search(r'([A-Z]+)(\d+)\s*\*\s*([A-Z]+)(\d+)', f_str.upper().replace('$', ''))
                        if f_str.startswith("=") and mult_match and "SUM" not in f_str.upper():
                            c1, r1, c2, r2 = mult_match.groups()
                            try:
                                v1 = float(ws_data.cell(row=int(r1), column=openpyxl.utils.column_index_from_string(c1)).value or 0)
                                v2 = float(ws_data.cell(row=int(r2), column=openpyxl.utils.column_index_from_string(c2)).value or 0)
                                expected_prod = custom_round(v1 * v2, 0) if "ROUND" in f_str.upper() else (v1 * v2)
                                actual_val = float(d_val or 0)
                                if abs(actual_val - expected_prod) > 1.0:
                                    row_calc_errors[col_idx].append(f"公式運算結果異常 (快取為 {actual_val}，重算約為 {expected_prod:.2f})")
                            except:
                                pass

                # 逐格檢查格式、字體、合併儲存格與邊框
                for col in range(1, max_col + 1):
                    cell_formula = ws_formula.cell(row=row, column=col)
                    cell_data = ws_data.cell(row=row, column=col)
                    coord = cell_formula.coordinate
                    
                    is_phantom = False
                    eff_bottom_row = row
                    merge_range_str = "無"
                    
                    for mr in merged_ranges:
                        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
                        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
                            if row == min_row and col == min_col:
                                eff_bottom_row = mr_max_row
                                merge_range_str = str(mr)
                            else:
                                is_phantom = True
                            break

                    if is_phantom:
                        map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    if cell_data.value is None and cell_formula.value is None:
                        if row_calc_errors[col]:
                            err_msg = "、".join(row_calc_errors[col])
                            map_data[(row, col)] = {'status': 'error', 'text': '(空)', 'error': err_msg}
                            error_list.append(f"❌ {coord}: {err_msg}")
                        else:
                            map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    val_str = str(cell_data.value if cell_data.value is not None else "").strip()
                    formula_str = str(cell_formula.value if cell_formula.value is not None else "").strip()
                    font = cell_formula.font
                    border = cell_formula.border
                    font_size = font.size if font else None
                    is_bold = bool(font and font.bold)
                    has_bottom_border = bool(border and border.bottom and border.bottom.style)

                    # 記錄至結構探測器 (方便檢視建物1118.xlsm所有儲存格設定)
                    if len(structure_info) < 250:
                        structure_info.append({
                            "coord": coord,
                            "val": val_str[:25],
                            "formula": formula_str[:35],
                            "font_size": font_size,
                            "bold": is_bold,
                            "merge_info": merge_range_str,
                            "bottom_border": border.bottom.style if has_bottom_border else "無"
                        })
                    
                    errors = row_calc_errors[col].copy()
                    is_special = row_special_flags[col]

                    # 字體大小檢核
                    t_start, t_end = BUILDING_RULES["title_rows"]
                    if col == 1 and row < data_start_row:
                        pass
                    elif t_start <= row <= t_end and col <= BUILDING_RULES["title_max_col"]:
                        if font_size not in [BUILDING_RULES["title_font_size"], 14, 16, 18]:
                            errors.append(f"標題字體應為 {BUILDING_RULES['title_font_size']}pt (目前 {font_size}pt)")
                    else:
                        if font_size not in [BUILDING_RULES["default_font_size"], BUILDING_RULES["small_font_size"]]:
                            errors.append(f"字體大小異常 (目前 {font_size}pt，應為 12pt 或 10pt)")

                    # 明細區底線檢核
                    if row >= data_start_row and not any(kw in val_str for kw in ["備註", "說明", "調查人", "複核", "隊長"]):
                        # 依據有效底部列檢查是否有底線
                        bottom_cell_border = ws_formula.cell(row=eff_bottom_row, column=col).border
                        if not (bottom_cell_border and bottom_cell_border.bottom and bottom_cell_border.bottom.style):
                            errors.append("缺少底部格線")

                    display_text = val_str[:4] + '..' if len(val_str) > 4 else (val_str if val_str else formula_str[:4])
                    
                    if errors:
                        err_msg = "、".join(errors)
                        status = 'special' if is_special else 'error'
                        map_data[(row, col)] = {'status': status, 'text': display_text, 'error': err_msg}
                        error_list.append(f"{'🟪' if is_special else '❌'} {coord}: {err_msg}")
                    else:
                        map_data[(row, col)] = {'status': 'valid', 'text': display_text}
            
            sheets_data[sheet_name] = {
                'map_data': map_data,
                'error_list': error_list,
                'correction_list': correction_list,
                'structure_info': structure_info,
                'print_area_str': print_area_str,
                'max_row': max_row,
                'max_col': max_col,
                'excess_rows': excess_rows
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
    if record_id and record_id in history_db:
        return render_template_string(
            HTML_TEMPLATE,
            record=history_db[record_id],
            history=history_db,
            current_id=record_id,
            get_col_letter=openpyxl.utils.get_column_letter
        )
    return render_template_string(HTML_TEMPLATE, record=None, history=history_db, current_id=None)

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