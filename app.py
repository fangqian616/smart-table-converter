#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能表格格式转换系统 - Streamlit Web应用
上传源数据 + 目标模板/指令，AI自动生成映射规则，一键转换输出
v2.0: 可视化列映射配置，无需手写JSON
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import json
import os
import tempfile
import time
from datetime import datetime
from pathlib import Path

# ═══════════════════════════════════════
# i18n
# ═══════════════════════════════════════
LANG = {
    "zh": {
        "sidebar_title": "表格格式转换",
        "sidebar_caption": "上传源数据 + 指定目标格式，可视化映射转换",
        "ai_section": "AI映射（可选）",
        "ai_key_label": "API Key",
        "ai_key_help": "填入后可AI自动生成映射规则，留空则手动配置",
        "api_base_label": "API Base URL",
        "api_base_help": "默认DeepSeek官方，也可替换兼容接口",
        "api_model_label": "模型名称",
        "preset_section": "预设管理",
        "load_preset": "加载预设",
        "no_preset": "（不使用）",
        "no_preset_hint": "暂无预设，转换后可保存",
        "main_title": "🔄 智能表格格式转换系统",
        "main_desc": "上传源数据表格，指定目标格式，可视化配置列映射，一键转换输出。无需写代码或JSON。",
        "step1_title": "① 上传源数据",
        "upload_source": "上传源数据文件",
        "upload_source_help": "支持 Excel 和 CSV 格式",
        "read_success": "读取成功：{} 行 × {} 列",
        "col_list": "列名列表",
        "sample_preview": "前5行预览",
        "read_fail": "读取失败：{}",
        "step2_title": "② 指定目标列",
        "target_mode_label": "目标列来源",
        "mode_manual": "✏️ 手动输入",
        "mode_template": "📎 上传模板文件",
        "mode_instruction": "📝 自然语言描述（需API）",
        "manual_label": "输入目标列名",
        "manual_help": "每行一个列名，例如：\n订单号\n日期\n客户名\n金额",
        "template_label": "上传目标模板",
        "template_help": "上传目标格式的Excel模板，自动读取列头",
        "template_ok": "模板读取成功",
        "template_cols": "Sheet「{}」列头：",
        "template_fail": "模板读取失败：{}",
        "instruction_label": "描述目标格式",
        "instruction_placeholder": "例：输出列：订单号(文本)、日期(YYYY-MM-DD)、客户名、金额(保留2位小数)、品类、备注\n日期从YYYYMMDD转成YYYY-MM-DD...",
        "instruction_help": "用自然语言描述目标列名和格式要求（需要API Key）",
        "target_cols_preview": "目标列（{}个）",
        "step3_title": "③ 配置列映射",
        "mapping_hint": "为每个目标列选择对应的源列和转换方式。源列下拉框中可选择源数据中的任意列。",
        "target_col_header": "目标列",
        "src_col_header": "源列",
        "transform_header": "转换方式",
        "no_mapping": "⚠️ 请先在上方指定目标列",
        "transform_none": "直接复制",
        "transform_text": "转文本",
        "transform_strip": "去空格",
        "transform_upper": "转大写",
        "transform_lower": "转小写",
        "transform_round2": "保留2位小数",
        "transform_round0": "取整",
        "transform_date_ymd": "日期→YYYY-MM-DD",
        "transform_date_dmy": "日期→DD/MM/YYYY",
        "transform_prefix": "加前缀...",
        "transform_suffix": "加后缀...",
        "transform_custom": "自定义转换...",
        "not_mapped": "（不映射）",
        "prefix_label": "前缀内容",
        "suffix_label": "后缀内容",
        "custom_transform_label": "自定义转换指令",
        "custom_transform_help": "如: date:YYYYMMDD->YYYY-MM-DD  或  regex:模式:替换",
        "sort_section": "排序设置（可选）",
        "sort_col_label": "排序列",
        "sort_order_asc": "升序",
        "sort_order_desc": "降序",
        "sort_none": "不排序",
        "filter_section": "过滤设置（可选）",
        "filter_col_label": "过滤列",
        "filter_op_label": "条件",
        "filter_val_label": "值",
        "filter_eq": "等于",
        "filter_neq": "不等于",
        "filter_contains": "包含",
        "filter_not_empty": "不为空",
        "add_filter": "➕ 添加过滤条件",
        "remove_filter": "🗑 移除",
        "step4_title": "④ 预览与转换",
        "source_rows": "源数据行数",
        "result_rows": "转换后行数",
        "target_cols": "目标列数",
        "preview_title": "转换结果预览（前20行）",
        "gen_excel_btn": "📥 生成Excel文件",
        "gen_excel_running": "正在生成...",
        "dl_excel_btn": "⬇️ 下载转换结果",
        "dl_csv_btn": "⬇️ 下载CSV",
        "gen_csv_btn": "📥 导出CSV文件",
        "output_filename": "转换结果_{}",
        "save_preset_title": "💾 保存为预设",
        "preset_name_label": "预设名称",
        "preset_name_ph": "例：发票格式转换",
        "preset_desc_label": "预设描述",
        "preset_desc_ph": "简要说明用途",
        "save_btn": "保存预设",
        "preset_saved": "预设「{}」已保存！",
        "convert_fail": "转换执行失败：{}",
        "error_detail": "错误详情",
        "ai_btn": "🤖 AI自动映射",
        "ai_running": "AI正在分析列映射...",
        "ai_ok": "AI映射完成！已自动填入下方配置，可手动调整。",
        "ai_fail": "AI生成失败：{}",
        "ai_no_key": "AI映射需要API Key，请在侧边栏填入。你也可以直接在下方手动选择映射。",
        "intro_title": "📤 使用流程",
        "intro_1": "1. **上传源数据** — Excel/CSV文件",
        "intro_2": "2. **指定目标列** — 三种方式：",
        "intro_2a": "   - ✏️ 手动输入目标列名（每行一个）",
        "intro_2b": "   - 📎 上传模板Excel自动读取列头",
        "intro_2c": "   - 📝 自然语言描述（需API）",
        "intro_3": "3. **配置映射** — 可视化下拉选择，或AI一键映射",
        "intro_4": "4. **预览确认** — 查看转换结果",
        "intro_5": "5. **一键导出** — 下载Excel或CSV",
        "intro_features": "**支持的转换：**",
        "intro_f1": "- 列名映射、格式转换（日期/数字/文本）",
        "intro_f2": "- 自定义前缀/后缀、正则替换",
        "intro_f3": "- 行过滤、排序、预设保存",
        "result_sheet": "转换结果",
    },
    "en": {
        "sidebar_title": "Table Converter",
        "sidebar_caption": "Upload source data + target format, visual mapping",
        "ai_section": "AI Mapping (Optional)",
        "ai_key_label": "API Key",
        "ai_key_help": "Enter to enable AI auto-mapping; leave empty for manual config",
        "api_base_label": "API Base URL",
        "api_base_help": "Default: DeepSeek official; replace with compatible endpoint",
        "api_model_label": "Model Name",
        "preset_section": "Presets",
        "load_preset": "Load Preset",
        "no_preset": "(None)",
        "no_preset_hint": "No presets yet; save one after conversion",
        "main_title": "🔄 Smart Table Format Converter",
        "main_desc": "Upload source data, specify target format, configure column mapping visually, convert and export. No code or JSON needed.",
        "step1_title": "① Upload Source Data",
        "upload_source": "Upload source data file",
        "upload_source_help": "Supports Excel and CSV formats",
        "read_success": "Read success: {} rows × {} columns",
        "col_list": "Columns",
        "sample_preview": "First 5 rows preview",
        "read_fail": "Read failed: {}",
        "step2_title": "② Specify Target Columns",
        "target_mode_label": "Target columns source",
        "mode_manual": "✏️ Manual Input",
        "mode_template": "📎 Upload Template",
        "mode_instruction": "📝 Natural Language (API needed)",
        "manual_label": "Enter target column names",
        "manual_help": "One column name per line, e.g.:\nOrder No.\nDate\nCustomer\nAmount",
        "template_label": "Upload target template",
        "template_help": "Upload target Excel template; column headers will be auto-read",
        "template_ok": "Template read successfully",
        "template_cols": "Sheet '{}' columns:",
        "template_fail": "Template read failed: {}",
        "instruction_label": "Describe target format",
        "instruction_placeholder": "e.g. Output columns: Order No. (text), Date (YYYY-MM-DD), Customer, Amount (2 decimals)...",
        "instruction_help": "Describe target columns and format in plain text (requires API Key)",
        "target_cols_preview": "Target columns ({})",
        "step3_title": "③ Configure Column Mapping",
        "mapping_hint": "For each target column, select the corresponding source column and transform type.",
        "target_col_header": "Target",
        "src_col_header": "Source",
        "transform_header": "Transform",
        "no_mapping": "⚠️ Please specify target columns above first",
        "transform_none": "Copy as-is",
        "transform_text": "To text",
        "transform_strip": "Trim whitespace",
        "transform_upper": "Uppercase",
        "transform_lower": "Lowercase",
        "transform_round2": "Round to 2 decimals",
        "transform_round0": "Round to integer",
        "transform_date_ymd": "Date → YYYY-MM-DD",
        "transform_date_dmy": "Date → DD/MM/YYYY",
        "transform_prefix": "Add prefix...",
        "transform_suffix": "Add suffix...",
        "transform_custom": "Custom transform...",
        "not_mapped": "(Not mapped)",
        "prefix_label": "Prefix text",
        "suffix_label": "Suffix text",
        "custom_transform_label": "Custom transform",
        "custom_transform_help": "e.g. date:YYYYMMDD->YYYY-MM-DD  or  regex:pattern:replacement",
        "sort_section": "Sort Settings (Optional)",
        "sort_col_label": "Sort column",
        "sort_order_asc": "Ascending",
        "sort_order_desc": "Descending",
        "sort_none": "No sort",
        "filter_section": "Filter Settings (Optional)",
        "filter_col_label": "Filter column",
        "filter_op_label": "Condition",
        "filter_val_label": "Value",
        "filter_eq": "Equals",
        "filter_neq": "Not equals",
        "filter_contains": "Contains",
        "filter_not_empty": "Not empty",
        "add_filter": "➕ Add filter",
        "remove_filter": "🗑 Remove",
        "step4_title": "④ Preview & Convert",
        "source_rows": "Source rows",
        "result_rows": "Result rows",
        "target_cols": "Target columns",
        "preview_title": "Conversion preview (first 20 rows)",
        "gen_excel_btn": "📥 Generate Excel",
        "gen_excel_running": "Generating...",
        "dl_excel_btn": "⬇️ Download Excel",
        "dl_csv_btn": "⬇️ Download CSV",
        "gen_csv_btn": "📥 Export CSV",
        "output_filename": "converted_{}",
        "save_preset_title": "💾 Save as Preset",
        "preset_name_label": "Preset name",
        "preset_name_ph": "e.g. Invoice Format",
        "preset_desc_label": "Preset description",
        "preset_desc_ph": "Brief description of usage",
        "save_btn": "Save Preset",
        "preset_saved": "Preset '{}' saved!",
        "convert_fail": "Conversion failed: {}",
        "error_detail": "Error details",
        "ai_btn": "🤖 AI Auto-Map",
        "ai_running": "AI is analyzing column mapping...",
        "ai_ok": "AI mapping done! Auto-filled below, you can adjust manually.",
        "ai_fail": "AI generation failed: {}",
        "ai_no_key": "AI mapping requires an API Key in the sidebar. You can also configure mapping manually below.",
        "intro_title": "📤 How to Use",
        "intro_1": "1. **Upload source data** — Excel or CSV file",
        "intro_2": "2. **Specify target columns** — 3 options:",
        "intro_2a": "   - ✏️ Type target column names (one per line)",
        "intro_2b": "   - 📎 Upload template to auto-read column headers",
        "intro_2c": "   - 📝 Natural language description (requires API)",
        "intro_3": "3. **Configure mapping** — Visual dropdowns, or AI auto-map",
        "intro_4": "4. **Preview & confirm** — View converted data",
        "intro_5": "5. **Export** — Download Excel or CSV",
        "intro_features": "**Supported transforms:**",
        "intro_f1": "- Column mapping, format conversion (date/number/text)",
        "intro_f2": "- Custom prefix/suffix, regex replacement",
        "intro_f3": "- Row filtering, sorting, preset saving",
        "result_sheet": "Converted",
    },
}

def t(key, *args):
    lang = st.session_state.get("lang", "zh")
    text = LANG.get(lang, LANG["zh"]).get(key, key)
    if args:
        text = text.format(*args)
    return text

# Transform dropdown options mapped to internal transform values
TRANSFORM_OPTIONS = {
    "zh": {
        "直接复制": "none",
        "转文本": "text",
        "去空格": "strip",
        "转大写": "upper",
        "转小写": "lower",
        "保留2位小数": "round:2",
        "取整": "round:0",
        "日期→YYYY-MM-DD": "date:YYYYMMDD->YYYY-MM-DD",
        "日期→DD/MM/YYYY": "date:YYYYMMDD->DD/MM/YYYY",
        "加前缀...": "prefix:",
        "加后缀...": "suffix:",
        "自定义转换...": "custom",
    },
    "en": {
        "Copy as-is": "none",
        "To text": "text",
        "Trim whitespace": "strip",
        "Uppercase": "upper",
        "Lowercase": "lower",
        "Round to 2 decimals": "round:2",
        "Round to integer": "round:0",
        "Date → YYYY-MM-DD": "date:YYYYMMDD->YYYY-MM-DD",
        "Date → DD/MM/YYYY": "date:YYYYMMDD->DD/MM/YYYY",
        "Add prefix...": "prefix:",
        "Add suffix...": "suffix:",
        "Custom transform...": "custom",
    },
}

# ── 页面配置 ──
st.set_page_config(
    page_title="Smart Table Converter",
    page_icon="🔄",
    layout="wide",
)

PRESET_DIR = Path(__file__).parent / "presets"
PRESET_DIR.mkdir(exist_ok=True)

st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; color: #2C3E50; }
    .mapping-ok { color: #27AE60; font-weight: 600; }
    .mapping-warn { color: #E67E22; font-weight: 600; }
    .mapping-miss { color: #E74C3C; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════

def read_source_file(file_obj, encoding="utf-8"):
    name = file_obj.name.lower()
    suffix = Path(name).suffix
    if suffix == ".csv":
        return pd.read_csv(file_obj, encoding=encoding)
    elif suffix in (".xls", ".xlsx", ".xlsm"):
        return pd.read_excel(file_obj, engine="openpyxl" if suffix != ".xls" else None)
    else:
        try:
            return pd.read_excel(file_obj)
        except:
            return pd.read_csv(file_obj, encoding=encoding)

def read_template_columns(file_obj):
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    result = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = []
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            if cell.value is not None:
                headers.append({"column": cell.column, "name": str(cell.value)})
        if headers:
            result[sheet_name] = headers
    wb.close()
    return result

def get_sample_data(df, n=5):
    rows = []
    for _, row in df.head(n).iterrows():
        row_dict = {}
        for col in df.columns:
            val = row[col]
            row_dict[str(col)] = str(val) if pd.notna(val) else ""
        rows.append(row_dict)
    return rows

def call_ai_for_mapping(api_key, api_base, model, source_columns, source_sample, target_columns, instruction=None):
    import requests
    lang = st.session_state.get("lang", "zh")
    
    target_info = f"\nTarget columns: {json.dumps(target_columns, ensure_ascii=False)}"
    instruction_text = f"\nUser instructions: {instruction.strip()}" if instruction and instruction.strip() else ""
    
    if lang == "en":
        prompt = f"""You are a data format conversion expert. Map source columns to target columns.

Source columns: {json.dumps(source_columns, ensure_ascii=False)}
Sample data: {json.dumps(source_sample, ensure_ascii=False, indent=2)}
{target_info}{instruction_text}

Output ONLY a JSON array. Each element: {{"source": "source_col", "target": "target_col", "transform": "transform_type"}}
If a target column has no matching source, set source to "".
Transform types: none, text, strip, upper, lower, round:2, round:0, date:YYYYMMDD->YYYY-MM-DD, prefix:X, suffix:X, regex:PATTERN:REPLACEMENT
No explanation, just the JSON array."""
    else:
        prompt = f"""你是数据格式转换专家。将源列映射到目标列。

源列: {json.dumps(source_columns, ensure_ascii=False)}
样本数据: {json.dumps(source_sample, ensure_ascii=False, indent=2)}
{target_info}{instruction_text}

只输出JSON数组，每个元素: {{"source": "源列名", "target": "目标列名", "transform": "转换方式"}}
如果目标列无对应源列，source设为""。
transform可选: none, text, strip, upper, lower, round:2, round:0, date:YYYYMMDD->YYYY-MM-DD, prefix:前缀, suffix:后缀, regex:模式:替换
不要解释，只输出JSON数组。"""

    resp = requests.post(
        f"{api_base}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"model": model, "messages": [{"role": "user", "content": prompt}], "temperature": 0.3, "max_tokens": 2000},
        timeout=60,
    )
    if resp.status_code != 200:
        raise Exception(f"API error: {resp.status_code}")
    content = resp.json()['choices'][0]['message']['content'].strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1] if "\n" in content else content[3:]
        if content.endswith("```"):
            content = content[:-3]
    return json.loads(content.strip())

def apply_transform(value, transform):
    if not transform or transform == "none":
        return value
    if transform == "text":
        return str(value) if value is not None else ""
    if transform == "strip":
        return str(value).strip() if value is not None else ""
    if transform == "upper":
        return str(value).upper() if value is not None else ""
    if transform == "lower":
        return str(value).lower() if value is not None else ""
    if transform.startswith("round:"):
        try:
            n = int(transform.split(":")[1])
            return round(float(value), n)
        except:
            return value
    if transform.startswith("date:"):
        fmt_part = transform[5:]
        if "->" in fmt_part:
            src_fmt, dst_fmt = fmt_part.split("->", 1)
            src_fmt = src_fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            dst_fmt = dst_fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            try:
                from datetime import datetime as dt
                if isinstance(value, pd.Timestamp):
                    return value.strftime(dst_fmt)
                return dt.strptime(str(value), src_fmt).strftime(dst_fmt)
            except:
                return value
    if transform.startswith("prefix:"):
        prefix = transform.split(":", 1)[1]
        return f"{prefix}{value}" if value is not None else value
    if transform.startswith("suffix:"):
        suffix = transform.split(":", 1)[1]
        return f"{value}{suffix}" if value is not None else value
    if transform.startswith("regex:"):
        parts = transform.split(":", 2)
        if len(parts) >= 3:
            import re
            try:
                return re.sub(parts[1], parts[2], str(value))
            except:
                return value
    return value

def resolve_transform(display_name, custom_prefix="", custom_suffix="", custom_transform=""):
    """从下拉框显示名转换为内部transform值"""
    lang = st.session_state.get("lang", "zh")
    opts = TRANSFORM_OPTIONS.get(lang, TRANSFORM_OPTIONS["zh"])
    internal = opts.get(display_name, "none")
    if internal == "prefix:" and custom_prefix:
        return f"prefix:{custom_prefix}"
    if internal == "suffix:" and custom_suffix:
        return f"suffix:{custom_suffix}"
    if internal == "custom" and custom_transform:
        return custom_transform
    if internal in ("prefix:", "suffix:", "custom"):
        return "none"  # 没填具体值就当none
    return internal

def execute_conversion(df, mappings_config, sort_config=None, filters_config=None):
    """mappings_config: list of {target, source, transform}"""
    result_data = {}
    target_columns = []
    for m in mappings_config:
        target = m.get("target", "")
        source = m.get("source", "")
        transform = m.get("transform", "none")
        if not target:
            continue
        target_columns.append(target)
        if source and source in df.columns:
            result_data[target] = df[source].apply(lambda v: apply_transform(v, transform))
        else:
            result_data[target] = ""
    
    result_df = pd.DataFrame(result_data, columns=target_columns)
    
    # Filters
    if filters_config:
        for f in filters_config:
            col = f.get("column", "")
            op = f.get("op", "eq")
            val = f.get("value", "")
            if col and col in result_df.columns:
                if op == "eq":
                    result_df = result_df[result_df[col].astype(str) == str(val)]
                elif op == "neq":
                    result_df = result_df[result_df[col].astype(str) != str(val)]
                elif op == "contains":
                    result_df = result_df[result_df[col].astype(str).str.contains(str(val), na=False)]
                elif op == "not_empty":
                    result_df = result_df[result_df[col].astype(str) != ""]
    
    # Sort
    if sort_config and sort_config.get("column"):
        sort_col = sort_config["column"]
        if sort_col in result_df.columns:
            result_df = result_df.sort_values(by=sort_col, ascending=(sort_config.get("order", "asc") == "asc"))
    
    return result_df.reset_index(drop=True)

def save_preset(name, mappings_config, sort_config, filters_config, description=""):
    preset = {
        "name": name,
        "description": description,
        "mappings": mappings_config,
        "sort": sort_config,
        "filters": filters_config,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    filepath = PRESET_DIR / f"{name}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(preset, f, ensure_ascii=False, indent=2)
    return filepath

def load_presets():
    presets = []
    for fp in PRESET_DIR.glob("*.json"):
        with open(fp, "r", encoding="utf-8") as f:
            presets.append(json.load(f))
    return presets

def generate_output_excel(df, template_path=None):
    tmp = tempfile.mktemp(suffix=".xlsx")
    if template_path:
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 2, column=c_idx + 1, value=val)
        wb.save(tmp)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = t("result_sheet")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
        for c_idx, col_name in enumerate(df.columns, 1):
            max_len = max(len(str(col_name)), df[col_name].astype(str).str.len().max() if len(df) > 0 else 0)
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 40)
        wb.save(tmp)
    return tmp


# ═══════════════════════════════════════
# 侧边栏
# ═══════════════════════════════════════
with st.sidebar:
    lang_col1, lang_col2 = st.columns(2)
    with lang_col1:
        if st.button("🇨🇳 中文", use_container_width=True, disabled=st.session_state.get("lang","zh")=="zh"):
            st.session_state["lang"] = "zh"
            st.rerun()
    with lang_col2:
        if st.button("🇬🇧 EN", use_container_width=True, disabled=st.session_state.get("lang","zh")=="en"):
            st.session_state["lang"] = "en"
            st.rerun()
    
    st.divider()
    
    st.image("https://img.icons8.com/fluency/96/convert-text.png", width=60)
    st.title(t("sidebar_title"))
    st.caption(t("sidebar_caption"))
    
    st.divider()
    
    st.subheader(t("ai_section"))
    deepseek_key = st.text_input(t("ai_key_label"), type="password", help=t("ai_key_help"))
    deepseek_base = st.text_input(t("api_base_label"), value="https://api.deepseek.com/v1", help=t("api_base_help"))
    deepseek_model = st.text_input(t("api_model_label"), value="deepseek-chat", help=t("api_key_help"))
    
    st.divider()
    
    st.subheader(t("preset_section"))
    existing_presets = load_presets()
    preset_names = [p["name"] for p in existing_presets]
    if preset_names:
        selected_preset = st.selectbox(t("load_preset"), [t("no_preset")] + preset_names)
    else:
        selected_preset = t("no_preset")
        st.caption(t("no_preset_hint"))

# ═══════════════════════════════════════
# 主区域
# ═══════════════════════════════════════
st.markdown(f'<p class="main-title">{t("main_title")}</p>', unsafe_allow_html=True)
st.markdown(t("main_desc"))

# ── Step 1: 上传源数据 ──
st.header(t("step1_title"))
source_file = st.file_uploader(t("upload_source"), type=["xlsx", "xls", "csv"], help=t("upload_source_help"))

if source_file:
    try:
        source_df = read_source_file(source_file)
        st.success(t("read_success", len(source_df), len(source_df.columns)))
        col1, col2 = st.columns(2)
        with col1:
            st.caption(t("col_list"))
            st.write(list(source_df.columns))
        with col2:
            st.caption(t("sample_preview"))
            st.dataframe(source_df.head(), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(t("read_fail", e))
        source_df = None
else:
    source_df = None

# ── Step 2: 指定目标列 ──
if source_df is not None:
    st.divider()
    st.header(t("step2_title"))
    
    target_mode = st.radio(t("target_mode_label"),
                           [t("mode_manual"), t("mode_template"), t("mode_instruction")],
                           horizontal=True)
    
    target_columns = []
    template_file_path = None
    
    if target_mode == t("mode_manual"):
        manual_text = st.text_area(t("manual_label"), height=150, help=t("manual_help"),
                                    placeholder="订单号\n日期\n客户名\n金额\n品类\n备注")
        if manual_text.strip():
            target_columns = [line.strip() for line in manual_text.strip().split("\n") if line.strip()]
    
    elif target_mode == t("mode_template"):
        template_file = st.file_uploader(t("template_label"), type=["xlsx", "xls"], key="template_upload", help=t("template_help"))
        if template_file:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(template_file.getvalue())
                    template_file_path = tmp.name
                target_columns_info = read_template_columns(template_file)
                st.success(t("template_ok"))
                for sheet, cols in target_columns_info.items():
                    st.caption(t("template_cols", sheet))
                    st.write([c["name"] for c in cols])
                    target_columns = [c["name"] for c in cols]
            except Exception as e:
                st.error(t("template_fail", e))
    
    elif target_mode == t("mode_instruction"):
        instruction_text = st.text_area(t("instruction_label"), value="", height=120,
                                         placeholder=t("instruction_placeholder"), help=t("instruction_help"))
        # AI自动提取目标列名
        if instruction_text and instruction_text.strip() and deepseek_key:
            if st.button("🔍 " + ("提取目标列" if st.session_state.get("lang","zh")=="zh" else "Extract columns"), key="extract_cols"):
                with st.spinner(t("ai_running")):
                    try:
                        import requests
                        lang = st.session_state.get("lang", "zh")
                        extract_prompt = f"""{"Extract target column names from the following description. Output ONLY a JSON array of column name strings, no explanation." if lang=="en" else "从以下描述中提取目标列名。只输出JSON数组（列名字符串），不要解释。"}

{instruction_text}"""
                        resp = requests.post(
                            f"{deepseek_base}/chat/completions",
                            headers={"Authorization": f"Bearer {deepseek_key}", "Content-Type": "application/json"},
                            json={"model": deepseek_model, "messages": [{"role": "user", "content": extract_prompt}], "temperature": 0.2, "max_tokens": 500},
                            timeout=30,
                        )
                        if resp.status_code == 200:
                            content = resp.json()['choices'][0]['message']['content'].strip()
                            if content.startswith("```"):
                                content = content.split("\n", 1)[1] if "\n" in content else content[3:]
                                if content.endswith("```"):
                                    content = content[:-3]
                            target_columns = json.loads(content.strip())
                            st.success(t("ai_ok"))
                        else:
                            st.error(t("ai_fail", f"HTTP {resp.status_code}"))
                    except Exception as e:
                        st.error(t("ai_fail", str(e)))
    
    # ── 预设加载 ──
    if selected_preset != t("no_preset"):
        preset_data = next((p for p in existing_presets if p["name"] == selected_preset), None)
        if preset_data:
            # 从预设恢复target_columns和mapping
            if preset_data.get("mappings"):
                target_columns = [m["target"] for m in preset_data["mappings"] if m.get("target")]
                st.info(t("preset_loaded", selected_preset))
    
    # 显示目标列
    if target_columns:
        with st.expander(t("target_cols_preview", len(target_columns)), expanded=True):
            st.write(target_columns)
    
    # ── Step 3: 可视化列映射 ──
    st.divider()
    st.header(t("step3_title"))
    
    if not target_columns:
        st.warning(t("no_mapping"))
    else:
        st.markdown(t("mapping_hint"))
        
        source_col_list = [t("not_mapped")] + [str(c) for c in source_df.columns]
        lang = st.session_state.get("lang", "zh")
        transform_display_names = list(TRANSFORM_OPTIONS.get(lang, TRANSFORM_OPTIONS["zh"]).keys())
        
        # 如果有预设，预填充
        preset_mappings = {}
        if selected_preset != t("no_preset"):
            preset_data = next((p for p in existing_presets if p["name"] == selected_preset), None)
            if preset_data and preset_data.get("mappings"):
                for m in preset_data["mappings"]:
                    preset_mappings[m.get("target", "")] = m
        
        # AI一键映射按钮
        if deepseek_key:
            if st.button(t("ai_btn")):
                with st.spinner(t("ai_running")):
                    try:
                        ai_result = call_ai_for_mapping(
                            deepseek_key, deepseek_base, deepseek_model,
                            [str(c) for c in source_df.columns],
                            get_sample_data(source_df),
                            target_columns,
                            instruction_text if target_mode == t("mode_instruction") else None
                        )
                        # 将AI结果写入session_state
                        if isinstance(ai_result, list):
                            st.session_state["ai_mapping"] = {m.get("target", ""): m for m in ai_result}
                            st.success(t("ai_ok"))
                    except Exception as e:
                        st.error(t("ai_fail", str(e)))
        else:
            st.info(t("ai_no_key"))
        
        # 获取AI映射缓存
        ai_mapping_cache = st.session_state.get("ai_mapping", {})
        
        # 可视化映射表
        mappings_config = []
        
        # 用columns布局展示映射
        for i, target_col in enumerate(target_columns):
            col_target, col_source, col_transform = st.columns([2, 3, 3])
            
            with col_target:
                st.markdown(f"**{target_col}**")
            
            # 预设/AI自动选择
            pre_selected_source = ""
            pre_selected_transform = "none"
            if target_col in preset_mappings:
                pre_selected_source = preset_mappings[target_col].get("source", "")
                pre_selected_transform = preset_mappings[target_col].get("transform", "none")
            elif target_col in ai_mapping_cache:
                pre_selected_source = ai_mapping_cache[target_col].get("source", "")
                pre_selected_transform = ai_mapping_cache[target_col].get("transform", "none")
            
            with col_source:
                # 智能默认：如果源列有同名就自动选上
                default_idx = 0
                if pre_selected_source and pre_selected_source in source_col_list:
                    default_idx = source_col_list.index(pre_selected_source)
                elif target_col in [str(c) for c in source_df.columns]:
                    default_idx = source_col_list.index(target_col)
                
                selected_source = st.selectbox(
                    f"src_{i}", source_col_list, index=default_idx,
                    key=f"src_col_{i}", label_visibility="collapsed"
                )
            
            with col_transform:
                # 找transform下拉默认值
                default_tf_idx = 0
                opts_internal = list(TRANSFORM_OPTIONS.get(lang, TRANSFORM_OPTIONS["zh"]).values())
                if pre_selected_transform in opts_internal:
                    default_tf_idx = opts_internal.index(pre_selected_transform)
                
                selected_transform_display = st.selectbox(
                    f"tf_{i}", transform_display_names, index=default_tf_idx,
                    key=f"transform_{i}", label_visibility="collapsed"
                )
            
            # 额外输入（前缀/后缀/自定义）
            internal_tf = TRANSFORM_OPTIONS.get(lang, TRANSFORM_OPTIONS["zh"]).get(selected_transform_display, "none")
            extra_val = ""
            if internal_tf == "prefix:":
                extra_val = st.text_input(t("prefix_label"), key=f"prefix_{i}", placeholder="输入前缀")
            elif internal_tf == "suffix:":
                extra_val = st.text_input(t("suffix_label"), key=f"suffix_{i}", placeholder="输入后缀")
            elif internal_tf == "custom":
                extra_val = st.text_input(t("custom_transform_label"), key=f"custom_{i}",
                                          placeholder=t("custom_transform_help"))
            
            # 解析最终transform
            final_transform = resolve_transform(selected_transform_display,
                                                 custom_prefix=extra_val if internal_tf == "prefix:" else "",
                                                 custom_suffix=extra_val if internal_tf == "suffix:" else "",
                                                 custom_transform=extra_val if internal_tf == "custom" else "")
            
            # 记录映射
            source_name = "" if selected_source == t("not_mapped") else selected_source
            mappings_config.append({
                "target": target_col,
                "source": source_name,
                "transform": final_transform
            })
        
        # ── 排序 & 过滤 ──
        with st.expander(t("sort_section")):
            sort_col_options = [t("sort_none")] + target_columns
            sort_col = st.selectbox(t("sort_col_label"), sort_col_options, key="sort_col")
            sort_order = st.radio(t("sort_order_asc") + " / " + t("sort_order_desc"),
                                  [t("sort_order_asc"), t("sort_order_desc")], horizontal=True, key="sort_order")
            
            sort_config = None
            if sort_col != t("sort_none"):
                sort_config = {"column": sort_col, "order": "asc" if sort_order == t("sort_order_asc") else "desc"}
        
        with st.expander(t("filter_section")):
            if "filter_count" not in st.session_state:
                st.session_state["filter_count"] = 0
            
            if st.button(t("add_filter")):
                st.session_state["filter_count"] += 1
                st.rerun()
            
            filters_config = []
            for fi in range(st.session_state["filter_count"]):
                fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
                with fc1:
                    f_col = st.selectbox(t("filter_col_label"), target_columns, key=f"filter_col_{fi}")
                with fc2:
                    f_op = st.selectbox(t("filter_op_label"),
                                        [t("filter_eq"), t("filter_neq"), t("filter_contains"), t("filter_not_empty")],
                                        key=f"filter_op_{fi}")
                with fc3:
                    f_val = st.text_input(t("filter_val_label"), key=f"filter_val_{fi}")
                with fc4:
                    if st.button(t("remove_filter"), key=f"del_filter_{fi}"):
                        st.session_state["filter_count"] = max(0, st.session_state["filter_count"] - 1)
                        st.rerun()
                
                op_map = {t("filter_eq"): "eq", t("filter_neq"): "neq", t("filter_contains"): "contains", t("filter_not_empty"): "not_empty"}
                filters_config.append({"column": f_col, "op": op_map.get(f_op, "eq"), "value": f_val})
        
        # ── Step 4: 预览与转换 ──
        st.divider()
        st.header(t("step4_title"))
        
        try:
            result_df = execute_conversion(source_df, mappings_config, sort_config=sort_config if target_columns else None,
                                           filters_config=filters_config if target_columns else None)
            
            col1, col2, col3 = st.columns(3)
            with col1: st.metric(t("source_rows"), f"{len(source_df)}")
            with col2: st.metric(t("result_rows"), f"{len(result_df)}")
            with col3: st.metric(t("target_cols"), f"{len(result_df.columns)}")
            
            st.subheader(t("preview_title"))
            st.dataframe(result_df.head(20), use_container_width=True, hide_index=True)
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                if st.button(t("gen_excel_btn")):
                    with st.spinner(t("gen_excel_running")):
                        output_path = generate_output_excel(result_df, template_path=template_file_path)
                        with open(output_path, "rb") as f:
                            st.download_button(label=t("dl_excel_btn"), data=f.read(),
                                               file_name=f"{t('output_filename', datetime.now().strftime('%Y%m%d_%H%M'))}.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            with col_dl2:
                if st.button(t("gen_csv_btn")):
                    csv_data = result_df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(label=t("dl_csv_btn"), data=csv_data,
                                       file_name=f"{t('output_filename', datetime.now().strftime('%Y%m%d_%H%M'))}.csv", mime="text/csv")
            
            st.divider()
            st.subheader(t("save_preset_title"))
            col_s1, col_s2 = st.columns([1, 2])
            with col_s1: preset_name = st.text_input(t("preset_name_label"), placeholder=t("preset_name_ph"))
            with col_s2: preset_desc = st.text_input(t("preset_desc_label"), placeholder=t("preset_desc_ph"))
            
            if st.button(t("save_btn")) and preset_name:
                save_preset(preset_name, mappings_config, 
                           sort_config if target_columns else None,
                           filters_config if target_columns else None, preset_desc)
                st.success(t("preset_saved", preset_name))
        
        except Exception as e:
            st.error(t("convert_fail", e))
            with st.expander(t("error_detail")):
                st.code(str(e))

else:
    st.divider()
    st.info(f"""### {t('intro_title')}

{t('intro_1')}
{t('intro_2')}
{t('intro_2a')}
{t('intro_2b')}
{t('intro_2c')}
{t('intro_3')}
{t('intro_4')}
{t('intro_5')}

{t('intro_features')}
{t('intro_f1')}
{t('intro_f2')}
{t('intro_f3')}
""")
