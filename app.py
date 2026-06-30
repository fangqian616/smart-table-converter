#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能表格格式转换系统 - Streamlit Web应用
上传源数据 + 目标模板/指令，AI自动生成映射规则，一键转换输出
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

# ── 页面配置 ──
st.set_page_config(
    page_title="智能表格格式转换",
    page_icon="🔄",
    layout="wide",
)

# ── 常量 ──
PRESET_DIR = Path(__file__).parent / "presets"
PRESET_DIR.mkdir(exist_ok=True)
SAMPLE_ROWS = 5

# ── 自定义CSS ──
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
    """读取源数据文件，返回DataFrame"""
    name = file_obj.name.lower()
    suffix = Path(name).suffix
    
    if suffix == ".csv":
        return pd.read_csv(file_obj, encoding=encoding)
    elif suffix in (".xls", ".xlsx", ".xlsm"):
        return pd.read_excel(file_obj, engine="openpyxl" if suffix != ".xls" else None)
    else:
        # 尝试当Excel读
        try:
            return pd.read_excel(file_obj)
        except:
            return pd.read_csv(file_obj, encoding=encoding)


def read_template_columns(file_obj):
    """读取模板文件的列头信息"""
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


def get_sample_data(df, n=SAMPLE_ROWS):
    """获取前n行样本数据，转为可读文本"""
    rows = []
    for _, row in df.head(n).iterrows():
        row_dict = {}
        for col in df.columns:
            val = row[col]
            row_dict[str(col)] = str(val) if pd.notna(val) else ""
        rows.append(row_dict)
    return rows


def call_deepseek_for_mapping(api_key, api_base, source_columns, source_sample,
                               target_columns=None, instruction=None):
    """调用DeepSeek API生成映射规则"""
    import requests
    
    source_info = {
        "columns": source_columns,
        "sample_data": source_sample,
    }
    
    target_info = ""
    if target_columns:
        target_info = f"\n目标格式列头:\n{json.dumps(target_columns, ensure_ascii=False, indent=2)}"
    
    instruction_text = ""
    if instruction and instruction.strip():
        instruction_text = f"\n用户指令:\n{instruction.strip()}"
    
    prompt = f"""你是数据格式转换专家。根据源数据列和目标格式，生成列映射规则。

源数据信息:
列名列表: {json.dumps(source_columns, ensure_ascii=False)}
样本数据(前{len(source_sample)}行):
{json.dumps(source_sample, ensure_ascii=False, indent=2)}
{target_info}{instruction_text}

请输出严格的JSON格式映射规则，不要输出任何其他内容。格式如下:
{{
  "mappings": [
    {{"source": "源列名", "target": "目标列名", "transform": "转换方式"}}
  ],
  "computed": [
    {{"target": "目标列名", "formula": "计算公式描述"}}
  ],
  "filters": [],
  "sort": {{"column": "排序列", "order": "asc或desc"}}
}}

transform可选值:
- "none": 直接复制
- "date:源格式->目标格式": 日期格式转换，如 "date:YYYYMMDD->YYYY-MM-DD"
- "round:N": 四舍五入保留N位小数
- "text": 转为文本格式
- "strip": 去除首尾空格
- "upper"/"lower": 大写/小写
- "prefix:前缀"/"suffix:后缀": 添加前缀/后缀
- "regex:正则表达式:替换值": 正则替换

如果某目标列在源数据中找不到对应，放入computed用formula描述如何计算，或留空source。
只输出JSON，不要输出解释。"""

    resp = requests.post(
        f"{api_base}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "max_tokens": 2000,
        },
        timeout=60,
    )
    
    if resp.status_code != 200:
        raise Exception(f"API调用失败: {resp.status_code} - {resp.text}")
    
    content = resp.json()['choices'][0]['message']['content']
    
    # 提取JSON
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1] if "\n" in content else content[3:]
        if content.endswith("```"):
            content = content[:-3]
    content = content.strip()
    
    return json.loads(content)


def apply_transform(value, transform):
    """应用单个转换规则"""
    if transform == "none" or not transform:
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
        n = int(transform.split(":")[1])
        try:
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
                if isinstance(value, (pd.Timestamp,)):
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
            pattern, replacement = parts[1], parts[2]
            try:
                return re.sub(pattern, replacement, str(value))
            except:
                return value
    
    return value


def execute_conversion(df, mapping_rules):
    """根据映射规则执行转换，返回新的DataFrame"""
    mappings = mapping_rules.get("mappings", [])
    computed = mapping_rules.get("computed", [])
    filters = mapping_rules.get("filters", [])
    sort_config = mapping_rules.get("sort", {})
    
    result_data = {}
    target_columns = []
    
    # 处理直接映射列
    for m in mappings:
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
    
    # 处理计算列
    for c in computed:
        target = c.get("target", "")
        formula = c.get("formula", "")
        if not target:
            continue
        if target not in target_columns:
            target_columns.append(target)
        
        # 简单公式解析
        result_data[target] = _eval_formula(formula, df, result_data)
    
    result_df = pd.DataFrame(result_data, columns=target_columns)
    
    # 过滤
    for f in filters:
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
    
    # 排序
    if sort_config:
        sort_col = sort_config.get("column", "")
        sort_order = sort_config.get("order", "asc")
        if sort_col and sort_col in result_df.columns:
            result_df = result_df.sort_values(by=sort_col, ascending=(sort_order == "asc"))
    
    result_df = result_df.reset_index(drop=True)
    return result_df


def _eval_formula(formula, source_df, result_data):
    """简单公式求值"""
    # concat(A, '-', B)
    import re
    
    if formula.startswith("concat(") and formula.endswith(")"):
        inner = formula[7:-1]
        parts = [p.strip().strip("'\"") for p in inner.split(",")]
        series_list = []
        for p in parts:
            if p in source_df.columns:
                series_list.append(source_df[p].astype(str))
            elif p in result_data:
                series_list.append(result_data[p].astype(str) if isinstance(result_data[p], pd.Series) else pd.Series([str(result_data[p])]*len(source_df)))
            else:
                series_list.append(pd.Series([p]*len(source_df)))
        
        result = series_list[0]
        for s in series_list[1:]:
            result = result + s
        return result
    
    # 直接列引用
    if formula in source_df.columns:
        return source_df[formula]
    if formula in result_data and isinstance(result_data[formula], pd.Series):
        return result_data[formula]
    
    # 固定值
    return pd.Series([formula]*len(source_df))


def save_preset(name, mapping_rules, description=""):
    """保存映射预设"""
    preset = {
        "name": name,
        "description": description,
        "rules": mapping_rules,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    filepath = PRESET_DIR / f"{name}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(preset, f, ensure_ascii=False, indent=2)
    return filepath


def load_presets():
    """加载所有预设"""
    presets = []
    for fp in PRESET_DIR.glob("*.json"):
        with open(fp, "r", encoding="utf-8") as f:
            presets.append(json.load(f))
    return presets


def delete_preset(name):
    """删除预设"""
    filepath = PRESET_DIR / f"{name}.json"
    if filepath.exists():
        filepath.unlink()


def generate_output_excel(df, template_path=None):
    """生成输出的Excel文件"""
    tmp = tempfile.mktemp(suffix=".xlsx")
    
    if template_path:
        # 基于模板输出（保留模板格式）
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        # 从第2行开始写入数据
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 2, column=c_idx + 1, value=val)
        wb.save(tmp)
    else:
        # 全新生成
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "转换结果"
        
        # 表头样式
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font_white = Font(bold=True, size=11, color="FFFFFF")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        
        # 写表头
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        
        # 写数据
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
        
        # 自动列宽
        for c_idx, col_name in enumerate(df.columns, 1):
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 40)
        
        wb.save(tmp)
    
    return tmp


# ═══════════════════════════════════════
# 侧边栏
# ═══════════════════════════════════════
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/convert-text.png", width=60)
    st.title("表格格式转换")
    st.caption("上传源数据 + 指定目标格式，AI自动映射转换")
    
    st.divider()
    
    # DeepSeek API
    st.subheader("AI映射（可选）")
    deepseek_key = st.text_input("DeepSeek API Key", type="password",
                                  help="填入后可AI自动生成映射规则，留空则需手动配置")
    deepseek_base = st.text_input("API Base URL", value="https://api.deepseek.com/v1",
                                   help="默认DeepSeek官方，也可替换兼容接口")
    
    st.divider()
    
    # 预设管理
    st.subheader("预设管理")
    existing_presets = load_presets()
    preset_names = [p["name"] for p in existing_presets]
    if preset_names:
        selected_preset = st.selectbox("加载预设", ["（不使用）"] + preset_names)
    else:
        selected_preset = "（不使用）"
        st.caption("暂无预设，转换后可保存")

# ═══════════════════════════════════════
# 主区域
# ═══════════════════════════════════════
st.markdown('<p class="main-title">🔄 智能表格格式转换系统</p>', unsafe_allow_html=True)
st.markdown("上传源数据表格，指定目标格式（模板/指令/预设），AI自动生成列映射规则，预览确认后一键转换输出。")

# ── Step 1: 上传源数据 ──
st.header("① 上传源数据")
source_file = st.file_uploader("上传源数据文件", type=["xlsx", "xls", "csv"],
                                help="支持 Excel 和 CSV 格式")

if source_file:
    try:
        source_df = read_source_file(source_file)
        st.success(f"读取成功：{len(source_df)} 行 × {len(source_df.columns)} 列")
        
        col1, col2 = st.columns(2)
        with col1:
            st.caption("列名列表")
            st.write(list(source_df.columns))
        with col2:
            st.caption("前5行预览")
            st.dataframe(source_df.head(), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"读取失败：{e}")
        source_df = None
else:
    source_df = None

# ── Step 2: 指定目标格式 ──
if source_df is not None:
    st.divider()
    st.header("② 指定目标格式")
    
    target_mode = st.radio("目标格式来源", 
                           ["📝 自然语言指令", "📎 上传模板文件", "💾 使用预设"],
                           horizontal=True)
    
    target_columns_info = None
    instruction_text = None
    template_file_path = None
    mapping_rules = None
    
    if target_mode == "📝 自然语言指令":
        instruction_text = st.text_area(
            "描述目标格式",
            value="",
            height=120,
            placeholder='例：输出列：订单号(文本)、日期(YYYY-MM-DD)、客户名、金额(保留2位小数)、品类、备注\n日期从YYYYMMDD转成YYYY-MM-DD，金额四舍五入保留2位，按日期升序排列',
            help="用自然语言描述目标列名、格式要求和转换规则"
        )
    
    elif target_mode == "📎 上传模板文件":
        template_file = st.file_uploader("上传目标模板", type=["xlsx", "xls"],
                                          key="template_upload",
                                          help="上传目标格式的Excel模板，系统自动读取列头")
        if template_file:
            try:
                # 保存模板到临时文件
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(template_file.getvalue())
                    template_file_path = tmp.name
                
                target_columns_info = read_template_columns(template_file)
                st.success("模板读取成功")
                for sheet, cols in target_columns_info.items():
                    st.caption(f"Sheet「{sheet}」列头：")
                    st.write([c["name"] for c in cols])
            except Exception as e:
                st.error(f"模板读取失败：{e}")
    
    elif target_mode == "💾 使用预设":
        if selected_preset != "（不使用）":
            preset_data = next((p for p in existing_presets if p["name"] == selected_preset), None)
            if preset_data:
                mapping_rules = preset_data["rules"]
                st.success(f"已加载预设「{selected_preset}」")
                st.caption(f"描述：{preset_data.get('description', '无')}")
                with st.expander("查看映射规则"):
                    st.json(mapping_rules)
        else:
            st.info("请在侧边栏选择一个预设")
    
    # ── Step 3: 生成映射规则 ──
    st.divider()
    st.header("③ 映射规则")
    
    need_ai = (target_mode in ["📝 自然语言指令", "📎 上传模板文件"]) and not mapping_rules
    
    if need_ai:
        if not deepseek_key:
            st.warning("AI映射需要DeepSeek API Key，请在侧边栏填入。你也可以在下方手动编辑JSON规则。")
        
        if st.button("🤖 AI生成映射规则", disabled=not deepseek_key or (target_mode == "📝 自然语言指令" and not instruction_text.strip())):
            with st.spinner("AI正在分析列映射..."):
                try:
                    source_columns = [str(c) for c in source_df.columns]
                    source_sample = get_sample_data(source_df)
                    
                    mapping_rules = call_deepseek_for_mapping(
                        deepseek_key, deepseek_base,
                        source_columns, source_sample,
                        target_columns=target_columns_info,
                        instruction=instruction_text
                    )
                    st.success("映射规则生成完成！")
                except Exception as e:
                    st.error(f"AI生成失败：{e}")
                    mapping_rules = None
    
    # 映射规则编辑区
    if mapping_rules:
        # 映射对照表
        st.subheader("列映射对照")
        mappings = mapping_rules.get("mappings", [])
        computed = mapping_rules.get("computed", [])
        
        map_data = []
        for m in mappings:
            source = m.get("source", "")
            target = m.get("target", "")
            transform = m.get("transform", "none")
            status = "✅ 直接映射" if source and source in source_df.columns else ("⚠️ 列名不匹配" if source else "❌ 未指定源列")
            map_data.append({
                "源列": source if source else "（空）",
                "→": "→",
                "目标列": target,
                "转换": transform,
                "状态": status,
            })
        for c in computed:
            map_data.append({
                "源列": "计算列",
                "→": "→",
                "目标列": c.get("target", ""),
                "转换": c.get("formula", ""),
                "状态": "🧮 计算生成",
            })
        
        if map_data:
            st.dataframe(map_data, use_container_width=True, hide_index=True)
        
        # JSON编辑
        with st.expander("编辑映射规则（JSON）"):
            rules_json = st.text_area("映射规则JSON", value=json.dumps(mapping_rules, ensure_ascii=False, indent=2), height=300, key="rules_json_edit")
            if st.button("应用修改"):
                try:
                    mapping_rules = json.loads(rules_json)
                    st.success("规则已更新")
                    st.rerun()
                except json.JSONDecodeError as e:
                    st.error(f"JSON格式错误：{e}")
    else:
        # 手动输入JSON
        st.info("没有AI生成的规则？可以在下方手动输入映射规则JSON：")
        manual_json = st.text_area("手动输入映射规则", value='{\n  "mappings": [],\n  "computed": [],\n  "filters": [],\n  "sort": {}\n}', height=200, key="manual_json")
        if st.button("应用手动规则"):
            try:
                mapping_rules = json.loads(manual_json)
                st.success("规则已应用")
            except json.JSONDecodeError as e:
                st.error(f"JSON格式错误：{e}")
    
    # ── Step 4: 预览与转换 ──
    if mapping_rules:
        st.divider()
        st.header("④ 预览与转换")
        
        try:
            result_df = execute_conversion(source_df, mapping_rules)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("源数据行数", f"{len(source_df)}")
            with col2:
                st.metric("转换后行数", f"{len(result_df)}")
            with col3:
                st.metric("目标列数", f"{len(result_df.columns)}")
            
            st.subheader("转换结果预览（前20行）")
            st.dataframe(result_df.head(20), use_container_width=True, hide_index=True)
            
            # 下载
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                if st.button("📥 生成Excel文件"):
                    with st.spinner("正在生成..."):
                        output_path = generate_output_excel(result_df, template_path=template_file_path)
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="⬇️ 下载转换结果",
                                data=f.read(),
                                file_name=f"转换结果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
            
            with col_dl2:
                if st.button("📥 导出CSV文件"):
                    csv_data = result_df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        label="⬇️ 下载CSV",
                        data=csv_data,
                        file_name=f"转换结果_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv"
                    )
            
            # 保存预设
            st.divider()
            st.subheader("💾 保存为预设")
            col_s1, col_s2 = st.columns([1, 2])
            with col_s1:
                preset_name = st.text_input("预设名称", placeholder="例：发票格式转换")
            with col_s2:
                preset_desc = st.text_input("预设描述", placeholder="简要说明用途")
            
            if st.button("保存预设") and preset_name:
                save_preset(preset_name, mapping_rules, preset_desc)
                st.success(f"预设「{preset_name}」已保存！")
        
        except Exception as e:
            st.error(f"转换执行失败：{e}")
            with st.expander("错误详情"):
                st.code(str(e))

else:
    # 没有上传文件
    st.divider()
    st.info("""
    ### 📤 使用流程
    
    1. **上传源数据** — Excel/CSV文件
    2. **指定目标格式** — 三种方式：
       - 📝 自然语言指令描述目标列和格式要求
       - 📎 上传目标模板Excel，自动读取列头
       - 💾 使用之前保存的映射预设
    3. **AI生成映射** — 自动匹配源列→目标列+转换规则
    4. **预览确认** — 查看映射对照表和转换预览
    5. **一键导出** — 下载Excel或CSV
    
    **支持的转换：**
    - 列名映射、格式转换（日期/数字/文本）
    - 新增计算列、行过滤、排序
    - 基于模板保留原格式输出
    """)
