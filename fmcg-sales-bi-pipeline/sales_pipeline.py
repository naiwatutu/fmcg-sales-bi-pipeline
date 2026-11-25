# -*- coding: utf-8 -*-
"""
月度销售数据自动合并与专业快消 BI 多载体汇报系统（v4 · All-in-One · 零硬编码）
兼容环境：GitHub / Workbuddy / 本地命令行 / Google Colab

核心分析模型：
  - Revenue Bridge 营收桥接瀑布归因（存量增减 / 阶段断单 / 破零恢复 / 首发新品）
  - PVM 量价结构拆解（销量变动影响 vs 件均单价/结构变动影响）
  - 休闲自主 vs OEM/KA 客户集中度诊断（CR3 / CR5 / HHI 指数）
  - 新品 Cohort 留存与 4 维趋势洞察 + 连续 3 月僵尸 SKU 清退与退货预警
  - 无手工改配置的选单式核心品持续追踪（watchlist.json）

常用运行示例：
  1. 日常仅合并数据（默认 -f none，不出报告，保持总表纯净）：
     python sales_pipeline.py
  2. 出报告前先看《待选追踪清单》（Workbuddy / 无头环境第一步）：
     python sales_pipeline.py --preview
  3. 勾选核心品并按需输出独立高颜值报告（支持 md / html / pptx / xlsx / all）：
     python sales_pipeline.py --report-only --month 202610 -f all
     python sales_pipeline.py --report-only --month 202610 -f html,pptx --add-track 1,8,9,10,20,21
  4. 本地交互终端一键向导：
     python sales_pipeline.py -f all -i
"""
import os
import re
import io
import sys
import json
import shutil
import argparse
from datetime import datetime
from collections import Counter
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule

# ==============================================================================
# 第一部分：极简外置纠偏规则加载 & 主表动态词典自学习（零硬编码产品/赛道/品牌）
# ==============================================================================
DEFAULT_BASE_DIR = os.environ.get("SALES_BASE_DIR", os.path.abspath("./data"))

BUILTIN_FALLBACK_RULES = {
    "new_product_window_months": 3,
    "track_aliases": {"蔬菜": "蔬菜类"},
    "brand_extract_priority": ["天润一生缘"],
    "brand_aliases": {"美one": "美ONE"},
    "subcat_aliases": {},
}


def load_rules(base_dir: str, cli_rules_path: str = None) -> dict:
    """加载可选外置 rules.json；缺失或损坏时自动回退至内置极简纠偏"""
    candidates = [
        cli_rules_path,
        os.path.join(base_dir, "rules.json"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules.json"),
    ]
    rules = {
        k: (v.copy() if isinstance(v, (dict, list)) else v)
        for k, v in BUILTIN_FALLBACK_RULES.items()
    }
    for path in candidates:
        if path and os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    user_rules = json.load(f)
                for k in rules:
                    if k in user_rules and isinstance(user_rules[k], type(rules[k])):
                        rules[k] = user_rules[k]
                print(f"    - 已加载外置纠偏规则：{path}")
                return rules
            except Exception as e:
                print(f"⚠️ 警告：读取 {path} 失败（{e}），已自动回退至内置极简规则。")
    return rules


def normalize_text(val) -> str:
    s = str(val if val is not None else "").strip()
    return "" if s in ["", "nan", "None", "0", "0.0", "-"] else s


def clean_display_name(name: str, max_len: int = None) -> str:
    """优化报表展示名称：去除冗余的‘成品’前缀，保留规格与口味，不生硬截断"""
    s = normalize_text(name)
    s = re.sub(r"^成品", "", s).strip()
    if max_len and len(s) > max_len:
        return s[: max_len - 1] + "…"
    return s or normalize_text(name)


def parse_yyyymm(val) -> int:
    s = str(val if val is not None else "").strip()
    m = re.match(r"^(20\d{2})(0[1-9]|1[0-2])(?:\.0)?$", s)
    if m:
        return int(f"{m.group(1)}{m.group(2)}")
    m2 = re.match(r"^(20\d{2})[-/年](1[0-2]|0?[1-9])(?!\d)", s)
    if m2:
        return int(f"{m2.group(1)}{int(m2.group(2)):02d}")
    return None


def shift_yyyymm(ym: int, delta_months: int) -> int:
    y, m = ym // 100, ym % 100
    total = y * 12 + (m - 1) + delta_months
    return (total // 12) * 100 + (total % 12 + 1)


def apply_brand_rule(product_name: str, raw_brand: str, rules: dict) -> str:
    p_name = normalize_text(product_name)
    b_raw = normalize_text(raw_brand)
    for pri in rules.get("brand_extract_priority", []):
        if pri and (pri in p_name or pri in b_raw):
            return pri
    if b_raw:
        return rules.get("brand_aliases", {}).get(b_raw, b_raw)
    return ""


def apply_track_rule(raw_track: str, rules: dict) -> str:
    t = normalize_text(raw_track)
    return rules.get("track_aliases", {}).get(t, t) if t else "待补录"


def apply_subcat_rule(raw_sub: str, rules: dict) -> str:
    s = normalize_text(raw_sub)
    return rules.get("subcat_aliases", {}).get(s, s) if s else "待补录"


def learn_master_vocab(ws_main, data_start_row: int, max_row: int, rules: dict) -> dict:
    brands = set(rules.get("brand_extract_priority", []))
    subcat_track_votes = {}
    for r in range(data_start_row, max_row + 1):
        if not clean_code(ws_main.cell(row=r, column=1).value):
            continue
        p_name = normalize_text(ws_main.cell(row=r, column=2).value)
        track = apply_track_rule(ws_main.cell(row=r, column=4).value, rules)
        brand = apply_brand_rule(p_name, ws_main.cell(row=r, column=13).value, rules)
        subcat = apply_subcat_rule(ws_main.cell(row=r, column=14).value, rules)
        if brand and brand not in ["其他", "待补录"]:
            brands.add(brand)
        if subcat and subcat not in ["其他", "待补录"] and len(subcat) >= 2:
            subcat_track_votes.setdefault(subcat, Counter())[track] += 1

    subcat_to_track = {sc: cnt.most_common(1)[0][0] for sc, cnt in subcat_track_votes.items()}
    return {
        "brands_sorted": sorted(brands, key=len, reverse=True),
        "subcats_sorted": sorted(subcat_to_track.keys(), key=len, reverse=True),
        "subcat_to_track": subcat_to_track,
    }


def infer_new_sku_from_vocab(product_name: str, channel: str, vocab: dict, rules: dict) -> dict:
    p_name = normalize_text(product_name)
    brand = apply_brand_rule(p_name, "", rules)
    if not brand:
        for b in vocab["brands_sorted"]:
            if b in p_name:
                brand = b
                break
    brand = brand or "待补录"

    subcat, track = "待补录", "待补录"
    for sc in vocab["subcats_sorted"]:
        if sc in p_name:
            subcat = sc
            track = vocab["subcat_to_track"].get(sc, "待补录")
            break

    sales_limit = "专供品" if channel == "OEM" else "待补录"
    return {
        "storage_temp": "常温",
        "track": track,
        "brand": brand,
        "sub_cat": subcat,
        "sales_limit": sales_limit,
    }


# ==============================================================================
# 第二部分：三级智能月份提取、发货表自适应解析与主表幂等合并引擎
# ==============================================================================
def clean_code(val) -> str:
    s = str(val if val is not None else "").strip()
    if not s or s in ["nan", "None"] or any(k in s for k in ["空白", "总计", "合计", "小计"]):
        return None
    try:
        return str(int(float(s)))
    except Exception:
        return s


def safe_float(val) -> float:
    num = pd.to_numeric(val, errors="coerce")
    return 0.0 if pd.isna(num) else float(num)


def detect_layout(ws):
    header_row_1, header_row_2 = 1, 2
    for r in range(1, 10):
        val = str(ws.cell(row=r, column=1).value or "").strip()
        if val in ["商品编码", "存货编码", "商品唯一识别码"]:
            header_row_1 = r
            break
    for r in range(header_row_1, header_row_1 + 5):
        for c in range(16, min(40, ws.max_column + 1)):
            if "发货数量" in str(ws.cell(row=r, column=c).value or ""):
                header_row_2 = r
                break
    return header_row_1, header_row_2, header_row_2 + 1


def collect_month_columns(ws, header_row_1: int) -> dict:
    month_col_map = {}
    for c in range(18, ws.max_column + 1):
        m_str = str(ws.cell(row=header_row_1, column=c).value or "").strip()
        if re.match(r"^20\d{2}(0[1-9]|1[0-2])$", m_str):
            month_col_map[int(m_str)] = (c, c + 1)
    return month_col_map


def extract_month_tag(xx_file: str, oem_file: str, ws_main, cli_month: int = None) -> int:
    if cli_month:
        return int(cli_month)

    for fpath in [xx_file, oem_file]:
        if not fpath:
            continue
        try:
            df_head = pd.read_excel(fpath, sheet_name=0, header=None, nrows=8)
            text_blob = " ".join(df_head.astype(str).values.flatten())
            m_iso = re.search(r"(20\d{2})[-/年](1[0-2]|0?[1-9])(?!\d)", text_blob)
            if m_iso:
                return int(f"{m_iso.group(1)}{int(m_iso.group(2)):02d}")
            m_ym = re.search(r"\b(20\d{2})(0[1-9]|1[0-2])\b", text_blob)
            if m_ym:
                return int(m_ym.group(0))
        except Exception:
            pass

    h1, _, _ = detect_layout(ws_main)
    existing_months = sorted(collect_month_columns(ws_main, h1).keys())
    base_year = (existing_months[-1] // 100) if existing_months else 2026

    for fpath in [xx_file, oem_file]:
        if not fpath:
            continue
        fname = os.path.basename(fpath)
        m_full = re.search(r"(20\d{2})(0[1-9]|1[0-2])", fname)
        if m_full:
            return int(m_full.group(0))
        m_short = re.search(r"(?<!\d)(1[0-2]|0?[1-9])\s*月", fname)
        if m_short:
            return int(f"{base_year}{int(m_short.group(1)):02d}")

    raise ValueError("无法从表内日期或文件名自动识别月份，请通过 --month YYYYMM 指定。")


def smart_extract_invoice(file_path: str, source_channel: str) -> pd.DataFrame:
    xl = pd.ExcelFile(file_path)
    target_sheet = next((s for s in xl.sheet_names if "发货" in s or "Sheet" in s), xl.sheet_names[0])
    df_raw = pd.read_excel(file_path, sheet_name=target_sheet, header=None)

    header_idx = 0
    for i in range(min(15, len(df_raw))):
        if "存货编码" in [str(x).strip() for x in df_raw.iloc[i].tolist()]:
            header_idx = i
            break

    df = pd.read_excel(file_path, sheet_name=target_sheet, skiprows=header_idx + 1, header=None)
    raw_cols = [str(x).strip() for x in df_raw.iloc[header_idx].tolist()]
    df.columns = [c if c and c != "nan" else f"Col_{i}" for i, c in enumerate(raw_cols)]

    col_map = {}
    for c in df.columns:
        if c == "存货编码":
            col_map[c] = "Code"
        elif c in ["存货名称", "产品名称"] and "Name" not in col_map.values():
            col_map[c] = "Name"
        elif c == "规格型号":
            col_map[c] = "Spec"
        elif "发货数量" in c and "Qty" not in col_map.values():
            col_map[c] = "Qty"
        elif "发货金额" in c and "Amt" not in col_map.values():
            col_map[c] = "Amt"
        elif c in ["客户", "客户名称", "购货单位"] and "Customer" not in col_map.values():
            col_map[c] = "Customer"

    df_ex = df[list(col_map.keys())].copy().rename(columns=col_map)
    if "Customer" not in df_ex.columns:
        df_ex["Customer"] = "未标注客户"
    df_ex["Customer"] = df_ex["Customer"].fillna("未标注客户").astype(str).str.strip()
    df_ex["Code"] = df_ex["Code"].apply(clean_code)
    df_ex = df_ex.dropna(subset=["Code"]).copy()
    df_ex["Qty"] = pd.to_numeric(df_ex["Qty"], errors="coerce").fillna(0.0)
    df_ex["Amt"] = pd.to_numeric(df_ex["Amt"], errors="coerce").fillna(0.0)
    df_ex["Channel"] = source_channel
    return df_ex


def build_master_dataframe(ws_main, data_start_row: int, end_row: int, month_col_map: dict,
                           month_tag: int, rules: dict, xx_summary: dict = None, oem_summary: dict = None) -> pd.DataFrame:
    xx_summary = xx_summary or {}
    oem_summary = oem_summary or {}
    records = []
    sorted_months = sorted(month_col_map.keys())
    for r in range(data_start_row, end_row + 1):
        c_val = clean_code(ws_main.cell(row=r, column=1).value)
        if not c_val:
            continue
        p_name = normalize_text(ws_main.cell(row=r, column=2).value)
        track = apply_track_rule(ws_main.cell(row=r, column=4).value, rules)
        brand = apply_brand_rule(p_name, ws_main.cell(row=r, column=13).value, rules) or "待补录"
        subcat = apply_subcat_rule(ws_main.cell(row=r, column=14).value, rules)
        raw_first = ws_main.cell(row=r, column=15).value
        first_ym = parse_yyyymm(raw_first)

        rec = {
            "Code": c_val, "Name": p_name, "Disp_Name": clean_display_name(p_name),
            "Track": track, "Spec": normalize_text(ws_main.cell(row=r, column=5).value),
            "Brand": brand, "SubCat": subcat,
            "First_Month_Raw": str(raw_first or "").strip(),
            "First_Month_YM": first_ym,
            "Curr_XX_Amt": round(float(xx_summary[c_val]["Amt"]), 2) if c_val in xx_summary else 0.0,
            "Curr_OEM_Amt": round(float(oem_summary[c_val]["Amt"]), 2) if c_val in oem_summary else 0.0,
            "Curr_XX_Qty": round(float(xx_summary[c_val]["Qty"]), 4) if c_val in xx_summary else 0.0,
            "Curr_OEM_Qty": round(float(oem_summary[c_val]["Qty"]), 4) if c_val in oem_summary else 0.0,
        }
        tot_q, tot_a = 0.0, 0.0
        for m_int in sorted_months:
            q_col, a_col = month_col_map[m_int]
            mq = safe_float(ws_main.cell(row=r, column=q_col).value)
            ma = safe_float(ws_main.cell(row=r, column=a_col).value)
            rec[f"Qty_{m_int}"] = mq
            rec[f"Amt_{m_int}"] = ma
            if m_int <= month_tag:
                tot_q += mq
                tot_a += ma
        rec["Curr_Qty"] = rec.get(f"Qty_{month_tag}", 0.0)
        rec["Curr_Amt"] = rec.get(f"Amt_{month_tag}", 0.0)
        rec["YTD_Qty"] = round(tot_q, 4)
        rec["YTD_Amt"] = round(tot_a, 2)
        records.append(rec)
    return pd.DataFrame(records)


def merge_monthly_data(master_file: str, xx_file: str, oem_file: str, rules: dict, cli_month: int = None) -> dict:
    wb = openpyxl.load_workbook(master_file, data_only=False)
    main_sheet_name = next((s for s in wb.sheetnames if any(k in s for k in ["年", "业绩", "总表"])), wb.sheetnames[0])
    ws_main = wb[main_sheet_name]

    month_tag = extract_month_tag(xx_file, oem_file, ws_main, cli_month)
    df_xx = smart_extract_invoice(xx_file, "休闲")
    df_oem = smart_extract_invoice(oem_file, "OEM")
    df_all_curr = pd.concat([df_xx, df_oem], ignore_index=True)

    summary_curr = df_all_curr.groupby("Code")[["Qty", "Amt"]].sum().to_dict("index")
    xx_summary = df_xx.groupby("Code")[["Qty", "Amt"]].sum().to_dict("index")
    oem_summary = df_oem.groupby("Code")[["Qty", "Amt"]].sum().to_dict("index")

    header_row_1, header_row_2, data_start_row = detect_layout(ws_main)
    existing_codes = {}
    true_max_row = data_start_row - 1
    for r in range(data_start_row, ws_main.max_row + 1):
        c_val = clean_code(ws_main.cell(row=r, column=1).value)
        if c_val:
            true_max_row = r
            existing_codes[c_val] = r

    vocab = learn_master_vocab(ws_main, data_start_row, true_max_row, rules)
    unique_curr = df_all_curr.drop_duplicates(subset=["Code"])
    new_items_df = unique_curr[~unique_curr["Code"].isin(existing_codes.keys())].copy()

    current_row = true_max_row + 1
    new_sku_records = []
    for _, item in new_items_df.iterrows():
        attrs = infer_new_sku_from_vocab(item["Name"], item["Channel"], vocab, rules)
        ws_main.cell(row=current_row, column=1, value=item["Code"])
        ws_main.cell(row=current_row, column=2, value=item["Name"])
        ws_main.cell(row=current_row, column=3, value=attrs["storage_temp"])
        ws_main.cell(row=current_row, column=4, value=attrs["track"])
        ws_main.cell(row=current_row, column=5, value=item["Spec"])
        ws_main.cell(row=current_row, column=12, value=attrs["sales_limit"])
        ws_main.cell(row=current_row, column=13, value=attrs["brand"])
        ws_main.cell(row=current_row, column=14, value=attrs["sub_cat"])
        ws_main.cell(row=current_row, column=15, value=month_tag)
        existing_codes[item["Code"]] = current_row
        new_sku_records.append({
            "Code": item["Code"], "Name": item["Name"], "Spec": item["Spec"],
            "Channel": item["Channel"], "Brand": attrs["brand"],
            "Track": attrs["track"], "SubCat": attrs["sub_cat"],
            "Qty": float(summary_curr[item["Code"]]["Qty"]),
            "Amt": float(summary_curr[item["Code"]]["Amt"]),
        })
        current_row += 1

    last_data_row = current_row - 1

    month_col_map = collect_month_columns(ws_main, header_row_1)
    is_overwrite = month_tag in month_col_map
    if is_overwrite:
        target_qty_col = month_col_map[month_tag][0]
    else:
        target_qty_col = None
        for c in range(ws_main.max_column, 15, -1):
            if str(ws_main.cell(row=header_row_2, column=c).value or "").strip() == "发货金额":
                target_qty_col = c + 1
                break
        if target_qty_col is None:
            target_qty_col = 18

    target_amt_col = target_qty_col + 1
    ws_main.cell(row=header_row_1, column=target_qty_col, value=month_tag)
    ws_main.cell(row=header_row_2, column=target_qty_col, value="发货数量")
    ws_main.cell(row=header_row_2, column=target_amt_col, value="发货金额")
    month_col_map[month_tag] = (target_qty_col, target_amt_col)

    sorted_months = sorted(month_col_map.keys())
    qty_letters = [get_column_letter(month_col_map[m][0]) for m in sorted_months]
    amt_letters = [get_column_letter(month_col_map[m][1]) for m in sorted_months]

    for r in range(data_start_row, last_data_row + 1):
        c_val = clean_code(ws_main.cell(row=r, column=1).value)
        if not c_val:
            continue
        qty = round(float(summary_curr[c_val]["Qty"]), 4) if c_val in summary_curr else 0.0
        amt = round(float(summary_curr[c_val]["Amt"]), 2) if c_val in summary_curr else 0.0
        ws_main.cell(row=r, column=target_qty_col, value=qty if qty != 0 else None)
        ws_main.cell(row=r, column=target_amt_col, value=amt if amt != 0 else None)
        if qty_letters:
            ws_main.cell(row=r, column=16, value="=" + "+".join([f"{let}{r}" for let in qty_letters]))
        if amt_letters:
            ws_main.cell(row=r, column=17, value="=" + "+".join([f"{let}{r}" for let in amt_letters]))

    df_master = build_master_dataframe(
        ws_main, data_start_row, last_data_row, month_col_map, month_tag, rules, xx_summary, oem_summary
    )
    if not new_sku_records:
        df_new_curr = df_master[df_master["First_Month_YM"] == month_tag]
        for _, r in df_new_curr.iterrows():
            ch = "OEM" if r["Curr_OEM_Amt"] > r["Curr_XX_Amt"] else "休闲"
            new_sku_records.append({
                "Code": r["Code"], "Name": r["Name"], "Spec": r["Spec"], "Channel": ch,
                "Brand": r["Brand"], "Track": r["Track"], "SubCat": r["SubCat"],
                "Qty": r["Curr_Qty"], "Amt": r["Curr_Amt"],
            })

    return {
        "wb": wb, "month_tag": month_tag, "is_overwrite": is_overwrite,
        "sorted_months": sorted_months, "df_master": df_master,
        "df_curr_detail": df_all_curr, "df_new_skus": pd.DataFrame(new_sku_records),
        "has_detail_channel": True,
    }


def load_data_for_report_only(master_file: str, search_dirs: list, rules: dict, cli_month: int = None) -> dict:
    wb = openpyxl.load_workbook(master_file, data_only=False)
    main_sheet_name = next((s for s in wb.sheetnames if any(k in s for k in ["年", "业绩", "总表"])), wb.sheetnames[0])
    ws_main = wb[main_sheet_name]
    header_row_1, _, data_start_row = detect_layout(ws_main)
    month_col_map = collect_month_columns(ws_main, header_row_1)
    if not month_col_map:
        raise ValueError("总表中未识别到任何 YYYYMM 格式月份列。")

    sorted_months = sorted(month_col_map.keys())
    month_tag = int(cli_month) if cli_month else sorted_months[-1]
    if month_tag not in month_col_map:
        raise ValueError(f"总表中不存在月份 {month_tag}，当前已有月份：{sorted_months}")

    last_data_row = data_start_row - 1
    for r in range(data_start_row, ws_main.max_row + 1):
        if clean_code(ws_main.cell(row=r, column=1).value):
            last_data_row = r

    xx_file, oem_file = None, None
    m_str, m_short = str(month_tag), f"{month_tag % 100}月"
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if not f.endswith(".xlsx") or f.startswith("~$"):
                continue
            if m_str in f or m_short in f:
                if "休闲" in f and not xx_file:
                    xx_file = os.path.join(d, f)
                elif ("OEM" in f or "oem" in f) and not oem_file:
                    oem_file = os.path.join(d, f)

    df_all_curr = pd.DataFrame(columns=["Code", "Name", "Spec", "Qty", "Amt", "Customer", "Channel"])
    xx_summary, oem_summary = {}, {}
    has_detail = False
    if xx_file or oem_file:
        parts = []
        if xx_file:
            df_xx = smart_extract_invoice(xx_file, "休闲")
            xx_summary = df_xx.groupby("Code")[["Qty", "Amt"]].sum().to_dict("index")
            parts.append(df_xx)
        if oem_file:
            df_oem = smart_extract_invoice(oem_file, "OEM")
            oem_summary = df_oem.groupby("Code")[["Qty", "Amt"]].sum().to_dict("index")
            parts.append(df_oem)
        df_all_curr = pd.concat(parts, ignore_index=True)
        has_detail = True

    df_master = build_master_dataframe(
        ws_main, data_start_row, last_data_row, month_col_map, month_tag, rules, xx_summary, oem_summary
    )
    df_new = df_master[df_master["First_Month_YM"] == month_tag].copy()
    new_sku_records = []
    for _, r in df_new.iterrows():
        ch = "OEM" if r["Curr_OEM_Amt"] > r["Curr_XX_Amt"] else "休闲"
        new_sku_records.append({
            "Code": r["Code"], "Name": r["Name"], "Spec": r["Spec"], "Channel": ch,
            "Brand": r["Brand"], "Track": r["Track"], "SubCat": r["SubCat"],
            "Qty": r["Curr_Qty"], "Amt": r["Curr_Amt"],
        })
    return {
        "wb": None, "month_tag": month_tag, "is_overwrite": False,
        "sorted_months": sorted_months, "df_master": df_master,
        "df_curr_detail": df_all_curr, "df_new_skus": pd.DataFrame(new_sku_records),
        "has_detail_channel": has_detail,
    }


# ==============================================================================
# 第三部分：无手工改配置的【核心品追踪清单】持久化管理与出报告前预检选单
# ==============================================================================
def get_watchlist_path(base_dir: str) -> str:
    return os.path.join(base_dir, "watchlist.json")


def get_preview_cache_path(base_dir: str) -> str:
    return os.path.join(base_dir, "last_preview.json")


def load_watchlist(base_dir: str) -> list:
    path = get_watchlist_path(base_dir)
    if not os.path.isfile(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        items = data.get("tracked_codes", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
        seen, out = set(), []
        for c in items:
            cc = clean_code(c)
            if cc and cc not in seen:
                seen.add(cc)
                out.append(cc)
        return out
    except Exception:
        return []


def save_watchlist(base_dir: str, codes: list, df_master: pd.DataFrame):
    code_to_name = dict(zip(df_master["Code"], df_master["Name"]))
    seen, clean_codes = set(), []
    for c in codes:
        cc = clean_code(c)
        if cc and cc not in seen:
            seen.add(cc)
            clean_codes.append(cc)
    payload = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tracked_codes": clean_codes,
        "items": [{"code": c, "name": code_to_name.get(c, "")} for c in clean_codes],
    }
    with open(get_watchlist_path(base_dir), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def build_tracking_preview(data_pack: dict, base_dir: str, rules: dict) -> dict:
    df = data_pack["df_master"].copy()
    curr_m = data_pack["month_tag"]
    win_n = max(1, int(rules.get("new_product_window_months", 3)))
    win_months = [shift_yyyymm(curr_m, -i) for i in range(win_n)]
    tracked_codes = load_watchlist(base_dir)
    tracked_set = set(tracked_codes)

    entries, used_codes, idx = [], set(), 1

    def add_group(group_key: str, rows_df: pd.DataFrame):
        nonlocal idx
        for _, r in rows_df.iterrows():
            code = r["Code"]
            if code in used_codes:
                continue
            used_codes.add(code)
            entries.append({
                "idx": idx, "group": group_key, "code": code, "name": r["Name"],
                "brand": r["Brand"], "subcat": r["SubCat"],
                "first_ym": int(r["First_Month_YM"]) if pd.notna(r["First_Month_YM"]) else None,
                "curr_amt_wan": round(r["Curr_Amt"] / 1e4, 2),
                "ytd_amt_wan": round(r["YTD_Amt"] / 1e4, 2),
                "is_tracked": code in tracked_set,
            })
            idx += 1

    if tracked_codes:
        t_df = df[df["Code"].isin(tracked_set)].copy()
        t_df["_ord"] = t_df["Code"].map({c: i for i, c in enumerate(tracked_codes)})
        add_group("tracked", t_df.sort_values("_ord"))

    df_curr_new = df[df["First_Month_YM"] == curr_m].sort_values("Curr_Amt", ascending=False)
    add_group("curr_new", df_curr_new)

    older_win = [m for m in win_months if m != curr_m]
    if older_win:
        df_recent = df[df["First_Month_YM"].isin(older_win)].sort_values(["Curr_Amt", "YTD_Amt"], ascending=False).head(12)
        add_group("recent_new", df_recent)

    add_group("star_candidates", df.sort_values("YTD_Amt", ascending=False).head(12))

    cache_payload = {
        "month_tag": curr_m,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "index_map": {str(e["idx"]): {"code": e["code"], "name": e["name"]} for e in entries},
        "curr_new_codes": df_curr_new["Code"].tolist(),
        "recent_new_codes": df[df["First_Month_YM"].isin(win_months)].sort_values("Curr_Amt", ascending=False)["Code"].tolist(),
    }
    with open(get_preview_cache_path(base_dir), "w", encoding="utf-8") as f:
        json.dump(cache_payload, f, ensure_ascii=False, indent=2)

    return {"month_tag": curr_m, "window_months": win_months, "entries": entries, "tracked_codes": tracked_codes}


def print_tracking_preview(preview: dict):
    curr_m, entries = preview["month_tag"], preview["entries"]
    print(f"\n================ 📋 {curr_m} 出报告前《新品与核心品待选追踪清单》 ================")
    group_titles = [
        ("tracked", "【① 当前持续追踪中（如需移出可输入 -编号 或 --remove-track）】"),
        ("curr_new", f"【② 当月首发新品（{curr_m} 首次上市，输入 new 可一键全选）】"),
        ("recent_new", f"【③ 近 {len(preview['window_months'])} 个月次新孵化池 TOP 候选（观察复购爬坡）】"),
        ("star_candidates", "【④ 全盘年度 TOP 明星候选品（大盘压舱石）】"),
    ]
    for gk, title in group_titles:
        sub = [e for e in entries if e["group"] == gk]
        print(f"\n{title}")
        if not sub:
            print("    （暂无）")
            continue
        for e in sub:
            tag = "★已追踪" if e["is_tracked"] else "  未追踪"
            fm_txt = f"上市:{e['first_ym']}" if e["first_ym"] else "老品"
            print(
                f"  [{e['idx']:>2d}] {tag} | {e['code']} | {e['name']} "
                f"（{fm_txt} | 当月:{e['curr_amt_wan']:,.2f}万 | 累计:{e['ytd_amt_wan']:,.2f}万）"
            )
    print("\n👉 快捷指令：")
    print("   - 按编号追加追踪并出全套报告：python sales_pipeline.py --report-only -f all --add-track 1,8,9,10,20")
    print("   - 一键全选本月首发新品追踪  ：python sales_pipeline.py --report-only -f all --add-track new")
    print("   - 暂不更改追踪直接出报告    ：python sales_pipeline.py --report-only -f all")
    print("========================================================================\n")


def resolve_selection_tokens(tokens_str: str, base_dir: str, df_master: pd.DataFrame, curr_m: int, rules: dict) -> list:
    if not tokens_str:
        return []
    cache = {}
    cache_path = get_preview_cache_path(base_dir)
    if os.path.isfile(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                cache = json.load(f)
        except Exception:
            cache = {}

    index_map = cache.get("index_map", {})
    valid_codes = set(df_master["Code"].tolist())
    win_n = max(1, int(rules.get("new_product_window_months", 3)))
    win_months = [shift_yyyymm(curr_m, -i) for i in range(win_n)]

    resolved = []
    for raw_tok in re.split(r"[,，\s]+", tokens_str.strip()):
        tok = raw_tok.strip()
        if not tok:
            continue
        low = tok.lower()
        if low == "new":
            resolved.extend(df_master[df_master["First_Month_YM"] == curr_m]["Code"].tolist())
        elif low == "recent":
            resolved.extend(
                df_master[df_master["First_Month_YM"].isin(win_months)]
                .sort_values("Curr_Amt", ascending=False)["Code"].head(10).tolist()
            )
        elif tok in index_map:
            resolved.append(index_map[tok]["code"])
        elif clean_code(tok) in valid_codes:
            resolved.append(clean_code(tok))
        else:
            matches = df_master[df_master["Name"].str.contains(re.escape(tok), case=False, na=False)]
            if not matches.empty:
                resolved.extend(matches.sort_values("YTD_Amt", ascending=False)["Code"].head(5).tolist())
            else:
                print(f"⚠️ 未匹配到追踪项：{tok}")

    seen, dedup = set(), []
    for c in resolved:
        if c and c not in seen:
            seen.add(c)
            dedup.append(c)
    return dedup


def apply_cli_track_updates(args, data_pack: dict, base_dir: str, rules: dict) -> list:
    df_m, curr_m = data_pack["df_master"], data_pack["month_tag"]
    current = load_watchlist(base_dir)
    changed = False

    if args.clear_track:
        current, changed = [], True
    if args.set_track:
        current, changed = resolve_selection_tokens(args.set_track, base_dir, df_m, curr_m, rules), True
    if args.add_track:
        for c in resolve_selection_tokens(args.add_track, base_dir, df_m, curr_m, rules):
            if c not in current:
                current.append(c)
        changed = True
    if args.remove_track:
        to_rem = set(resolve_selection_tokens(args.remove_track, base_dir, df_m, curr_m, rules))
        current, changed = [c for c in current if c not in to_rem], True

    if changed:
        save_watchlist(base_dir, current, df_m)
        print(f">>> [追踪名单已更新] 当前持续追踪核心品共 {len(current)} 款（已写入 {get_watchlist_path(base_dir)}）")
    return current


def run_interactive_watchlist_wizard(data_pack: dict, base_dir: str, rules: dict) -> list:
    preview = build_tracking_preview(data_pack, base_dir, rules)
    print_tracking_preview(preview)
    current = list(preview["tracked_codes"])
    try:
        ans = input("请输入要追加的编号/关键词（如 1,3,8 或 new，移除用 -编号，清空输 clear，回车保持现状）：").strip()
    except EOFError:
        return current
    if not ans:
        return current
    if ans.lower() == "clear":
        save_watchlist(base_dir, [], data_pack["df_master"])
        return []

    add_tokens, rem_tokens = [], []
    for part in re.split(r"[,，\s]+", ans):
        p = part.strip()
        if not p:
            continue
        if p.startswith("-") and len(p) > 1:
            rem_tokens.append(p[1:])
        else:
            add_tokens.append(p)

    df_m, curr_m = data_pack["df_master"], data_pack["month_tag"]
    if add_tokens:
        for c in resolve_selection_tokens(",".join(add_tokens), base_dir, df_m, curr_m, rules):
            if c not in current:
                current.append(c)
    if rem_tokens:
        rem_set = set(resolve_selection_tokens(",".join(rem_tokens), base_dir, df_m, curr_m, rules))
        current = [c for c in current if c not in rem_set]

    save_watchlist(base_dir, current, df_m)
    print(f">>> [追踪名单已保存] 当前持续追踪核心品共 {len(current)} 款。")
    return current


# ==============================================================================
# 第四部分：专业快消经营分析引擎（Revenue Bridge / PVM / CR5&HHI / 僵尸SKU / 新品Cohort）
# ==============================================================================
def safe_mom(curr: float, prev: float, has_prev_month: bool = True) -> str:
    if not has_prev_month:
        return "首月无环比"
    if prev <= 0:
        return "新动销/破零" if curr > 0 else "—"
    pct = (curr - prev) / prev * 100.0
    return f"{'+' if pct > 0 else ''}{pct:.1f}%"


def classify_bcg(curr_wan: float, prev_wan: float, median_scale: float, has_prev: bool = True) -> str:
    if not has_prev:
        return "首月基准"
    if prev_wan <= 0 and curr_wan > 0:
        return "🌱 潜力爬坡(破零)"
    mom = ((curr_wan - prev_wan) / prev_wan * 100.0) if prev_wan > 0 else 0.0
    scale_cut = max(20.0, median_scale)
    if curr_wan >= scale_cut:
        return "★ 明星增肌区" if mom >= 10.0 else ("⚠️ 支柱回调区" if mom <= -10.0 else "💰 现金牛稳健区")
    return "🌱 潜力爬坡区" if mom >= 15.0 else "🔍 长尾观察区"


def compute_revenue_bridge(df: pd.DataFrame, curr_m: int, has_prev: bool) -> dict:
    """营收桥接瀑布归因：将上月到当月的净增减严格拆解为 5 个互斥驱动因子"""
    prev_tot = float(df["Prev_Amt"].sum())
    curr_tot = float(df["Curr_Amt"].sum())
    net_diff = curr_tot - prev_tot
    if not has_prev:
        return {"has_prev": False, "prev_tot": 0.0, "curr_tot": curr_tot, "net_diff": curr_tot, "items": []}

    is_new = df["First_Month_YM"] == curr_m
    is_churn = (~is_new) & (df["Prev_Amt"] > 0) & (df["Curr_Amt"] <= 0)
    is_rebound = (~is_new) & (df["Prev_Amt"] <= 0) & (df["Curr_Amt"] > 0)
    is_grow = (~is_new) & (df["Prev_Amt"] > 0) & (df["Curr_Amt"] > 0) & (df["Curr_Amt"] >= df["Prev_Amt"])
    is_drop = (~is_new) & (df["Prev_Amt"] > 0) & (df["Curr_Amt"] > 0) & (df["Curr_Amt"] < df["Prev_Amt"])
    is_other = ~(is_new | is_churn | is_rebound | is_grow | is_drop)

    items = [
        {"key": "new", "label": "本月首发新品增量", "val": float((df.loc[is_new, "Curr_Amt"] - df.loc[is_new, "Prev_Amt"]).sum()), "cnt": int(is_new.sum())},
        {"key": "rebound", "label": "老品破零/复采增量", "val": float((df.loc[is_rebound, "Curr_Amt"] - df.loc[is_rebound, "Prev_Amt"]).sum()), "cnt": int(is_rebound.sum())},
        {"key": "grow", "label": "在售品自然增长", "val": float((df.loc[is_grow, "Curr_Amt"] - df.loc[is_grow, "Prev_Amt"]).sum()), "cnt": int(is_grow.sum())},
        {"key": "drop", "label": "在售品自然回调", "val": float((df.loc[is_drop, "Curr_Amt"] - df.loc[is_drop, "Prev_Amt"]).sum()), "cnt": int(is_drop.sum())},
        {"key": "churn", "label": "上月动销本月断单/冲红", "val": float((df.loc[is_churn, "Curr_Amt"] - df.loc[is_churn, "Prev_Amt"]).sum()), "cnt": int(is_churn.sum())},
    ]
    other_v = float((df.loc[is_other, "Curr_Amt"] - df.loc[is_other, "Prev_Amt"]).sum())
    if abs(other_v) >= 1.0:
        items.append({"key": "other", "label": "历史退货冲红净变动", "val": other_v, "cnt": int(is_other.sum())})

    top_drags = df.sort_values("Diff_Amt", ascending=True).head(5)
    top_lifts = df.sort_values("Diff_Amt", ascending=False).head(5)
    return {
        "has_prev": True, "prev_tot": prev_tot, "curr_tot": curr_tot, "net_diff": net_diff,
        "items": items, "top_drags": top_drags, "top_lifts": top_lifts,
    }


def compute_pvm_analysis(df: pd.DataFrame, has_prev: bool) -> dict:
    """PVM 量价拆解：拆出【发货件数变动影响】与【件均单价/产品结构变动影响】"""
    curr_q, curr_a = float(df["Curr_Qty"].sum()), float(df["Curr_Amt"].sum())
    curr_asp = (curr_a / curr_q) if curr_q > 0 else 0.0
    if not has_prev:
        return {"has_prev": False, "curr_q": curr_q, "curr_asp": curr_asp}

    prev_q, prev_a = float(df["Prev_Qty"].sum()), float(df["Prev_Amt"].sum())
    prev_asp = (prev_a / prev_q) if prev_q > 0 else 0.0
    net_diff = curr_a - prev_a
    vol_effect = (curr_q - prev_q) * prev_asp
    price_mix_effect = (curr_asp - prev_asp) * curr_q
    q_mom = ((curr_q - prev_q) / prev_q * 100.0) if prev_q > 0 else 0.0
    p_mom = ((curr_asp - prev_asp) / prev_asp * 100.0) if prev_asp > 0 else 0.0

    driver = "销量变动主导" if abs(vol_effect) >= abs(price_mix_effect) else "单价/结构变动主导"
    return {
        "has_prev": True,
        "prev_q": prev_q, "curr_q": curr_q, "q_mom_pct": round(q_mom, 1),
        "prev_asp": round(prev_asp, 2), "curr_asp": round(curr_asp, 2), "p_mom_pct": round(p_mom, 1),
        "net_diff": net_diff, "vol_effect": vol_effect, "price_mix_effect": price_mix_effect,
        "driver": driver,
    }


def compute_customer_concentration(cust_df: pd.DataFrame) -> dict:
    """量化计算全盘、休闲、OEM 三口径的客户集中度（CR3 / CR5 / HHI 指数）"""
    if cust_df.empty:
        return {"available": False, "scopes": []}

    def calc_scope(label: str, sub: pd.DataFrame) -> dict:
        pos = sub[sub["Amt"] > 0].sort_values("Amt", ascending=False)
        tot = float(pos["Amt"].sum())
        if tot <= 0 or pos.empty:
            return {"scope": label, "cust_cnt": 0, "cr3": 0.0, "cr5": 0.0, "hhi": 0, "risk": "无数据", "top1": "—", "top1_pct": 0.0}
        shares = pos["Amt"] / tot
        cr3 = float(shares.head(3).sum() * 100.0)
        cr5 = float(shares.head(5).sum() * 100.0)
        hhi = int(round(float(((shares * 100.0) ** 2).sum())))
        risk = "🔴 高度集中(强依赖)" if hhi >= 2500 else ("🟡 中度集中(需关注)" if hhi >= 1500 else "🟢 分散健康")
        return {
            "scope": label, "cust_cnt": len(pos),
            "cr3": round(cr3, 1), "cr5": round(cr5, 1), "hhi": hhi, "risk": risk,
            "top1": str(pos.iloc[0]["Customer"]), "top1_pct": round(float(shares.iloc[0] * 100.0), 1),
        }

    scopes = [calc_scope("全盘汇总", cust_df)]
    for ch_key, ch_lbl in [("休闲", "休闲自主板块"), ("OEM", "OEM与KA定制板块")]:
        sub = cust_df[cust_df["Channel"] == ch_key]
        if not sub.empty:
            scopes.append(calc_scope(ch_lbl, sub))
    return {"available": True, "scopes": scopes}


def compute_sku_health(df: pd.DataFrame, months: list) -> dict:
    """SKU 健康度诊断：分离当月0动销与【连续3个月零动销僵尸SKU】及冲红退货率"""
    tail_n = min(3, len(months))
    tail_ms = months[-tail_n:]
    tail_cols = [f"Amt_{m}" for m in tail_ms if f"Amt_{m}" in df.columns]

    if tail_cols:
        max_recent = df[tail_cols].max(axis=1)
        is_mature = ~df["First_Month_YM"].isin(tail_ms)
        zombie_df = df[is_mature & (max_recent <= 0)].sort_values("YTD_Amt", ascending=True).copy()
    else:
        zombie_df = pd.DataFrame(columns=df.columns)

    pos_amt = float(df[df["Curr_Amt"] > 0]["Curr_Amt"].sum())
    neg_amt = float(df[df["Curr_Amt"] < 0]["Curr_Amt"].sum())
    return_ratio = (abs(neg_amt) / pos_amt * 100.0) if pos_amt > 0 else 0.0
    active_cnt = int((df["Curr_Amt"] > 0).sum())
    active_rate = (active_cnt / len(df) * 100.0) if len(df) > 0 else 0.0

    return {
        "tail_n": tail_n, "tail_months": tail_ms,
        "active_cnt": active_cnt, "active_rate": round(active_rate, 1),
        "zombie_df": zombie_df, "zombie_cnt": len(zombie_df),
        "neg_amt": neg_amt, "return_ratio": round(return_ratio, 2),
    }


def compute_new_product_insights(df: pd.DataFrame, curr_m: int, sorted_months: list, rules: dict) -> dict:
    win_n = max(1, int(rules.get("new_product_window_months", 3)))
    win_months = sorted([shift_yyyymm(curr_m, -i) for i in range(win_n)])
    avail_win_months = [m for m in win_months if m in sorted_months]

    df_pool = df[df["First_Month_YM"].isin(win_months)].copy()
    df_curr_new = df_pool[df_pool["First_Month_YM"] == curr_m].sort_values("Curr_Amt", ascending=False).copy()
    tot_all_amt = df["Curr_Amt"].sum() or 1.0

    cohort_rows = []
    for fm in reversed(win_months):
        sub = df_pool[df_pool["First_Month_YM"] == fm]
        n_skus = len(sub)
        active_curr = int((sub["Curr_Amt"] > 0).sum())
        curr_amt = float(sub["Curr_Amt"].sum())
        launch_amt = float(sub[f"Amt_{fm}"].sum()) if f"Amt_{fm}" in sub.columns else curr_amt
        ytd_amt = float(sub["YTD_Amt"].sum())
        retention_rate = (active_curr / n_skus * 100.0) if n_skus > 0 else 0.0
        amt_retention = (curr_amt / launch_amt * 100.0) if launch_amt > 0 else 0.0
        cohort_rows.append({
            "Batch": str(fm), "Batch_Int": fm, "Role": "当月首发" if fm == curr_m else "次新留存",
            "SKU_Count": n_skus, "Active_Curr": active_curr,
            "Retention_Pct": round(retention_rate, 1),
            "Launch_Amt": launch_amt, "Curr_Amt": curr_amt, "YTD_Amt": ytd_amt,
            "Amt_Retention_Pct": round(amt_retention, 1),
        })
    cohort_df = pd.DataFrame(cohort_rows)

    ramp_horses, status_list = [], []
    for _, r in df_pool.iterrows():
        fm = int(r["First_Month_YM"])
        if fm == curr_m:
            st = "🆕 当月首发" if r["Curr_Amt"] > 0 else "🆕 首发待起量"
        else:
            seq_months = [m for m in avail_win_months if m >= fm]
            vals = [float(r.get(f"Amt_{m}", 0.0)) for m in seq_months]
            if len(vals) >= 2 and all(vals[i] > vals[i - 1] and vals[i] > 0 for i in range(1, len(vals))):
                st = "🚀 连续爬坡"
                ramp_horses.append({"Name": r["Disp_Name"], "Batch": fm, "Vals": vals, "Curr_Amt": r["Curr_Amt"]})
            elif r["Curr_Amt"] <= 0:
                st = "⚠️ 本月冲红" if r["Curr_Amt"] < 0 else "⏸️ 本月断单"
            elif len(vals) >= 2 and vals[-1] < vals[-2] * 0.7:
                st = "📉 冲高回调"
            else:
                st = "🔄 稳定复购"
        status_list.append(st)
    if not df_pool.empty:
        df_pool["Trend_Status"] = status_list
        df_pool = df_pool.sort_values(["Curr_Amt", "YTD_Amt"], ascending=False)
    else:
        df_pool["Trend_Status"] = []

    pool_curr_amt = float(df_pool["Curr_Amt"].sum()) if not df_pool.empty else 0.0
    pool_ytd_amt = float(df_pool["YTD_Amt"].sum()) if not df_pool.empty else 0.0
    pool_active = int((df_pool["Curr_Amt"] > 0).sum()) if not df_pool.empty else 0
    curr_new_amt = float(df_curr_new["Curr_Amt"].sum()) if not df_curr_new.empty else 0.0
    pool_share = pool_curr_amt / tot_all_amt * 100.0
    pool_active_rate = (pool_active / len(df_pool) * 100.0) if len(df_pool) > 0 else 0.0

    insights = [
        f"整体贡献：近{win_n}月新品池共 {len(df_pool)} 款（当月动销 {pool_active} 款，动销率 {pool_active_rate:.1f}%），"
        f"创收 {pool_curr_amt/1e4:,.2f} 万元（占全盘 {pool_share:.1f}%），其中 {curr_m} 首发 {len(df_curr_new)} 款首单 {curr_new_amt/1e4:,.2f} 万元。"
    ]
    older_cohorts = cohort_df[(cohort_df["Batch_Int"] < curr_m) & (cohort_df["SKU_Count"] > 0)]
    if not older_cohorts.empty:
        weak = older_cohorts[older_cohorts["Retention_Pct"] < 40.0]
        if not weak.empty:
            w_strs = [f"{r['Batch']}批次({int(r['Active_Curr'])}/{int(r['SKU_Count'])}款,{r['Retention_Pct']:.1f}%)" for _, r in weak.iterrows()]
            insights.append(f"批次留存预警：{'、'.join(w_strs)} 在本月出现明显复购断层，需核查渠道首单后终端周转与补货排期。")
        else:
            best_c = older_cohorts.sort_values("Retention_Pct", ascending=False).iloc[0]
            insights.append(f"批次留存稳健：{best_c['Batch']}批次当月动销留存率 {best_c['Retention_Pct']:.1f}%，贡献 {best_c['Curr_Amt']/1e4:,.2f} 万元。")
    else:
        insights.append("批次留存：当前窗口内仅有首发批次，次月将自动计算留存率。")

    if ramp_horses:
        ramp_horses.sort(key=lambda x: x["Curr_Amt"], reverse=True)
        top_h = "、".join([f"{h['Name']}({h['Curr_Amt']/1e4:.2f}万)" for h in ramp_horses[:3]])
        insights.append(f"连续爬坡黑马：共 {len(ramp_horses)} 款次新品实现连续正增长放量，代表单品为 {top_h}，建议重点倾斜铺货资源。")
    else:
        insights.append("爬坡态势：受大盘周期波动影响，近3月次新品未现连续两月正增单品，建议定向拉动复购。")

    if pool_curr_amt > 0 and len(df_pool) >= 2:
        top2 = df_pool.head(2)
        top2_share = float(top2["Curr_Amt"].sum()) / pool_curr_amt * 100.0
        top2_names = " + ".join([str(n)[:16] for n in top2["Disp_Name"].tolist()])
        if top2_share >= 60.0:
            insights.append(f"单点集中风险：前2款大单品（{top2_names}）占新品池营收 {top2_share:.1f}%，高度依赖头部 KA 排期，需拓宽腰部矩阵。")
        else:
            insights.append(f"结构分布均衡：前2款新品占比 {top2_share:.1f}%，新品收入呈多点开花态势。")
    else:
        insights.append("结构分布：当月新品尚处冷启动期，重点跟进首单二返转化。")

    return {
        "window_months": avail_win_months, "pool_df": df_pool, "curr_new_df": df_curr_new,
        "cohort_df": cohort_df, "pool_curr_amt": pool_curr_amt, "pool_ytd_amt": pool_ytd_amt,
        "pool_share_pct": round(pool_share, 1), "pool_active": pool_active, "insights": insights,
    }


def build_action_titles(stats: dict) -> dict:
    """咨询级 Action Title 自动生成引擎：每页直接亮出核心经营判断而非目录名称"""
    m_str = f"{int(str(stats['curr_m'])[4:])}月"
    tot_w = stats["curr_tot"] / 1e4
    mom = safe_mom(stats["curr_tot"], stats["prev_tot"], stats["has_prev"])
    pvm, bridge, npd, conc, health = stats["pvm"], stats["bridge"], stats["new_prod"], stats["conc"], stats["health"]

    if stats["has_detail"]:
        xx_p = stats["xx_tot"] / (stats["curr_tot"] or 1) * 100
        oem_p = stats["oem_tot"] / (stats["curr_tot"] or 1) * 100
        t_trend = f"{m_str}总营收 ¥{tot_w:,.1f}万（环比 {mom}），休闲自主贡献 {xx_p:.1f}% 稳守基本盘，OEM 占 {oem_p:.1f}%"
    else:
        t_trend = f"{m_str}总营收 ¥{tot_w:,.1f}万（环比 {mom}），全年累计突破 ¥{stats['ytd_tot']/1e4:,.1f}万"

    if bridge["has_prev"] and pvm["has_prev"]:
        main_item = sorted(bridge["items"], key=lambda x: abs(x["val"]), reverse=True)[0]
        t_bridge = f"净变动 {bridge['net_diff']/1e4:+,.1f}万：{pvm['driver']}，核心因子为【{main_item['label']}】({main_item['val']/1e4:+,.1f}万)"
    else:
        t_bridge = f"{m_str}为首期基准月，下月起自动生成营收桥接瀑布与 PVM 量价拆解"

    t_new = f"近{len(npd['window_months'])}月新品池贡献 ¥{npd['pool_curr_amt']/1e4:,.1f}万（占全盘 {npd['pool_share_pct']:.1f}%），需同步盯紧次新留存与爬坡"

    top_tr = stats["track_df"].iloc[0] if not stats["track_df"].empty else None
    t_track = (
        f"【{top_tr['Track']}】以 {top_tr['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}% 份额领跑全盘，环比 {safe_mom(top_tr['Curr_Amt'], top_tr['Prev_Amt'], stats['has_prev'])}"
        if top_tr is not None else "全量品类赛道动态表现"
    )

    if conc["available"] and conc["scopes"]:
        sc0 = conc["scopes"][0]
        t_brand = f"全盘前3大客户集中度 CR3={sc0['cr3']:.1f}%（HHI={sc0['hhi']}，{sc0['risk']}）"
    else:
        top_b = stats["brand_df"].iloc[0] if not stats["brand_df"].empty else None
        t_brand = f"头部品牌【{top_b['Brand']}】贡献 {top_b['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}%，持续领跑品牌矩阵" if top_b is not None else "核心品牌矩阵"

    t_risk = f"全盘动销率 {health['active_rate']:.1f}%，检出连续 {health['tail_n']} 月零动销僵尸品 {health['zombie_cnt']} 款、冲红 {len(stats['neg_df'])} 款"

    return {
        "trend": t_trend, "bridge": t_bridge, "new_prod": t_new,
        "track": t_track, "brand": t_brand, "risk": t_risk,
    }


def compute_analytics(data_pack: dict, rules: dict, tracked_codes: list) -> dict:
    df = data_pack["df_master"].copy()
    df_det = data_pack["df_curr_detail"].copy()
    months = [m for m in data_pack["sorted_months"] if m <= data_pack["month_tag"]]
    curr_m = data_pack["month_tag"]
    prev_m = months[months.index(curr_m) - 1] if curr_m in months and months.index(curr_m) > 0 else None
    has_prev = prev_m is not None

    df["Prev_Amt"] = df[f"Amt_{prev_m}"].fillna(0.0) if has_prev else 0.0
    df["Prev_Qty"] = df[f"Qty_{prev_m}"].fillna(0.0) if has_prev else 0.0
    df["Diff_Amt"] = df["Curr_Amt"] - df["Prev_Amt"]

    curr_tot, prev_tot, ytd_tot = float(df["Curr_Amt"].sum()), float(df["Prev_Amt"].sum()), float(df["YTD_Amt"].sum())
    xx_tot, oem_tot = float(df["Curr_XX_Amt"].sum()), float(df["Curr_OEM_Amt"].sum())

    df_pos = df[df["Curr_Amt"] > 0].sort_values("Curr_Amt", ascending=False).copy()
    df_pos["Cum_Share"] = df_pos["Curr_Amt"].cumsum() / (df_pos["Curr_Amt"].sum() or 1.0)
    a_skus = int((df_pos["Cum_Share"] <= 0.80).sum() + (1 if len(df_pos) > 0 else 0))
    b_skus = int(((df_pos["Cum_Share"] > 0.80) & (df_pos["Cum_Share"] <= 0.95)).sum())
    c_skus = max(0, len(df_pos) - a_skus - b_skus)
    neg_df = df[df["Curr_Amt"] < 0].sort_values("Curr_Amt")

    track_grp = df.groupby("Track", as_index=False).agg(
        SKU_Count=("Code", "count"), Active_SKU=("Curr_Amt", lambda s: int((s > 0).sum())),
        Curr_Amt=("Curr_Amt", "sum"), Prev_Amt=("Prev_Amt", "sum"), YTD_Amt=("YTD_Amt", "sum"),
    ).sort_values("Curr_Amt", ascending=False)

    sub_grp = df.groupby(["SubCat", "Track"], as_index=False).agg(
        Active_SKU=("Curr_Amt", lambda s: int((s > 0).sum())),
        Curr_Amt=("Curr_Amt", "sum"), Prev_Amt=("Prev_Amt", "sum"), YTD_Amt=("YTD_Amt", "sum"),
    ).sort_values("Curr_Amt", ascending=False)
    active_sub_wan = sub_grp[sub_grp["Curr_Amt"] > 0]["Curr_Amt"] / 1e4
    med_scale = float(active_sub_wan.quantile(0.65)) if len(active_sub_wan) > 0 else 25.0
    sub_grp["BCG"] = sub_grp.apply(
        lambda r: classify_bcg(r["Curr_Amt"] / 1e4, r["Prev_Amt"] / 1e4, med_scale, has_prev), axis=1
    )

    brand_grp = df.groupby("Brand", as_index=False).agg(
        Total_SKU=("Code", "count"), Active_SKU=("Curr_Amt", lambda s: int((s > 0).sum())),
        Curr_Amt=("Curr_Amt", "sum"), Prev_Amt=("Prev_Amt", "sum"), YTD_Amt=("YTD_Amt", "sum"),
    ).sort_values("Curr_Amt", ascending=False)

    if data_pack["has_detail_channel"] and not df_det.empty:
        cust_grp = df_det.groupby(["Channel", "Customer"], as_index=False).agg(
            SKU_Count=("Code", "nunique"), Qty=("Qty", "sum"), Amt=("Amt", "sum"),
        ).sort_values("Amt", ascending=False)
    else:
        cust_grp = pd.DataFrame(columns=["Channel", "Customer", "SKU_Count", "Qty", "Amt"])

    trend_series = [{"Month": str(m), "Amt_Wan": round(df[f"Amt_{m}"].sum() / 1e4, 2)} for m in months]
    pending_df = df[(df["Brand"] == "待补录") | (df["Track"] == "待补录") | (df["SubCat"] == "待补录")]

    if tracked_codes:
        wl_df = df[df["Code"].isin(set(tracked_codes))].copy()
        wl_df["_ord"] = wl_df["Code"].map({c: i for i, c in enumerate(tracked_codes)})
        wl_df = wl_df.sort_values("_ord").drop(columns=["_ord"])
    else:
        wl_df = pd.DataFrame(columns=df.columns)

    stats = {
        "curr_m": curr_m, "prev_m": prev_m, "has_prev": has_prev, "months": months,
        "has_detail": data_pack["has_detail_channel"],
        "curr_tot": curr_tot, "prev_tot": prev_tot, "ytd_tot": ytd_tot,
        "xx_tot": xx_tot, "oem_tot": oem_tot,
        "xx_active": int((df["Curr_XX_Amt"] > 0).sum()), "oem_active": int((df["Curr_OEM_Amt"] > 0).sum()),
        "abc": {"total": len(df), "A": a_skus, "B": b_skus, "C": c_skus, "zero": int((df["Curr_Amt"] == 0).sum()), "neg": len(neg_df)},
        "track_df": track_grp, "sub_df": sub_grp, "brand_df": brand_grp, "cust_df": cust_grp,
        "top_skus": df.sort_values("Curr_Amt", ascending=False).head(12),
        "neg_df": neg_df, "new_skus": data_pack["df_new_skus"], "pending_df": pending_df,
        "trend_df": pd.DataFrame(trend_series),
        "new_prod": compute_new_product_insights(df, curr_m, months, rules),
        "watchlist_df": wl_df,
        "bridge": compute_revenue_bridge(df, curr_m, has_prev),
        "pvm": compute_pvm_analysis(df, has_prev),
        "conc": compute_customer_concentration(cust_grp),
        "health": compute_sku_health(df, months),
    }
    stats["titles"] = build_action_titles(stats)
    return stats


# ==============================================================================
# 第五部分：四大专业独立载体导出器（MD / ECharts BI HTML / 4表Excel / 咨询级PPTX）
# ==============================================================================
def export_markdown_report(stats: dict, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    m, hp = stats["curr_m"], stats["has_prev"]
    tot_w, prev_w, ytd_w = stats["curr_tot"] / 1e4, stats["prev_tot"] / 1e4, stats["ytd_tot"] / 1e4
    abc, npd, wl_df = stats["abc"], stats["new_prod"], stats["watchlist_df"]
    bridge, pvm, conc, health = stats["bridge"], stats["pvm"], stats["conc"], stats["health"]

    lines = [
        f"# {m} 月度销售经营专业分析报告（快消 BI 归因版）", "",
        f"> **核心判断**：{stats['titles']['trend']}；{stats['titles']['bridge']}。", "",
        "## 一、 经营大盘、营收桥接归因与 PVM 量价拆解",
        f"- **当月总营收**：**{tot_w:,.2f} 万元**（上期 {prev_w:,.2f} 万元，环比 **{safe_mom(stats['curr_tot'], stats['prev_tot'], hp)}**），本年累计（YTD）**{ytd_w:,.2f} 万元**。",
    ]
    if stats["has_detail"]:
        lines.append(f"- **商业板块**：休闲自主 **{stats['xx_tot']/1e4:,.2f} 万元**（{stats['xx_tot']/(stats['curr_tot'] or 1)*100:.1f}%，动销 {stats['xx_active']} 款）；OEM/KA 定制 **{stats['oem_tot']/1e4:,.2f} 万元**（{stats['oem_tot']/(stats['curr_tot'] or 1)*100:.1f}%，动销 {stats['oem_active']} 款）。")
    lines.append(f"- **SKU 动销健康度**：总 SKU {abc['total']} 款，当月动销率 **{health['active_rate']:.1f}%**（A 类核心品 {abc['A']} 款撑起 80% 营收），连续 {health['tail_n']} 月零动销僵尸品 **{health['zombie_cnt']} 款**，冲红退货占比 **{health['return_ratio']:.2f}%**。")

    if bridge["has_prev"]:
        lines.extend([
            "", "### 1. 营收桥接归因（Revenue Bridge：精准对平月度净变动）",
            "| 驱动因子 | 涉及SKU数 | 影响金额(万元) | 性质 |",
            "| :--- | :---: | :---: | :---: |",
            f"| **上期基数（{stats['prev_m']}）** | — | **{bridge['prev_tot']/1e4:,.2f}** | 基准 |",
        ])
        for it in bridge["items"]:
            sign_tag = "🟢 正向拉动" if it["val"] >= 0 else "🔴 负向拖累"
            lines.append(f"| {it['label']} | {it['cnt']} | **{it['val']/1e4:+,.2f}** | {sign_tag} |")
        lines.append(f"| **本期实绩（{m}）** | — | **{bridge['curr_tot']/1e4:,.2f}** | 净变动 {bridge['net_diff']/1e4:+,.2f}万 |")

    if pvm["has_prev"]:
        lines.extend([
            "", "### 2. PVM 量价结构拆解（Volume vs Price/Mix）",
            f"- **总发货件数**：{pvm['prev_q']:,.1f} $\\rightarrow$ **{pvm['curr_q']:,.1f} 件**（环比 **{pvm['q_mom_pct']:+.1f}%**），**销量变动影响：{pvm['vol_effect']/1e4:+,.2f} 万元**。",
            f"- **件均单价（ASP）**：¥{pvm['prev_asp']:,.2f} $\\rightarrow$ **¥{pvm['curr_asp']:,.2f}/件**（环比 **{pvm['p_mom_pct']:+.1f}%**），**单价/结构变动影响：{pvm['price_mix_effect']/1e4:+,.2f} 万元**。",
            f"- **拆解结论**：本月营收波动由 **{pvm['driver']}**。",
        ])

    lines.extend(["", "## 二、 新品销售专项表现与趋势洞察"])
    for idx, ins in enumerate(npd["insights"], 1):
        lines.append(f"{idx}. {ins}")
    lines.extend([
        "", "| 上市批次 | 定位 | 批次SKU | 当月动销 | 留存率 | 首发月(万) | 当月(万) | 累计(万) |",
        "| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |",
    ])
    for _, r in npd["cohort_df"].iterrows():
        lines.append(f"| **{r['Batch']}** | {r['Role']} | {int(r['SKU_Count'])} | {int(r['Active_Curr'])} | **{r['Retention_Pct']:.1f}%** | {r['Launch_Amt']/1e4:,.2f} | **{r['Curr_Amt']/1e4:,.2f}** | {r['YTD_Amt']/1e4:,.2f} |")

    if not wl_df.empty:
        recent_ms = stats["months"][-6:]
        m_hdrs = " | ".join([str(x) for x in recent_ms])
        m_algn = " | ".join([":---:" for _ in recent_ms])
        lines.extend([
            "", f"## 三、 核心品月度持续追踪（共 {len(wl_df)} 款 | 万元）",
            f"| 编码 | 核心单品 | 品牌 | {m_hdrs} | 环比 | 全年累计(万) |",
            f"| :---: | :--- | :---: | {m_algn} | :---: | :---: |",
        ])
        for _, r in wl_df.iterrows():
            m_vals = " | ".join([f"{r.get(f'Amt_{wm}', 0.0)/1e4:,.2f}" for wm in recent_ms])
            lines.append(f"| `{r['Code']}` | **{r['Disp_Name']}** | {r['Brand']} | {m_vals} | {safe_mom(r['Curr_Amt'], r['Prev_Amt'], hp)} | **{r['YTD_Amt']/1e4:,.2f}** |")

    lines.extend([
        "", "## 四、 品类赛道、BCG 矩阵与客户集中度诊断",
        "| 品类赛道 | 动销/存量 | 当月(万) | 上月(万) | 环比 | 全年累计(万) | 占比 |",
        "| :--- | :---: | :---: | :---: | :---: | :---: | :---: |",
    ])
    for _, r in stats["track_df"].iterrows():
        lines.append(f"| **{r['Track']}** | {int(r['Active_SKU'])}/{int(r['SKU_Count'])} | **{r['Curr_Amt']/1e4:,.2f}** | {r['Prev_Amt']/1e4:,.2f} | {safe_mom(r['Curr_Amt'], r['Prev_Amt'], hp)} | {r['YTD_Amt']/1e4:,.2f} | {r['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}% |")

    if conc["available"]:
        lines.extend([
            "", "### 客户集中度风控（CR3 / CR5 / HHI 指数）",
            "| 统计口径 | 动销客户数 | CR3(前3占比) | CR5(前5占比) | HHI指数 | 第一大客户(占比) | 风险评级 |",
            "| :--- | :---: | :---: | :---: | :---: | :--- | :---: |",
        ])
        for sc in conc["scopes"]:
            lines.append(f"| **{sc['scope']}** | {sc['cust_cnt']} | **{sc['cr3']:.1f}%** | {sc['cr5']:.1f}% | `{sc['hhi']}` | {sc['top1']} ({sc['top1_pct']:.1f}%) | {sc['risk']} |")

    lines.extend([
        "", "## 五、 风险预警：僵尸 SKU 清退与负向退货冲红",
        f"- **连续 {health['tail_n']} 个月零动销（建议评估清退）**：共 **{health['zombie_cnt']} 款**（详细清单见独立 Excel 看板《04_僵尸品与冲红预警》）。",
    ])
    if not stats["neg_df"].empty:
        lines.append(f"- **当月负向退货/冲红预警（共 {len(stats['neg_df'])} 款，合计 {health['neg_amt']/1e4:,.2f} 万元）**：")
        for _, r in stats["neg_df"].head(8).iterrows():
            lines.append(f"  - **{r['Disp_Name']}**（`{r['Code']}`）：数量 `{r['Curr_Qty']}`，金额 `{r['Curr_Amt']:,.2f} 元`。")

    md_path = os.path.join(out_dir, f"{m}_月度经营分析报告.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return md_path


def export_html_report(stats: dict, out_dir: str) -> str:
    """现代商务 BI 交互式驾驶舱（Apache ECharts 5 + 4大板块Tab切换 + 实时表格搜索）"""
    os.makedirs(out_dir, exist_ok=True)
    m, hp = stats["curr_m"], stats["has_prev"]
    ym_str = f"{str(m)[:4]}年{int(str(m)[4:]):02d}月"
    tot_w, prev_w, ytd_w = stats["curr_tot"] / 1e4, stats["prev_tot"] / 1e4, stats["ytd_tot"] / 1e4
    xx_w, oem_w = stats["xx_tot"] / 1e4, stats["oem_tot"] / 1e4
    mom_str = safe_mom(stats["curr_tot"], stats["prev_tot"], hp)
    mom_cls = "up" if (hp and stats["curr_tot"] >= stats["prev_tot"]) else "down"
    npd, wl_df, bridge, pvm, conc, health = (
        stats["new_prod"], stats["watchlist_df"], stats["bridge"], stats["pvm"], stats["conc"], stats["health"]
    )

    def mom_badge(c, p):
        s = safe_mom(c, p, hp)
        cls = "up" if s.startswith("+") else ("down" if s.startswith("-") else "neu")
        return f"<span class='pill {cls}'>{s}</span>"

    def to_rows(df_in, fn):
        return "".join(["<tr>" + "".join([f"<td>{c}</td>" for c in fn(r)]) + "</tr>" for _, r in df_in.iterrows()])

    tr = stats["trend_df"]
    chart_data = {
        "trendMonths": tr["Month"].tolist(),
        "trendVals": tr["Amt_Wan"].tolist(),
        "hasDetail": stats["has_detail"],
        "channelPie": [
            {"name": "休闲自主", "value": round(max(0.0, xx_w), 2)},
            {"name": "OEM/KA定制", "value": round(max(0.0, oem_w), 2)},
        ],
        "trackNames": stats["track_df"]["Track"].head(8).tolist()[::-1],
        "trackCurr": [round(v / 1e4, 2) for v in stats["track_df"]["Curr_Amt"].head(8).tolist()][::-1],
        "trackPrev": [round(v / 1e4, 2) for v in stats["track_df"]["Prev_Amt"].head(8).tolist()][::-1],
        "bridge": {"hasPrev": bridge["has_prev"], "labels": [], "base": [], "pos": [], "neg": [], "total": []},
        "wlMonths": [f"{str(x)[4:]}月" for x in stats["months"][-8:]],
        "wlSeries": [],
    }
    if bridge["has_prev"]:
        b_lbls = [f"{stats['prev_m']}基数"] + [it["label"] for it in bridge["items"]] + [f"{m}实绩"]
        running = round(bridge["prev_tot"] / 1e4, 2)
        base_a, pos_a, neg_a, tot_a = [0], ["-"], ["-"], [running]
        for it in bridge["items"]:
            v = round(it["val"] / 1e4, 2)
            if v >= 0:
                base_a.append(round(running, 2))
                pos_a.append(v)
                neg_a.append("-")
            else:
                base_a.append(round(running + v, 2))
                pos_a.append("-")
                neg_a.append(abs(v))
            tot_a.append("-")
            running += v
        base_a.append(0)
        pos_a.append("-")
        neg_a.append("-")
        tot_a.append(round(bridge["curr_tot"] / 1e4, 2))
        chart_data["bridge"] = {"hasPrev": True, "labels": b_lbls, "base": base_a, "pos": pos_a, "neg": neg_a, "total": tot_a}

    for _, r in wl_df.head(8).iterrows():
        chart_data["wlSeries"].append({
            "name": clean_display_name(r["Name"], 18),
            "data": [round(r.get(f"Amt_{wm}", 0.0) / 1e4, 2) for wm in stats["months"][-8:]],
        })

    ins_html = "".join([f"<li>{x}</li>" for x in npd["insights"]])
    win_ms = npd["window_months"]
    np_th = "".join([f"<th>{wm}(万)</th>" for wm in win_ms])
    np_rows = to_rows(npd["pool_df"].head(20), lambda r: (
        [f"<b>{r['Disp_Name']}</b>", r["Brand"], int(r["First_Month_YM"])]
        + [f"{r.get(f'Amt_{wm}', 0.0)/1e4:,.2f}" for wm in win_ms]
        + [f"<b>{r['YTD_Amt']/1e4:,.2f}</b>", f"<span class='tag'>{r['Trend_Status']}</span>"]
    ))
    cohort_rows = to_rows(npd["cohort_df"], lambda r: [
        f"<b>{r['Batch']}</b>", f"<span class='tag'>{r['Role']}</span>", int(r["SKU_Count"]), int(r["Active_Curr"]),
        f"<b>{r['Retention_Pct']:.1f}%</b>", f"{r['Launch_Amt']/1e4:,.2f}", f"<b>{r['Curr_Amt']/1e4:,.2f}</b>", f"{r['YTD_Amt']/1e4:,.2f}"
    ])

    recent_ms = stats["months"][-6:]
    wl_th = "".join([f"<th>{wm}</th>" for wm in recent_ms])
    wl_rows = to_rows(wl_df, lambda r: (
        [f"<b>{r['Disp_Name']}</b>", r["Brand"]]
        + [f"{r.get(f'Amt_{wm}', 0.0)/1e4:,.2f}" for wm in recent_ms]
        + [mom_badge(r["Curr_Amt"], r["Prev_Amt"]), f"<b>{r['YTD_Amt']/1e4:,.2f}</b>"]
    )) if not wl_df.empty else "<tr><td colspan='10' style='text-align:center;color:#64748b;'>暂未勾选核心品（运行 --preview 可一键勾选并持续追踪）</td></tr>"

    sub_rows = to_rows(stats["sub_df"].head(15), lambda r: [
        f"<b>{r['SubCat']}</b>", r["Track"], int(r["Active_SKU"]),
        f"<b>{r['Curr_Amt']/1e4:,.2f}</b>", f"{r['Prev_Amt']/1e4:,.2f}", mom_badge(r["Curr_Amt"], r["Prev_Amt"]),
        f"{r['YTD_Amt']/1e4:,.2f}", f"<span class='tag'>{r['BCG']}</span>"
    ])
    brand_rows = to_rows(stats["brand_df"].head(15), lambda r: [
        f"<b>{r['Brand']}</b>", f"{int(r['Active_SKU'])}/{int(r['Total_SKU'])}",
        f"<b>{r['Curr_Amt']/1e4:,.2f}</b>", f"{r['Prev_Amt']/1e4:,.2f}", mom_badge(r["Curr_Amt"], r["Prev_Amt"]),
        f"{r['YTD_Amt']/1e4:,.2f}", f"{r['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}%"
    ])
    conc_rows = "".join([
        f"<tr><td><b>{sc['scope']}</b></td><td>{sc['cust_cnt']}</td><td><b>{sc['cr3']:.1f}%</b></td><td>{sc['cr5']:.1f}%</td><td>{sc['hhi']}</td><td>{sc['top1']} ({sc['top1_pct']:.1f}%)</td><td>{sc['risk']}</td></tr>"
        for sc in conc.get("scopes", [])
    ]) if conc["available"] else "<tr><td colspan='7' style='text-align:center;color:#64748b;'>离线补出模式未检索到客户明细原表</td></tr>"

    neg_rows = to_rows(stats["neg_df"].head(15), lambda r: [
        f"<code>{r['Code']}</code>", f"<b>{r['Disp_Name']}</b>", r["Brand"], f"{r['Curr_Qty']:,.1f}", f"<span class='pill down'>¥{r['Curr_Amt']:,.2f}</span>"
    ]) if not stats["neg_df"].empty else "<tr><td colspan='5' style='text-align:center;color:#16a34a;'>本月无负向退货冲红记录</td></tr>"

    zomb_rows = to_rows(health["zombie_df"].head(25), lambda r: [
        f"<code>{r['Code']}</code>", f"<b>{r['Disp_Name']}</b>", r["Brand"], r["Track"], f"{r['YTD_Amt']/1e4:,.2f}"
    ]) if not health["zombie_df"].empty else "<tr><td colspan='5' style='text-align:center;color:#16a34a;'>无连续 3 月零动销僵尸品</td></tr>"

    pvm_banner = (
        f"<div class='pvm-grid'>"
        f"<div class='pvm-box'><div class='pl'>发货总件数变动</div><div class='pv'>{pvm['curr_q']:,.0f} 件 <span class='pill {'up' if pvm['q_mom_pct']>=0 else 'down'}'>{pvm['q_mom_pct']:+.1f}%</span></div><div class='ps'>销量影响：<b>{pvm['vol_effect']/1e4:+,.2f} 万</b></div></div>"
        f"<div class='pvm-box'><div class='pl'>件均单价 (ASP) 变动</div><div class='pv'>¥{pvm['curr_asp']:,.2f}/件 <span class='pill {'up' if pvm['p_mom_pct']>=0 else 'down'}'>{pvm['p_mom_pct']:+.1f}%</span></div><div class='ps'>单价/结构影响：<b>{pvm['price_mix_effect']/1e4:+,.2f} 万</b></div></div>"
        f"<div class='pvm-box'><div class='pl'>PVM 归因主驱动判定</div><div class='pv' style='color:#1d4ed8;'>{pvm['driver']}</div><div class='ps'>上期均价 ¥{pvm['prev_asp']:,.2f} → 本期 ¥{pvm['curr_asp']:,.2f}</div></div>"
        f"</div>"
    ) if pvm["has_prev"] else ""

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{ym_str} 快消经营分析 BI 驾驶舱</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<style>
  :root {{ --navy:#0f172a; --blue:#1d4ed8; --amber:#d97706; --bg:#f1f5f9; --card:#ffffff; --border:#e2e8f0; --text:#1e293b; --muted:#64748b; }}
  * {{ box-sizing:border-box; }}
  body {{ font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif; background:var(--bg); color:var(--text); margin:0; }}
  .topbar {{ background:linear-gradient(135deg,#0f172a 0%,#1e3a8a 100%); color:#fff; padding:20px 28px 16px; }}
  .topbar-inner {{ max-width:1320px; margin:0 auto; display:flex; flex-wrap:wrap; justify-content:space-between; align-items:flex-end; gap:12px; }}
  .topbar h1 {{ margin:0; font-size:21px; font-weight:700; letter-spacing:.3px; }}
  .topbar .sub {{ margin:6px 0 0; font-size:13px; color:#cbd5e1; }}
  .tabs {{ display:flex; gap:8px; margin-top:14px; flex-wrap:wrap; }}
  .tab-btn {{ background:rgba(255,255,255,.12); color:#e2e8f0; border:1px solid rgba(255,255,255,.18); padding:7px 15px; border-radius:999px; cursor:pointer; font-size:13px; font-weight:600; transition:.15s; }}
  .tab-btn:hover {{ background:rgba(255,255,255,.22); }}
  .tab-btn.active {{ background:#fff; color:#0f172a; border-color:#fff; }}
  .wrap {{ max-width:1320px; margin:20px auto 36px; padding:0 20px; }}
  .panel {{ display:none; }}
  .panel.active {{ display:block; }}
  .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:14px; margin-bottom:18px; }}
  .kpi {{ background:var(--card); border-radius:12px; padding:16px 18px; border:1px solid var(--border); box-shadow:0 1px 2px rgba(15,23,42,.04); position:relative; overflow:hidden; }}
  .kpi::before {{ content:""; position:absolute; top:0; left:0; right:0; height:3px; background:var(--blue); }}
  .kpi.amber::before {{ background:var(--amber); }}
  .kpi.teal::before {{ background:#0d9488; }}
  .kpi.rose::before {{ background:#e11d48; }}
  .kpi .lbl {{ font-size:12px; color:var(--muted); font-weight:600; }}
  .kpi .val {{ font-size:26px; font-weight:800; color:var(--navy); margin:6px 0 4px; font-variant-numeric:tabular-nums; }}
  .kpi .foot {{ font-size:12px; color:#475569; display:flex; align-items:center; gap:6px; flex-wrap:wrap; font-variant-numeric:tabular-nums; }}
  .grid-2 {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(460px,1fr)); gap:16px; margin-bottom:16px; }}
  .sec {{ background:var(--card); border-radius:12px; padding:18px 20px; border:1px solid var(--border); box-shadow:0 1px 2px rgba(15,23,42,.04); margin-bottom:16px; overflow-x:auto; }}
  .sec-hd {{ display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:12px; flex-wrap:wrap; }}
  .sec h2 {{ font-size:15px; margin:0; color:var(--navy); border-left:4px solid var(--blue); padding-left:10px; }}
  .sec.warm h2 {{ border-left-color:var(--amber); color:#92400e; }}
  .action-bar {{ background:#eff6ff; border-left:4px solid var(--blue); color:#1e3a8a; padding:10px 14px; border-radius:8px; font-size:13px; font-weight:600; margin-bottom:16px; }}
  .insights {{ background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:12px 16px 12px 30px; margin:0 0 14px; font-size:13px; line-height:1.7; color:#451a03; }}
  .chart-box {{ width:100%; height:330px; }}
  .pvm-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(240px,1fr)); gap:12px; margin-bottom:14px; }}
  .pvm-box {{ background:#f8fafc; border:1px solid var(--border); border-radius:10px; padding:12px 14px; }}
  .pvm-box .pl {{ font-size:12px; color:var(--muted); }}
  .pvm-box .pv {{ font-size:18px; font-weight:700; margin:4px 0; font-variant-numeric:tabular-nums; }}
  .pvm-box .ps {{ font-size:12px; color:#475569; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; font-variant-numeric:tabular-nums; }}
  th, td {{ padding:9px 11px; border-bottom:1px solid var(--border); text-align:right; }}
  th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) {{ text-align:left; }}
  th {{ background:#f8fafc; color:#334155; font-weight:600; white-space:nowrap; }}
  tr:hover td {{ background:#f8fafc; }}
  .pill {{ display:inline-block; padding:2px 7px; border-radius:999px; font-size:11.5px; font-weight:700; font-variant-numeric:tabular-nums; }}
  .pill.up {{ background:#dcfce7; color:#15803d; }}
  .pill.down {{ background:#fee2e2; color:#b91c1c; }}
  .pill.neu {{ background:#e2e8f0; color:#475569; }}
  .tag {{ background:#eff6ff; color:#1d4ed8; padding:2px 7px; border-radius:5px; font-size:12px; white-space:nowrap; }}
  .search-inp {{ border:1px solid #cbd5e1; border-radius:8px; padding:6px 11px; font-size:12.5px; width:220px; outline:none; }}
  .search-inp:focus {{ border-color:var(--blue); }}
  .offline-tip {{ display:none; background:#fef2f2; color:#991b1b; padding:10px 14px; border-radius:8px; font-size:12.5px; margin-bottom:14px; }}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-inner">
    <div>
      <h1>📊 {ym_str} 快消销售业绩经营 BI 驾驶舱</h1>
      <div class="sub">动态数据驱动 · Revenue Bridge 瀑布归因 · PVM 量价拆解 · 新品 Cohort 洞察 · 核心品追踪</div>
      <div class="tabs">
        <button class="tab-btn active" onclick="switchTab('overview', this)">1. 总览与量价归因</button>
        <button class="tab-btn" onclick="switchTab('growth', this)">2. 新品洞察与核心品追踪</button>
        <button class="tab-btn" onclick="switchTab('structure', this)">3. 品类矩阵与客户集中度</button>
        <button class="tab-btn" onclick="switchTab('risk', this)">4. 僵尸品清退与冲红预警</button>
      </div>
    </div>
    <div style="font-size:12px;color:#cbd5e1;text-align:right;">统计口径：截至 {m}<br>请用 Chrome / Edge 浏览器打开以获得完整交互图表</div>
  </div>
</div>

<div class="wrap">
  <div id="offlineTip" class="offline-tip">⚠️ 当前处于无网络环境，未能加载 ECharts 图表脚本；所有专业分析数据表格与归因结论仍完整可用。</div>

  <!-- TAB 1: 总览与归因 -->
  <div id="tab-overview" class="panel active">
    <div class="action-bar">💡 本期经营结论：{stats['titles']['trend']} ｜ {stats['titles']['bridge']}</div>
    <div class="kpis">
      <div class="kpi"><div class="lbl">当月总营收</div><div class="val">¥{tot_w:,.2f}万</div><div class="foot">环比上月 <span class="pill {mom_cls}">{mom_str}</span> (上期 ¥{prev_w:,.2f}万)</div></div>
      <div class="kpi amber"><div class="lbl">近{len(win_ms)}月新品池贡献</div><div class="val">¥{npd['pool_curr_amt']/1e4:,.2f}万</div><div class="foot">占全盘 <b>{npd['pool_share_pct']:.1f}%</b> · 动销 {npd['pool_active']}/{len(npd['pool_df'])} 款</div></div>
      <div class="kpi teal"><div class="lbl">本年累计营收 (YTD)</div><div class="val">¥{ytd_w:,.2f}万</div><div class="foot">A 类核心品 <b>{stats['abc']['A']}</b> 款 (占80%营收)</div></div>
      <div class="kpi rose"><div class="lbl">SKU 健康度与退货率</div><div class="val">{health['active_rate']:.1f}% 动销</div><div class="foot">连续{health['tail_n']}月僵尸品 <b>{health['zombie_cnt']}</b> 款 · 冲红率 {health['return_ratio']:.2f}%</div></div>
    </div>

    <div class="sec">
      <div class="sec-hd"><h2>营收桥接瀑布归因（Revenue Bridge）与 PVM 量价拆解</h2></div>
      {pvm_banner}
      <div id="chartBridge" class="chart-box"></div>
    </div>

    <div class="grid-2">
      <div class="sec">
        <div class="sec-hd"><h2>全年各月总营收走势（万元）</h2></div>
        <div id="chartTrend" class="chart-box"></div>
      </div>
      <div class="sec">
        <div class="sec-hd"><h2>休闲自主 vs OEM/KA 定制板块占比</h2></div>
        <div id="chartChannel" class="chart-box"></div>
      </div>
    </div>
  </div>

  <!-- TAB 2: 新品与核心品 -->
  <div id="tab-growth" class="panel">
    <div class="sec warm">
      <div class="sec-hd"><h2>🚀 新品销售专项与趋势简要分析（近 {len(win_ms)} 个月批次）</h2></div>
      <ul class="insights">{ins_html}</ul>
      <table><thead><tr><th>上市批次</th><th>角色</th><th>批次SKU</th><th>当月动销</th><th>动销留存率</th><th>首发月(万)</th><th>当月(万)</th><th>累计(万)</th></tr></thead><tbody>{cohort_rows}</tbody></table>
    </div>
    <div class="sec warm">
      <div class="sec-hd">
        <h2>新品池爬坡明细（TOP 20）</h2>
        <input class="search-inp" placeholder="搜索新品名/品牌/状态…" oninput="filterTable('tblNew', this.value)">
      </div>
      <table id="tblNew"><thead><tr><th>新品名称</th><th>品牌</th><th>上市批次</th>{np_th}<th>累计(万)</th><th>趋势状态</th></tr></thead><tbody>{np_rows}</tbody></table>
    </div>
    <div class="sec">
      <div class="sec-hd"><h2>🎯 核心品月度持续追踪（共 {len(wl_df)} 款）</h2></div>
      <div id="chartWatchlist" class="chart-box" style="height:350px;margin-bottom:14px;"></div>
      <table><thead><tr><th>核心追踪单品</th><th>品牌</th>{wl_th}<th>当月环比</th><th>全年累计(万)</th></tr></thead><tbody>{wl_rows}</tbody></table>
    </div>
  </div>

  <!-- TAB 3: 品类与客户 -->
  <div id="tab-structure" class="panel">
    <div class="sec">
      <div class="sec-hd"><h2>全量品类赛道当月 vs 上月对比（万元）</h2></div>
      <div id="chartTrack" class="chart-box"></div>
    </div>
    <div class="sec">
      <div class="sec-hd">
        <h2>细分品类（类别2）波士顿矩阵 TOP 15</h2>
        <input class="search-inp" placeholder="搜索细分品类/象限…" oninput="filterTable('tblSub', this.value)">
      </div>
      <table id="tblSub"><thead><tr><th>细分品类</th><th>所属赛道</th><th>动销SKU</th><th>当月(万)</th><th>上月(万)</th><th>环比</th><th>累计(万)</th><th>BCG定位</th></tr></thead><tbody>{sub_rows}</tbody></table>
    </div>
    <div class="sec">
      <div class="sec-hd"><h2>客户集中度风控（CR3 / CR5 / HHI 指数）</h2></div>
      <table><thead><tr><th>统计口径</th><th>动销客户</th><th>CR3(前3)</th><th>CR5(前5)</th><th>HHI指数</th><th>第一大客户(占比)</th><th>风险评级</th></tr></thead><tbody>{conc_rows}</tbody></table>
    </div>
    <div class="sec">
      <div class="sec-hd">
        <h2>核心品牌矩阵 TOP 15（独立核算）</h2>
        <input class="search-inp" placeholder="搜索品牌…" oninput="filterTable('tblBrand', this.value)">
      </div>
      <table id="tblBrand"><thead><tr><th>品牌名称</th><th>动销/总SKU</th><th>当月(万)</th><th>上月(万)</th><th>环比</th><th>累计(万)</th><th>当月占比</th></tr></thead><tbody>{brand_rows}</tbody></table>
    </div>
  </div>

  <!-- TAB 4: 风险与清退 -->
  <div id="tab-risk" class="panel">
    <div class="action-bar">🛡️ SKU 健康诊断：{stats['titles']['risk']}</div>
    <div class="sec">
      <div class="sec-hd"><h2>⚠️ 当月负向退货 / 冲红异常明细（共 {len(stats['neg_df'])} 款）</h2></div>
      <table><thead><tr><th>编码</th><th>产品名称</th><th>品牌</th><th>冲红数量</th><th>冲红金额</th></tr></thead><tbody>{neg_rows}</tbody></table>
    </div>
    <div class="sec">
      <div class="sec-hd">
        <h2>🧹 连续 {health['tail_n']} 个月零动销僵尸 SKU 清退候选（前 25 款 / 共 {health['zombie_cnt']} 款）</h2>
        <input class="search-inp" placeholder="搜索僵尸品/品牌/赛道…" oninput="filterTable('tblZombie', this.value)">
      </div>
      <table id="tblZombie"><thead><tr><th>编码</th><th>产品名称</th><th>品牌</th><th>所属赛道</th><th>本年累计(万)</th></tr></thead><tbody>{zomb_rows}</tbody></table>
    </div>
  </div>
</div>

<script>
const DATA = {json.dumps(chart_data, ensure_ascii=False)};
const charts = [];
function switchTab(id, btn) {{
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  btn.classList.add('active');
  setTimeout(() => charts.forEach(c => c && c.resize()), 60);
}}
function filterTable(tblId, kw) {{
  const q = kw.trim().toLowerCase();
  document.querySelectorAll('#' + tblId + ' tbody tr').forEach(tr => {{
    tr.style.display = tr.innerText.toLowerCase().includes(q) ? '' : 'none';
  }});
}}
if (typeof echarts === 'undefined') {{
  document.getElementById('offlineTip').style.display = 'block';
}} else {{
  const cBridge = echarts.init(document.getElementById('chartBridge'));
  if (DATA.bridge.hasPrev) {{
    cBridge.setOption({{
      tooltip: {{ trigger: 'axis', axisPointer: {{ type: 'shadow' }} }},
      grid: {{ left: 55, right: 25, top: 36, bottom: 45 }},
      xAxis: {{ type: 'category', data: DATA.bridge.labels, axisLabel: {{ interval: 0, fontSize: 11 }} }},
      yAxis: {{ type: 'value', name: '万元' }},
      series: [
        {{ name: '辅助', type: 'bar', stack: 'wf', itemStyle: {{ color: 'transparent' }}, emphasis: {{ itemStyle: {{ color: 'transparent' }} }}, data: DATA.bridge.base }},
        {{ name: '基准/实绩', type: 'bar', stack: 'wf', itemStyle: {{ color: '#1e3a8a', borderRadius: [4,4,0,0] }}, label: {{ show: true, position: 'top', formatter: '{{c}}' }}, data: DATA.bridge.total }},
        {{ name: '正向增量', type: 'bar', stack: 'wf', itemStyle: {{ color: '#16a34a', borderRadius: [4,4,0,0] }}, label: {{ show: true, position: 'top', formatter: '+{{c}}' }}, data: DATA.bridge.pos }},
        {{ name: '负向拖累', type: 'bar', stack: 'wf', itemStyle: {{ color: '#dc2626', borderRadius: [4,4,0,0] }}, label: {{ show: true, position: 'bottom', formatter: '-{{c}}' }}, data: DATA.bridge.neg }}
      ]
    }});
  }}
  charts.push(cBridge);

  const cTrend = echarts.init(document.getElementById('chartTrend'));
  cTrend.setOption({{
    tooltip: {{ trigger: 'axis' }},
    grid: {{ left: 50, right: 20, top: 30, bottom: 40 }},
    xAxis: {{ type: 'category', data: DATA.trendMonths }},
    yAxis: {{ type: 'value', name: '万元' }},
    series: [
      {{ name: '月营收', type: 'bar', data: DATA.trendVals, itemStyle: {{ color: '#2563eb', borderRadius: [5,5,0,0] }}, barMaxWidth: 32 }},
      {{ name: '趋势线', type: 'line', data: DATA.trendVals, smooth: true, itemStyle: {{ color: '#d97706' }}, lineStyle: {{ width: 2.5 }} }}
    ]
  }});
  charts.push(cTrend);

  const cCh = echarts.init(document.getElementById('chartChannel'));
  cCh.setOption({{
    tooltip: {{ trigger: 'item', formatter: '{{b}}: ¥{{c}}万 ({{d}}%)' }},
    legend: {{ bottom: 0 }},
    series: [{{
      type: 'pie', radius: ['48%', '74%'], avoidLabelOverlap: true,
      itemStyle: {{ borderRadius: 6, borderColor: '#fff', borderWidth: 2 }},
      label: {{ formatter: '{{b}}\\n{{d}}%' }},
      color: ['#1e3a8a', '#0ea5e9'],
      data: DATA.hasDetail ? DATA.channelPie : [{{ name: '无渠道明细', value: 1 }}]
    }}]
  }});
  charts.push(cCh);

  const cTrack = echarts.init(document.getElementById('chartTrack'));
  cTrack.setOption({{
    tooltip: {{ trigger: 'axis', axisPointer: {{ type: 'shadow' }} }},
    legend: {{ top: 0 }},
    grid: {{ left: 165, right: 30, top: 34, bottom: 25 }},
    xAxis: {{ type: 'value', name: '万元' }},
    yAxis: {{ type: 'category', data: DATA.trackNames }},
    series: [
      {{ name: '当月(万)', type: 'bar', data: DATA.trackCurr, itemStyle: {{ color: '#1e3a8a', borderRadius: [0,4,4,0] }} }},
      {{ name: '上月(万)', type: 'bar', data: DATA.trackPrev, itemStyle: {{ color: '#94a3b8', borderRadius: [0,4,4,0] }} }}
    ]
  }});
  charts.push(cTrack);

  const cWl = echarts.init(document.getElementById('chartWatchlist'));
  if (DATA.wlSeries.length > 0) {{
    cWl.setOption({{
      tooltip: {{ trigger: 'axis' }},
      legend: {{ type: 'scroll', bottom: 0 }},
      grid: {{ left: 50, right: 20, top: 20, bottom: 45 }},
      xAxis: {{ type: 'category', data: DATA.wlMonths }},
      yAxis: {{ type: 'value', name: '万元' }},
      series: DATA.wlSeries.map(s => ({{ name: s.name, type: 'line', smooth: true, symbolSize: 7, data: s.data }}))
    }});
  }} else {{
    document.getElementById('chartWatchlist').style.display = 'none';
  }}
  charts.push(cWl);
  window.addEventListener('resize', () => charts.forEach(c => c && c.resize()));
}}
</script>
</body>
</html>"""
    html_path = os.path.join(out_dir, f"{m}_月度经营看板.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    return html_path


def export_standalone_excel(stats: dict, out_dir: str) -> str:
    """财务级 4 工作表独立经营看板（总览与量价归因 / 新品留存 / 核心品追踪 / 僵尸与冲红预警）"""
    os.makedirs(out_dir, exist_ok=True)
    m, hp = stats["curr_m"], stats["has_prev"]
    npd, wl_df, bridge, pvm, conc, health = (
        stats["new_prod"], stats["watchlist_df"], stats["bridge"], stats["pvm"], stats["conc"], stats["health"]
    )
    wb = openpyxl.Workbook()

    navy_fill = PatternFill("solid", fgColor="1B365D")
    amber_fill = PatternFill("solid", fgColor="92400E")
    rose_fill = PatternFill("solid", fgColor="9F1239")
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")
    hdr_font = Font(name="Microsoft YaHei", color="FFFFFF", bold=True, size=10)
    sec_font = Font(name="Microsoft YaHei", size=11, bold=True, color="1B365D")
    title_font = Font(name="Microsoft YaHei", size=13, bold=True, color="1B365D")
    body_font = Font(name="Microsoft YaHei", size=10, color="1E293B")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def fmt_tbl(ws, hr, er, nc, fill=navy_fill, num_cols=None):
        num_cols = num_cols or []
        for c in range(1, nc + 1):
            cell = ws.cell(hr, c)
            cell.fill, cell.font, cell.border = fill, hdr_font, border
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for r in range(hr + 1, er + 1):
            for c in range(1, nc + 1):
                cell = ws.cell(r, c)
                cell.font, cell.border = body_font, border
                if r % 2 == 0:
                    cell.fill = zebra_fill
                if c in num_cols:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c in (1, 2):
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

    def finish_sheet(ws, freeze_row=4, max_w=36):
        ws.freeze_panes = f"A{freeze_row}"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        for col in ws.columns:
            let = get_column_letter(col[0].column)
            best = 10
            for cell in col:
                if cell.value is not None:
                    best = max(best, min(max_w, len(str(cell.value)) * 1.55 + 3))
            ws.column_dimensions[let].width = best

    # Sheet 1: 总览与量价归因
    ws1 = wb.active
    ws1.title = "01_总览与量价归因"
    ws1.append([f"{m} 快消月度经营总览与营收归因（{stats['titles']['trend']}）"])
    ws1.cell(1, 1).font = title_font
    ws1.append([
        "当月营收(万)", round(stats["curr_tot"] / 1e4, 2), "环比", safe_mom(stats["curr_tot"], stats["prev_tot"], hp),
        "全年累计(万)", round(stats["ytd_tot"] / 1e4, 2), "动销率", f"{health['active_rate']:.1f}%",
        "僵尸品数", health["zombie_cnt"],
    ])
    ws1.append([])

    if bridge["has_prev"]:
        ws1.append(["【一、营收桥接瀑布归因（Revenue Bridge）与 PVM 量价拆解】"])
        ws1.cell(ws1.max_row, 1).font = sec_font
        ws1.append(["归因因子", "涉及SKU数", "影响金额(万元)", "PVM 维度", "上期", "本期", "变动影响(万元)"])
        bh = ws1.max_row
        pvm_rows = [
            ("发货件数(Volume)", round(pvm["prev_q"], 1), round(pvm["curr_q"], 1), round(pvm["vol_effect"] / 1e4, 2)),
            ("件均单价(ASP/Mix)", pvm["prev_asp"], pvm["curr_asp"], round(pvm["price_mix_effect"] / 1e4, 2)),
            ("主驱动判定", "—", pvm["driver"], round(pvm["net_diff"] / 1e4, 2)),
        ]
        for idx, it in enumerate(bridge["items"]):
            pr = pvm_rows[idx] if idx < len(pvm_rows) else ("", "", "", "")
            ws1.append([it["label"], it["cnt"], round(it["val"] / 1e4, 2), pr[0], pr[1], pr[2], pr[3]])
        fmt_tbl(ws1, bh, ws1.max_row, 7, num_cols=[3, 7])
        ws1.append([])

    ws1.append(["【二、全量品类赛道表现】"])
    ws1.cell(ws1.max_row, 1).font = sec_font
    ws1.append(["品类赛道", "存量SKU", "动销SKU", "当月(万元)", "上月(万元)", "环比增速", "全年累计(万元)", "当月占比"])
    th1 = ws1.max_row
    for _, r in stats["track_df"].iterrows():
        ws1.append([
            r["Track"], int(r["SKU_Count"]), int(r["Active_SKU"]),
            round(r["Curr_Amt"] / 1e4, 2), round(r["Prev_Amt"] / 1e4, 2),
            safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp), round(r["YTD_Amt"] / 1e4, 2),
            f"{r['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}%"
        ])
    te1 = ws1.max_row
    fmt_tbl(ws1, th1, te1, 8, num_cols=[4, 5, 7])
    ws1.append([])

    if conc["available"]:
        ws1.append(["【三、客户集中度风控（CR3 / CR5 / HHI）】"])
        ws1.cell(ws1.max_row, 1).font = sec_font
        ws1.append(["统计口径", "动销客户数", "CR3(%)", "CR5(%)", "HHI指数", "第一大客户", "风险评级"])
        ch = ws1.max_row
        for sc in conc["scopes"]:
            ws1.append([sc["scope"], sc["cust_cnt"], sc["cr3"], sc["cr5"], sc["hhi"], f"{sc['top1']} ({sc['top1_pct']}%)", sc["risk"]])
        fmt_tbl(ws1, ch, ws1.max_row, 7, num_cols=[3, 4])

    bar1 = BarChart()
    bar1.title = f"{m} 各品类赛道当月 vs 上月营收(万元)"
    bar1.add_data(Reference(ws1, min_col=4, max_col=5, min_row=th1, max_row=te1), titles_from_data=True)
    bar1.set_categories(Reference(ws1, min_col=1, min_row=th1 + 1, max_row=te1))
    bar1.width, bar1.height = 17, 10
    ws1.add_chart(bar1, "J4")
    finish_sheet(ws1, freeze_row=4)

    # Sheet 2: 新品专项与留存
    ws2 = wb.create_sheet("02_新品专项与留存")
    ws2.append([f"{m} 新品销售专项与趋势洞察（近 {len(npd['window_months'])} 个月批次）"])
    ws2.cell(1, 1).font = title_font
    for idx, ins in enumerate(npd["insights"], 1):
        ws2.append([f"洞察 {idx}：{ins}"])
        ws2.cell(ws2.max_row, 1).font = Font(name="Microsoft YaHei", size=10, color="78350F")
    ws2.append([])
    ws2.append(["上市批次", "角色定位", "批次SKU数", "当月动销数", "留存率(%)", "首发月(万元)", "当月(万元)", "累计(万元)"])
    c_h = ws2.max_row
    for _, r in npd["cohort_df"].iterrows():
        ws2.append([r["Batch"], r["Role"], int(r["SKU_Count"]), int(r["Active_Curr"]), r["Retention_Pct"], round(r["Launch_Amt"] / 1e4, 2), round(r["Curr_Amt"] / 1e4, 2), round(r["YTD_Amt"] / 1e4, 2)])
    c_e = ws2.max_row
    fmt_tbl(ws2, c_h, c_e, 8, fill=amber_fill, num_cols=[5, 6, 7, 8])
    ws2.append([])

    win_ms = npd["window_months"]
    ws2.append(["商品编码", "新品名称", "品牌", "上市批次"] + [f"{wm}月(万)" for wm in win_ms] + ["累计(万元)", "趋势状态"])
    np_h = ws2.max_row
    for _, r in npd["pool_df"].iterrows():
        ws2.append([r["Code"], r["Disp_Name"], r["Brand"], int(r["First_Month_YM"])] + [round(r.get(f"Amt_{wm}", 0.0) / 1e4, 2) for wm in win_ms] + [round(r["YTD_Amt"] / 1e4, 2), r["Trend_Status"]])
    np_e = ws2.max_row
    fmt_tbl(ws2, np_h, np_e, 6 + len(win_ms), num_cols=list(range(5, 5 + len(win_ms) + 1)))
    if np_e > np_h:
        col_let = get_column_letter(4 + len(win_ms))
        ws2.conditional_formatting.add(f"{col_let}{np_h+1}:{col_let}{np_e}", ColorScaleRule(start_type="min", start_color="FEF3C7", end_type="max", end_color="34D399"))
    finish_sheet(ws2, freeze_row=np_h + 1)

    # Sheet 3: 核心品月度追踪
    ws3 = wb.create_sheet("03_核心品月度追踪")
    ws3.append([f"{m} 核心品月度持续追踪（共 {len(wl_df)} 款 | 单位：万元）"])
    ws3.cell(1, 1).font = title_font
    ws3.append([])
    all_ms = stats["months"]
    ws3.append(["商品编码", "核心追踪产品名称", "品牌", "细分品类"] + [str(wm) for wm in all_ms] + ["当月环比", "全年累计(万元)"])
    wl_h = ws3.max_row
    if not wl_df.empty:
        for _, r in wl_df.iterrows():
            ws3.append([r["Code"], r["Disp_Name"], r["Brand"], r["SubCat"]] + [round(r.get(f"Amt_{wm}", 0.0) / 1e4, 2) for wm in all_ms] + [safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp), round(r["YTD_Amt"] / 1e4, 2)])
        wl_e = ws3.max_row
        fmt_tbl(ws3, wl_h, wl_e, 6 + len(all_ms), num_cols=list(range(5, 5 + len(all_ms))) + [6 + len(all_ms)])
        plot_end = min(wl_e, wl_h + 6)
        lc = LineChart()
        lc.title = "核心品月度营收曲线（万元，前6款）"
        lc.add_data(Reference(ws3, min_col=5, max_col=4 + len(all_ms), min_row=wl_h + 1, max_row=plot_end), from_rows=True, titles_from_data=False)
        lc.set_categories(Reference(ws3, min_col=5, max_col=4 + len(all_ms), min_row=wl_h, max_row=wl_h))
        for si, s in enumerate(lc.series):
            s.tx = openpyxl.chart.series.SeriesLabel(v=str(ws3.cell(wl_h + 1 + si, 2).value)[:18])
        lc.width, lc.height = 22, 11
        ws3.add_chart(lc, f"B{wl_e + 3}")
    else:
        ws3.append(["—", "暂未勾选核心品，运行 --preview 可一键勾选。"])
    finish_sheet(ws3, freeze_row=wl_h + 1)

    # Sheet 4: 僵尸品与冲红预警
    ws4 = wb.create_sheet("04_僵尸品与冲红预警")
    ws4.append([f"{m} SKU 健康度与风险预警（{stats['titles']['risk']}）"])
    ws4.cell(1, 1).font = title_font
    ws4.append([])
    ws4.append(["【一、当月负向退货/冲红预警】"])
    ws4.cell(ws4.max_row, 1).font = sec_font
    ws4.append(["商品编码", "产品名称", "品牌", "冲红数量", "冲红金额(元)"])
    nh = ws4.max_row
    for _, r in stats["neg_df"].iterrows():
        ws4.append([r["Code"], r["Disp_Name"], r["Brand"], round(r["Curr_Qty"], 2), round(r["Curr_Amt"], 2)])
    fmt_tbl(ws4, nh, max(nh, ws4.max_row), 5, fill=rose_fill, num_cols=[4, 5])
    ws4.append([])
    ws4.append([f"【二、连续 {health['tail_n']} 个月零动销僵尸 SKU 清退候选（共 {health['zombie_cnt']} 款）】"])
    ws4.cell(ws4.max_row, 1).font = sec_font
    ws4.append(["商品编码", "产品名称", "品牌", "所属赛道", "本年累计(万元)"])
    zh = ws4.max_row
    for _, r in health["zombie_df"].iterrows():
        ws4.append([r["Code"], r["Disp_Name"], r["Brand"], r["Track"], round(r["YTD_Amt"] / 1e4, 2)])
    fmt_tbl(ws4, zh, max(zh, ws4.max_row), 5, num_cols=[5])
    finish_sheet(ws4, freeze_row=4)

    out_path = os.path.join(out_dir, f"{m}_独立经营分析看板.xlsx")
    wb.save(out_path)
    return out_path


def _configure_cjk_font():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm

    for fp in [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
    ]:
        if os.path.isfile(fp):
            try:
                fm.fontManager.addfont(fp)
            except Exception:
                pass

    installed = {f.name for f in fm.fontManager.ttflist}
    chosen = [f for f in ["Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "Noto Sans CJK JP", "WenQuanYi Zen Hei", "SimHei"] if f in installed]
    if chosen:
        plt.rcParams["font.sans-serif"] = chosen + ["DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    return plt


def export_pptx_report(stats: dict, out_dir: str) -> str:
    """咨询级高管汇报 PPTX（Action Title + 营收瀑布 + PVM + 双量级核心品折线 + 零截断换行表）"""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
        from pptx.enum.shapes import MSO_SHAPE
    except ImportError as e:
        raise RuntimeError(f"生成 PPTX 需安装依赖：pip install python-pptx matplotlib（缺失：{e}）")

    plt = _configure_cjk_font()
    os.makedirs(out_dir, exist_ok=True)
    m, hp = stats["curr_m"], stats["has_prev"]
    ym_title = f"{str(m)[:4]}年{int(str(m)[4:]):02d}月"
    tot_w, prev_w, ytd_w = stats["curr_tot"] / 1e4, stats["prev_tot"] / 1e4, stats["ytd_tot"] / 1e4
    mom_str = safe_mom(stats["curr_tot"], stats["prev_tot"], hp)
    npd, wl_df, bridge, pvm, conc, health, titles = (
        stats["new_prod"], stats["watchlist_df"], stats["bridge"], stats["pvm"], stats["conc"], stats["health"], stats["titles"]
    )

    NAVY = RGBColor(0x0F, 0x17, 0x2A)
    BLUE = RGBColor(0x1D, 0x4E, 0xD8)
    TEAL = RGBColor(0x0F, 0x76, 0x6E)
    AMBER = RGBColor(0xB4, 0x53, 0x09)
    ROSE = RGBColor(0xBE, 0x12, 0x3C)
    SLATE = RGBColor(0x33, 0x41, 0x55)
    MUTED = RGBColor(0x64, 0x74, 0x8B)
    BG_LIGHT = RGBColor(0xF8, 0xFA, 0xFC)
    BG_WARM = RGBColor(0xFF, 0xFB, 0xEB)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def add_action_header(slide, kicker: str, action_title: str):
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
        bar.fill.solid()
        bar.fill.fore_color.rgb = WHITE
        bar.line.fill.background()
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.02), Inches(12.333), Inches(0.03))
        accent.fill.solid()
        accent.fill.fore_color.rgb = BLUE
        accent.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(11.0), Inches(0.88))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = kicker
        p0.font.size, p0.font.bold, p0.font.color.rgb = Pt(10.5), True, BLUE
        p1 = tf.add_paragraph()
        p1.text = action_title
        p1.font.size, p1.font.bold, p1.font.color.rgb = Pt(18.5), True, NAVY
        p1.space_before = Pt(2)

        tb2 = slide.shapes.add_textbox(Inches(11.3), Inches(0.14), Inches(1.55), Inches(0.35))
        p2 = tb2.text_frame.paragraphs[0]
        p2.text, p2.alignment = f"{ym_title} BI", PP_ALIGN.RIGHT
        p2.font.size, p2.font.color.rgb = Pt(10), MUTED

    def add_wrapped_table(slide, left, top, width, height, headers, rows_data, col_ratios=None, font_size=9.5, hdr_rgb=NAVY):
        tbl_shape = slide.shapes.add_table(len(rows_data) + 1, len(headers), Inches(left), Inches(top), Inches(width), Inches(height))
        tbl = tbl_shape.table
        if col_ratios:
            tot_r = sum(col_ratios)
            for ci, r_val in enumerate(col_ratios):
                tbl.columns[ci].width = Inches(width * (r_val / tot_r))
        for ci, ht in enumerate(headers):
            cell = tbl.cell(0, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hdr_rgb
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.word_wrap = True
            p = cell.text_frame.paragraphs[0]
            p.text = str(ht)
            p.font.size, p.font.bold, p.font.color.rgb = Pt(font_size), True, WHITE
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT
        for ri, rv in enumerate(rows_data, 1):
            bg = BG_LIGHT if ri % 2 == 0 else WHITE
            for ci, val in enumerate(rv):
                cell = tbl.cell(ri, ci)
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.text_frame.word_wrap = True
                p = cell.text_frame.paragraphs[0]
                p.text = str(val)
                p.font.size, p.font.color.rgb = Pt(font_size), SLATE
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT
        return tbl

    def fig_stream(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=240, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        buf.seek(0)
        return buf

    # ---------- P1: 封面与核心 KPI ----------
    s1 = prs.slides.add_slide(blank)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(3.85))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.fill.background()
    tf = s1.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(11.7), Inches(2.2)).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"{ym_title} 销售业绩经营分析报告"
    p.font.size, p.font.bold, p.font.color.rgb = Pt(36), True, WHITE
    p2 = tf.add_paragraph()
    p2.text = f"{titles['trend']}"
    p2.font.size, p2.font.color.rgb, p2.space_before = Pt(15), RGBColor(0x93, 0xC5, 0xFD), Pt(12)
    p3 = tf.add_paragraph()
    p3.text = "Revenue Bridge 瀑布归因 · PVM 量价拆解 · 客户集中度 CR3/HHI · 新品留存与核心品追踪"
    p3.font.size, p3.font.color.rgb, p3.space_before = Pt(11.5), RGBColor(0xCB, 0xD5, 0xE1), Pt(10)

    kpis = [
        ("当月总营收", f"¥{tot_w:,.2f}万", f"环比：{mom_str}（上期 ¥{prev_w:,.2f}万）", NAVY),
        ("本年累计 (YTD)", f"¥{ytd_w:,.2f}万", f"A类核心品 {stats['abc']['A']} 款撑起80%营收", BLUE),
        (f"近{len(npd['window_months'])}月新品池", f"¥{npd['pool_curr_amt']/1e4:,.2f}万", f"占全盘 {npd['pool_share_pct']:.1f}% · 动销 {npd['pool_active']}/{len(npd['pool_df'])}款", AMBER),
        ("SKU 健康与风控", f"{health['active_rate']:.1f}% 动销", f"僵尸品 {health['zombie_cnt']}款 · 冲红率 {health['return_ratio']:.2f}%", ROSE),
    ]
    for i, (lbl, val, sub, col) in enumerate(kpis):
        x = 0.65 + i * 3.08
        card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(4.35), Inches(2.85), Inches(2.35))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
        ctf = s1.shapes.add_textbox(Inches(x + 0.18), Inches(4.52), Inches(2.5), Inches(2.0)).text_frame
        ctf.word_wrap = True
        cp0 = ctf.paragraphs[0]
        cp0.text, cp0.font.size, cp0.font.color.rgb = lbl, Pt(12), MUTED
        cp1 = ctf.add_paragraph()
        cp1.text, cp1.font.size, cp1.font.bold, cp1.font.color.rgb, cp1.space_before = val, Pt(26), True, col, Pt(8)
        cp2 = ctf.add_paragraph()
        cp2.text, cp2.font.size, cp2.font.color.rgb, cp2.space_before = sub, Pt(10.5), SLATE, Pt(10)

    # ---------- P2: 营收桥接瀑布归因 + PVM 量价拆解 ----------
    if bridge["has_prev"]:
        s_brg = prs.slides.add_slide(blank)
        add_action_header(s_brg, "01 | 营收变动归因（REVENUE BRIDGE & PVM）", titles["bridge"])

        fig_w, ax_w = plt.subplots(figsize=(7.6, 4.7))
        labels_w = [f"{stats['prev_m']}\n基数"] + [it["label"].replace("上月动销本月", "上月动销\n本月") for it in bridge["items"]] + [f"{m}\n实绩"]
        running = bridge["prev_tot"] / 1e4
        bottoms, heights, colors, texts = [0.0], [running], ["#1E3A8A"], [f"{running:.1f}"]
        for it in bridge["items"]:
            v = it["val"] / 1e4
            if v >= 0:
                bottoms.append(running)
                heights.append(v)
                colors.append("#16A34A")
                texts.append(f"+{v:.1f}")
            else:
                bottoms.append(running + v)
                heights.append(abs(v))
                colors.append("#DC2626")
                texts.append(f"{v:.1f}")
            running += v
        bottoms.append(0.0)
        heights.append(bridge["curr_tot"] / 1e4)
        colors.append("#0F172A")
        texts.append(f"{bridge['curr_tot']/1e4:.1f}")

        xs = range(len(labels_w))
        ax_w.bar(xs, heights, bottom=bottoms, color=colors, width=0.56)
        max_y = max([b + h for b, h in zip(bottoms, heights)] or [100])
        for xi, b, h, tx in zip(xs, bottoms, heights, texts):
            ax_w.text(xi, b + h + max_y * 0.02, tx, ha="center", va="bottom", fontsize=9, fontweight="bold")
        ax_w.set_xticks(list(xs))
        ax_w.set_xticklabels(labels_w, fontsize=8.5)
        ax_w.set_ylim(0, max_y * 1.16)
        ax_w.set_ylabel("万元", fontsize=9.5)
        ax_w.set_title("月度营收桥接瀑布拆解（万元）", fontsize=12, fontweight="bold")
        ax_w.spines["top"].set_visible(False)
        ax_w.spines["right"].set_visible(False)
        s_brg.shapes.add_picture(fig_stream(fig_w), Inches(0.45), Inches(1.25), width=Inches(7.7))

        p_card = s_brg.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.35), Inches(1.3), Inches(4.5), Inches(2.45))
        p_card.fill.solid()
        p_card.fill.fore_color.rgb = BG_LIGHT
        p_card.line.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
        ptf = s_brg.shapes.add_textbox(Inches(8.55), Inches(1.42), Inches(4.15), Inches(2.25)).text_frame
        ptf.word_wrap = True
        pp0 = ptf.paragraphs[0]
        pp0.text = f"📐 PVM 量价结构拆解（{pvm['driver']}）"
        pp0.font.size, pp0.font.bold, pp0.font.color.rgb = Pt(12), True, BLUE
        for ln in [
            f"• 发货件数：{pvm['prev_q']:,.0f} → {pvm['curr_q']:,.0f} 件（{pvm['q_mom_pct']:+.1f}%）",
            f"  └ 销量变动影响：{pvm['vol_effect']/1e4:+,.2f} 万元",
            f"• 件均单价：¥{pvm['prev_asp']:,.2f} → ¥{pvm['curr_asp']:,.2f}/件（{pvm['p_mom_pct']:+.1f}%）",
            f"  └ 单价/结构影响：{pvm['price_mix_effect']/1e4:+,.2f} 万元",
        ]:
            pp = ptf.add_paragraph()
            pp.text, pp.font.size, pp.font.color.rgb, pp.space_before = ln, Pt(10.2), SLATE, Pt(5)

        drag_rows = [
            [r["Disp_Name"], f"{r['Prev_Amt']/1e4:.1f}", f"{r['Curr_Amt']/1e4:.1f}", f"{r['Diff_Amt']/1e4:+.1f}"]
            for _, r in bridge["top_drags"].head(5).iterrows()
        ]
        add_wrapped_table(
            s_brg, 8.35, 3.95, 4.5, 3.05,
            ["本月主要缩量/断单品 TOP 5", "上月", "当月", "净变动"],
            drag_rows, col_ratios=[4.6, 1.2, 1.2, 1.4], font_size=8.8, hdr_rgb=ROSE
        )

    # ---------- P3: 商业板块与全年趋势 ----------
    s2 = prs.slides.add_slide(blank)
    add_action_header(s2, "02 | 全年趋势与商业板块结构", titles["trend"])
    tr = stats["trend_df"]
    fig, ax = plt.subplots(figsize=(6.6, 4.5))
    bars = ax.bar(tr["Month"], tr["Amt_Wan"], color="#2563EB", width=0.56)
    ax.plot(tr["Month"], tr["Amt_Wan"], color="#D97706", marker="o", linewidth=2.2)
    for b in bars:
        h = b.get_height()
        ax.text(b.get_x() + b.get_width() / 2, h + max(tr["Amt_Wan"].max() * 0.02, 10), f"{h:.0f}", ha="center", va="bottom", fontsize=9)
    ax.set_title("本年各月总营收趋势（万元）", fontsize=12.5, fontweight="bold", pad=10)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.xticks(rotation=30, fontsize=9)
    s2.shapes.add_picture(fig_stream(fig), Inches(0.5), Inches(1.3), width=Inches(6.9))

    if stats["has_detail"] and (stats["xx_tot"] > 0 or stats["oem_tot"] > 0):
        fig2, ax2 = plt.subplots(figsize=(4.8, 3.4))
        vals = [max(0.0, stats["xx_tot"] / 1e4), max(0.0, stats["oem_tot"] / 1e4)]
        ax2.pie(vals, labels=[f"休闲自主\n{vals[0]:.1f}万", f"OEM定制\n{vals[1]:.1f}万"],
                autopct="%1.1f%%", startangle=90, colors=["#1E3A8A", "#0EA5E9"],
                wedgeprops=dict(width=0.42, edgecolor="white"), textprops={"fontsize": 11})
        ax2.set_title(f"{ym_title} 双商业板块占比", fontsize=12, fontweight="bold")
        s2.shapes.add_picture(fig_stream(fig2), Inches(7.7), Inches(1.25), width=Inches(4.9))
        btf = s2.shapes.add_textbox(Inches(7.75), Inches(5.0), Inches(5.0), Inches(1.9)).text_frame
        btf.word_wrap = True
        for ln in [
            f"• 休闲自主：¥{vals[0]:,.2f}万（动销 {stats['xx_active']} 款 SKU）",
            f"• OEM/KA定制：¥{vals[1]:,.2f}万（动销 {stats['oem_active']} 款 SKU）",
            f"• 新品池贡献：¥{npd['pool_curr_amt']/1e4:,.2f}万（占全盘 {npd['pool_share_pct']:.1f}%）",
        ]:
            bp = btf.add_paragraph()
            bp.text, bp.font.size, bp.font.color.rgb, bp.space_after = ln, Pt(12), SLATE, Pt(6)

    # ---------- P4: 新品专项与趋势洞察 ----------
    s_np = prs.slides.add_slide(blank)
    add_action_header(s_np, "03 | 新品专项表现与 COHORT 留存洞察", titles["new_prod"])
    cdf = npd["cohort_df"]
    fig_np, ax_np = plt.subplots(figsize=(6.0, 3.1))
    xp = range(len(cdf))
    ax_np.bar([x - 0.18 for x in xp], (cdf["Launch_Amt"] / 1e4).tolist(), width=0.35, label="首发月营收(万)", color="#94A3B8")
    ax_np.bar([x + 0.18 for x in xp], (cdf["Curr_Amt"] / 1e4).tolist(), width=0.35, label="当月营收(万)", color="#D97706")
    ax_np.set_xticks(list(xp))
    ax_np.set_xticklabels([f"{r['Batch']}\n({r['Retention_Pct']:.0f}%留存)" for _, r in cdf.iterrows()], fontsize=9.5)
    ax_np.set_title("近3月上市批次：首发月 vs 当月营收（万元）", fontsize=11.5, fontweight="bold")
    ax_np.legend(fontsize=8.5)
    ax_np.spines["top"].set_visible(False)
    ax_np.spines["right"].set_visible(False)
    s_np.shapes.add_picture(fig_stream(fig_np), Inches(0.45), Inches(1.2), width=Inches(6.1))

    np_rows = [[r["Disp_Name"], int(r["First_Month_YM"]), f"{r['Curr_Amt']/1e4:,.2f}", f"{r['YTD_Amt']/1e4:,.2f}", r["Trend_Status"]] for _, r in npd["pool_df"].head(6).iterrows()]
    if np_rows:
        add_wrapped_table(
            s_np, 0.45, 4.45, 6.25, 2.75,
            ["重点新品（TOP 6）", "批次", "当月", "累计", "趋势状态"],
            np_rows, col_ratios=[4.5, 1.3, 1.2, 1.2, 2.0], font_size=8.6, hdr_rgb=AMBER
        )

    r_card = s_np.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.25), Inches(5.95), Inches(5.9))
    r_card.fill.solid()
    r_card.fill.fore_color.rgb = BG_WARM
    r_card.line.color.rgb = RGBColor(0xFD, 0xE6, 0x8A)
    itf = s_np.shapes.add_textbox(Inches(7.1), Inches(1.42), Inches(5.55), Inches(5.55)).text_frame
    itf.word_wrap = True
    ih = itf.paragraphs[0]
    ih.text = "📌 新品趋势简要分析（自动诊断结论）"
    ih.font.size, ih.font.bold, ih.font.color.rgb = Pt(13.5), True, AMBER
    for idx, ins in enumerate(npd["insights"], 1):
        ip = itf.add_paragraph()
        ip.text, ip.font.size, ip.font.color.rgb, ip.space_before = f"{idx}. {ins}", Pt(11.2), SLATE, Pt(12)

    # ---------- P5: 核心品双量级分面折线追踪（动态生成） ----------
    if not wl_df.empty:
        s_wl = prs.slides.add_slide(blank)
        add_action_header(
            s_wl, "04 | 核心品月度持续追踪（双量级分面视图）",
            f"持续追踪 {len(wl_df)} 款重点单品，当月合计创收 ¥{wl_df['Curr_Amt'].sum()/1e4:,.2f}万（按量级自动分面防曲线压扁）"
        )
        plot_ms = stats["months"][-8:]
        x_lbls = [f"{str(x)[4:]}月" for x in plot_ms]
        top_wl = wl_df.head(8).copy()
        peak_vals = [max([r.get(f"Amt_{wm}", 0.0) / 1e4 for wm in plot_ms] or [0.0]) for _, r in top_wl.iterrows()]
        top_wl["_peak"] = peak_vals
        max_p, min_p = max(peak_vals or [1.0]), min([v for v in peak_vals if v > 0] or [1.0])
        need_split = (len(top_wl) >= 2) and (max_p / max(min_p, 0.1) >= 4.0)

        palette = ["#1E3A8A", "#D97706", "#0F766E", "#2563EB", "#DC2626", "#7C3AED", "#0284C7", "#65A30D"]
        if need_split:
            cut = max_p * 0.25
            hi_df = top_wl[top_wl["_peak"] >= cut]
            lo_df = top_wl[top_wl["_peak"] < cut]
            fig_wl, (ax1, ax2) = plt.subplots(2, 1, figsize=(6.3, 5.2), sharex=True)
            for i, (_, r) in enumerate(hi_df.iterrows()):
                ax1.plot(x_lbls, [r.get(f"Amt_{wm}", 0.0) / 1e4 for wm in plot_ms], marker="o", linewidth=2, color=palette[i % len(palette)], label=clean_display_name(r["Name"], 16))
            ax1.set_title("主力大单品月度曲线（高量级，万元）", fontsize=10.5, fontweight="bold")
            ax1.legend(loc="upper left", fontsize=7.8)
            ax1.grid(axis="y", linestyle="--", alpha=0.4)
            ax1.spines["top"].set_visible(False)
            ax1.spines["right"].set_visible(False)

            for i, (_, r) in enumerate(lo_df.iterrows()):
                ax2.plot(x_lbls, [r.get(f"Amt_{wm}", 0.0) / 1e4 for wm in plot_ms], marker="s", linewidth=2, color=palette[(i + 3) % len(palette)], label=clean_display_name(r["Name"], 16))
            ax2.set_title("成长与新品孵化曲线（中低量级独立纵轴，万元）", fontsize=10.5, fontweight="bold")
            ax2.legend(loc="upper left", fontsize=7.8)
            ax2.grid(axis="y", linestyle="--", alpha=0.4)
            ax2.spines["top"].set_visible(False)
            ax2.spines["right"].set_visible(False)
            plt.tight_layout()
        else:
            fig_wl, ax1 = plt.subplots(figsize=(6.3, 5.0))
            for i, (_, r) in enumerate(top_wl.iterrows()):
                ax1.plot(x_lbls, [r.get(f"Amt_{wm}", 0.0) / 1e4 for wm in plot_ms], marker="o", linewidth=2, color=palette[i % len(palette)], label=clean_display_name(r["Name"], 16))
            ax1.set_title("核心追踪单品月度营收曲线（万元）", fontsize=11.5, fontweight="bold")
            ax1.legend(loc="upper left", fontsize=8)
            ax1.grid(axis="y", linestyle="--", alpha=0.4)
            ax1.spines["top"].set_visible(False)
            ax1.spines["right"].set_visible(False)
        s_wl.shapes.add_picture(fig_stream(fig_wl), Inches(0.45), Inches(1.25), width=Inches(6.35))

        wl_rows = [[r["Disp_Name"], f"{r['Prev_Amt']/1e4:,.2f}", f"{r['Curr_Amt']/1e4:,.2f}", safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp), f"{r['YTD_Amt']/1e4:,.2f}"] for _, r in wl_df.head(10).iterrows()]
        add_wrapped_table(
            s_wl, 6.95, 1.3, 5.9, 5.4,
            ["核心追踪单品", "上月", "当月", "环比", "累计"],
            wl_rows, col_ratios=[4.6, 1.2, 1.2, 1.4, 1.3], font_size=9.0, hdr_rgb=TEAL
        )

    # ---------- P6: 品类赛道与 BCG 矩阵 ----------
    s_tr = prs.slides.add_slide(blank)
    add_action_header(s_tr, "05 | 动态品类赛道对比", titles["track"])
    tdf = stats["track_df"].head(8)
    fig3, ax3 = plt.subplots(figsize=(6.2, 4.8))
    y_pos = range(len(tdf))
    ax3.barh([y + 0.18 for y in y_pos], (tdf["Curr_Amt"] / 1e4).tolist()[::-1], height=0.35, label="当月(万)", color="#1E3A8A")
    if hp:
        ax3.barh([y - 0.18 for y in y_pos], (tdf["Prev_Amt"] / 1e4).tolist()[::-1], height=0.35, label="上月(万)", color="#94A3B8")
    ax3.set_yticks(list(y_pos))
    ax3.set_yticklabels([str(x)[:14] for x in tdf["Track"]][::-1], fontsize=9.8)
    ax3.legend(loc="lower right", fontsize=9)
    ax3.set_title("各品类赛道当月 vs 上月营收（万元）", fontsize=12, fontweight="bold")
    ax3.spines["top"].set_visible(False)
    ax3.spines["right"].set_visible(False)
    s_tr.shapes.add_picture(fig_stream(fig3), Inches(0.45), Inches(1.25), width=Inches(6.2))
    t_rows = [[r["Track"], int(r["Active_SKU"]), f"{r['Curr_Amt']/1e4:,.2f}", f"{r['Prev_Amt']/1e4:,.2f}", safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp), f"{r['Curr_Amt']/(stats['curr_tot'] or 1)*100:.1f}%"] for _, r in tdf.iterrows()]
    add_wrapped_table(s_tr, 6.85, 1.35, 6.0, 5.0, ["品类赛道", "动销", "当月", "上月", "环比", "占比"], t_rows, col_ratios=[3.6, 0.9, 1.2, 1.2, 1.3, 1.0], font_size=9.5)

    # ---------- P7: 细分品类 BCG 矩阵 ----------
    s_bcg = prs.slides.add_slide(blank)
    add_action_header(s_bcg, "06 | 细分品类（类别2）波士顿矩阵诊断", f"前 10 大细分品类动态象限评估（按当月规模与环比增速自动分层）")
    sub_rows = [[f"{idx}. {r['SubCat']}", r["Track"], int(r["Active_SKU"]), f"{r['Curr_Amt']/1e4:,.2f}", f"{r['Prev_Amt']/1e4:,.2f}", safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp), f"{r['YTD_Amt']/1e4:,.2f}", r["BCG"]] for idx, (_, r) in enumerate(stats["sub_df"].head(10).iterrows(), 1)]
    add_wrapped_table(s_bcg, 0.5, 1.3, 12.3, 5.6, ["细分品类", "所属赛道", "动销", "当月(万)", "上月(万)", "环比", "累计(万)", "BCG定位"], sub_rows, col_ratios=[2.4, 2.8, 0.8, 1.2, 1.2, 1.2, 1.3, 2.0], font_size=9.8)

    # ---------- P8: 核心品牌与客户集中度风控 ----------
    s_br = prs.slides.add_slide(blank)
    add_action_header(s_br, "07 | 核心品牌矩阵与客户集中度风控（CR3/HHI）", titles["brand"])
    bdf = stats["brand_df"].head(10)
    fig5, ax5 = plt.subplots(figsize=(6.0, 4.7))
    b_vals = (bdf["Curr_Amt"] / 1e4).tolist()[::-1]
    bars5 = ax5.barh([str(x)[:10] for x in bdf["Brand"]][::-1], b_vals, color="#2563EB", height=0.6)
    for b in bars5:
        w = b.get_width()
        ax5.text(w + max(max(b_vals or [1]) * 0.02, 2), b.get_y() + b.get_height() / 2, f"{w:.1f}万", va="center", fontsize=9)
    ax5.set_title("当月品牌营收 TOP 10（万元）", fontsize=12, fontweight="bold")
    ax5.spines["top"].set_visible(False)
    ax5.spines["right"].set_visible(False)
    ax5.set_xlim(0, max(b_vals or [1]) * 1.18)
    s_br.shapes.add_picture(fig_stream(fig5), Inches(0.45), Inches(1.25), width=Inches(5.9))

    if conc["available"]:
        conc_t = [[sc["scope"], sc["cust_cnt"], f"{sc['cr3']:.1f}%", f"{sc['cr5']:.1f}%", sc["hhi"], sc["risk"]] for sc in conc["scopes"]]
        add_wrapped_table(s_br, 6.55, 1.3, 6.3, 1.9, ["集中度口径", "客户", "CR3", "CR5", "HHI", "风险评级"], conc_t, col_ratios=[2.2, 0.9, 1.1, 1.1, 1.0, 2.3], font_size=9.0, hdr_rgb=BLUE)
    if not stats["cust_df"].empty:
        c_rows = [[f"[{r['Channel']}] {r['Customer']}", int(r["SKU_Count"]), f"{r['Amt']/1e4:,.2f}", f"{r['Amt']/(stats['curr_tot'] or 1)*100:.1f}%"] for _, r in stats["cust_df"].head(6).iterrows()]
        add_wrapped_table(s_br, 6.55, 3.4, 6.3, 3.6, ["当月头部客户 TOP 6", "SKU", "当月(万)", "占比"], c_rows, col_ratios=[5.0, 0.9, 1.3, 1.1], font_size=9.0)

    # ---------- P9: 明星单品与僵尸/退货风险预警 ----------
    s_end = prs.slides.add_slide(blank)
    add_action_header(s_end, "08 | 明星单品 TOP 10 与 SKU 健康风控", titles["risk"])
    sku_rows = [[f"{idx}. {r['Disp_Name']}", r["Brand"], f"{r['Curr_Amt']/1e4:,.2f}", safe_mom(r["Curr_Amt"], r["Prev_Amt"], hp)] for idx, (_, r) in enumerate(stats["top_skus"].head(10).iterrows(), 1)]
    add_wrapped_table(s_end, 0.45, 1.3, 7.65, 5.7, ["明星单品 TOP 10", "品牌", "当月(万)", "环比"], sku_rows, col_ratios=[5.2, 1.3, 1.2, 1.3], font_size=9.0)

    panel = s_end.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), Inches(1.3), Inches(4.55), Inches(5.7))
    panel.fill.solid()
    panel.fill.fore_color.rgb = BG_LIGHT
    panel.line.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
    ptf = s_end.shapes.add_textbox(Inches(8.5), Inches(1.45), Inches(4.2), Inches(5.4)).text_frame
    ptf.word_wrap = True
    p_h1 = ptf.paragraphs[0]
    p_h1.text = f"⚠️ 负向冲红 / 退货预警（{len(stats['neg_df'])} 款 | 占比 {health['return_ratio']:.2f}%）"
    p_h1.font.size, p_h1.font.bold, p_h1.font.color.rgb = Pt(11.5), True, ROSE
    if not stats["neg_df"].empty:
        for _, r in stats["neg_df"].head(5).iterrows():
            rp = ptf.add_paragraph()
            rp.text, rp.font.size, rp.font.color.rgb, rp.space_before = f"• {r['Disp_Name']}：¥{r['Curr_Amt']:,.0f}", Pt(9.2), SLATE, Pt(4)
    else:
        rp = ptf.add_paragraph()
        rp.text, rp.font.size, rp.font.color.rgb = "• 本月无负向冲红异常", Pt(9.5), MUTED

    p_h2 = ptf.add_paragraph()
    p_h2.text = f"🧹 连续 {health['tail_n']} 月零动销僵尸品（共 {health['zombie_cnt']} 款）"
    p_h2.font.size, p_h2.font.bold, p_h2.font.color.rgb, p_h2.space_before = Pt(11.5), True, NAVY, Pt(14)
    zp = ptf.add_paragraph()
    zp.text = f"建议结合研发与渠道对这 {health['zombie_cnt']} 款成熟期零动销 SKU 发起淘汰评审，完整清单见独立 Excel《04_僵尸品与冲红预警》。"
    zp.font.size, zp.font.color.rgb, zp.space_before = Pt(9.5), SLATE, Pt(4)

    if not stats["pending_df"].empty:
        p_h3 = ptf.add_paragraph()
        p_h3.text = f"📝 属性待补录提醒：{len(stats['pending_df'])} 款 SKU 含【待补录】，请在总表补充品牌/品类以自动沉淀词库。"
        p_h3.font.size, p_h3.font.color.rgb, p_h3.space_before = Pt(9.2), AMBER, Pt(12)

    pptx_path = os.path.join(out_dir, f"{m}_月度经营汇报.pptx")
    prs.save(pptx_path)
    return pptx_path


def generate_independent_reports(data_pack: dict, formats: list, out_dir: str, rules: dict, tracked_codes: list) -> list:
    norm = [f.strip().lower() for f in formats if f.strip()]
    if not norm or "none" in norm:
        return []
    if "all" in norm:
        norm = ["md", "html", "pptx", "xlsx"]

    stats = compute_analytics(data_pack, rules, tracked_codes)
    outputs = []
    if "md" in norm:
        outputs.append(export_markdown_report(stats, out_dir))
    if "html" in norm:
        outputs.append(export_html_report(stats, out_dir))
    if "xlsx" in norm:
        outputs.append(export_standalone_excel(stats, out_dir))
    if "pptx" in norm:
        outputs.append(export_pptx_report(stats, out_dir))
    return outputs


# ==============================================================================
# 第六部分：命令行与 Workbuddy 调度主入口
# ==============================================================================
def locate_files(base_dir: str, input_dir: str):
    base_files = os.listdir(base_dir) if os.path.isdir(base_dir) else []
    in_files = os.listdir(input_dir) if os.path.isdir(input_dir) else []

    master = next((
        os.path.join(base_dir, f) for f in sorted(base_files)
        if any(k in f for k in ["商品信息", "销售", "总表"]) and f.endswith(".xlsx") and not f.startswith("~$")
    ), None)
    xx = next((
        os.path.join(input_dir, f) for f in sorted(in_files)
        if "休闲" in f and f.endswith(".xlsx") and not f.startswith("~$")
    ), None)
    oem = next((
        os.path.join(input_dir, f) for f in sorted(in_files)
        if ("OEM" in f or "oem" in f) and f.endswith(".xlsx") and not f.startswith("~$")
    ), None)
    return master, xx, oem


def main():
    parser = argparse.ArgumentParser(description="月度销售数据自动合并与专业快消 BI 多载体报告系统（v4）")
    parser.add_argument("--base-dir", type=str, default=DEFAULT_BASE_DIR, help="数据根目录（默认 ./data）")
    parser.add_argument("--month", type=int, default=None, help="可选：手动指定月份 YYYYMM（默认自动识别）")
    parser.add_argument("-f", "--format", type=str, default="none", help="独立报告格式：none(默认) / md / html / pptx / xlsx / all")
    parser.add_argument("--report-only", action="store_true", help="纯出报告模式：跳过合并，直接基于总表补出报告")
    parser.add_argument("--preview", action="store_true", help="出报告前先打印《待选追踪清单》并缓存序号（适配 Workbuddy）")
    parser.add_argument("-i", "--interactive", action="store_true", help="本地交互向导模式")
    parser.add_argument("--add-track", type=str, default=None, help="追加持续追踪品：支持预检序号、编码、产品关键词或 new/recent")
    parser.add_argument("--set-track", type=str, default=None, help="覆盖重设持续追踪品清单")
    parser.add_argument("--remove-track", type=str, default=None, help="从追踪清单移出指定项")
    parser.add_argument("--clear-track", action="store_true", help="清空核心品追踪清单")
    parser.add_argument("--rules", type=str, default=None, help="可选：指定外置 rules.json 路径")
    parser.add_argument("--no-archive", action="store_true", help="调试模式：不移动待处理原表")
    args, _ = parser.parse_known_args()

    base_dir = os.path.abspath(args.base_dir)
    in_dir = os.path.join(base_dir, "待处理发货表")
    arch_dir = os.path.join(base_dir, "01_已归档原表")
    bak_dir = os.path.join(base_dir, "02_历史总表备份")
    rep_dir = os.path.join(base_dir, "03_月度分析报告")
    for d in [in_dir, arch_dir, bak_dir, rep_dir]:
        os.makedirs(d, exist_ok=True)

    rules = load_rules(base_dir, args.rules)
    master_file, xx_file, oem_file = locate_files(base_dir, in_dir)
    chosen_formats = [x.strip().lower() for x in args.format.split(",") if x.strip()]
    want_report = bool(chosen_formats) and ("none" not in chosen_formats)
    has_track_cli = any([args.add_track, args.set_track, args.remove_track, args.clear_track])

    if args.report_only or (args.preview and not (xx_file and oem_file)):
        if not master_file:
            print(f"❌ 未在 {base_dir} 找到基准总表。")
            return
        data_pack = load_data_for_report_only(master_file, [arch_dir, in_dir], rules, cli_month=args.month)
        if has_track_cli:
            tracked_codes = apply_cli_track_updates(args, data_pack, base_dir, rules)
        elif args.preview:
            print_tracking_preview(build_tracking_preview(data_pack, base_dir, rules))
            return
        elif args.interactive or (want_report and sys.stdin.isatty()):
            tracked_codes = run_interactive_watchlist_wizard(data_pack, base_dir, rules)
        else:
            tracked_codes = load_watchlist(base_dir)

        if not want_report:
            if not has_track_cli:
                print("⚠️ 请通过 -f 指定报告格式（如 -f html,pptx 或 -f all），或使用 --preview 查看待选追踪清单。")
            return

        outputs = generate_independent_reports(data_pack, chosen_formats, rep_dir, rules, tracked_codes)
        print(f">>> 已按需生成 {data_pack['month_tag']} 专业独立报告（追踪核心品 {len(tracked_codes)} 款）：")
        for p in outputs:
            print(f"    - {p}")
        return

    if not (master_file and xx_file and oem_file):
        print("❌ 未凑齐所需表格，请检查目录：")
        print(f"  - 基准总表 ({base_dir}): {master_file or '未找到'}")
        print(f"  - 休闲发货表 ({in_dir}): {xx_file or '未找到'}")
        print(f"  - OEM发货表 ({in_dir}): {oem_file or '未找到'}")
        return

    print(f">>> 1. 已定位文件：\n    总表: {os.path.basename(master_file)}\n    休闲: {os.path.basename(xx_file)}\n    OEM : {os.path.basename(oem_file)}")
    data_pack = merge_monthly_data(master_file, xx_file, oem_file, rules, cli_month=args.month)
    m_tag = data_pack["month_tag"]
    mode_str = "原位覆盖更新" if data_pack["is_overwrite"] else "末尾新增月份列"
    print(f">>> 2. 智能识别月份：【{m_tag}】（写入模式：{mode_str}，新增SKU：{len(data_pack['df_new_skus'])} 款）")

    data_pack["wb"].save(master_file)
    backup_path = os.path.join(bak_dir, f"销售业绩总表_{m_tag}快照备份.xlsx")
    data_pack["wb"].save(backup_path)
    print(f">>> 3. 纯净总表已原地更新并备份至：{backup_path}")

    if not args.no_archive:
        shutil.move(xx_file, os.path.join(arch_dir, f"{m_tag}_{os.path.basename(xx_file)}"))
        shutil.move(oem_file, os.path.join(arch_dir, f"{m_tag}_{os.path.basename(oem_file)}"))
        print(f">>> 4. 原始发货表已自动归档至：{arch_dir}")

    if args.preview:
        print_tracking_preview(build_tracking_preview(data_pack, base_dir, rules))
        return

    if has_track_cli:
        tracked_codes = apply_cli_track_updates(args, data_pack, base_dir, rules)
    elif want_report and (args.interactive or sys.stdin.isatty()):
        tracked_codes = run_interactive_watchlist_wizard(data_pack, base_dir, rules)
    else:
        tracked_codes = load_watchlist(base_dir)

    if want_report:
        outputs = generate_independent_reports(data_pack, chosen_formats, rep_dir, rules, tracked_codes)
        print(f">>> 5. 已按需生成专业独立月度分析报告（追踪核心品 {len(tracked_codes)} 款）：")
        for out_p in outputs:
            print(f"    - {out_p}")
    else:
        print(">>> 5. 默认未生成分析报告（如需出报告前先选追踪品，请运行 --preview 或追加 -f all -i）。")

    print("🎉 全部流程执行完毕！")


if __name__ == "__main__":
    main()
