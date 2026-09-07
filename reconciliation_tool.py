import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import tempfile
import os


# ========== 页面配置 ==========
st.set_page_config(page_title="智能对账工具", layout="wide")
st.title("🤖 智能对账工具")
st.markdown(
    "上传**直联支付单**（银行流水）与**付款表**（财务账），"
    "自动按金额+收款人匹配，逐条标注匹配结果"
)


# ==================================================================
#  工具函数
# ==================================================================
def parse_amount(val):
    """
    万能金额解析：支持常规数字、带千分位(1,000.00)、文本、空值等。
    返回 float，解析失败返回 NaN。
    """
    if val is None:
        return np.nan
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return np.nan
    # 去掉千分位逗号、全角字符、空格
    s = s.replace(",", "").replace("，", "").replace(" ", "").replace("　", "")
    try:
        return float(s)
    except ValueError:
        return np.nan


def parse_date(val):
    """
    万能日期解析：支持 2026-09-07 / 2026/09/07 / 2026-09-07 18:41:10 / Excel序列号(如46266) 等。
    返回 'YYYY-MM-DD' 字符串，失败返回 None。
    """
    if val is None:
        return None
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return None

    # 尝试常见格式（只取日期部分）
    date_part = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(date_part, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    # Excel 日期序列号 (>=40000 且 <=60000)
    try:
        num = float(date_part)
        if 40000 <= num <= 60000:
            base = datetime(1899, 12, 30)
            return (base + pd.Timedelta(days=num)).strftime("%Y-%m-%d")
    except ValueError:
        pass

    return None


def find_header_row(df_raw, keywords=("供应商名称",)):
    """自动寻找表头行（含有关键字的行），返回行号"""
    for i, row in df_raw.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v))
        if all(kw in row_str for kw in keywords):
            return i
    # fallback：只匹配一个关键词
    for i, row in df_raw.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v))
        if "供应商名称" in row_str or "收款人名称" in row_str:
            return i
    return None


def name_similarity(n1, n2):
    """名称相似度判断：完全匹配 / 包含 / 去括号后匹配"""
    if pd.isna(n1) or pd.isna(n2):
        return False
    s1 = str(n1).strip().replace(" ", "").replace("（", "(").replace("）", ")")
    s2 = str(n2).strip().replace(" ", "").replace("（", "(").replace("）", ")")
    if s1 == s2:
        return True
    if s1 in s2 or s2 in s1:
        return True
    # 去掉括号内容后匹配（处理"原：xxx"等标注）
    s1_clean = re.sub(r"[\(（].*?[\)）]", "", s1).strip()
    s2_clean = re.sub(r"[\(（].*?[\)）]", "", s2).strip()
    if s1_clean and s2_clean and (s1_clean in s2_clean or s2_clean in s1_clean):
        return True
    return False


# ==================================================================
#  解析：直联支付单（银行流水）
# ==================================================================
def parse_bank_table(raw_df):
    """从原始 DataFrame 中解析直联支付单，返回标准化 DataFrame"""
    header_row = find_header_row(raw_df, keywords=("收款人名称",))
    if header_row is None:
        return pd.DataFrame()

    headers = raw_df.iloc[header_row].tolist()
    data = raw_df.iloc[header_row + 1:].copy()
    data.columns = [str(h).strip() for h in headers]
    data = data.dropna(how="all")

    records = []
    for _, row in data.iterrows():
        payee = str(row.get("收款人名称", "")).strip()
        amount = parse_amount(row.get("付款金额", ""))
        if payee == "" or payee == "nan" or pd.isna(amount):
            continue
        records.append({
            "收款人名称": payee,
            "付款金额": round(amount, 2),
            "单据编号": str(row.get("单据编号", "")).strip(),
            "单据日期": parse_date(row.get("单据日期", "")),
            "摘要": str(row.get("摘要", "")).strip(),
        })
    return pd.DataFrame(records)


# ==================================================================
#  解析：付款表（财务账）—— 自动识别日期对应 Sheet
# ==================================================================
def parse_payment_table(uploaded_file, target_date_str=None):
    """
    读取付款表 Excel，自动选择与目标日期匹配的 Sheet（如 9.7 和 9.7电子）。
    如果没有指定日期，则使用所有非汇总 Sheet。
    返回标准化 DataFrame。
    """
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    try:
        tmp.write(uploaded_file.read())
    finally:
        tmp.close()

    try:
        xls = pd.ExcelFile(tmp.name)
        sheets = xls.sheet_names

        # ---- 确定要处理的 Sheet ----
        if target_date_str:
            month = str(int(target_date_str[5:7]))
            day = str(int(target_date_str[8:10]))
            patterns = [f"{month}.{day}", f"{month}.{day}电子",
                        f"{month}-{day}", f"{month}-{day}电子"]

            selected = []
            for sn in sheets:
                sn_lower = sn.strip().lower()
                for p in patterns:
                    if sn_lower == p or sn_lower.startswith(p + " ") or sn_lower.startswith(p + "\t"):
                        selected.append(sn)
                        break
            # 如果精确匹配没找到，尝试包含匹配
            if not selected:
                for sn in sheets:
                    sn_lower = sn.strip().lower()
                    if f"{month}.{day}" in sn_lower:
                        selected.append(sn)

            if not selected:
                st.warning(
                    f"⚠️ 付款表中未找到日期 {target_date_str} 对应的 Sheet"
                    f"（期望如 '9.7' / '9.7电子'），将使用全部非汇总 Sheet"
                )
                selected = [s for s in sheets if "汇总" not in s]
        else:
            selected = [s for s in sheets if "汇总" not in s]

        uploaded_file.seek(0)  # 重置文件指针

        all_dfs = []
        for sn in selected:
            df_raw = pd.read_excel(tmp.name, dtype=str, sheet_name=sn, header=None)
            header_row = find_header_row(df_raw, keywords=("供应商名称",))
            if header_row is None:
                continue

            headers = [str(h).strip() for h in df_raw.iloc[header_row].tolist()]
            df = df_raw.iloc[header_row + 1:].copy()
            df.columns = headers
            df = df.dropna(how="all")

            # 统一列名
            rename_map = {}
            for col in df.columns:
                col_str = str(col).strip()
                if "供应商名称" in col_str:
                    rename_map[col] = "供应商名称"
                elif "含税金额" in col_str or "申请付款金额" in col_str:
                    rename_map[col] = "金额"
                elif "付款日期" in col_str:
                    rename_map[col] = "付款日期"
                elif "到财务时间" in col_str:
                    rename_map[col] = "到财务时间"
            df = df.rename(columns=rename_map)

            # 提取有效行
            for _, row in df.iterrows():
                supplier = str(row.get("供应商名称", "")).strip()
                amount = parse_amount(row.get("金额", ""))
                if supplier == "" or supplier == "nan" or pd.isna(amount):
                    continue
                all_dfs.append({
                    "供应商名称": supplier,
                    "金额": round(amount, 2),
                    "付款日期": parse_date(row.get("付款日期", "")) or parse_date(row.get("到财务时间", "")),
                    "_source_sheet": sn,
                })
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

    return pd.DataFrame(all_dfs)


# ==================================================================
#  主流程
# ==================================================================
col1, col2 = st.columns(2)
with col1:
    uploaded_bank = st.file_uploader("📂 上传直联支付单（银行流水）", type=["xlsx", "xls"], key="bank")
with col2:
    uploaded_account = st.file_uploader("📂 上传付款表（财务账）", type=["xlsx", "xls"], key="account")

if uploaded_bank and uploaded_account:
    # ---------- 解析银行流水 ----------
    bank_raw = pd.read_excel(uploaded_bank, dtype=str, header=None)
    bank_df = parse_bank_table(bank_raw)

    if bank_df.empty:
        st.error("❌ 直联支付单解析失败，请确认文件包含 '收款人名称' 和 '付款金额' 列")
        st.stop()

    # ---------- 自动识别日期 ----------
    all_dates = bank_df["单据日期"].dropna().unique()
    if len(all_dates) > 0:
        target_date = all_dates[0]
        st.success(f"📅 自动识别对账日期：**{target_date}**")
    else:
        target_date = None
        st.info("ℹ️ 未能自动识别日期，将对付款表全部 Sheet 进行匹配")

    # ---------- 解析付款表（根据日期自动选 Sheet） ----------
    account_df = parse_payment_table(uploaded_account, target_date)

    if account_df.empty:
        st.error("❌ 付款表解析失败，未找到有效数据（需含 '供应商名称' 和 '金额' 列）")
        st.stop()

    st.info(
        f"📋 付款表共加载 **{len(account_df)}** 条有效记录，"
        f"来源 Sheet: {', '.join(account_df['_source_sheet'].unique())}"
    )

    # ---------- 核心匹配 ----------
    bank_matched = []
    account_used = set()

    for _, b_row in bank_df.iterrows():
        b_amount = b_row["付款金额"]
        b_payee = b_row["收款人名称"]

        # 1) 先按金额精确匹配，且未使用过的
        candidates = account_df[
            (account_df["金额"] == b_amount) &
            (~account_df.index.isin(account_used))
        ]

        # 2) 再按名称相似度筛选
        matched_idx = None
        for a_idx, a_row in candidates.iterrows():
            if name_similarity(b_payee, a_row["供应商名称"]):
                matched_idx = a_idx
                break

        if matched_idx is not None:
            account_used.add(matched_idx)
            bank_matched.append({
                "收款人名称": b_row["收款人名称"],
                "付款金额": b_row["付款金额"],
                "单据编号": b_row["单据编号"],
                "摘要": b_row["摘要"],
                "匹配状态": "✅ 匹配成功",
                "匹配供应商": account_df.loc[matched_idx, "供应商名称"],
                "匹配金额": account_df.loc[matched_idx, "金额"],
            })
        else:
            bank_matched.append({
                "收款人名称": b_row["收款人名称"],
                "付款金额": b_row["付款金额"],
                "单据编号": b_row["单据编号"],
                "摘要": b_row["摘要"],
                "匹配状态": "⚠️ 未匹配",
                "匹配供应商": "",
                "匹配金额": "",
            })

    result_df = pd.DataFrame(bank_matched)

    # 财务账中未匹配的记录
    unmatched_account = account_df[~account_df.index.isin(account_used)]

    # ---------- 展示指标 ----------
    total_bank = len(bank_df)
    matched_count = (result_df["匹配状态"] == "✅ 匹配成功").sum()
    match_rate = matched_count / total_bank if total_bank > 0 else 0

    st.divider()
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        st.metric("📊 匹配成功率", f"{match_rate:.1%}")
    with col_m2:
        st.metric("✅ 匹配成功", f"{matched_count} 条")
    with col_m3:
        st.metric("⚠️ 银行侧未匹配", f"{total_bank - matched_count} 条")
    with col_m4:
        st.metric("📋 财务账未匹配", f"{len(unmatched_account)} 条")

    # ---------- Tab 展示结果 ----------
    tab1, tab2, tab3 = st.tabs(["✅ 匹配成功", "⚠️ 银行侧未匹配", "📋 财务账未匹配"])

    with tab1:
        matched_df = result_df[result_df["匹配状态"] == "✅ 匹配成功"]
        st.write(f"共 **{len(matched_df)}** 条匹配记录")
        display_cols = ["收款人名称", "付款金额", "单据编号", "摘要", "匹配供应商", "匹配金额"]
        st.dataframe(matched_df[display_cols], use_container_width=True)

    with tab2:
        unmatched_bank_df = result_df[result_df["匹配状态"] == "⚠️ 未匹配"]
        st.write(f"共 **{len(unmatched_bank_df)}** 条")
        if len(unmatched_bank_df) > 0:
            display_cols = ["收款人名称", "付款金额", "单据编号", "摘要"]
            st.dataframe(unmatched_bank_df[display_cols], use_container_width=True)
        else:
            st.success("🎉 全部匹配成功，无异常")

    with tab3:
        st.write(f"共 **{len(unmatched_account)}** 条（付款表中有，但直联支付单中未找到对应记录）")
        if len(unmatched_account) > 0:
            display_cols = ["供应商名称", "金额", "付款日期", "_source_sheet"]
            st.dataframe(unmatched_account[display_cols], use_container_width=True)
        else:
            st.success("🎉 财务账全部匹配完毕")

    # ---------- 导出 ----------
    st.divider()
    st.subheader("📥 导出对账结果")
    export_df = result_df[
        ["收款人名称", "付款金额", "单据编号", "摘要", "匹配状态", "匹配供应商", "匹配金额"]
    ].copy()
    unmatched_acc_export = unmatched_account[["供应商名称", "金额", "付款日期", "_source_sheet"]].rename(
        columns={"供应商名称": "收款人名称", "金额": "付款金额", "_source_sheet": "来源Sheet"}
    )
    full_export = pd.concat([export_df, unmatched_acc_export], ignore_index=True)

    csv = full_export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="📥 下载完整对账结果 CSV",
        data=csv,
        file_name=f"对账结果_{target_date or 'all'}.csv",
        mime="text/csv",
    )

else:
    st.info("👆 请上传两个文件开始对账")
    with st.expander("📖 使用说明"):
        st.markdown("""
        ### 操作流程
        1. **左侧**上传「直联支付单列表.xlsx」（银行流水，含 收款人名称 + 付款金额）
        2. **右侧**上传「付款表.xlsx」（财务账，含 供应商名称 + 含税金额/申请付款金额）
        3. 系统**自动识别日期**，匹配对应 Sheet（如 `9.7` 和 `9.7电子`）
        4. 按 **金额 + 名称** 双向匹配，逐条标注 ✅ / ⚠️

        ### 智能兼容
        - **金额**：兼容常规格式 `1000`、`1,000.00` 千分位、`1 000` 空格等
        - **日期**：兼容 `2026-09-07`、`2026/09/07`、`2026-09-07 18:41:10`、Excel序列号 `46266` 等
        - **名称**：兼容括号差异（全角/半角）、简称、"原名"标注等

        ### 匹配逻辑
        1. 先按 **金额精确匹配**
        2. 再按 **名称智能匹配**（完全相同 / 包含关系 / 去括号后匹配）
        3. 每条记录只用一次（避免重复匹配）
        """)
