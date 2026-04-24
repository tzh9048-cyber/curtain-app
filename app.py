from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st


# =========================
# 轻量级客服辅助查询软件
# 技术栈：Streamlit + pandas + openpyxl
# =========================

def _normalize_text(x) -> str:
    """
    将单元格内容安全转成字符串，并做基础清洗。
    - None/NaN -> ""
    - 去掉首尾空格
    """
    if x is None:
        return ""
    try:
        # pandas 的缺失值通常是 NaN（float），直接转字符串会变成 'nan'
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def _ensure_required_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    """
    尝试确保 DataFrame 中包含核心列，并返回缺失列清单。
    期望列名（与需求一致）：
    - 产品名称
    - 价格（元/米）
    - 克重（g/m²）
    - 门幅
    - 颜色
    """
    required = ["产品名称", "价格（元/米）", "克重（g/m²）", "门幅", "颜色"]
    missing = [c for c in required if c not in df.columns]

    # 如果 Excel 列名有轻微差异，做一些“尽力匹配”（不改变你的原始文件）
    # 例如：价格(元/米)、克重(g/m2) 等
    rename_map = {}
    if "价格(元/米)" in df.columns and "价格（元/米）" not in df.columns:
        rename_map["价格(元/米)"] = "价格（元/米）"
    if "克重(g/m²)" in df.columns and "克重（g/m²）" not in df.columns:
        rename_map["克重(g/m²)"] = "克重（g/m²）"
    if "克重(g/m2)" in df.columns and "克重（g/m²）" not in df.columns:
        rename_map["克重(g/m2)"] = "克重（g/m²）"
    if "门幅(米)" in df.columns and "门幅" not in df.columns:
        rename_map["门幅(米)"] = "门幅"

    if rename_map:
        df = df.rename(columns=rename_map)
        missing = [c for c in required if c not in df.columns]

    return df, missing


def _safe_series_get(row: pd.Series, col: str) -> str:
    return _normalize_text(row.get(col))


def _prepare_sheet_df(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    """
    针对单个工作表做“尽力而为”的标准化：
    - 处理列名轻微差异（复用 _ensure_required_columns 的 rename 逻辑）
    - 返回缺失的核心列清单（但不会中断运行）
    """
    df, missing_required = _ensure_required_columns(df_raw)

    # 将所有列尽量转成可检索的字符串，避免 .str.contains 报错
    df = df.copy()
    for col in df.columns:
        try:
            df[col] = df[col].map(_normalize_text)
        except Exception:
            # 极端情况下（比如列里是复杂对象），保底转 str
            df[col] = df[col].astype(str).map(_normalize_text)

    return df, missing_required


def main() -> None:
    st.set_page_config(page_title="客服辅助查询软件", page_icon="🧾", layout="wide")

    st.title("客服辅助查询软件")

    with st.sidebar:
        uploaded_file = st.sidebar.file_uploader(
            "📂 请上传最新的产品资料 Excel",
            type=["xlsx", "xls"],
        )
        st.divider()

        st.subheader("🧮 快速报价计算器")
        # 侧边栏只保留快速报价计算器模块

        w_width = st.number_input(
            "窗户宽度（米）",
            min_value=0.0,
            value=3.0,
            step=0.1,
            key="calc_window_width_m",
        )
        w_height = st.number_input(
            "窗户高度（米）",
            min_value=0.0,
            value=2.5,
            step=0.1,
            key="calc_window_height_m",
        )

        pricing_type = st.selectbox(
            "计价类型",
            options=[
                "布艺窗帘 (按宽度×倍数)",
                "特殊帘 (按面积 平米)",
            ],
            index=0,
            key="calc_pricing_type",
        )

        if pricing_type == "布艺窗帘 (按宽度×倍数)":
            fold = st.number_input(
                "褶皱倍数",
                min_value=0.0,
                value=2.0,
                step=0.1,
                key="calc_fold_multiplier",
            )
            fabric_unit = st.number_input(
                "布料单价（元/米）",
                min_value=0.0,
                value=0.0,
                step=1.0,
                key="calc_fabric_unit_price",
            )
            rail_unit = st.number_input(
                "轨道/罗马杆单价（元/米）",
                min_value=0.0,
                value=35.0,
                step=1.0,
                key="calc_rail_unit_price",
            )

            if st.button("计算预估总价", use_container_width=True, key="calc_btn_fabric"):
                fabric_cost = w_width * fold * fabric_unit
                rail_cost = w_width * rail_unit
                total = fabric_cost + rail_cost
                st.success(f"**预估总价：{total:.2f} 元**")
                st.info(
                    "明细：\n"
                    f"- 布料费用 = 窗宽({w_width:.2f}) × 褶皱倍数({fold:.2f}) × 布料单价({fabric_unit:.2f}) = {fabric_cost:.2f} 元\n"
                    f"- 轨道费用 = 窗宽({w_width:.2f}) × 轨道单价({rail_unit:.2f}) = {rail_cost:.2f} 元\n"
                    f"- 总价 = {fabric_cost:.2f} + {rail_cost:.2f} = {total:.2f} 元"
                )
        else:
            area_unit = st.number_input(
                "产品单价（元/平米）",
                min_value=0.0,
                value=0.0,
                step=1.0,
                key="calc_area_unit_price",
            )

            if st.button("计算预估总价", use_container_width=True, key="calc_btn_area"):
                area = w_width * w_height
                total = area * area_unit
                st.success(f"**预估总价：{total:.2f} 元**")
                st.info(
                    "明细：\n"
                    f"- 面积 = 窗宽({w_width:.2f}) × 窗高({w_height:.2f}) = {area:.2f} 平米\n"
                    f"- 总价 = 面积({area:.2f}) × 单价({area_unit:.2f}) = {total:.2f} 元"
                )

        st.divider()

    # 读取数据：改为手动上传 Excel（仅修改入口，不影响后续展示逻辑）
    if uploaded_file is None:
        st.info("👈 请先在左侧菜单上传产品资料 Excel 文件，即可开始查询和报价。")
        return

    try:
        # 读取所有工作表，保持“产品大类（Sheet）”选择与后续逻辑不变
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception:
        st.error("读取上传文件失败，请检查文件是否为有效的 Excel。")
        return

    sheet_names = [name for name in sheets.keys() if name is not None]
    if not sheet_names:
        st.error("该 Excel 未读取到任何工作表（Sheet）。请确认文件内容是否正常。")
        return

    selected_sheet = st.selectbox("产品大类", options=sheet_names, index=0)

    df_raw = sheets.get(selected_sheet)
    if df_raw is None:
        st.error("所选工作表为空或读取失败，请切换其他工作表重试。")
        return

    df, _missing = _prepare_sheet_df(df_raw)

    query = st.text_input("搜索", placeholder="输入关键字")

    # 过滤
    if query.strip():
        q = query.strip().lower()
        # 固定在常用字段里检索，保持界面简洁
        preferred = [c for c in ["产品名称", "颜色", "规格", "成分", "成份"] if c in df.columns]
        selected_cols = preferred if preferred else list(df.columns)
        mask = False
        for col in selected_cols:
            if col not in df.columns:
                continue
            mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)
        result_full = df[mask].copy()
    else:
        result_full = df.copy()

    # 匹配结果数量（使用完整结果，保证能读取“淘宝链接/图片链接”等列）
    st.metric("匹配结果", value=len(result_full))

    st.divider()

    # 结果展示：逐条卡片（更适合客服一键复制）
    if len(result_full) == 0:
        st.warning("没有匹配到结果，请换个关键字试试。")
        return

    # 限制一次渲染过多卡片导致卡顿（客服场景通常不需要一下看几千条）
    max_render = 200
    if len(result_full) > max_render:
        result_to_render = result_full.head(max_render)
    else:
        result_to_render = result_full

    # 这里的 result_to_render 保留完整列，便于安全判断 “淘宝链接/图片链接”
    for idx, row in result_to_render.iterrows():
        with st.container(border=True):
            col1, col2 = st.columns([2, 1])

            with col1:
                name = _safe_series_get(row, "产品名称")
                color = _safe_series_get(row, "颜色")
                if not name:
                    first_col = result_to_render.columns[0] if len(result_to_render.columns) > 0 else ""
                    name = _safe_series_get(row, first_col) if first_col else "（未命名产品）"

                st.markdown(f"**{name}**" + (f"  ·  **{color}**" if color else ""))

                # 基础参数（门幅、克重、价格等）—— 空值直接跳过
                meta_parts = []
                for label, col in [
                    ("门幅", "门幅"),
                    ("克重", "克重（g/m²）"),
                    ("价格", "价格（元/米）"),
                    ("规格", "规格"),
                    ("成分", "成分"),
                    ("成份", "成份"),
                ]:
                    if col in result_to_render.columns:
                        v = _safe_series_get(row, col)
                        if v:
                            if col == "价格（元/米）":
                                meta_parts.append(f"{label}：{v} 元/米")
                            else:
                                meta_parts.append(f"{label}：{v}")

                if meta_parts:
                    st.write("｜".join(meta_parts))

                # 淘宝链接：安全判断列名 + 当前行有值才展示
                if "淘宝链接" in result_to_render.columns:
                    taobao_link = _safe_series_get(row, "淘宝链接")
                    if taobao_link:
                        st.markdown(
                            f'<a href="{taobao_link}" target="_blank" '
                            f'style="color:#1E88E5; text-decoration: underline;">'
                            f'🔗 淘宝产品详情（点击跳转）'
                            f"</a>",
                            unsafe_allow_html=True,
                        )
                        st.code(taobao_link, language="text")

            with col2:
                # 图片链接：安全判断列名 + 当前行有值才展示
                if "图片链接" in result_to_render.columns:
                    image_url = _safe_series_get(row, "图片链接")
                    if image_url:
                        try:
                            st.image(image_url, use_container_width=True)
                        except Exception:
                            pass

if __name__ == "__main__":
    main()
