from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st


# =========================
# 轻量级客服辅助查询软件
# 技术栈：Streamlit + pandas + openpyxl
# =========================

# 产品资料 Excel：部署到云端时使用仓库内相对路径
DEFAULT_EXCEL_FILENAME = "古思特产品资料-2026.4.3.xlsx"

# 内部访问密码（可按需修改）
INTERNAL_ACCESS_PASSWORD = "123456"


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


@st.cache_data(show_spinner=False)
def load_products_from_path(excel_path: str) -> dict:
    """
    从本地路径读取 Excel。
    - 使用 cache_data 避免每次输入搜索都重新读取文件
    """
    # sheet_name=None -> 读取所有工作表，返回 {sheet_name: DataFrame}
    return pd.read_excel(excel_path, sheet_name=None)


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

    # -------------------------
    # 简单登录拦截：输入内部访问密码
    # -------------------------
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "auth_error" not in st.session_state:
        st.session_state.auth_error = None

    if not st.session_state.authenticated:
        st.title("内部访问验证")
        st.caption("请输入内部访问密码后继续使用。")

        with st.form("login_form", clear_on_submit=False):
            pwd = st.text_input("请输入内部访问密码", type="password")
            submitted = st.form_submit_button("进入")

        if submitted:
            if pwd == INTERNAL_ACCESS_PASSWORD:
                st.session_state.authenticated = True
                st.session_state.auth_error = None
                st.rerun()
            else:
                st.session_state.auth_error = "密码错误，请输入正确的内部密码"

        if st.session_state.auth_error:
            st.error(st.session_state.auth_error)

        return

    st.title("客服辅助查询软件")

    with st.sidebar:
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

    # 读取数据
    excel_path = Path(DEFAULT_EXCEL_FILENAME)
    if not excel_path.exists():
        st.error("未找到产品资料文件，请联系管理员")
        return
    try:
        sheets = load_products_from_path(str(excel_path))
    except Exception:
        st.error("未找到产品资料文件，请联系管理员")
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

    # 优先展示需要的列
    preferred_show_cols = ["产品名称", "颜色", "门幅", "克重（g/m²）", "价格（元/米）"]
    show_cols = [c for c in preferred_show_cols if c in result_full.columns]
    if show_cols:
        result_show = result_full[show_cols].copy()
    else:
        result_show = result_full.copy()

    st.metric("匹配结果", value=len(result_show))

    st.divider()

    # 结果展示：逐条卡片（更适合客服一键复制）
    if len(result_show) == 0:
        st.warning("没有匹配到结果，请换个关键字试试。")
        return

    # 限制一次渲染过多卡片导致卡顿（客服场景通常不需要一下看几千条）
    max_render = 200
    if len(result_show) > max_render:
        result_to_render = result_show.head(max_render)
    else:
        result_to_render = result_show

    for idx, row in result_to_render.iterrows():
        with st.container(border=True):
            # 用完整行数据读取“淘宝链接”等更多字段
            full_row_obj = result_full.loc[idx] if idx in result_full.index else row
            if isinstance(full_row_obj, pd.DataFrame):
                full_row = full_row_obj.iloc[0]
            else:
                full_row = full_row_obj

            left_col, right_col = st.columns([2, 1])

            with left_col:
                name = _safe_series_get(row, "产品名称")
                color = _safe_series_get(row, "颜色")
                if not name:
                    # 兜底：用第一列做标题（避免不同表头时空白）
                    first_col = result_to_render.columns[0] if len(result_to_render.columns) > 0 else ""
                    name = _safe_series_get(row, first_col) if first_col else "（未命名产品）"

                st.markdown(
                    f"**{name}**" + (f"  ·  **{color}**" if color else "")
                )

                # 核心参数（空值跳过）
                meta_parts = []
                v = _safe_series_get(full_row, "门幅") if "门幅" in df.columns else ""
                if v:
                    meta_parts.append(f"门幅：{v}")
                v = _safe_series_get(full_row, "克重（g/m²）") if "克重（g/m²）" in df.columns else ""
                if v:
                    meta_parts.append(f"克重：{v}")
                v = _safe_series_get(full_row, "价格（元/米）") if "价格（元/米）" in df.columns else ""
                if v:
                    meta_parts.append(f"价格：{v} 元/米")
                if "规格" in df.columns:
                    v = _safe_series_get(full_row, "规格")
                    if v:
                        meta_parts.append(f"规格：{v}")
                comp_col = "成分" if "成分" in df.columns else ("成份" if "成份" in df.columns else None)
                if comp_col:
                    v = _safe_series_get(full_row, comp_col)
                    if v:
                        meta_parts.append(f"成分：{v}")
                if meta_parts:
                    st.write("｜".join(meta_parts))

                taobao_link = _safe_series_get(full_row, "淘宝链接") if "淘宝链接" in df.columns else ""
                if taobao_link:
                    st.link_button(
                        "🛒 点击前往淘宝查看（主图/视频/买家秀）",
                        taobao_link,
                        use_container_width=True,
                    )

            with right_col:
                image_url = _safe_series_get(full_row, "图片链接") if "图片链接" in df.columns else ""
                if image_url:
                    try:
                        st.image(image_url, use_container_width=True)
                    except Exception:
                        # 链接失效或加载失败时不影响主流程
                        pass

if __name__ == "__main__":
    main()
