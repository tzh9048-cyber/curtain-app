import json
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# =========================
# 轻量级客服辅助查询软件
# 技术栈：Streamlit + pandas + openpyxl
# =========================

# 你提供的产品 Excel 默认路径（可在侧边栏改）
DEFAULT_EXCEL_PATH = r"C:\Users\Administrator\Desktop\古思特产品资料-2026.4.3.xlsx"

# 话术模板（可在界面里随时修改）
DEFAULT_TEMPLATE = (
    "您好，这款【产品名称】【颜色】的门幅是【门幅】，克重【克重】，"
    "目前的价格是【价格】元/米。请问您需要定做多宽的尺寸呢？"
)


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


@st.cache_data(show_spinner=False)
def load_products_from_upload(file_bytes: bytes, filename: str) -> dict:
    """
    从上传文件读取 Excel。
    - filename 参与缓存 key，避免不同文件内容但同 bytes 结构导致缓存混淆
    """
    # pandas 可以直接读取字节流；这里用 BytesIO 包装
    from io import BytesIO

    return pd.read_excel(BytesIO(file_bytes), sheet_name=None)


def build_script(template: str, row: pd.Series) -> str:
    """
    将一条产品记录按模板拼接成标准话术。
    模板占位符采用中文中括号形式：【产品名称】等。
    """
    mapping = {
        "产品名称": _normalize_text(row.get("产品名称")),
        "颜色": _normalize_text(row.get("颜色")),
        "门幅": _normalize_text(row.get("门幅")),
        "克重": _normalize_text(row.get("克重（g/m²）")),
        "价格": _normalize_text(row.get("价格（元/米）")),
    }

    script = template
    # 逐个替换，保证客服可直观看懂模板
    script = script.replace("【产品名称】", mapping["产品名称"])
    script = script.replace("【颜色】", mapping["颜色"])
    script = script.replace("【门幅】", mapping["门幅"])
    script = script.replace("【克重】", mapping["克重"])
    script = script.replace("【价格】", mapping["价格"])
    return script


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


def copy_to_clipboard_js(text: str) -> None:
    """
    通过浏览器 JS 写入剪贴板。
    Streamlit 运行在浏览器里，因此最稳妥的是用 navigator.clipboard。
    """
    payload = json.dumps(text, ensure_ascii=False)
    components.html(
        f"""
        <script>
          (async () => {{
            try {{
              await navigator.clipboard.writeText({payload});
            }} catch (e) {{
              console.error(e);
            }}
          }})();
        </script>
        """,
        height=0,
    )


def main() -> None:
    st.set_page_config(page_title="客服辅助查询软件", page_icon="🧾", layout="wide")

    st.title("客服辅助查询软件")
    st.caption("支持 Excel 自动读取/上传、模糊搜索、结果展示、一键复制标准话术。")

    with st.sidebar:
        st.subheader("数据来源")
        mode = st.radio(
            "请选择加载方式",
            options=["读取本地路径（推荐）", "上传 Excel"],
            index=0,
        )

        excel_path = st.text_input("本地 Excel 路径", value=DEFAULT_EXCEL_PATH)
        uploaded = st.file_uploader(
            "上传 Excel（.xlsx/.xls）",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
        )

        st.divider()
        st.subheader("话术模板")
        template = st.text_area("可按需修改模板", value=DEFAULT_TEMPLATE, height=120)
        st.caption("可用占位符：【产品名称】【颜色】【门幅】【克重】【价格】")

    # 读取数据
    sheets: Optional[dict] = None
    load_error: Optional[str] = None

    if mode == "上传 Excel":
        if uploaded is None:
            st.info("请在左侧上传你的 `products.xlsx`（或你的产品资料 Excel）。")
        else:
            try:
                sheets = load_products_from_upload(uploaded.getvalue(), uploaded.name)
            except Exception as e:
                load_error = f"读取上传文件失败：{e}"
    else:
        # 本地路径读取
        try:
            if not excel_path.strip():
                st.info("请在左侧填写 Excel 路径，或切换到“上传 Excel”。")
            else:
                p = Path(excel_path)
                if not p.exists():
                    st.warning("未找到该路径的 Excel 文件。请检查路径是否正确，或使用上传方式。")
                else:
                    sheets = load_products_from_path(excel_path)
        except Exception as e:
            load_error = f"读取本地文件失败：{e}"

    if load_error:
        st.error(load_error)
        return

    if sheets is None:
        return

    sheet_names = [name for name in sheets.keys() if name is not None]
    if not sheet_names:
        st.error("该 Excel 未读取到任何工作表（Sheet）。请确认文件内容是否正常。")
        return

    # 侧边栏：产品大类（工作表）导航
    with st.sidebar:
        st.divider()
        st.subheader("产品大类")
        selected_sheet = st.selectbox(
            "请选择产品大类",
            options=sheet_names,
            index=0,
        )

    df_raw = sheets.get(selected_sheet)
    if df_raw is None:
        st.error("所选工作表为空或读取失败，请切换其他工作表重试。")
        return

    df, missing = _prepare_sheet_df(df_raw)
    if missing:
        st.warning(
            "当前工作表缺少部分标准列，已进入兼容模式：仍可搜索与展示，但标准话术/固定字段可能显示为空。\n\n"
            f"缺失列：{', '.join(missing)}"
        )

    st.subheader("模糊搜索")

    def _resolve_optional_col(*candidates: str) -> Optional[str]:
        for c in candidates:
            if c in df.columns:
                return c
        return None

    optional_cols = {
        "规格": _resolve_optional_col("规格"),
        "成分": _resolve_optional_col("成分", "成份"),
    }

    selectable = [("产品名称", "产品名称"), ("颜色", "颜色")]
    if optional_cols["规格"] is not None:
        selectable.append(("规格", optional_cols["规格"]))
    if optional_cols["成分"] is not None:
        selectable.append(("成分", optional_cols["成分"]))

    selected_labels = st.multiselect(
        "检索字段",
        options=[x[0] for x in selectable],
        default=[x[0] for x in selectable],
        help="可多选。若 Excel 中不存在“规格/成分”列，会自动不显示该选项。",
    )
    selected_cols = [col for label, col in selectable if label in set(selected_labels)]

    query = st.text_input(
        "输入关键字（支持模糊匹配）",
        placeholder="例如：遮光 / 奶茶 / 高级灰 / 法式 / 珍珠白 ...",
    )

    # 过滤
    if query.strip():
        q = query.strip().lower()
        # 选中的字段任意命中即可
        if not selected_cols:
            # 兜底：优先常用列，否则用所有列
            preferred = [c for c in ["产品名称", "颜色"] if c in df.columns]
            selected_cols = preferred if preferred else list(df.columns)
        mask = False
        for col in selected_cols:
            if col not in df.columns:
                continue
            mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)
        result = df[mask].copy()
    else:
        result = df.copy()

    # 优先展示需要的列
    preferred_show_cols = ["产品名称", "颜色", "门幅", "克重（g/m²）", "价格（元/米）"]
    show_cols = [c for c in preferred_show_cols if c in result.columns]
    if show_cols:
        result = result[show_cols]

    left, right = st.columns([1, 1])
    with left:
        st.metric("匹配结果", value=len(result))
    with right:
        st.caption("提示：结果很多时建议先输入关键字缩小范围。")

    st.divider()

    # 结果展示：逐条卡片（更适合客服一键复制）
    if len(result) == 0:
        st.warning("没有匹配到结果，请换个关键字试试。")
        return

    # 限制一次渲染过多卡片导致卡顿（客服场景通常不需要一下看几千条）
    max_render = 200
    if len(result) > max_render:
        st.info(f"结果较多（{len(result)} 条），仅展示前 {max_render} 条。建议继续输入关键字缩小范围。")
        result_to_render = result.head(max_render)
    else:
        result_to_render = result

    # 点击“复制话术”后，用 JS 自动复制
    if "copy_text" not in st.session_state:
        st.session_state.copy_text = None

    for idx, row in result_to_render.iterrows():
        with st.container(border=True):
            c1, c2, c3 = st.columns([0.55, 0.25, 0.20])

            with c1:
                name = _safe_series_get(row, "产品名称")
                color = _safe_series_get(row, "颜色")
                if not name:
                    # 兜底：用第一列做标题（避免不同表头时空白）
                    first_col = result_to_render.columns[0] if len(result_to_render.columns) > 0 else ""
                    name = _safe_series_get(row, first_col) if first_col else "（未命名产品）"

                st.markdown(
                    f"**{name}**" + (f"  ·  **{color}**" if color else "")
                )

                # 仅展示存在的关键字段，避免 KeyError
                meta_parts = []
                if "门幅" in result_to_render.columns or "门幅" in df.columns:
                    v = _safe_series_get(row, "门幅")
                    if v:
                        meta_parts.append(f"门幅：{v}")
                if "克重（g/m²）" in result_to_render.columns or "克重（g/m²）" in df.columns:
                    v = _safe_series_get(row, "克重（g/m²）")
                    if v:
                        meta_parts.append(f"克重：{v}")
                if "价格（元/米）" in result_to_render.columns or "价格（元/米）" in df.columns:
                    v = _safe_series_get(row, "价格（元/米）")
                    if v:
                        meta_parts.append(f"价格：{v} 元/米")
                if meta_parts:
                    st.write("｜".join(meta_parts))

            with c2:
                script = build_script(template, row)
                st.text_area("预览话术", value=script, height=90, key=f"preview_{idx}")

            with c3:
                if st.button("复制话术", key=f"copy_{idx}", use_container_width=True):
                    st.session_state.copy_text = script
                    st.toast("已复制到剪贴板", icon="✅")

    # 放在循环后统一执行复制（避免重复插入 JS）
    if st.session_state.copy_text:
        copy_to_clipboard_js(st.session_state.copy_text)
        # 复制完成后清空，避免每次 rerun 都重复复制
        st.session_state.copy_text = None


if __name__ == "__main__":
    main()
