import streamlit as st
from io import BytesIO
from engine import generate_ranking, load_genre_catalog

st.set_page_config(page_title="月銷售排行產生器", page_icon="📊", layout="wide")
st.title("📊 月銷售排行產生器")
st.caption("上傳各平台原始報表，自動整合成銷售排行 Excel")

# ── 月份 & 範本 ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    month = st.text_input("當月月份（格式 YYYYMM）", placeholder="例：202604")
with col2:
    template_file = st.file_uploader(
        "上個月的銷售排行 Excel（選填，用來保留歷史分頁）",
        type=["xlsx"],
        key="template",
    )

st.divider()

# ── 書種目錄上傳 ──────────────────────────────────────────────────────────────
st.subheader("書種目錄")
catalog_file = st.file_uploader(
    "product_fixed.xlsx（BW 書種目錄，選填，沒上傳則書種欄位留空）",
    type=["xlsx"],
    key="catalog",
)

@st.cache_resource(show_spinner="載入書種目錄中…")
def get_catalog_from_bytes(data: bytes):
    return load_genre_catalog(BytesIO(data))

genre_catalog = None
if catalog_file is not None:
    genre_catalog = get_catalog_from_bytes(catalog_file.read())
    st.caption(f"書種目錄已載入（{len(genre_catalog):,} 筆）")

st.divider()

# ── 各平台上傳 ────────────────────────────────────────────────────────────────
st.subheader("各平台報表上傳")

c1, c2 = st.columns(2)
with c1:
    google_file = st.file_uploader(
        "Google Play（*稅務試算.xlsx）",
        type=["xlsx", "xls"],
        key="google",
    )
    kobo_files = st.file_uploader(
        "Kobo（TW_BOOKWALKER / TW_GAEABOOKS / TW_KADOKAWA 共 3 個 .xlsx）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="kobo",
    )
    readmoo_file = st.file_uploader(
        "Readmoo（台灣漫讀拆帳報表 YYYY-MM.xlsx）",
        type=["xlsx", "xls"],
        key="readmoo",
    )

with c2:
    bokuto_files = st.file_uploader(
        "博客來（YYYYMM_bookwalker.xls + YYYYMM_bookwalker1.xls 共 2 個）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="bokuto",
    )
    luanda_file = st.file_uploader(
        "亂搭（YYYYMM_台灣漫讀.xlsx）",
        type=["xlsx", "xls"],
        key="luanda",
    )
    pubu_files = st.file_uploader(
        "Pubu + 誠品 + PChome（共 3 個 .xls）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="pubu",
    )

st.divider()

# ── 產生按鈕 ─────────────────────────────────────────────────────────────────
st.subheader("產生排行")

ready = bool(month and len(month) == 6 and month.isdigit())
if not ready:
    st.info("請先填寫月份（6 位數字），再上傳各平台報表。")

if st.button("🚀 開始產生", disabled=not ready, type="primary"):
    progress_bar = st.progress(0, text="準備中…")
    status_text = st.empty()

    def on_progress(step, total, message):
        progress_bar.progress(step / total, text=message)
        status_text.text(message)

    try:
        result_bytes = generate_ranking(
            month=month,
            google_file=google_file,
            kobo_files=kobo_files if kobo_files else None,
            readmoo_file=readmoo_file,
            bokuto_files=bokuto_files if bokuto_files else None,
            luanda_file=luanda_file,
            pubu_files=pubu_files if pubu_files else None,
            template_file=template_file,
            genre_catalog=genre_catalog,
            progress_cb=on_progress,
        )
        progress_bar.progress(1.0, text="完成！")
        status_text.empty()
        st.success("✅ 銷售排行已產生！")

        from datetime import date
        today = date.today().strftime("%m%d")
        filename = f"{month}銷售排行_{today}.xlsx"

        st.download_button(
            label=f"⬇️ 下載 {filename}",
            data=result_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        progress_bar.empty()
        st.error(f"❌ 產生失敗：{e}")
        st.exception(e)
