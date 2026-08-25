# -*- coding: utf-8 -*-
"""FedEx Reprice Studio — 網頁版
計價引擎 100% 沿用桌面版 FedEx_RepricingTool.py(經 adapter 無頭載入),
這裡只做網頁介面:設定費率 → 上傳帳單 → 下載重新計價結果。
"""
import io
import json

import pandas as pd
import streamlit as st

st.set_page_config(page_title="FedEx Reprice Studio", page_icon="📦",
                   layout="wide")

# ---------------------------------------------------------------- 密碼閘門
# 在 Streamlit Cloud 的 Settings → Secrets 設 APP_PASSWORD = "你的密碼"
# 沒設就不擋(例如本機測試)。
_pw = ""
try:
    _pw = st.secrets.get("APP_PASSWORD", "")
except Exception:
    _pw = ""
if _pw:
    if not st.session_state.get("_authed"):
        st.title("📦 FedEx Reprice Studio")
        got = st.text_input("請輸入密碼 Password", type="password")
        if st.button("登入 Sign in") or got:
            if got == _pw:
                st.session_state["_authed"] = True
                st.rerun()
            elif got:
                st.error("密碼不對")
        st.stop()

import adapter  # noqa: E402  (放在密碼之後,沒登入就不載引擎)

if "engine" not in st.session_state:
    st.session_state["engine"] = adapter.new_app()
app = st.session_state["engine"]

st.markdown(
    "<h2 style='color:#4D148C;margin-bottom:0'>📦 FedEx Reprice Studio</h2>"
    "<p style='color:#888;margin-top:2px'>計價引擎與桌面版完全相同 · "
    "設定檔 billing_tool_config.json 可雙向互通</p>",
    unsafe_allow_html=True)

TABLE_LABELS = {
    "base_rate_table": "運費表 Base Rates",
    "ahs_rate_table": "AHS 附加費",
    "das_rate_table": "DAS 附加費",
    "oversize_rate_table": "Oversize 附加費",
    "residential_rate_table": "住宅附加費",
    "signature_rate_table": "簽收費",
    "general_demand_table": "旺季費 General",
    "ahs_demand_table": "旺季費 AHS",
    "oversize_demand_table": "旺季費 Oversize",
}

# ---------------------------------------------------------------- 側欄:設定檔
with st.sidebar:
    st.subheader("⚙️ 設定檔")
    up = st.file_uploader("匯入設定 JSON(桌面版存的檔)", type=["json"],
                          key="cfg_up")
    if up is not None and st.session_state.get("_cfg_loaded") != up.name + str(up.size):
        try:
            adapter.load_config_json(app, up.getvalue().decode("utf-8"))
            st.session_state["_cfg_loaded"] = up.name + str(up.size)
            st.success("設定已載入:" + up.name)
        except Exception as e:
            st.error(f"讀不動這個設定檔:{e}")
    st.download_button("⬇️ 下載目前設定 JSON",
                       data=adapter.dump_config_json(app),
                       file_name="billing_tool_config.json",
                       mime="application/json",
                       use_container_width=True)
    st.caption("下載後可直接給桌面版 Load Config 使用,反之亦然。")
    st.divider()
    try:
        st.download_button("📄 下載範例帳單格式",
                           data=adapter.make_billing_sample(app),
                           file_name="billing_sample.xlsx",
                           use_container_width=True)
    except Exception:
        pass

tab_run, tab_rules, tab_tables, tab_help = st.tabs(
    ["▶️ 執行 Reprice", "🧮 全域規則", "📋 費率表", "📖 說明"])

# ---------------------------------------------------------------- 執行
with tab_run:
    st.subheader("上傳 FedEx 帳單,重新計價")
    billing = st.file_uploader("FedEx Billing 檔(CSV 或 XLSX)",
                               type=["csv", "xlsx", "xls"])
    col1, col2 = st.columns([1, 3])
    with col1:
        go = st.button("🚀 Generate Reprice Report", type="primary",
                       disabled=billing is None, use_container_width=True)
    if go and billing is not None:
        with st.spinner("計價中…"):
            try:
                out, summary = adapter.run_repricing(
                    app, billing.getvalue(), billing.name)
                st.session_state["_result"] = out
                st.session_state["_result_name"] = (
                    billing.name.rsplit(".", 1)[0] + "_repriced.xlsx")
                st.session_state["_summary"] = summary
            except adapter.RunError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"執行失敗:{e}")
    if st.session_state.get("_result"):
        st.success(st.session_state.get("_summary") or "完成")
        st.download_button("⬇️ 下載重新計價結果 (xlsx)",
                           data=st.session_state["_result"],
                           file_name=st.session_state["_result_name"],
                           mime=("application/vnd.openxmlformats-officedocument"
                                 ".spreadsheetml.sheet"),
                           type="primary")
        try:
            prev = pd.read_excel(io.BytesIO(st.session_state["_result"]))
            st.dataframe(prev, use_container_width=True, height=420)
        except Exception:
            pass

# ---------------------------------------------------------------- 全域規則
with tab_rules:
    st.subheader("全域規則(與桌面版 Rules 頁相同)")
    cols = st.columns(3)
    for i, (attr, label, typ) in enumerate(adapter.GLOBAL_RULES):
        var = getattr(app, attr)
        with cols[i % 3]:
            val = st.number_input(label, value=float(var.get()),
                                  key="rule_" + attr, format="%.3f")
            var.set(val)
    st.divider()
    fcols = st.columns(4)
    for i, (attr, label) in enumerate(adapter.GLOBAL_FLAGS):
        var = getattr(app, attr)
        with fcols[i % 4]:
            val = st.checkbox(label, value=bool(var.get()), key="flag_" + attr)
            var.set(val)
    st.caption("改完直接生效;要保存請到左側下載設定 JSON。")

# ---------------------------------------------------------------- 費率表
with tab_tables:
    st.subheader("費率表(可直接編輯,列數用左下角 + 增加)")
    pick = st.selectbox("選擇要編輯的表",
                        list(TABLE_LABELS),
                        format_func=lambda k: TABLE_LABELS[k])
    table = getattr(app, pick)
    df = pd.DataFrame(table.get_rows(), columns=table.columns)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True,
                            height=480, key="ed_" + pick)
    table.load_rows(edited.fillna("").astype(str).values.tolist())
    st.caption(f"{TABLE_LABELS[pick]}:目前 {len(table.rows)} 列。"
               "改完記得到左側下載設定 JSON 存起來。")

# ---------------------------------------------------------------- 說明
with tab_help:
    st.markdown("""
### 使用流程
1. 左側 **匯入設定 JSON** — 直接用桌面版存出來的 `billing_tool_config.json`(沒有的話,系統已載入預設費率,也可在「費率表」分頁直接編)。
2. **執行 Reprice** 分頁 — 上傳 FedEx 帳單(CSV / XLSX),按 Generate。
3. 下載 `_repriced.xlsx` 結果,格式與桌面版輸出完全相同。

### 與桌面版的關係
- 這個網站**直接執行桌面版的計價程式**(`FedEx_RepricingTool.py`,一行未改),不是重寫 — 兩邊算出來一定一樣。
- 設定 JSON 同一格式,可雙向匯入匯出。
- 桌面版更新時,把新的 `FedEx_RepricingTool.py` 覆蓋上來重新部署即可。

### 注意
- 這裡的設定改動只存在**這個瀏覽器分頁的工作階段**;要保存請下載設定 JSON。
- 帳單資料會上傳到伺服器計算後即丟棄,不落地保存;仍建議設好 APP_PASSWORD 密碼。
""")
