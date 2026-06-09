# -*- coding: utf-8 -*-
"""
永寬化學股份有限公司
ISO 50001 重大能源使用設備 (SEUs) 網頁版管理系統 - UI/UX 優化版
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64, json, os, math
from io import BytesIO
from PIL import Image
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# 頁面基本設定與客製化 CSS 樣式
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="永寬化學 - ISO 50001 SEUs 管理系統", layout="wide", page_icon="⚡")

# 全域樣式優化：優化字體、按鈕與區塊陰影
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    html, body, [data-testid="stWidgetLabel"] { font-family: 'Noto Sans TC', sans-serif !important; }
    .stButton>button { border-radius: 6px; font-weight: 500; }
    .stTabs [data-baseweb="tab"] { font-size: 16px; font-weight: 500; }
    div[data-testid="stExpander"] { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

DB_FILE = "energy_db.json"
EXCEL_FILE = "永寬化學_ISO50001_設備清單.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# Helper: 置中表格 (保留原邏輯，優化寬度與外觀)
# ─────────────────────────────────────────────────────────────────────────────
def centered_table(df):
    """將 DataFrame 轉為文字置中的 HTML 表格"""
    styles = """
    <style>
    .ctable { width:100%; border-collapse:collapse; font-size:14px; margin-bottom:15px; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
    .ctable th {
        background:#1a3a5c; color:#fff; padding:12px 14px;
        text-align:center; font-weight:600; border:1px solid #334155;
    }
    .ctable td {
        padding:10px 14px; text-align:center;
        border:1px solid #e2e8f0; color:#1e293b;
    }
    .ctable tr:nth-child(even) td { background:#f8fafc; }
    .ctable tr:hover td { background:#f1f5f9; }
    </style>
    """
    rows_html = ""
    for _, row in df.iterrows():
        tds = ""
        for val in row.values:
            if isinstance(val, float):
                if val == int(val):
                    tds += f"<td>{int(val)}</td>"
                else:
                    tds += f"<td>{val:.2f}</td>"
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                tds += f"<td>{val:,}</td>"
            else:
                tds += f"<td>{str(val)}</td>"
        rows_html += f"<tr>{tds}</tr>"

    cols_html = "".join([f"<th>{col}</th>" for col in df.columns])
    table_html = f"{styles}<table class='ctable'><thead><tr>{cols_html}</tr></thead><tbody>{rows_html}</tbody></table>"
    return table_html

# ─────────────────────────────────────────────────────────────────────────────
# 資料處理核心邏輯 (讀取/載入/儲存)
# ─────────────────────────────────────────────────────────────────────────────
def get_builtin_data():
    """內建示範資料"""
    return [
        {"設備編號": "PUMP-01", "設備名稱": "1號冷卻水泵", "區域/製程": "A棟公用區", "馬力/功率(kW或HP)": "15 kW", "全年運轉時數(小時)": 6000, "照片Base64": ""},
        {"設備編號": "COMP-01", "設備名稱": "甲班空壓機", "區域/製程": "B棟生產線", "馬力/功率(kW或HP)": "50 HP", "全年運轉時數(小時)": 7200, "照片Base64": ""},
        {"設備編號": "CH-01", "設備名稱": "中央冰水主機", "區域/製程": "研發大樓", "馬力/功率(kW或HP)": "80 kW", "全年運轉時數(小時)": 4000, "照片Base64": ""},
        {"設備編號": "EX-01", "設備名稱": "現場排風機組", "區域/製程": "C棟包裝區", "馬力/功率(kW或HP)": "5.5 kW", "全年運轉時數(小時)": 2500, "照片Base64": ""},
        {"設備編號": "BLO-01", "設備名稱": "鍋爐送風機", "區域/製程": "熱能動力區", "馬力/功率(kW或HP)": "22 kW", "全年運轉時數(小時)": 3000, "照片Base64": ""},
    ]

def parse_kw(power_str):
    """解析馬力/功率字串為 kW 數值"""
    try:
        s = str(power_str).upper().strip()
        if "HP" in s:
            val = float(s.replace("HP", "").strip())
            return val * 0.746
        elif "KW" in s:
            return float(s.replace("KW", "").strip())
        else:
            return float(s)
    except:
        return 0.0

def load_json():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return get_builtin_data()
    return get_builtin_data()

def save_json(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ─────────────────────────────────────────────────────────────────────────────
# 系統初始化
# ─────────────────────────────────────────────────────────────────────────────
if "db" not in st.session_state:
    st.session_state["db"] = load_json()

# 🔄 將 JSON 轉換成 Pandas DataFrame 並計算能耗
raw_df = pd.DataFrame(st.session_state["db"])
if raw_df.empty:
    raw_df = pd.DataFrame(columns=["設備編號", "設備名稱", "區域/製程", "馬力/功率(kW或HP)", "全年運轉時數(小時)", "照片Base64"])

kw_list = [parse_kw(x) for x in raw_df["馬力/功率(kW或HP)"]]
kwh_list = [kw * float(hours) for kw, hours in zip(kw_list, raw_df["全年運轉時數(小時)"])]
raw_df["估算功率(kW)"] = kw_list
raw_df["年用電量(kWh)"] = kwh_list

total_kwh = raw_df["年用電量(kWh)"].sum()
if total_kwh > 0:
    raw_df["用電佔比(%)"] = (raw_df["年用電量(kWh)"] / total_kwh) * 100
else:
    raw_df["用電佔比(%)"] = 0.0

# 排序並計算 80/20 法則累計百分比
df_sorted = raw_df.sort_values(by="年用電量(kWh)", ascending=False).reset_index(drop=True)
cum_sum = 0.0
cum_pct_list = []
seu_list = []

for idx, row in df_sorted.iterrows():
    cum_sum += row["用電佔比(%)"]
    cum_pct_list.append(cum_sum)
    # 只要前一個節點累計不超過 80%，或者該設備本身就是讓能耗突破 80% 的關鍵設備，皆定義為 SEU
    if idx == 0:
        seu_list.append(True)
    elif cum_pct_list[max(0, idx-1)] <= 80.0:
        seu_list.append(True)
    else:
        seu_list.append(False)

df_sorted["累計百分比(%)"] = cum_pct_list
df_sorted["是否為SEU"] = seu_list

seu_count = sum(seu_list)
seu_kwh = df_sorted[df_sorted["是否為SEU"]]["年用電量(kWh)"].sum()
seu_ratio = seu_kwh / total_kwh if total_kwh > 0 else 0.0

# ─────────────────────────────────────────────────────────────────────────────
# 側邊欄設計 (Sidebar Layout)
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<h2 style='color:#1a3a5c; margin-bottom:5px;'>🏢 永寬化學</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748b; font-size:14px; margin-top:0;'>ISO 50001 能源管理系統</p>", unsafe_allow_html=True)
    st.write("---")
    
    # 系統選單切換
    menu = st.radio(
        "📂 系統主選單",
        ["📊 首頁看板 & SEU 辨識", "➕ 新增/管理設備檔案", "📈 模擬 24H 負載曲線", "🔄 從 Excel 重新載入"],
        index=0
    )
    
    st.write("---")
    st.markdown("⚙️ **系統進階設定**")
    # 將修改模式改為更美觀的 st.toggle
    edit_mode = st.toggle("✏️ 開啟資料修改模式", value=st.session_state.get("edit_mode", False))
    st.session_state["edit_mode"] = edit_mode
    
    st.write("---")
    st.caption(f"系統版本: v2.5 (UI 改版) | 💡 建議能耗涵蓋率應大於 80%")

# ── 安全防錯提示：若開啟修改模式，在主畫面上方跳出醒目通知 ──
if st.session_state["edit_mode"]:
    st.warning("🚨 **目前處於［修改模式］**：您可以進行欄位編輯、刪除設備。完成操作後建議關閉此模式以防誤觸。")

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面一：首頁看板 & SEU 辨識
# ─────────────────────────────────────────────────────────────────────────────
if "首頁看板" in menu:
    st.title("⚡ 重大能源使用設備 (SEUs) 管理戰情室")
    st.markdown("本系統依據 **ISO 50001:2018** 條文要求，進行全廠設備能耗審查與重大能源使用（SEU）之排定與鑑別。")
    
    # 優化頂部三大 KPI 大盤卡片
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(
            f"""
            <div style="background-color:#f1f5f9; padding:22px; border-radius:12px; border-left: 6px solid #1a3a5c; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                <p style="margin:0; font-size:14px; color:#64748b; font-weight:bold;">⚡ 全廠設備估算總用電量</p>
                <h2 style="margin:8px 0 0 0; color:#1a3a5c; font-size:28px;">{int(total_kwh):,} <span style="font-size:16px;">kWh/年</span></h2>
            </div>
            """, unsafe_allow_html=True
        )
    with m2:
        st.markdown(
            f"""
            <div style="background-color:#fef2f2; padding:22px; border-radius:12px; border-left: 6px solid #ef4444; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                <p style="margin:0; font-size:14px; color:#991b1b; font-weight:bold;">🚨 已辨識重大能源使用 (SEUs)</p>
                <h2 style="margin:8px 0 0 0; color:#b91c1c; font-size:28px;">{seu_count} <span style="font-size:16px;">台設備</span></h2>
            </div>
            """, unsafe_allow_html=True
        )
    with m3:
        # 根據是否達標 (>80%) 自動變換邊框顏色
        border_color = "#22c55e" if seu_ratio >= 0.8 else "#f59e0b"
        text_color = "#166534" if seu_ratio >= 0.8 else "#9a3412"
        bg_color = "#f0fdf4" if seu_ratio >= 0.8 else "#fffbeb"
        st.markdown(
            f"""
            <div style="background-color:{bg_color}; padding:22px; border-radius:12px; border-left: 6px solid {border_color}; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                <p style="margin:0; font-size:14px; color:{text_color}; font-weight:bold;">📊 SEUs 能耗總涵蓋率</p>
                <h2 style="margin:8px 0 0 0; color:{text_color}; font-size:28px;">{seu_ratio * 100:.1f}% <span style="font-size:14px; color:#64748b; font-weight:normal;">(法規目標: >80%)</span></h2>
            </div>
            """, unsafe_allow_html=True
        )
    
    st.write("")
    
    # 建立左右兩欄：左邊放專業巴雷托圖與統計，右邊放表格
    col_chart, col_table = st.columns([11, 10])
    
    with col_chart:
        if not df_sorted.empty:
            # 🎨 重構一體化巴雷托分析圖 (Pareto Chart)
            fig = go.Figure()
            # 1. 各設備能耗長條圖 (企業藍)
            fig.add_trace(go.Bar(
                x=df_sorted["設備編號"] + " " + df_sorted["設備名稱"],
                y=df_sorted["年用電量(kWh)"],
                name="年用電量 (kWh)",
                marker_color="#1a3a5c",
                hovertemplate="設備: %{x}<br>能耗: %{y:,.0f} kWh<br>"
            ))
            # 2. 累計百分比折線圖 (警示紅)
            fig.add_trace(go.Scatter(
                x=df_sorted["設備編號"] + " " + df_sorted["設備名稱"],
                y=df_sorted["累計百分比(%)"],
                name="累計百分比 (%)",
                yaxis="y2",
                line=dict(color="#ef4444", width=3, dash="dash"),
                hovertemplate="累計佔比: %{y:.1f}%<br>"
            ))
            # 3. 畫出 80% 門檻輔助線
            fig.add_hline(y=80, yref="y2", line_color="#22c55e", line_width=2, 
                          annotation_text="ISO 80% 門檻線", annotation_position="top left",
                          annotation_font=dict(color="#22c55e", size=12))
            
            fig.update_layout(
                title=dict(text="<b>📊 全廠設備能耗巴雷托分析圖 (80/20 法則)</b>", font=dict(size=16)),
                xaxis=dict(tickangle=35, title="設備名稱"),
                yaxis=dict(title="年用電量 (kWh)", titlefont=dict(color="#1a3a5c")),
                yaxis2=dict(title="累計百分比 (%)", titlefont=dict(color="#ef4444"), overlaying="y", side="right", range=[0, 105]),
                margin=dict(l=40, r=40, t=50, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                hovermode="x unified",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("目前無設備資料，請先新增資料。")
            
    with col_table:
        st.markdown("##### 📄 全廠設備能耗審查與 SEU 鑑別清單")
        # 隱藏照片欄位後，輸出美化置中表格
        display_df = df_sorted.copy()
        if "照片Base64" in display_df.columns:
            display_df = display_df.drop(columns=["照片Base64"])
        
        # 將布林值轉為直觀的亮眼文字
        display_df["是否為SEU"] = display_df["是否為SEU"].map({True: "🚨 重大能耗 (SEU)", False: "📄 一般設備"})
        st.markdown(centered_table(display_df), unsafe_allow_html=True)
        
    st.write("---")
    
    # ⚙️ 設備檔案與現場照片巡檢：使用 Expander 摺疊優化介面密度
    st.markdown("### 🔍 設備現場檔案與照片巡檢")
    st.markdown("點擊下方各設備摺疊面板，可查看詳細能源參數及現場點檢照片。")
    
    for index, row in df_sorted.iterrows():
        is_seu_label = "🚨 [SEU 重大]" if row["是否為SEU"] else "📄 [一般設備]"
        card_title = f"{is_seu_label} {row['設備編號']} - {row['設備名稱']} (年用電: {int(row['年用電量(kWh)']):,} kWh)"
        
        with st.expander(card_title, expanded=False):
            c_txt, c_img = st.columns([3, 2])
            with c_txt:
                st.markdown(f"**📍 區域/製程：** {row['區域/製程']}")
                st.markdown(f"**⚡ 馬力/功率：** {row['馬力/功率(kW或HP)']}")
                st.markdown(f"**⏱️ 全年運轉時數：** {int(row['全年運轉時數(小時)']):,} 小時")
                st.markdown(f"**📈 估算功率：** {row['估算功率(kW)']:.2f} kW")
                st.markdown(f"**📊 總能耗佔比：** {row['用電佔比(%)']:.2f}% (全廠排名第 {index+1} 名)")
                
                if row["是否為SEU"]:
                    st.error("💡 **ISO 50001 管理指引**：此設備已列為重大能源使用。建議排定日常操作操作指引（OI）、設定能源操作目標、並優先導入高效能馬達(IE4/IE5)或加裝變頻器監控。")
                else:
                    st.info("💡 **管理指引**：此設備屬一般能耗，維持現行維護與自主點檢即可。")
            
            with c_img:
                img_b64 = row.get("照片Base64", "")
                if isinstance(img_b64, str) and img_b64.strip():
                    try:
                        img_data = base64.b64decode(img_b64)
                        img = Image.open(BytesIO(img_data))
                        st.image(img, caption=f"{row['設備名稱']} 現場照片", use_container_width=True)
                    except:
                        st.caption("❌ 現場照片解碼失敗")
                else:
                    st.caption("📷 尚未上傳該設備之現場實照")

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面二：新增/管理設備檔案
# ─────────────────────────────────────────────────────────────────────────────
elif "新增" in menu:
    st.title("➕ 新增與管理設備檔案")
    
    # 修改與刪除區塊
    if st.session_state["edit_mode"]:
        st.markdown("### ✏️ 修改/刪除既有設備")
        if not st.session_state["db"]:
            st.info("目前資料庫為空。")
        else:
            del_list = [f"{x['設備編號']} - {x['設備名稱']}" for x in st.session_state["db"]]
            target_del = st.selectbox("請選擇要管理的設備", del_list)
            target_idx = del_list.index(target_del)
            
            col_btn1, col_btn2 = st.columns([1, 4])
            with col_btn1:
                if st.button("❌ 確認刪除此設備", type="primary"):
                    removed = st.session_state["db"].pop(target_idx)
                    save_json(st.session_state["db"])
                    st.success(f"✅ 已成功刪除設備：{removed['設備編號']}")
                    st.rerun()
            st.write("---")

    # 新增區塊
    st.markdown("### 📥 錄入新設備或更新檔案")
    with st.form("add_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            new_id = st.text_input("設備編號 * (例: PUMP-02)", placeholder="必填且不可重複")
            new_name = st.text_input("設備名稱 * (例: 2號冷卻水泵)", placeholder="必填")
            new_area = st.text_input("區域/製程 (例: A棟廠房)", placeholder="選填")
        with c2:
            new_power = st.text_input("馬力/功率 * (例: 15 kW 或 20 HP)", placeholder="必填，請輸入數值與單位")
            new_hours = st.number_input("全年運轉時數(小時) *", min_value=0, max_value=8760, value=2000)
            new_file = st.file_uploader("上傳設備現場照片", type=["jpg", "png", "jpeg"])
            
        submit = st.form_submit_submit = st.form_submit_button("💾 儲存並新增至系統資料庫", type="primary")
        
        if submit:
            if not new_id.strip() or not new_name.strip() or not new_power.strip():
                st.error("❌ 請注意：設備編號、名稱與馬力/功率為必填欄位！")
            else:
                # 檢查重複
                id_exists = any(x["設備編號"].upper() == new_id.upper() for x in st.session_state["db"])
                if id_exists:
                    st.error(f"❌ 設備編號「{new_id}」已存在系統中，請使用其他編號。")
                else:
                    b64_str = ""
                    if new_file is not None:
                        try:
                            # 壓縮並轉換照片
                            img = Image.open(new_file)
                            img.thumbnail((600, 600))
                            buffer = BytesIO()
                            img.save(buffer, format="JPEG", quality=85)
                            b64_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
                        except Exception as e:
                            st.error(f"照片處理失敗: {e}")
                            
                    new_item = {
                        "設備編號": new_id.strip(),
                        "設備名稱": new_name.strip(),
                        "區域/製程": new_area.strip() if new_area.strip() else "未分類",
                        "馬力/功率(kW或HP)": new_power.strip(),
                        "全年運轉時數(小時)": int(new_hours),
                        "照片Base64": b64_str
                    }
                    st.session_state["db"].append(new_item)
                    save_json(st.session_state["db"])
                    st.success(f"🎉 成功加入新設備：[{new_id}] {new_name}")
                    st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面三：模擬 24H 負載曲線
# ─────────────────────────────────────────────────────────────────────────────
elif "負載曲線" in menu:
    st.title("📈 模擬 24H 廠區動態用電負載曲線")
    st.markdown("依據現有設備檔案之運轉功率，在此頁面您可以自訂各設備在「一天 24 小時中」的啟停稼動狀況，以模擬全廠的總電力負載走勢。")
    
    if df_sorted.empty:
        st.info("目前無設備資料，請先前往新增設備。")
    else:
        st.markdown("#### ⏱️ 自訂各設備 24H 稼動時段 (時段打勾代表該小時運轉)")
        
        hours_cols = [f"{h:02d}:00" for h in range(24)]
        schedule_dict = {}
        
        # 透過多功能摺疊面板提供自訂時段，避免介面爆炸
        for idx, row in df_sorted.iterrows():
            with st.expander(f"⚙️ 設定時段：{row['設備編號']} - {row['設備名稱']} (功率: {row['估算功率(kW)']:.1f} kW)", expanded=(idx==0)):
                st.caption("請勾選該設備在哪些小時區間會正常啟動開機：")
                
                # 快速勾選捷徑
                c_all, c_clear, _ = st.columns([1, 1, 8])
                state_key_all = f"all_{row['設備編號']}"
                
                # 24個小時分為4排顯示，版面更整齊
                checked_hours = []
                r1 = st.columns(6)
                r2 = st.columns(6)
                r3 = st.columns(6)
                r4 = st.columns(6)
                all_rows = r1 + r2 + r3 + r4
                
                for h in range(24):
                    with all_rows[h]:
                        # 預設全選，方便使用者
                        is_checked = st.checkbox(f"{h:02d}:00", value=True, key=f"cb_{row['設備編號']}_{h}")
                        if is_checked:
                            checked_hours.append(h)
                schedule_dict[row["設備編號"]] = checked_hours

        # 計算 24 小時每小時總負載
        hourly_load = [0.0] * 24
        seu_hourly_load = [0.0] * 24
        
        for idx, row in df_sorted.iterrows():
            eq_id = row["設備編號"]
            kw = row["估算功率(kW)"]
            is_seu = row["是否為SEU"]
            active_hours = schedule_dict.get(eq_id, list(range(24)))
            
            for h in active_hours:
                hourly_load[h] += kw
                if is_seu:
                    seu_hourly_load[h] += kw
                    
        df_load = pd.DataFrame({
            "時間時段": hours_cols,
            "全廠總用電負載 (kW)": hourly_load,
            "重大能耗設備(SEUs)負載 (kW)": seu_hourly_load
        })
        
        st.write("---")
        st.markdown("#### 📊 模擬 24H 負載曲線視覺化看板")
        
        # 繪製漂亮的24H負載曲線圖
        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=df_load["時間時段"], y=df_load["全廠總用電負載 (kW)"],
            mode="lines+markers", name="全廠總電力負載 (kW)",
            line=dict(color="#1a3a5c", width=3),
            marker=dict(size=6)
        ))
        fig_line.add_trace(go.Scatter(
            x=df_load["時間時段"], y=df_load["重大能耗設備(SEUs)負載 (kW)"],
            mode="lines+markers", name="SEUs 貢獻負載 (kW)",
            line=dict(color="#ef4444", width=2, dash="dot"),
            marker=dict(size=4)
        ))
        
        fig_line.update_layout(
            title="<b>⏱️ 全廠動態電力需求模擬曲線 (24 Hours Load Profile)</b>",
            xaxis=dict(title="時間區間", gridcolor="#e2e8f0"),
            yaxis=dict(title="電力需求容量 (kW)", gridcolor="#e2e8f0"),
            hovermode="x unified",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        c_chart, c_metrics = st.columns([3, 1])
        with c_chart:
            st.plotly_chart(fig_line, use_container_width=True)
        with c_metrics:
            max_load = max(hourly_load)
            max_hour = hourly_load.index(max_load)
            avg_load = sum(hourly_load) / 24
            
            st.markdown("##### 📈 負載指標分析")
            st.info(f"**⚡ 最高用電尖峰：**\n{max_load:.2f} kW\n(發生於 {max_hour:02d}:00)")
            st.success(f"**📊 平均電力需求：**\n{avg_load:.2f} kW")
            st.metric(label="📉 廠區用電負載因數", value=f"{(avg_load/max_load*100 if max_load>0 else 0):.1f} %")
            
            st.download_button(
                "⬇️ 下載模擬負載數據 CSV",
                df_load.to_csv(index=False).encode("utf-8-sig"),
                "YuanKuan_24H_Load.csv", "text/csv",
                use_container_width=True
            )

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面四：從 Excel 重新載入
# ─────────────────────────────────────────────────────────────────────────────
elif "Excel" in menu:
    st.subheader("🔄 從系統外部 Excel 檔案重新同步")
    st.markdown("若您在廠內有維護一份包含欄位（`設備編號`、`設備名稱`、`區域/製程`、`馬力/功率(kW或HP)`、`全年運轉時數(小時)`）的 Excel 表，可置於目錄下重新覆蓋導入。")
    st.error("⚠️ 注意：此重置導入操作將完全覆蓋您在網頁版上做過的所有單筆修改與照片，請謹慎操作！")

    if not os.path.exists(EXCEL_FILE):
        st.warning(f"⚠️ 未在雲端根目錄下偵測到指定的 Excel 檔案：`{EXCEL_FILE}`，目前使用系統內建資料庫。")
        if st.session_state["edit_mode"]:
            if st.button("🔄 改為強制重置載入系統內建示範資料", type="primary"):
                st.session_state["db"] = get_builtin_data()
                save_json(st.session_state["db"])
                st.success(f"✅ 已成功重置！共載入 {len(st.session_state['db'])} 筆標準示範設備。")
                st.rerun()
        else:
            st.info("請在側邊欄開啟「✏️ 修改模式」才能執行重置載入功能。")
    else:
        st.success(f"✅ 成功找到指定 Excel 來源檔案：{EXCEL_FILE}")
        if st.session_state["edit_mode"]:
            if st.button("🔄 確認自 Excel 覆蓋更新系統資料庫", type="primary"):
                try:
                    df_ex = pd.read_excel(EXCEL_FILE)
                    # 確保必要欄位存在
                    needed = ["設備編號", "設備名稱", "區域/製程", "馬力/功率(kW或HP)", "全年運轉時數(小時)"]
                    for c in needed:
                        if c not in df_ex.columns:
                            df_ex[c] = "" if c != "全年運轉時數(小時)" else 2000
                    
                    new_db = []
                    for _, r in df_ex.iterrows():
                        new_db.append({
                            "設備編號": str(r["設備編號"]).strip(),
                            "設備名稱": str(r["設備名稱"]).strip(),
                            "區域/製程": str(r["區域/製程"]).strip(),
                            "馬力/功率(kW或HP)": str(r["馬力/功率(kW或HP)"]).strip(),
                            "全年運轉時數(小時)": int(r["全年運轉時數(小時)"]) if pd.notnull(r["全年運轉時數(小時)"]) else 2000,
                            "照片Base64": ""
                        })
                    st.session_state["db"] = new_db
                    save_json(new_db)
                    st.success(f"🎉 檔案匯入成功！已從 Excel 載入 {len(new_db)} 筆全新設備檔案。")
                    st.rerun()
                except Exception as e:
                    st.error(f"Excel 解析錯誤，請檢查欄位格式是否正確：{e}")
        else:
            st.info("請先在左側邊欄切換開啟「✏️ 修改模式」方可執行覆蓋載入。")