# -*- coding: utf-8 -*-
"""
永寬化學股份有限公司
ISO 50001 重大能源使用設備 (SEUs) 網頁版管理系統
v2.1 - 修正圓餅圖空值錯誤、強化資料讀取穩定性
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
# 0. 頁面設定
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="永寬化學 ISO 50001",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="⚡"
)

st.markdown("""
<style>
  [data-testid="stSidebar"] { background:#1a3a5c !important; }
  [data-testid="stSidebar"] * { color:#e2e8f0 !important; }
  [data-testid="stSidebar"] .stButton button {
    background:#2563a8; color:#fff; border-radius:8px; border:none; width:100%;
  }
  .kpi {
    background:#fff; border-radius:12px; padding:14px 10px;
    box-shadow:0 1px 6px rgba(0,0,0,.08); text-align:center;
    min-width:0; overflow:hidden;
  }
  .kpi-v { font-size:22px; font-weight:800; color:#1a3a5c;
            white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .kpi-l { font-size:11px; color:#64748b; margin-top:4px;
            line-height:1.3; word-break:keep-all; }
  .mode-edit {
    background:#fef3c7; border-left:5px solid #f59e0b;
    padding:12px 16px; border-radius:6px; color:#92400e; font-weight:600;
  }
  .mode-view {
    background:#eff6ff; border-left:5px solid #2563a8;
    padding:12px 16px; border-radius:6px; color:#1e40af; font-size:14px;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 1. 常數
# ─────────────────────────────────────────────────────────────────────────────
ADMIN_PASSWORD = "yk50001"        # ← 可自行修改密碼
TOTAL_KWH      = 1_911_641
FLOOR_AREA     = 12_378.27
EXCEL_FILE     = "重大能源使用設備評估表.xlsx"
DB_JSON        = "equipment_db.json"

SYSTEM_SHEETS = {
    "空壓系統": "表4-1、空壓系統",
    "空調系統": "表4-2、空調系統",
    "照明系統": "表4-3、照明系統",
    "製程系統": "表4-4、製程系統",
    "其他系統": "表4-5、其他系統",
}
SYSTEM_ICONS = {
    "空壓系統":"💨","空調系統":"❄️",
    "照明系統":"💡","製程系統":"⚙️","其他系統":"🔧",
}

# ─────────────────────────────────────────────────────────────────────────────
# 2. ISO 50001 計算公式
# ─────────────────────────────────────────────────────────────────────────────
def score_consumption(kwh):
    if kwh < 2500:   return 1
    elif kwh < 5500: return 2
    elif kwh < 7500: return 3
    elif kwh < 10000: return 4
    else:            return 5

def score_power(kw):
    if kw < 2.5:   return 1
    elif kw < 5.0: return 2
    elif kw < 7.5: return 3
    elif kw < 9.0: return 4
    else:          return 5

def calc_row(rec):
    try:
        kw   = float(rec.get("消耗功率(kW)") or 0)
        load = float(rec.get("負載率") or 0)
        hrs  = float(rec.get("運轉時數(hr/年)") or 0)
        qty  = float(rec.get("設備數量") or 1)
        crit = float(rec.get("自評重大性") or 3)
        kwh  = kw * load * hrs * qty
        sc   = round(score_consumption(kwh)*0.3 + score_power(kw)*0.4 + crit*0.3, 2)
        seu  = "A" if sc >= 4.0 else "-"
        return kwh, sc, seu
    except:
        return 0.0, 0.0, "-"

# ─────────────────────────────────────────────────────────────────────────────
# 3. Excel 讀取
# ─────────────────────────────────────────────────────────────────────────────
def _sf(v):
    try:
        if v is None: return None
        if isinstance(v, float) and math.isnan(v): return None
        return float(v)
    except:
        return None

def read_system(sheet, sys_label):
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)
    except Exception as e:
        st.warning(f"無法讀取工作表 {sheet}：{e}")
        return []

    lit   = "照明" in sheet
    kw_c  = 9  if lit else 7
    qty_c = 10 if lit else 8
    ld_c  = 11 if lit else 9
    hrs_c = 13 if lit else 11
    yr_c  = 14 if lit else 12
    age_c = 15 if lit else 13
    cr_c  = 20 if lit else 18

    skip = {
        '', 'nan', '設備名稱', '設備總耗電量', 'A級設備耗電量',
        'A', 'I', '設備總耗能量', 'A級設備耗能量'
    }

    def g(row, c):
        if c >= len(row): return None
        v = row.iloc[c]
        if v is None: return None
        try:
            if isinstance(v, float) and math.isnan(v): return None
        except: pass
        return v

    recs = []
    for i, row in df.iterrows():
        if i < 2: continue
        name = row.iloc[1] if len(row) > 1 else None
        if name is None: continue
        try:
            if isinstance(name, float) and math.isnan(name): continue
        except: pass
        if str(name).strip() in skip: continue

        recs.append({
            "系統別":         sys_label,
            "設備名稱":       str(g(row, 1) or ""),
            "設備編號":       str(g(row, 2) or ""),
            "設備型式":       str(g(row, 3) or ""),
            "設備部門":       str(g(row, 6 if lit else 4) or ""),
            "所在棟別":       str(g(row, 7 if lit else 5) or ""),
            "所在樓層":       str(g(row, 8 if lit else 6) or ""),
            "消耗功率(kW)":   _sf(g(row, kw_c)),
            "設備數量":       _sf(g(row, qty_c)),
            "負載率":         _sf(g(row, ld_c)),
            "運轉時數(hr/年)": _sf(g(row, hrs_c)),
            "設備年份":       _sf(g(row, yr_c)),
            "使用年數":       _sf(g(row, age_c)),
            "自評重大性":     _sf(g(row, cr_c)),
            "設備管理者":     str(g(row, 34) or ""),
            "外包商承攬商":   str(g(row, 35 if not lit else 37) or ""),
            "相關變數":       str(g(row, 36 if not lit else 38) or ""),
            "外觀照片":       None,
            "銘牌照片":       None,
        })
    return recs

# ─────────────────────────────────────────────────────────────────────────────
# 4. 持久化
# ─────────────────────────────────────────────────────────────────────────────
def load_json():
    if os.path.exists(DB_JSON):
        try:
            with open(DB_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return None

def save_json(data):
    with open(DB_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, default=str)

def get_builtin_data():
    """雲端環境或無 Excel 時使用的內建示範資料"""
    return [
        {"系統別":"空調系統","設備名稱":"冷氣機(管理部)","設備編號":"AC-001","設備型式":"分離式冷氣",
         "設備部門":"管理部","所在棟別":"一廠","所在樓層":"1F",
         "消耗功率(kW)":2.39,"設備數量":1,"負載率":0.85,"運轉時數(hr/年)":1936,
         "設備年份":2015,"使用年數":10,"自評重大性":3,
         "設備管理者":"設備組","外包商承攬商":"","相關變數":"","外觀照片":None,"銘牌照片":None},
        {"系統別":"空調系統","設備名稱":"冰水主機","設備編號":"CH-001","設備型式":"離心式冰水主機",
         "設備部門":"製程部","所在棟別":"二廠","所在樓層":"1F",
         "消耗功率(kW)":75.0,"設備數量":2,"負載率":0.90,"運轉時數(hr/年)":6000,
         "設備年份":2010,"使用年數":15,"自評重大性":5,
         "設備管理者":"設備組","外包商承攬商":"台灣約克","相關變數":"冷卻水塔","外觀照片":None,"銘牌照片":None},
        {"系統別":"空壓系統","設備名稱":"無油式壓縮機","設備編號":"A0154","設備型式":"螺旋式空壓機",
         "設備部門":"製程部","所在棟別":"一廠","所在樓層":"B1",
         "消耗功率(kW)":3.725,"設備數量":1,"負載率":0.90,"運轉時數(hr/年)":290,
         "設備年份":2018,"使用年數":7,"自評重大性":3,
         "設備管理者":"設備組","外包商承攬商":"","相關變數":"","外觀照片":None,"銘牌照片":None},
        {"系統別":"空壓系統","設備名稱":"活塞式空壓機","設備編號":"A0471","設備型式":"活塞式",
         "設備部門":"製程部","所在棟別":"三廠","所在樓層":"1F",
         "消耗功率(kW)":5.0,"設備數量":1,"負載率":0.90,"運轉時數(hr/年)":1100,
         "設備年份":2005,"使用年數":20,"自評重大性":4,
         "設備管理者":"設備組","外包商承攬商":"","相關變數":"","外觀照片":None,"銘牌照片":None},
        {"系統別":"製程系統","設備名稱":"製程攪拌機","設備編號":"A0436","設備型式":"立式攪拌機",
         "設備部門":"製程一部","所在棟別":"一廠","所在樓層":"2F",
         "消耗功率(kW)":4.0,"設備數量":1,"負載率":0.45,"運轉時數(hr/年)":1725,
         "設備年份":2012,"使用年數":13,"自評重大性":3,
         "設備管理者":"製程組","外包商承攬商":"","相關變數":"製程溫度","外觀照片":None,"銘牌照片":None},
        {"系統別":"製程系統","設備名稱":"反應槽加熱器","設備編號":"HT-001","設備型式":"電熱式",
         "設備部門":"製程一部","所在棟別":"一廠","所在樓層":"2F",
         "消耗功率(kW)":15.0,"設備數量":3,"負載率":0.80,"運轉時數(hr/年)":4500,
         "設備年份":2008,"使用年數":17,"自評重大性":5,
         "設備管理者":"製程組","外包商承攬商":"","相關變數":"反應溫度","外觀照片":None,"銘牌照片":None},
        {"系統別":"照明系統","設備名稱":"LED燈管(廠區)","設備編號":"LT-001","設備型式":"LED T8",
         "設備部門":"總務部","所在棟別":"一廠","所在樓層":"全棟",
         "消耗功率(kW)":0.018,"設備數量":200,"負載率":1.0,"運轉時數(hr/年)":3000,
         "設備年份":2020,"使用年數":5,"自評重大性":2,
         "設備管理者":"總務組","外包商承攬商":"","相關變數":"","外觀照片":None,"銘牌照片":None},
        {"系統別":"照明系統","設備名稱":"日光燈(辦公室)","設備編號":"LT-002","設備型式":"T5 螢光燈",
         "設備部門":"管理部","所在棟別":"辦公室","所在樓層":"3F",
         "消耗功率(kW)":0.028,"設備數量":50,"負載率":1.0,"運轉時數(hr/年)":2500,
         "設備年份":2012,"使用年數":13,"自評重大性":2,
         "設備管理者":"總務組","外包商承攬商":"","相關變數":"","外觀照片":None,"銘牌照片":None},
        {"系統別":"其他系統","設備名稱":"恆溫恆濕試驗機","設備編號":"F-029","設備型式":"恆溫恆濕箱",
         "設備部門":"品管部","所在棟別":"二廠","所在樓層":"2F",
         "消耗功率(kW)":2.6,"設備數量":1,"負載率":0.80,"運轉時數(hr/年)":3300,
         "設備年份":2016,"使用年數":9,"自評重大性":3,
         "設備管理者":"品管組","外包商承攬商":"","相關變數":"溫濕度","外觀照片":None,"銘牌照片":None},
        {"系統別":"其他系統","設備名稱":"廢水處理泵浦","設備編號":"WP-001","設備型式":"離心泵",
         "設備部門":"環安部","所在棟別":"廠外","所在樓層":"B1",
         "消耗功率(kW)":7.5,"設備數量":2,"負載率":0.75,"運轉時數(hr/年)":5000,
         "設備年份":2009,"使用年數":16,"自評重大性":4,
         "設備管理者":"環安組","外包商承攬商":"","相關變數":"廢水量","外觀照片":None,"銘牌照片":None},
    ]

def init_from_excel():
    if os.path.exists(EXCEL_FILE):
        # 本機環境：讀取真實 Excel
        recs = []
        for s, sh in SYSTEM_SHEETS.items():
            recs.extend(read_system(sh, s))
        return recs if recs else get_builtin_data()
    else:
        # 雲端環境：使用內建示範資料
        return get_builtin_data()

# ─────────────────────────────────────────────────────────────────────────────
# 5. Session 初始化
# ─────────────────────────────────────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "edit_mode" not in st.session_state:
    st.session_state["edit_mode"] = False
if "db" not in st.session_state:
    saved = load_json()
    st.session_state["db"] = saved if saved else init_from_excel()

def all_calc():
    rows = []
    for rec in st.session_state["db"]:
        r = dict(rec)
        kwh, sc, seu = calc_row(r)
        r.update({"_kwh": kwh, "_sc": sc, "_seu": seu})
        rows.append(r)
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# 6. Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center;padding:14px 0 10px'>
      <div style='font-size:34px'>⚡</div>
      <div style='font-size:16px;font-weight:800;color:#fff'>永寬化學</div>
      <div style='font-size:11px;color:#94a3b8;margin-top:2px'>ISO 50001 能源控制台</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.markdown("**🔐 系統權限**")
    if not st.session_state["logged_in"]:
        st.caption("預設唯讀。輸入密碼解鎖修改功能。")
        pwd = st.text_input("管理員密碼", type="password", key="pwd")
        if st.button("🔓 驗證並解鎖", use_container_width=True):
            if pwd == ADMIN_PASSWORD:
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("密碼錯誤！")
    else:
        st.success("✅ 高級系統管理員")
        if st.button("🔒 登出並鎖定", use_container_width=True):
            st.session_state.update(logged_in=False, edit_mode=False)
            st.rerun()

    st.divider()

    if st.session_state["logged_in"]:
        mode = st.radio(
            "操作模式",
            ["觀看模式（唯讀）", "修改模式（開放編輯）"],
            index=0
        )
        st.session_state["edit_mode"] = "修改" in mode
    else:
        st.info("👁️ 唯讀保護中")
        st.session_state["edit_mode"] = False

    st.divider()
    st.markdown("**📋 功能選單**")
    menu = st.radio("", [
        "全廠能耗儀表板",
        "設備盤查與照片管理",
        "評分標準說明",
        "每日負載分析",
        "從Excel重新載入",
    ], label_visibility="collapsed")

    st.divider()
    db_count = len(st.session_state["db"])
    st.caption(f"資料庫：{db_count} 台設備")
    st.caption(f"更新：{datetime.now().strftime('%Y/%m/%d %H:%M')}")

# ─────────────────────────────────────────────────────────────────────────────
# 7. 頂部標題 + 模式橫幅
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#f8fafc;padding:16px 22px;border-radius:10px;
            border-left:5px solid #1a3a5c;margin-bottom:18px'>
  <h1 style='margin:0;color:#1a3a5c;font-size:22px'>永寬化學股份有限公司</h1>
  <h3 style='margin:5px 0 0;color:#475569;font-size:14px;font-weight:normal'>
    重大能源使用設備 (SEUs) 網頁版管理系統
  </h3>
  <p style='margin:6px 0 0;color:#64748b;font-size:11px;line-height:1.6'>
    系統邊界：雲林縣斗六市榴南里　｜　樓地板面積：12,378.27 ㎡　｜　ISO 50001:2018
  </p>
</div>
""", unsafe_allow_html=True)

if st.session_state["edit_mode"]:
    st.markdown('<div class="mode-edit">✏️ <b>修改模式已啟用</b>：可新增設備、編輯數據、上傳照片。</div>',
                unsafe_allow_html=True)
else:
    st.markdown('<div class="mode-view">👁️ <b>唯讀觀看模式</b>：所有編輯功能已鎖定。管理員由左側輸入密碼解鎖。</div>',
                unsafe_allow_html=True)
st.markdown("")

# ─────────────────────────────────────────────────────────────────────────────
# Excel 檔案存在檢查
# ─────────────────────────────────────────────────────────────────────────────
if not os.path.exists(EXCEL_FILE) and len(st.session_state["db"]) == 0:
    st.warning("⚠️ 未偵測到 Excel 檔案，已載入內建示範資料。如需載入完整資料請將 Excel 放於同一資料夾。")

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面一：全廠能耗儀表板
# ─────────────────────────────────────────────────────────────────────────────
if "儀表板" in menu:
    rows = all_calc()

    if len(rows) == 0:
        st.warning("⚠️ 資料庫是空的，請前往「🔄 從Excel重新載入」重新載入資料。")
        st.stop()

    tot_kwh = sum(r["_kwh"] for r in rows)
    tot_a   = sum(1 for r in rows if r["_seu"] == "A")
    cov     = round(tot_kwh / TOTAL_KWH * 100, 1) if TOTAL_KWH else 0
    eui     = round(TOTAL_KWH / FLOOR_AREA, 2)

    # KPI 卡片 第一列
    k1, k2, k3 = st.columns(3)
    for col, val, lbl, color in [
        (k1, f"{TOTAL_KWH:,}", "全廠年實際總用電 (kWh)", "#1a3a5c"),
        (k2, str(eui),          "EUI 能源強度 (度/㎡·年)", "#2563a8"),
        (k3, str(len(rows)),    "已盤查設備總數（台）",    "#00c896"),
    ]:
        col.markdown(
            f'<div class="kpi"><div class="kpi-v" style="color:{color}">{val}</div>'
            f'<div class="kpi-l">{lbl}</div></div>',
            unsafe_allow_html=True
        )
    st.markdown("<div style='margin:8px 0'></div>", unsafe_allow_html=True)
    # KPI 卡片 第二列
    k4, k5, k6 = st.columns(3)
    for col, val, lbl, color in [
        (k4, str(tot_a),  "A 級重大耗能設備（台）", "#f59e0b"),
        (k5, f"{cov}%",   "盤查耗電覆蓋率（估算）", "#8b5cf6"),
        (k6, f"{FLOOR_AREA:,.0f} ㎡", "廠區樓地板面積",    "#64748b"),
    ]:
        col.markdown(
            f'<div class="kpi"><div class="kpi-v" style="color:{color}">{val}</div>'
            f'<div class="kpi-l">{lbl}</div></div>',
            unsafe_allow_html=True
        )
    st.markdown("<br>", unsafe_allow_html=True)

    # 系統彙總
    sys_agg = {}
    for r in rows:
        s = r.get("系統別", "其他")
        sys_agg.setdefault(s, {"kwh": 0, "a": 0, "n": 0})
        sys_agg[s]["kwh"] += r["_kwh"]
        sys_agg[s]["a"]   += 1 if r["_seu"] == "A" else 0
        sys_agg[s]["n"]   += 1

    df_sys = pd.DataFrame([
        {
            "系統別":       s,
            "耗電量(kWh/年)": round(v["kwh"], 0),
            "佔比(%)":      round(v["kwh"] / tot_kwh * 100, 1) if tot_kwh > 0 else 0,
            "設備數":       v["n"],
            "A級設備":      v["a"],
        }
        for s, v in sys_agg.items()
        if v["kwh"] > 0   # ← 只顯示有耗電量的系統
    ])

    ch1, ch2 = st.columns([1, 1])
    with ch1:
        if len(df_sys) > 0:
            fig_pie = px.pie(
                df_sys,
                values="耗電量(kWh/年)",
                names="系統別",
                hole=0.40,
                color_discrete_sequence=["#1a3a5c","#2563a8","#00c896","#f59e0b","#8b5cf6"]
            )
            fig_pie.update_traces(
                textinfo="percent",
                textposition="inside",
                hovertemplate="<b>%{label}</b><br>%{value:,.0f} kWh<br>%{percent}<extra></extra>"
            )
            fig_pie.update_layout(
                title=dict(text="各系統能耗佔比", x=0.5, font=dict(size=15)),
                legend=dict(
                    orientation="v",
                    yanchor="middle", y=0.5,
                    xanchor="left", x=1.02,
                    font=dict(size=13)
                ),
                margin=dict(t=50, b=20, l=20, r=120),
                height=380,
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("無耗電量資料，無法顯示圓餅圖。")

    with ch2:
        if len(df_sys) > 0:
            fig_bar = go.Figure(go.Bar(
                x=df_sys["系統別"],
                y=df_sys["耗電量(kWh/年)"],
                text=df_sys["耗電量(kWh/年)"].apply(lambda v: f"{v:,.0f}"),
                textposition="outside",
                marker_color=["#1a3a5c","#2563a8","#00c896","#f59e0b","#8b5cf6"],
                width=0.5,
            ))
            fig_bar.update_layout(
                title=dict(text="各系統年耗電量 (kWh)", x=0.5, font=dict(size=15)),
                height=380,
                margin=dict(t=50, b=60, l=60, r=20),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis_tickformat=",",
                xaxis=dict(tickfont=dict(size=13)),
                yaxis=dict(tickfont=dict(size=12)),
                bargap=0.3,
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()
    st.subheader("📋 各系統耗能摘要")
    if len(df_sys) > 0:
        df_show = df_sys.copy()
        df_show["耗電量(kWh/年)"] = df_show["耗電量(kWh/年)"].apply(lambda v: f"{v:,.0f}")
        df_show["佔比(%)"]        = df_show["佔比(%)"].apply(lambda v: f"{v:.1f}%")
        st.dataframe(df_show, hide_index=True, use_container_width=True)

    # A 級設備清單
    a_rows = sorted([r for r in rows if r["_seu"] == "A"],
                    key=lambda r: r["_kwh"], reverse=True)
    if a_rows:
        st.divider()
        st.subheader("⭐ A 級重大耗能設備（依耗電量排序）")
        st.dataframe(pd.DataFrame([{
            "系統":         r.get("系統別", ""),
            "設備名稱":     r.get("設備名稱", ""),
            "編號":         r.get("設備編號", ""),
            "部門":         r.get("設備部門", ""),
            "功率(kW)":    r.get("消耗功率(kW)", ""),
            "年耗電(kWh)": f"{r['_kwh']:,.0f}",
            "重大性評分":  r["_sc"],
            "管理者":       r.get("設備管理者", ""),
        } for r in a_rows]), hide_index=True, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面二：設備盤查與照片管理
# ─────────────────────────────────────────────────────────────────────────────
elif "設備盤查" in menu:

    # 新增設備表單（修改模式）
    if st.session_state["edit_mode"]:
        with st.expander("➕ 新增設備（展開填寫）", expanded=False):
            with st.form("add_form", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    in_sys  = st.selectbox("所屬系統", list(SYSTEM_SHEETS.keys()))
                    in_name = st.text_input("設備名稱 *")
                    in_id   = st.text_input("設備編號 *")
                    in_type = st.text_input("型式/說明")
                with c2:
                    in_dept = st.text_input("設備部門")
                    in_bldg = st.text_input("所在棟別")
                    in_kw   = st.number_input("消耗功率 (kW)", min_value=0.0, value=5.0, step=0.1)
                    in_qty  = st.number_input("設備數量 (台)", min_value=1, value=1)
                with c3:
                    in_load = st.slider("負載率", 0.1, 1.0, 0.85, 0.05)
                    in_hrs  = st.number_input("年運轉時數 (hr)", min_value=0.0, value=2000.0)
                    in_yr   = st.number_input("設備年份", 1990, 2030, 2015)
                    in_crit = st.slider("自評重大性 (1~5)", 1, 5, 3)
                pic1 = st.file_uploader("📷 設備外觀照片", type=["jpg","jpeg","png"])
                pic2 = st.file_uploader("🏷️ 銘牌照片",     type=["jpg","jpeg","png"])

                if st.form_submit_button("💾 提交寫入資料庫", use_container_width=True):
                    if not in_name or not in_id:
                        st.error("設備名稱與編號為必填！")
                    else:
                        new = {
                            "系統別": in_sys, "設備名稱": in_name, "設備編號": in_id,
                            "設備型式": in_type, "設備部門": in_dept, "所在棟別": in_bldg,
                            "消耗功率(kW)": in_kw, "設備數量": in_qty, "負載率": in_load,
                            "運轉時數(hr/年)": in_hrs, "設備年份": in_yr,
                            "使用年數": datetime.now().year - int(in_yr),
                            "自評重大性": in_crit,
                            "外觀照片": base64.b64encode(pic1.read()).decode() if pic1 else None,
                            "銘牌照片": base64.b64encode(pic2.read()).decode() if pic2 else None,
                        }
                        st.session_state["db"].append(new)
                        save_json(st.session_state["db"])
                        st.success(f"✅ 設備【{in_name}】已寫入！")
                        st.rerun()

    # 篩選列
    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1:
        kw_f  = st.text_input("🔍 搜尋設備名稱 / 編號 / 部門")
    with fc2:
        sys_f = st.selectbox("系統篩選", ["全部"] + list(SYSTEM_SHEETS.keys()))
    with fc3:
        seu_f = st.selectbox("重大性篩選", ["全部", "A 級重大設備", "一般設備"])

    rows     = all_calc()
    filtered = []
    for r in rows:
        if sys_f != "全部" and r.get("系統別") != sys_f:
            continue
        if seu_f == "A 級重大設備" and r["_seu"] != "A":
            continue
        if seu_f == "一般設備" and r["_seu"] == "A":
            continue
        if kw_f:
            s = f"{r.get('設備名稱','')} {r.get('設備編號','')} {r.get('設備部門','')}".lower()
            if kw_f.lower() not in s:
                continue
        filtered.append(r)

    st.caption(f"顯示 **{len(filtered)}** 筆（共 {len(rows)} 筆）")

    for loop_idx, r in enumerate(filtered):
        db_idx = next(
            (i for i, d in enumerate(st.session_state["db"])
             if str(d.get("設備名稱","")) == str(r.get("設備名稱",""))
             and str(d.get("設備編號","")) == str(r.get("設備編號",""))
             and str(d.get("系統別",""))   == str(r.get("系統別",""))),
            None
        )

        icon  = SYSTEM_ICONS.get(r.get("系統別",""), "🔧")
        a_tag = " ⭐A級" if r["_seu"] == "A" else ""
        title = (f"{icon} [{r.get('系統別','')}]  "
                 f"{r.get('設備名稱','')} ({r.get('設備編號','')})"
                 f"  ｜  {r['_kwh']:,.0f} kWh/年"
                 f"  評分 {r['_sc']}{a_tag}")

        with st.expander(title, expanded=False):
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("消耗功率",   f"{r.get('消耗功率(kW)','')} kW")
            d2.metric("年運轉時數", f"{float(r.get('運轉時數(hr/年)') or 0):,.0f} hr")
            d3.metric("使用年數",   f"{int(r.get('使用年數') or 0)} 年")
            d4.metric("重大性評分", r["_sc"])

            col_info, col_photo = st.columns([1, 1])
            with col_info:
                load_pct = f"{float(r.get('負載率',0))*100:.0f}%" if r.get('負載率') else "—"
                st.markdown(f"""
| 欄位 | 資料 |
|------|------|
| 部門 | {r.get('設備部門','—')} |
| 棟別/樓層 | {r.get('所在棟別','—')} / {r.get('所在樓層','—')} |
| 型式說明 | {r.get('設備型式','—')} |
| 數量 | {r.get('設備數量','—')} 台 |
| 負載率 | {load_pct} |
| 設備年份 | {r.get('設備年份','—')} |
| 管理者 | {r.get('設備管理者','—')} |
| 外包商 | {r.get('外包商承攬商','—')} |
| 相關變數 | {r.get('相關變數','—')} |
| **年耗電量** | **{r['_kwh']:,.0f} kWh** |
| **SEU 鑑別** | **{'⭐ A 級重大設備' if r['_seu']=='A' else '一般設備'}** |
""")
            with col_photo:
                st.markdown("**📷 設備影像**")
                ph1, ph2 = st.columns(2)
                with ph1:
                    st.caption("外觀照片")
                    if r.get("外觀照片"):
                        try:
                            st.image(Image.open(BytesIO(base64.b64decode(r["外觀照片"]))),
                                     use_container_width=True)
                        except:
                            st.warning("照片解碼失敗")
                    else:
                        st.info("未上傳")
                with ph2:
                    st.caption("銘牌照片")
                    if r.get("銘牌照片"):
                        try:
                            st.image(Image.open(BytesIO(base64.b64decode(r["銘牌照片"]))),
                                     use_container_width=True)
                        except:
                            st.warning("照片解碼失敗")
                    else:
                        st.info("未上傳")

            # 行內編輯（修改模式）
            if st.session_state["edit_mode"] and db_idx is not None:
                st.markdown("---")
                st.markdown("**✏️ 編輯此設備數據**")
                cur = st.session_state["db"][db_idx]
                with st.form(f"ef_{loop_idx}_{db_idx}"):
                    e1, e2, e3 = st.columns(3)
                    with e1:
                        e_name = st.text_input("設備名稱", value=str(cur.get("設備名稱","") or ""))
                        e_id   = st.text_input("設備編號", value=str(cur.get("設備編號","") or ""))
                        e_dept = st.text_input("部門",     value=str(cur.get("設備部門","") or ""))
                    with e2:
                        e_kw   = st.number_input("消耗功率(kW)",
                                    value=float(cur.get("消耗功率(kW)") or 0),
                                    min_value=0.0, step=0.1)
                        e_qty  = st.number_input("設備數量",
                                    value=int(cur.get("設備數量") or 1),
                                    min_value=1)
                        e_load = st.slider("負載率", 0.1, 1.0,
                                    float(cur.get("負載率") or 0.8), 0.05)
                    with e3:
                        e_hrs  = st.number_input("年運轉時數",
                                    value=float(cur.get("運轉時數(hr/年)") or 0),
                                    min_value=0.0)
                        e_crit = st.slider("自評重大性(1~5)", 1, 5,
                                    int(cur.get("自評重大性") or 3))
                        e_mgr  = st.text_input("設備管理者",
                                    value=str(cur.get("設備管理者","") or ""))

                    up1 = st.file_uploader("更新外觀照片（留空保留原圖）",
                                type=["jpg","jpeg","png"], key=f"u1_{loop_idx}_{db_idx}")
                    up2 = st.file_uploader("更新銘牌照片（留空保留原圖）",
                                type=["jpg","jpeg","png"], key=f"u2_{loop_idx}_{db_idx}")

                    sv, dl = st.columns([3, 1])
                    with sv:
                        save_ok = st.form_submit_button("💾 儲存變更", use_container_width=True)
                    with dl:
                        del_ok  = st.form_submit_button("🗑️ 刪除設備", use_container_width=True)

                    if save_ok:
                        st.session_state["db"][db_idx].update({
                            "設備名稱": e_name, "設備編號": e_id, "設備部門": e_dept,
                            "消耗功率(kW)": e_kw, "設備數量": e_qty, "負載率": e_load,
                            "運轉時數(hr/年)": e_hrs, "自評重大性": e_crit,
                            "設備管理者": e_mgr,
                        })
                        if up1:
                            st.session_state["db"][db_idx]["外觀照片"] = \
                                base64.b64encode(up1.read()).decode()
                        if up2:
                            st.session_state["db"][db_idx]["銘牌照片"] = \
                                base64.b64encode(up2.read()).decode()
                        save_json(st.session_state["db"])
                        st.success("✅ 已儲存！")
                        st.rerun()

                    if del_ok:
                        st.session_state["db"].pop(db_idx)
                        save_json(st.session_state["db"])
                        st.warning("已刪除。")
                        st.rerun()

    # 匯出 CSV
    st.divider()
    if filtered:
        exp_data = [
            {k: v for k, v in r.items()
             if k not in ("外觀照片","銘牌照片","_kwh","_sc","_seu")}
            for r in filtered
        ]
        csv = pd.DataFrame(exp_data).to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ 匯出篩選結果 CSV", csv,
            f"SEU_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv"
        )


# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面：評分標準說明
# ─────────────────────────────────────────────────────────────────────────────
elif "評分標準" in menu:
    st.subheader("📐 ISO 50001 重大能源使用設備評分標準")

    tab1, tab2 = st.tabs(["重大能源使用鑑別（A級）", "優先改善項目鑑別（I級）"])

    with tab1:
        st.markdown("#### 鑑別因子與權重")
        df_w1 = pd.DataFrame({
            "鑑別因子": ["設備耗能估比", "工廠自評重大性（設備管控評估）", "總計"],
            "估比":     ["50%", "50%", "100%"],
        })
        st.dataframe(df_w1, hide_index=True, use_container_width=False)

        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("##### 設備耗能估比評分")
            df_s1 = pd.DataFrame({
                "耗能估比範圍": ["— ～ 0.24%", "0.25% ～ 0.49%", "0.50% ～ 0.74%", "0.75% ～", "1.00% ～"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_s1, hide_index=True, use_container_width=True)

        with col2:
            st.markdown("##### 工廠自評重大性評分")
            df_s2 = pd.DataFrame({
                "評估等級": ["— ～ 1", "2 ～ 2", "3 ～ 3", "4 ～ 4", "5 ～ 5"],
                "說明": ["非重要管控項目", "", "需再評估", "", "既有或應該列入管控"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_s2, hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 重大能源使用鑑別級距")
        df_seu = pd.DataFrame({
            "評分範圍": ["0 ～ 1.9 分", "2 ～ 2.9 分", "3 ～ 3.9 分", "4 ～ 5.0 分"],
            "級別": ["-", "-", "-", "A"],
            "說明": ["一般設備", "一般設備", "一般設備", "重大能源使用設備（SEU）"],
        })
        st.dataframe(df_seu, hide_index=True, use_container_width=True)

        st.info("**計算公式：** 重大性評分 = 設備耗能估比分數 × 50% ＋ 工廠自評重大性分數 × 50%\n\n評分 ≥ 4.0 分 → 鑑別為 **A 級重大能源使用設備（SEU）**，需研提能源管理行動計畫並制訂操作規範。\n\n> 📝 備註：管控可為 SOP 或即時監控")

    with tab2:
        st.markdown("#### 鑑別因子與權重")
        df_w2 = pd.DataFrame({
            "鑑別因子":     ["設備耗能估比", "設備老舊度", "設備運轉度", "能效改善頻率", "改善執行難易度", "總計"],
            "估比":         ["15%", "30%", "5%", "20%", "30%", "100%"],
        })
        st.dataframe(df_w2, hide_index=True, use_container_width=False)

        st.markdown("<br>", unsafe_allow_html=True)
        col3, col4 = st.columns(2)

        with col3:
            st.markdown("##### 設備耗能估比評分")
            df_p1 = pd.DataFrame({
                "耗能估比範圍": ["0% ～ 0.1%", "0.1% ～ 0.1%", "0.1% ～ 0.4%", "0.5% ～ 1.0%", "1.0% ～"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_p1, hide_index=True, use_container_width=True)

            st.markdown("##### 設備老舊度評分")
            df_p2 = pd.DataFrame({
                "使用年數": ["0 ～ 4 年", "5 ～ 9 年", "10 ～ 14 年", "15 ～ 19 年", "20 年以上"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_p2, hide_index=True, use_container_width=True)

            st.markdown("##### 設備運轉度評分")
            df_p3 = pd.DataFrame({
                "年運轉時數": ["0 ～ 1,460 小時", "1,461 ～ 2,920 小時", "2,921 ～ 4,380 小時", "4,381 ～ 5,840 小時", "5,841 ～ 8,760 小時"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_p3, hide_index=True, use_container_width=True)

        with col4:
            st.markdown("##### 能效改善頻率評分")
            df_p4 = pd.DataFrame({
                "改善頻率": ["# ～ 1（5年內新機）", "2 ～ 2", "3 ～ 3（10年以上能效改善1次）", "4 ～ 4", "5 ～ 5（10年以上從未改善）"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_p4, hide_index=True, use_container_width=True)

            st.markdown("##### 改善執行難易度評分")
            df_p5 = pd.DataFrame({
                "難易度": ["# ～ 1（不會改善）", "2 ～ 2", "3 ～ 3（需再評估）", "4 ～ 4", "5 ～ 5（可立即改善）"],
                "分數": [1, 2, 3, 4, 5],
            })
            st.dataframe(df_p5, hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 優先改善項目鑑別級距")
        df_pri = pd.DataFrame({
            "評分範圍": ["0 ～ 1.9 分", "2 ～ 2.9 分", "3 ～ 3.9 分", "4 ～ 5.0 分"],
            "級別": ["-", "-", "-", "I"],
            "說明": ["一般設備", "一般設備", "一般設備", "優先改善項目"],
        })
        st.dataframe(df_pri, hide_index=True, use_container_width=True)

        st.info("**計算公式：** 優先改善評分 = 耗能估比分數 × 15% ＋ 老舊度分數 × 30% ＋ 運轉度分數 × 5% ＋ 改善頻率分數 × 20% ＋ 改善難易度分數 × 30%\n\n評分 ≥ 4.0 分 → 鑑別為 **I 級優先改善項目**，需優先執行能效改善作業。")

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面三：每日負載分析
# ─────────────────────────────────────────────────────────────────────────────
elif "負載" in menu:
    st.subheader("📈 全廠 24 小時電力負載曲線")
    st.info("統計區間：2024 年 6/11 – 10/18（共 95 天尖峰期）")

    df_load = pd.DataFrame({
        "時間(時)": list(range(1, 25)),
        "最高用電(kW)": [135,130,127,126,125,128,194,278,362,384,388,399,
                         377,386,389,381,372,302,226,158,142,141,136,135],
        "最低用電(kW)": [105,101,97,97,96,98,183,245,309,333,337,343,
                         319,329,333,325,316,257,196,120,110,108,104,103],
    })
    df_load["平均用電(kW)"] = (
        (df_load["最高用電(kW)"] + df_load["最低用電(kW)"]) / 2
    ).round(1)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_load["時間(時)"], y=df_load["最高用電(kW)"],
        name="最高", mode="lines+markers",
        line=dict(color="#ef4444", width=2)
    ))
    fig.add_trace(go.Scatter(
        x=df_load["時間(時)"], y=df_load["平均用電(kW)"],
        name="平均", mode="lines+markers",
        line=dict(color="#2563a8", width=2, dash="dash")
    ))
    fig.add_trace(go.Scatter(
        x=df_load["時間(時)"], y=df_load["最低用電(kW)"],
        name="最低", mode="lines+markers",
        line=dict(color="#22c55e", width=2),
        fill="tonexty", fillcolor="rgba(37,99,168,.06)"
    ))
    fig.update_layout(
        title="廠區 24 小時尖離峰電力負載分佈",
        xaxis=dict(title="時間（時）", tickmode="linear", tick0=1, dtick=1),
        yaxis_title="電力負載 (kW)",
        height=420,
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True)

    pk = df_load.loc[df_load["最高用電(kW)"].idxmax()]
    of = df_load.loc[df_load["最低用電(kW)"].idxmin()]
    pa, pb, pc = st.columns(3)
    pa.metric("🔺 尖峰時段", f"{int(pk['時間(時)'])}:00", f"{pk['最高用電(kW)']} kW")
    pb.metric("🔻 離峰時段", f"{int(of['時間(時)'])}:00", f"{of['最低用電(kW)']} kW")
    pc.metric("📊 負載差異",
              f"{pk['最高用電(kW)'] - of['最低用電(kW)']} kW",
              f"比值 {round(pk['最高用電(kW)']/of['最低用電(kW)'],2):.2f}x")

    st.divider()
    st.dataframe(df_load, hide_index=True, use_container_width=True)
    st.download_button(
        "⬇️ 下載負載曲線 CSV",
        df_load.to_csv(index=False).encode("utf-8-sig"),
        "YuanKuan_24H_Load.csv", "text/csv"
    )

# ─────────────────────────────────────────────────────────────────────────────
# ── 頁面四：從 Excel 重新載入
# ─────────────────────────────────────────────────────────────────────────────
elif "Excel" in menu:
    st.subheader("🔄 從 Excel 重新載入資料")
    st.warning("⚠️ 此操作將覆蓋所有已修改的設備資料與照片，請謹慎！")

    if not os.path.exists(EXCEL_FILE):
        st.warning(f"⚠️ 未偵測到 Excel 檔案（雲端環境），目前使用內建示範資料。")
        if st.session_state["edit_mode"]:
            if st.button("🔄 載入內建示範資料", type="primary"):
                st.session_state["db"] = get_builtin_data()
                save_json(st.session_state["db"])
                st.success(f"✅ 已載入內建資料！共 {len(st.session_state['db'])} 筆。")
                st.rerun()
        else:
            st.info("請切換至「✏️ 修改模式」才能執行重新載入。")
    else:
        st.success(f"✅ 找到 Excel 檔案：{EXCEL_FILE}")
        if st.session_state["edit_mode"]:
            if st.button("🔄 確認重新載入 Excel", type="primary"):
                st.session_state["db"] = init_from_excel()
                save_json(st.session_state["db"])
                st.success(f"✅ 已重新載入！共 {len(st.session_state['db'])} 筆。")
                st.rerun()
        else:
            st.info("請切換至「✏️ 修改模式」才能執行重新載入。")

    st.divider()
    st.markdown(f"目前資料庫共 **{len(st.session_state['db'])}** 筆設備")
    st.markdown(f"本地存檔：`{DB_JSON}`（{'✅ 已存在' if os.path.exists(DB_JSON) else '⚠️ 尚未建立'}）")
