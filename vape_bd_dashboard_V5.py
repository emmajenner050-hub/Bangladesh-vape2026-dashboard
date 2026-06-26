import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re

st.set_page_config(
    page_title="孟加拉国电子烟市场调研数据看板",
    page_icon="🇧🇩",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 汇率常量：1 孟加拉塔卡 = 0.056 人民币
EXCHANGE_RATE = 0.056

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #006A4E 0%, #004d38 100%);
        padding: 16px 20px; border-radius: 12px; color: white;
        text-align: center; box-shadow: 0 4px 15px rgba(0,106,78,0.25);
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; line-height: 1; }
    .metric-card .label { font-size: 0.8rem; opacity: 0.9; margin-top: 6px; }
    .metric-card-red {
        background: linear-gradient(135deg, #c0392b 0%, #922b21 100%);
        padding: 16px 20px; border-radius: 12px; color: white;
        text-align: center; box-shadow: 0 4px 15px rgba(192,57,43,0.25);
    }
    .metric-card-red .value { font-size: 2rem; font-weight: 700; line-height: 1; }
    .metric-card-red .label { font-size: 0.8rem; opacity: 0.9; margin-top: 6px; }
    .section-title {
        font-size: 1.05rem; font-weight: 600; color: #1e3a2f;
        margin: 18px 0 10px 0; padding-left: 10px;
        border-left: 4px solid #006A4E;
    }
    .insight-box {
        background: #e8f5f0; border-left: 4px solid #006A4E;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .warning-box {
        background: #fef3c7; border-left: 4px solid #d97706;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .danger-box {
        background: #fef2f2; border-left: 4px solid #ef4444;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .action-box {
        background: #f0fdf4; border-left: 4px solid #10b981;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .tag-chip {
        display: inline-block; background: #d1fae5; color: #065f46;
        padding: 2px 10px; border-radius: 999px;
        font-size: 0.78rem; margin: 2px;
    }
    .flavor-row {
        margin: 3px 0 6px 0; padding: 6px 10px;
        background: #f8fafc; border-radius: 6px;
        font-size: 0.85rem; color: #374151;
    }
    .summary-card {
        background: #e8f5f0; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 6px solid #006A4E;
    }
    .warning-card {
        background: #fff7ed; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 6px solid #f97316;
    }
    .danger-card {
        background: #fef2f2; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 6px solid #ef4444;
    }
    .reg-timeline {
        border-left: 3px solid #006A4E; padding-left: 16px; margin: 8px 0;
    }
    .reg-step {
        margin-bottom: 14px; position: relative;
    }
    .reg-step::before {
        content: ''; position: absolute; left: -22px; top: 4px;
        width: 12px; height: 12px; border-radius: 50%;
        background: #006A4E;
    }
    .reg-step-red::before {
        background: #ef4444 !important;
    }
    .reg-step-yellow::before {
        background: #f59e0b !important;
    }
</style>
""", unsafe_allow_html=True)

COLORS = {
    "primary":   "#006A4E",
    "secondary": "#004d38",
    "red":       "#F42A41",
    "china":     "#ef4444",
    "us":        "#3b82f6",
    "uk":        "#8b5cf6",
    "malaysia":  "#f59e0b",
    "japan":     "#ec4899",
    "ice":       "#38bdf8",
    "tobacco":   "#92400e",
    "fruit":     "#22c55e",
    "sweet":     "#ec4899",
    "drink":     "#06b6d4",
    "candy":     "#8b5cf6",
    "menthol":   "#34d399",
    "dessert":   "#f97316",
    "other":     "#94a3b8",
    "elfbar":    "#006A4E",
    "voltbar":   "#F42A41",
}
CAT_COLOR_MAP = {
    "水果": COLORS["fruit"],
    "烟草": COLORS["tobacco"],
    "甜点": COLORS["dessert"],
    "饮料": COLORS["drink"],
    "糖果": COLORS["candy"],
    "薄荷": COLORS["menthol"],
    "其他": COLORS["other"],
}
COUNTRY_COLOR_MAP = {
    "中国":   COLORS["china"],
    "美国":   COLORS["us"],
    "英国":   COLORS["uk"],
    "马来西亚": COLORS["malaysia"],
    "日本":   COLORS["japan"],
}
T = "plotly_white"

import os
_DIR = os.path.dirname(os.path.abspath(__file__))
# ⚠️ 请根据实际文件名修改下面的文件名
# 原文件名为：孟加拉数据梳理—Erin—20260624.xlsx（中文破折号）
# 代码中使用下划线，您可以将 Excel 重命名，或修改下方变量
EXCEL_FILE = os.path.join(_DIR, "孟加拉数据梳理_Erin_20260624.xlsx")


def parse_price(p):
    if pd.isna(p):
        return None
    s = str(p).replace("\xa0", "").replace("৳", "").replace(",", "").strip()
    m = re.search(r"[\d\.]+", s)
    if m:
        try:
            return float(m.group())
        except:
            return None
    return None


def split_tags(tag_str):
    if pd.isna(tag_str):
        return []
    return [p.strip().lower() for p in str(tag_str).split(",") if p.strip()]


def flavor_complexity(n):
    if n <= 1:
        return "单一口味"
    elif n == 2:
        return "双重复合"
    else:
        return "三重以上复合"


@st.cache_data
def load_data():
    xl = pd.ExcelFile(EXCEL_FILE)

    # ---- popular sell (主数据库) ----
    ps = pd.read_excel(xl, sheet_name="popular sell")
    ps.columns = ps.columns.str.strip()
    ps["产品类型"] = ps["产品类型"].str.strip()
    ps["产品类型"] = ps["产品类型"].replace({"E-Liquid": "E-liquid"})
    ps["含冰/薄荷"] = ps["含冰/薄荷"].fillna("不确定")
    ps["含烟草"] = ps["含烟草"].fillna("否")
    ps["口味标签列表"] = ps["口味标签"].apply(split_tags)
    ps["价格_数值"] = ps["价格"].apply(parse_price)
    ps["口味标签数"] = ps["口味标签列表"].apply(len)
    ps["口味复杂度"] = ps["口味标签数"].apply(flavor_complexity)

    def nic_list(v):
        if pd.isna(v):
            return []
        nums = re.findall(r"\d+", str(v))
        return [int(x) for x in nums]

    ps["尼古丁档位列表"] = ps["尼古丁浓度mg"].apply(nic_list)
    ps["尼古丁最高浓度"] = ps["尼古丁档位列表"].apply(lambda l: max(l) if l else None)

    # ---- brand popular (热销一次性品牌专项) ----
    bp = pd.read_excel(xl, sheet_name="brand popular")
    bp.columns = bp.columns.str.strip()
    bp["含冰/薄荷"] = bp["含冰/薄荷"].fillna("不确定")
    bp["口味标签列表"] = bp["口味标签"].apply(split_tags)
    bp["价格_数值"] = bp["价格"].apply(parse_price)
    bp["口味标签数"] = bp["口味标签列表"].apply(len)
    bp["口味复杂度"] = bp["口味标签数"].apply(flavor_complexity)

    # ---- daraz flavor freq ----
    daraz_raw = pd.read_excel(xl, sheet_name="vape_products_daraz", header=None)
    daraz = pd.DataFrame({
        "口味名称": daraz_raw.iloc[1:, 0].values,
        "口味元素": daraz_raw.iloc[1:, 1].values,
        "出现次数": pd.to_numeric(daraz_raw.iloc[1:, 2], errors="coerce"),
    }).dropna(subset=["出现次数"])
    daraz["出现次数"] = daraz["出现次数"].astype(int)

    # ---- google trends ----
    trends = pd.read_excel(xl, sheet_name="year_googletrend_hotbrand-2026")
    trends.columns = trends.columns.str.strip()

    # ---- related searches ----
    elfbar_rel = pd.read_excel(xl, sheet_name="searched_with_top-elfbar", header=None)
    elfbar_rel = pd.DataFrame({
        "搜索词": elfbar_rel.iloc[1:, 0].values,
        "搜索热度": pd.to_numeric(elfbar_rel.iloc[1:, 1], errors="coerce"),
        "增长率": pd.to_numeric(elfbar_rel.iloc[1:, 2], errors="coerce"),
    }).dropna(subset=["搜索热度"])

    voltbar_rel = pd.read_excel(xl, sheet_name="searched_with_top-voltbar", header=None)
    voltbar_rel = pd.DataFrame({
        "搜索词": voltbar_rel.iloc[1:, 0].values,
        "搜索热度": pd.to_numeric(voltbar_rel.iloc[1:, 1], errors="coerce"),
        "增长率": pd.to_numeric(voltbar_rel.iloc[1:, 2], errors="coerce"),
    }).dropna(subset=["搜索热度"])

    return ps, bp, daraz, trends, elfbar_rel, voltbar_rel


ps, bp, daraz, trends, elfbar_rel, voltbar_rel = load_data()

# ---- Sidebar ----
with st.sidebar:
    st.markdown("## 🔍 筛选（主数据库）")
    st.markdown("---")
    all_types = sorted(ps["产品类型"].unique().tolist())
    sel_types = st.multiselect("产品类型", all_types, default=all_types)
    all_countries = sorted(ps["品牌来源国"].dropna().unique().tolist())
    sel_countries = st.multiselect("品牌来源国", all_countries, default=all_countries)
    all_cats = sorted(ps["分类"].dropna().unique().tolist())
    sel_cats = st.multiselect("口味分类", all_cats, default=all_cats)
    all_sites = sorted(ps["网站名称"].dropna().unique().tolist())
    sel_sites = st.multiselect("数据来源网站", all_sites, default=all_sites)
    st.markdown("---")
    st.caption("📅 数据采集：2026-06")
    st.caption("数据看板制作人：Erin")
    st.caption(f"主数据库：{ps['网站名称'].nunique()} 个网站 · {len(ps)} 条记录")

fdf = ps[
    ps["产品类型"].isin(sel_types) &
    ps["品牌来源国"].isin(sel_countries) &
    ps["分类"].isin(sel_cats) &
    ps["网站名称"].isin(sel_sites)
].copy()

if fdf.empty:
    st.error("⚠️ 当前筛选条件下无数据，请放宽筛选范围。")
    st.stop()

st.markdown("# 🇧🇩 孟加拉国电子烟市场调研数据看板")
st.caption(
    f"筛选后：{len(fdf)} 条零售网站记录 ｜ 品牌 {fdf['品牌'].nunique()} 个 ｜"
    f" 网站 {fdf['网站名称'].nunique()} 个 ｜ 独立口味 {fdf['口味名称'].nunique()} 个"
)
st.markdown("---")

tabs = st.tabs([
    "📊 总览",
    "🍭 口味热榜",
    "🔖 口味元素",
    "🥇 热销一次性品牌",
    "🏷️ 品牌分析",
    "💉 产品规格",
    "💰 价格分析",
    "📝 总结洞察",
])

# ============================================================
# TAB 0  总览
# ============================================================
with tabs[0]:
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    kpis = [
        ("总零售记录数",       len(fdf),                      False),
        ("品牌数",         fdf["品牌"].nunique(),          False),
        ("独立口味数",     fdf["口味名称"].nunique(),      False),
        ("含冰/薄荷占比",  f"{(fdf['含冰/薄荷']=='是').sum()/len(fdf):.0%}", True),
        ("烟油品类占比",   f"{(fdf['产品类型']=='E-liquid').sum()/len(fdf):.0%}", True),
        ("覆盖网站数",     fdf["网站名称"].nunique(),      False),
    ]
    for col, (label, val, red) in zip([c1, c2, c3, c4, c5, c6], kpis):
        cls = "metric-card-red" if red else "metric-card"
        col.markdown(
            f'<div class="{cls}"><div class="value">{val}</div>'
            f'<div class="label">{label}</div></div>',
            unsafe_allow_html=True
        )
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<p class="section-title">⚖️ 法规动态（2025.12 — 2026.04）</p>', unsafe_allow_html=True)
    col_reg1, col_reg2, col_reg3 = st.columns(3)
    with col_reg1:
        st.markdown("""
        <div class="danger-box">
        <b>🚫 2025年12月30日：全面禁令</b><br>
        政府颁布《吸烟与烟草产品使用控制（修正）条例》，将电子烟、加热烟草产品、尼古丁袋全部纳入烟草定义并<b>全面禁止</b>生产、进口、出口、储存、销售、使用。<br><br>
        违者最高监禁 <b>6个月</b>，罚款最高 <b>50万塔卡（约$4,100）</b>。
        </div>
        """, unsafe_allow_html=True)
    with col_reg2:
        st.markdown("""
        <div class="action-box">
        <b>✅ 2026年4月：议会移除禁令</b><br>
        议会通过修正法案，<b>移除了禁止电子烟、电子烟具、尼古丁袋的条款</b>，允许受控市场存在。<br><br>
        修订后法律仍保留：广告限制、禁止向未成年人销售、公共场所使用限制。
        </div>
        """, unsafe_allow_html=True)
    with col_reg3:
        st.markdown("""
        <div class="warning-box">
        <b>⚠️ 后续走向：高度不确定</b><br>
        同期媒体报道显示，议会移除禁令的决定在<b>公共卫生界引发强烈反对</b>，后续是否再度收紧法规仍存在较大不确定性。<br><br>
        市场进入窗口期，但需持续关注政策动态。
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">口味分类分布</p>', unsafe_allow_html=True)
        cc = fdf["分类"].value_counts().reset_index()
        cc.columns = ["分类", "数量"]
        fig = px.bar(cc, x="分类", y="数量", color="分类",
                     color_discrete_map=CAT_COLOR_MAP, text="数量", template=T)
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=10, b=10),
                          xaxis_title="", yaxis_title="数量")
        st.plotly_chart(fig, use_container_width=True)
        top_cat = fdf["分类"].value_counts().idxmax()
        top_pct = fdf["分类"].value_counts().iloc[0] / len(fdf)
        st.caption(f"水果口味占绝对主导（{top_pct:.0%}），其次是糖果与烟草，薄荷占比较低。")

    with col_r:
        st.markdown('<p class="section-title">产品类型 × 口味分类</p>', unsafe_allow_html=True)
        cross = fdf.groupby(["产品类型", "分类"]).size().reset_index(name="数量")
        fig = px.bar(cross, x="产品类型", y="数量", color="分类",
                     barmode="stack", color_discrete_map=CAT_COLOR_MAP, template=T)
        fig.update_layout(xaxis_title="", yaxis_title="数量",
                          legend_title_text="分类", margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("一次性电子烟以水果与糖果口味为主；烟油（E-liquid）则有更高比例的烟草口味。")

    st.markdown("---")
    st.markdown('<p class="section-title">📈 热销品牌 Google 搜索热度趋势（孟加拉国，2025.06 — 2026.06）</p>',
                unsafe_allow_html=True)

    fig_tr = go.Figure()
    fig_tr.add_trace(go.Scatter(
        x=trends["Time"], y=trends["elfbar"],
        name="Elf Bar", mode="lines+markers",
        line=dict(color=COLORS["elfbar"], width=2),
        marker=dict(size=4)
    ))
    fig_tr.add_trace(go.Scatter(
        x=trends["Time"], y=trends["voltbar"],
        name="Voltbar", mode="lines+markers",
        line=dict(color=COLORS["voltbar"], width=2),
        marker=dict(size=4)
    ))
    fig_tr.update_layout(
        template=T, height=280,
        margin=dict(t=10, b=10),
        xaxis_title="", yaxis_title="搜索热度（相对值）",
        legend=dict(orientation="h", y=1.1)
    )
    st.plotly_chart(fig_tr, use_container_width=True)

    col_tr1, col_tr2 = st.columns(2)
    with col_tr1:
        elfbar_avg = trends["elfbar"].mean()
        voltbar_avg = trends["voltbar"].mean()
        st.markdown(f"""
        <div class="insight-box">
        <b>🔍 在孟加拉地区俩品牌搜索热度对比</b><br>
        • Voltbar 年均搜索热度 <b>{voltbar_avg:.1f}</b>，<b>略高于</b> Elf Bar（{elfbar_avg:.1f}）<br>
        • 两个品牌均出现周期性搜索低谷（热度为0），可能与库存/上架节奏有关<br>
        </div>
        """, unsafe_allow_html=True)
    with col_tr2:
        col_e, col_v = st.columns(2)
        with col_e:
            st.markdown('<p class="section-title">Elf Bar 关联搜索词</p>', unsafe_allow_html=True)
            ef_top = elfbar_rel.sort_values("搜索热度", ascending=False).head(8)
            fig_ef = px.bar(ef_top, x="搜索热度", y="搜索词", orientation="h",
                            text="搜索热度", color_discrete_sequence=[COLORS["elfbar"]], template=T)
            fig_ef.update_traces(textposition="outside")
            fig_ef.update_layout(height=280, margin=dict(t=5, b=5),
                                 xaxis_title="", yaxis_title="",
                                 yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_ef, use_container_width=True)
        with col_v:
            st.markdown('<p class="section-title">Voltbar 关联搜索词</p>', unsafe_allow_html=True)
            vb_top = voltbar_rel.sort_values("搜索热度", ascending=False).head(8)
            fig_vb = px.bar(vb_top, x="搜索热度", y="搜索词", orientation="h",
                            text="搜索热度", color_discrete_sequence=[COLORS["voltbar"]], template=T)
            fig_vb.update_traces(textposition="outside")
            fig_vb.update_layout(height=280, margin=dict(t=5, b=5),
                                 xaxis_title="", yaxis_title="",
                                 yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_vb, use_container_width=True)

    st.markdown("---")
    st.markdown('<p class="section-title">📋 数据来源与调研范围说明</p>', unsafe_allow_html=True)
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        st.markdown("""
        <div class="insight-box">
        <b>🛒 孟加拉最大电商平台 Daraz</b><br>
        关键词「vape」爬取10页共 <b>1,560条</b> 数据，筛选出电子烟相关 <b>620个SKU</b>，最高单品销量 <b>1.1K</b>。<br><br>
        本看板展示 Daraz 口味频次聚合汇总（Top 22 口味）。
        </div>
        """, unsafe_allow_html=True)
    with col_d2:
        st.markdown("""
        <div class="insight-box">
        <b>🏆 零售网站热销榜单</b><br>
        收集 <b>4个网站</b>（vapeempirebd / vapehubbd / vaporworldbd / vapershopbd）<br>
        两种产品类型（一次性、烟油）热销榜单前 <b>15名</b>产品及对应口味，共 <b>90条</b>记录。
        </div>
        """, unsafe_allow_html=True)
    with col_d3:
        st.markdown("""
        <div class="insight-box">
        <b>📊 热销品牌专项</b><br>
        通过 Google Trends 确认热度后，针对两大热销品牌 <b>Elf Bar RAYA D1（13K）</b> 和 <b>Voltbar Cartridge（12K）</b>，在3个网站各采集前10名口味，共 <b>60条</b>记录。
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<p class="section-title">⚠️ 调研瓶颈与数据局限性</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="warning-box">
    <ul style="margin:0; padding-left:20px;">
        <li><b>宏观市场数据缺失：</b>主流研究机构（Euromonitor、Statista、ECigIntelligence、IMARC等）均未对孟加拉电子烟品类做单独的产品类型/渠道/口味占比拆解。</li>
        <li><b>用户基数极小（2017年数据）：</b>截至2017年，孟加拉国共约有22.3万电子烟用户。根据全球烟草使用趋势调查（GATS 2017，世卫组织引用），孟加拉国成人电子烟使用率约 <b>0.2%</b>（约22.3万现用用户），0.4% 曾尝试过——这是目前能查到的、唯一具有全国代表性的孟加拉电子烟使用数据，但已是2017年的数据，且未涉及产品类型或口味细分。</li>
        <li><b>时效性不足：</b>本调研所有零售数据采集于2026年6月，虽能反映当前货架情况，但无法捕捉快速变化的消费偏好与政策影响下的市场波动。</li>
        <li><b>样本覆盖局限：</b>仅覆盖4个零售网站的热销榜单及Daraz平台部分SKU，未涵盖线下渠道（如便利店、烟草店）及社交媒体分销，可能遗漏非线上主流口味。</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# TAB 1  口味热榜
# ============================================================
with tabs[1]:
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">全榜热门口味 Top 15</p>', unsafe_allow_html=True)
        fl_counts = fdf["口味名称"].value_counts()
        fl = fl_counts.head(15).reset_index()
        fl.columns = ["口味名称", "出现次数"]
        fl = fl.sort_values("出现次数", ascending=True)
        fig = px.bar(fl, x="出现次数", y="口味名称", orientation="h",
                     text="出现次数",
                     color="出现次数",
                     color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                     template=T)
        fig.update_traces(textposition="outside")
        fig.update_layout(height=500, margin=dict(t=10, b=10),
                          xaxis_title="", yaxis_title="",
                          coloraxis_showscale=False,
                          yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)
        if not fl_counts.empty:
            top1 = fl_counts.index[0]
            top1_cnt = fl_counts.iloc[0]
            st.caption(f"全市场最热门口味：VCT和Mango Lychee Bubblegum（均出现3次）。")

    with col_r:
        st.markdown('<p class="section-title">按产品类型查看 Top 15</p>', unsafe_allow_html=True)
        ptype = st.selectbox("选择产品类型", fdf["产品类型"].unique(), key="tab1_type")
        sub_fl = fdf[fdf["产品类型"] == ptype]["口味名称"].value_counts().head(15).reset_index()
        sub_fl.columns = ["口味名称", "出现次数"]
        sub_fl = sub_fl.sort_values("出现次数", ascending=True)
        ptype_color = COLORS["primary"] if "Disposable" in ptype else COLORS["red"]
        fig = px.bar(sub_fl, x="出现次数", y="口味名称", orientation="h",
                     text="出现次数",
                     color_discrete_sequence=[ptype_color], template=T)
        fig.update_traces(textposition="outside")
        fig.update_layout(height=500, margin=dict(t=10, b=10),
                          xaxis_title="", yaxis_title="",
                          yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<p class="section-title">含冰/薄荷分布</p>', unsafe_allow_html=True)
        ice_cnt = fdf["含冰/薄荷"].value_counts().reset_index()
        ice_cnt.columns = ["状态", "数量"]
        fig_ice = px.pie(ice_cnt, values="数量", names="状态",
                         color="状态",
                         color_discrete_map={"是": COLORS["ice"], "否": "#94a3b8", "不确定": "#e2e8f0"},
                         hole=0.45, template=T)
        fig_ice.update_traces(textinfo="percent+label")
        fig_ice.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig_ice, use_container_width=True)
        ice_yes = (fdf["含冰/薄荷"] == "是").sum()
        st.caption(f"明确含冰/薄荷产品 {ice_yes} 款（{ice_yes/len(fdf):.0%}），加上「不确定」项，冰感产品实际占比可能更高。")

    with col_b:
        st.markdown('<p class="section-title">口味复杂度分布</p>', unsafe_allow_html=True)
        comp = fdf["口味复杂度"].value_counts().reset_index()
        comp.columns = ["复杂度", "数量"]
        fig_comp = px.pie(comp, values="数量", names="复杂度",
                          color="复杂度",
                          color_discrete_map={
                              "单一口味":     "#94a3b8",
                              "双重复合":     COLORS["primary"],
                              "三重以上复合": COLORS["secondary"],
                          },
                          hole=0.45, template=T)
        fig_comp.update_traces(textinfo="percent+label")
        fig_comp.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig_comp, use_container_width=True)
        multi_pct = fdf["口味复杂度"].isin(["双重复合", "三重以上复合"]).sum() / len(fdf)
        st.caption(f"复合口味占比 {multi_pct:.0%}，孟加拉消费者偏好层次丰富的口感组合。")

    st.markdown('<p class="section-title">🔥 口味分类 × 含冰/薄荷热力图</p>', unsafe_allow_html=True)
    cross_ice = pd.crosstab(fdf["分类"], fdf["含冰/薄荷"])
    fig_hi = px.imshow(cross_ice, text_auto=True, aspect="auto",
                       color_continuous_scale=[[0, "#f0fdf4"], [1, COLORS["primary"]]],
                       template=T)
    fig_hi.update_layout(xaxis_title="", yaxis_title="",
                         coloraxis_showscale=False, height=260, margin=dict(t=10, b=10))
    st.plotly_chart(fig_hi, use_container_width=True)
    st.caption("水果类中冰感产品数量最多；烟草口味几乎不搭配冰感，符合传统口感偏好。")

    st.markdown('<p class="section-title">Top 15 口味详细成分标签</p>', unsafe_allow_html=True)
    fl_list = fl_counts.head(15).reset_index()
    fl_list.columns = ["口味名称", "出现次数"]
    cols2 = st.columns(2)
    for i, (_, row) in enumerate(fl_list.iterrows()):
        fname = row["口味名称"]
        tag_freq = Counter([t for tl in fdf[fdf["口味名称"] == fname]["口味标签列表"] for t in tl])
        tags_html = " ".join([f'<span class="tag-chip">{t}</span>' for t, _ in tag_freq.most_common()]) \
                    or '<span style="color:#9ca3af">无标签</span>'
        with cols2[i % 2]:
            st.markdown(
                f'<div class="flavor-row"><b>#{i+1} {fname}</b>（{row["出现次数"]}次）<br>{tags_html}</div>',
                unsafe_allow_html=True
            )


# ============================================================
# TAB 2  口味元素（V5 修复版）
# ============================================================
with tabs[2]:
    # ---- 计算零售端口味元素 ----
    all_tags = [t for tl in fdf["口味标签列表"] for t in tl]
    tag_counts = Counter(all_tags)

    # ---- 计算 Daraz 平台口味元素（加权：每条记录按"出现次数"累加） ----
    daraz_element_counts = Counter()
    for _, row in daraz.iterrows():
        val = row["口味元素"]
        weight = row["出现次数"]  # 权重 = 该口味的SKU数量
        if pd.isna(val):
            continue
        if isinstance(val, str):
            elements = [t.strip().lower() for t in val.split(",") if t.strip()]
        elif isinstance(val, list):
            elements = [t.strip().lower() for t in val if t.strip()]
        else:
            continue
        for elem in elements:
            daraz_element_counts[elem] += weight  # 加权累加

    # ---- 第1行：零售元素 Top 20（左） vs Daraz 元素 Top 20（右） ----
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">高频口味元素 Top 20（零售热销榜单）</p>', unsafe_allow_html=True)
        top_tags_df = pd.DataFrame(tag_counts.most_common(20), columns=["口味元素", "出现次数"])
        top_tags_df = top_tags_df.sort_values("出现次数", ascending=True)
        fig = px.bar(top_tags_df, x="出现次数", y="口味元素", orientation="h",
                     text="出现次数",
                     color="出现次数",
                     color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                     template=T)
        fig.update_traces(textposition="outside")
        fig.update_layout(coloraxis_showscale=False, height=560,
                          margin=dict(t=10, b=10), xaxis_title="", yaxis_title="",
                          yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)
        if tag_counts:
            top3 = [f"{t}（{c}）" for t, c in tag_counts.most_common(3)]
            st.caption(f"出现最高频的元素：{' / '.join(top3)}。热带水果是孟加拉市场最强口味信号。")

    with col_r:
        st.markdown('<p class="section-title">高频口味元素 Top 20（Daraz 平台，已加权）</p>', unsafe_allow_html=True)
        if daraz_element_counts:
            daraz_element_df = pd.DataFrame(daraz_element_counts.most_common(20), columns=["口味元素", "出现次数"])
            daraz_element_df = daraz_element_df.sort_values("出现次数", ascending=True)
            fig_daraz_el = px.bar(daraz_element_df, x="出现次数", y="口味元素", orientation="h",
                                  text="出现次数",
                                  color="出现次数",
                                  color_continuous_scale=[[0, "#fde68a"], [1, "#d97706"]],
                                  template=T)
            fig_daraz_el.update_traces(textposition="outside")
            fig_daraz_el.update_layout(coloraxis_showscale=False, height=560,
                                       margin=dict(t=10, b=10), xaxis_title="出现次数（SKU加权）",
                                       yaxis_title="",
                                       yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_daraz_el, use_container_width=True)
            daraz_top1 = daraz_element_counts.most_common(1)[0]
            st.caption(f"📌 统计逻辑：将「口味元素」列拆分后，按「出现次数」加权累加（如「Mango Ice」出现15次，则Mango+15、Ice+15）。")
            st.caption(f"Daraz 平台最高频元素：{daraz_top1[0]}（{daraz_top1[1]}次）。与零售榜单对比可发现略微有差异。")
        else:
            st.info("Daraz 平台口味元素数据为空")

    # ---- 第2行：零售 vs Daraz 口味元素 TOP 10 并排对比 ----
    st.markdown("---")
    st.markdown('<p class="section-title">📊 零售 vs Daraz 口味元素 TOP 10 对比</p>', unsafe_allow_html=True)
    if tag_counts and daraz_element_counts:
        retail_top10_df = pd.DataFrame(tag_counts.most_common(10), columns=["口味元素", "零售出现次数"])
        daraz_top10_df = pd.DataFrame(daraz_element_counts.most_common(10), columns=["口味元素", "Daraz出现次数"])
        combined_df = pd.merge(retail_top10_df, daraz_top10_df, on="口味元素", how="outer").fillna(0)
        combined_df["总次数"] = combined_df["零售出现次数"] + combined_df["Daraz出现次数"]
        combined_df = combined_df.sort_values("总次数", ascending=False).head(10)
        combined_melted = combined_df.melt(
            id_vars=["口味元素"],
            value_vars=["零售出现次数", "Daraz出现次数"],
            var_name="数据源",
            value_name="出现次数"
        )
        fig_cmp = px.bar(
            combined_melted,
            x="口味元素",
            y="出现次数",
            color="数据源",
            barmode="group",
            text="出现次数",
            color_discrete_map={"零售出现次数": COLORS["primary"], "Daraz出现次数": "#d97706"},
            template=T
        )
        fig_cmp.update_traces(textposition="outside")
        fig_cmp.update_layout(
            xaxis_title="",
            yaxis_title="出现次数",
            legend_title_text="",
            margin=dict(t=10, b=10),
            xaxis_tickangle=-20,
            height=400
        )
        st.plotly_chart(fig_cmp, use_container_width=True)
        st.caption("Mango 在两数据源均为顶流；Strawberry 在 Daraz 上架量显著更高。")
    else:
        st.info("零售或 Daraz 口味元素数据不足，无法对比。")

    # ---- 第3行：Daraz 口味名称 Top 25（左） vs 各分类高频元素（右） ----
    st.markdown("---")
    col_l2, col_r2 = st.columns(2)
    with col_l2:
        st.markdown('<p class="section-title">Daraz 平台口味频次 Top 22（电商搜索数据）</p>', unsafe_allow_html=True)
        dz = daraz.sort_values("出现次数", ascending=True)
        fig_dz = px.bar(dz, x="出现次数", y="口味名称", orientation="h",
                        text="出现次数",
                        color="出现次数",
                        color_continuous_scale=[[0, "#fde68a"], [1, "#d97706"]],
                        template=T)
        fig_dz.update_traces(textposition="outside")
        fig_dz.update_layout(coloraxis_showscale=False, height=560,
                             margin=dict(t=10, b=10), xaxis_title="出现次数（SKU数）",
                             yaxis_title="",
                             yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_dz, use_container_width=True)
        dz_top1 = daraz.sort_values("出现次数", ascending=False).iloc[0]
        st.caption(f"Daraz平台最热口味：{dz_top1['口味名称']}（{dz_top1['出现次数']}个SKU），Strawberry 在电商上架量领先。")

    with col_r2:
        st.markdown('<p class="section-title">各口味分类在零售网站的高频元素</p>', unsafe_allow_html=True)
        cats_list = fdf["分类"].value_counts().index.tolist()
        for i, cat in enumerate(cats_list):
            cdf = fdf[fdf["分类"] == cat]
            cat_tags_list = [t for tl in cdf["口味标签列表"] for t in tl]
            top6 = Counter(cat_tags_list).most_common(6)
            tags_html = " ".join([f'<span class="tag-chip">{t}（{c}）</span>' for t, c in top6])
            st.markdown(
                f'<div class="flavor-row"><b>{cat}</b>（{len(cdf)}款）&nbsp;&nbsp;{tags_html}</div>',
                unsafe_allow_html=True
            )

    # ---- 第4行：双元素组合（融合零售 + Daraz） ----
    st.markdown("---")
    st.markdown('<p class="section-title">🔗 （零售网站与Daraz平台）口味元素双元素组合频率（Top 12）</p>', unsafe_allow_html=True)
    st.markdown('<div class="insight-box">📐 <b>权重逻辑：</b>零售热销榜单每条记录计 1 次；Daraz 平台按「出现次数」列加权累加。两者合并后统一排序。</div>',
                unsafe_allow_html=True)

    combos = Counter()

    # 1) 零售端：每条记录计 1 次
    for tl in fdf["口味标签列表"]:
        uniq = list(set(tl))
        for ii in range(len(uniq)):
            for jj in range(ii + 1, len(uniq)):
                pair = tuple(sorted([uniq[ii], uniq[jj]]))
                combos[pair] += 1

    # 2) Daraz 端：按"出现次数"加权
    for _, row in daraz.iterrows():
        val = row["口味元素"]
        weight = row["出现次数"]
        if pd.isna(val):
            continue
        if isinstance(val, str):
            elements = [t.strip().lower() for t in val.split(",") if t.strip()]
        elif isinstance(val, list):
            elements = [t.strip().lower() for t in val if t.strip()]
        else:
            continue
        if len(elements) < 2:
            continue
        uniq = list(set(elements))
        for ii in range(len(uniq)):
            for jj in range(ii + 1, len(uniq)):
                pair = tuple(sorted([uniq[ii], uniq[jj]]))
                combos[pair] += weight  # 加权累加

    top_combos = combos.most_common(12)
    if top_combos:
        combo_df = pd.DataFrame(
            [(f"{a} + {b}", c) for (a, b), c in top_combos],
            columns=["组合", "出现次数"]
        ).sort_values("出现次数", ascending=True)
        fig_cb = px.bar(combo_df, x="出现次数", y="组合", orientation="h",
                        text="出现次数",
                        color="出现次数",
                        color_continuous_scale=[[0, "#d1fae5"], [1, "#065f46"]],
                        template=T)
        fig_cb.update_traces(textposition="outside")
        fig_cb.update_layout(coloraxis_showscale=False, height=420,
                             margin=dict(t=10, b=10),
                             yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_cb, use_container_width=True)
        top_combo_name = f"{top_combos[0][0][0]} + {top_combos[0][0][1]}"
        st.caption(f"最高频双元素组合：{top_combo_name}（{top_combos[0][1]}次）。融合零售榜单与 Daraz 加权数据，热带水果与冰感/莓果是黄金搭档。")
    else:
        st.info("暂无足够数据生成双元素组合。")


# ============================================================
# TAB 3  热销一次性品牌（brand popular）
# ============================================================
with tabs[3]:
    st.markdown("### 🥇 热销一次性品牌口味深度分析")
    st.markdown(
        '<div class="insight-box">本板块数据来源：<b>brand popular</b> 专项表，聚焦孟加拉市场搜索热度最高的两大热销一次性品牌——'
        '<b>Elf Bar RAYA D1（13,000 puffs）</b> 和 <b>Voltbar Cartridge（12,000 puffs）</b>，'
        '各在 3 个零售网站（vaporworldbd / vapehubbd / vapershopbd）采集前 10 名热销口味，共 60 条记录。</div>',
        unsafe_allow_html=True
    )

    brands_bp = ["Elf Bar", "Voltbar"]
    brand_colors_bp = {"Elf Bar": COLORS["elfbar"], "Voltbar": COLORS["voltbar"]}
    brand_models = {"Elf Bar": "RAYA D1 · 13,000 puffs", "Voltbar": "Cartridge · 12,000 puffs"}

    kpi_cols = st.columns(2)
    for col, brand in zip(kpi_cols, brands_bp):
        bdf = bp[bp["品牌"] == brand]
        ice_pct = (bdf["含冰/薄荷"] == "是").sum() / len(bdf)
        top_cat = bdf["分类"].value_counts().idxmax()
        avg_price = bdf["价格_数值"].mean()
        multi_pct = bdf["口味复杂度"].isin(["双重复合", "三重以上复合"]).sum() / len(bdf)
        cny_avg = avg_price * EXCHANGE_RATE
        col.markdown(
            f'<div class="metric-card" style="background:linear-gradient(135deg,{brand_colors_bp[brand]},{COLORS["secondary"]});">'
            f'<div class="value" style="font-size:1.4rem">{brand}</div>'
            f'<div class="label">{brand_models[brand]}<br>'
            f'均价 ৳{avg_price:,.0f}（≈¥{cny_avg:,.1f}）｜ 含冰 {ice_pct:.0%} ｜ 复合口味 {multi_pct:.0%}<br>'
            f'主要口味分类：{top_cat}</div></div>',
            unsafe_allow_html=True
        )
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<p class="section-title">📊 综合热度 Top 10（∑1/排名 × 出现网站数）</p>', unsafe_allow_html=True)
    sel_brand = st.selectbox(
        "选择品牌",
        brands_bp,
        format_func=lambda x: f"{x}  ·  {brand_models[x]}"
    )
    brand_data = bp[bp["品牌"] == sel_brand].copy()
    brand_data["排名_数值"] = pd.to_numeric(brand_data["排名"], errors="coerce")
    brand_data = brand_data.dropna(subset=["排名_数值"])

    flavor_scores = {}
    flavor_meta = {}
    for _, row in brand_data.iterrows():
        flavor = row["口味名称"]
        rank = row["排名_数值"]
        site = row["网站名称"]
        ice_flag = row["含冰/薄荷"]
        tags = row["口味标签列表"]
        score_inc = 1.0 / rank
        if flavor not in flavor_scores:
            flavor_scores[flavor] = 0.0
            flavor_meta[flavor] = {
                "sites": set(), "ice": ice_flag == "是",
                "tags": tags, "occurrences": 0,
                "best_rank": rank, "sum_inv_rank": 0.0,
                "cat": row["分类"],
            }
        flavor_scores[flavor] += score_inc
        flavor_meta[flavor]["sites"].add(site)
        flavor_meta[flavor]["occurrences"] += 1
        flavor_meta[flavor]["sum_inv_rank"] += score_inc
        if rank < flavor_meta[flavor]["best_rank"]:
            flavor_meta[flavor]["best_rank"] = rank

    final_scores = {f: s * len(flavor_meta[f]["sites"]) for f, s in flavor_scores.items()}
    top_flavors = pd.DataFrame([{
        "口味名称":  f,
        "综合得分":  round(final_scores[f], 4),
        "覆盖网站数": len(flavor_meta[f]["sites"]),
        "总上榜次数": flavor_meta[f]["occurrences"],
        "最高排名":  int(flavor_meta[f]["best_rank"]),
        "含冰/薄荷": "是" if flavor_meta[f]["ice"] else "否",
        "口味分类":  flavor_meta[f]["cat"],
        "主要标签":  ", ".join(flavor_meta[f]["tags"][:3]) if flavor_meta[f]["tags"] else "-",
    } for f in final_scores]).sort_values("综合得分", ascending=False).head(10)

    st.markdown(
        '<div class="insight-box">📐 <b>权重说明：</b> 综合得分 = Σ(1/排名) × 出现网站数。'
        '排名越靠前贡献越大，同时在多网站同时热销权重更高。</div>',
        unsafe_allow_html=True
    )

    fig_top = px.bar(
        top_flavors,
        x="综合得分", y="口味名称", orientation="h",
        text="综合得分",
        color="综合得分",
        color_continuous_scale=[[0, "#a7f3d0"], [1, brand_colors_bp[sel_brand]]],
        template=T,
        category_orders={"口味名称": top_flavors["口味名称"].tolist()}
    )
    fig_top.update_traces(texttemplate="%{text:.3f}", textposition="outside")
    fig_top.update_layout(
        height=420, xaxis_title="综合得分",
        yaxis_title="", yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False, margin=dict(t=10, b=10)
    )
    st.plotly_chart(fig_top, use_container_width=True)
    st.dataframe(top_flavors, use_container_width=True, hide_index=True)

    with st.expander("🔎 点击查看每个口味的详细来源网站"):
        for _, row in top_flavors.iterrows():
            flavor = row["口味名称"]
            sites = flavor_meta[flavor]["sites"]
            st.markdown(f"**{flavor}** → 出现网站：{', '.join(sites)}")

    st.markdown("---")
    st.markdown('<p class="section-title">两品牌口味元素对比热力图</p>', unsafe_allow_html=True)
    all_bp_tags = [t for tl in bp["口味标签列表"] for t in tl]
    top15_tags = [t for t, _ in Counter(all_bp_tags).most_common(15)]
    brand_tag_matrix = {}
    for b in brands_bp:
        bdf2 = bp[bp["品牌"] == b]
        btags = [t for tl in bdf2["口味标签列表"] for t in tl]
        brand_tag_matrix[b] = Counter(btags)
    matrix_df = pd.DataFrame(
        {b: [brand_tag_matrix[b].get(t, 0) for t in top15_tags] for b in brands_bp},
        index=top15_tags
    )
    fig_hm = px.imshow(matrix_df.T, text_auto=True, aspect="auto",
                       color_continuous_scale=[[0, "#f0fdf4"], [1, COLORS["primary"]]],
                       template=T)
    fig_hm.update_layout(xaxis_title="", yaxis_title="",
                          coloraxis_showscale=False, height=220, margin=dict(t=10, b=10))
    st.plotly_chart(fig_hm, use_container_width=True)

    common_tags = set(brand_tag_matrix["Elf Bar"].keys()) & set(brand_tag_matrix["Voltbar"].keys())
    if common_tags:
        top_common = sorted(common_tags,
                            key=lambda t: sum(brand_tag_matrix[b].get(t, 0) for b in brands_bp),
                            reverse=True)
        st.markdown(
            '<div class="insight-box"><b>🎯 两品牌共同口味元素：</b><br>' +
            " ".join([f'<span class="tag-chip">{t}</span>' for t in top_common[:12]]) +
            "</div>", unsafe_allow_html=True
        )

    st.markdown('<p class="section-title">口味分类分布对比（Elf Bar vs Voltbar）</p>', unsafe_allow_html=True)
    cat_cmp = bp.groupby(["品牌", "分类"]).size().reset_index(name="数量")
    fig_cat = px.bar(cat_cmp, x="分类", y="数量", color="品牌",
                     barmode="group",
                     color_discrete_map=brand_colors_bp,
                     template=T)
    fig_cat.update_layout(xaxis_title="", yaxis_title="数量",
                          legend_title_text="", margin=dict(t=10, b=10))
    st.plotly_chart(fig_cat, use_container_width=True)
    st.caption("两品牌均以水果口味为绝对主导。")


# ============================================================
# TAB 4  品牌分析
# ============================================================
with tabs[4]:
    st.markdown("## 🏷️ 品牌维度分析")

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">品牌出现频次 Top 12</p>', unsafe_allow_html=True)
        brand_cnt = fdf["品牌"].value_counts().head(12).reset_index()
        brand_cnt.columns = ["品牌", "出现次数"]
        brand_cnt = brand_cnt.sort_values("出现次数", ascending=True)
        fig_bc = px.bar(brand_cnt, x="出现次数", y="品牌", orientation="h",
                        text="出现次数",
                        color="出现次数",
                        color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                        template=T)
        fig_bc.update_traces(textposition="outside")
        fig_bc.update_layout(coloraxis_showscale=False, height=480,
                              margin=dict(t=10, b=10), xaxis_title="", yaxis_title="",
                              yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_bc, use_container_width=True)
        top_brand = fdf["品牌"].value_counts().idxmax()
        top_brand_cnt = fdf["品牌"].value_counts().iloc[0]
        st.caption(f"{top_brand} 以 {top_brand_cnt} 次居首，是孟加拉一次性市场的绝对强势品牌。")

    with col_r:
        st.markdown('<p class="section-title">品牌来源国分布（产品数）</p>', unsafe_allow_html=True)
        oc = fdf["品牌来源国"].value_counts().reset_index()
        oc.columns = ["来源国", "数量"]
        fig_oc = px.pie(oc, values="数量", names="来源国",
                        color="来源国", color_discrete_map=COUNTRY_COLOR_MAP,
                        hole=0.45, template=T)
        fig_oc.update_traces(textinfo="percent+label", textposition="outside")
        fig_oc.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig_oc, use_container_width=True)
        top_country = fdf["品牌来源国"].value_counts().idxmax()
        top_country_pct = fdf["品牌来源国"].value_counts().iloc[0] / len(fdf)
        st.caption(f"{top_country}品牌占 {top_country_pct:.0%}，一次性市场由中国品牌主导；美国品牌（Vgod等）则主要在烟油品类发力。")

    st.markdown('<p class="section-title">网站 × 品牌出现频次矩阵</p>', unsafe_allow_html=True)
    top_brands_list = fdf["品牌"].value_counts().head(10).index.tolist()
    site_brand = pd.crosstab(fdf["网站名称"], fdf["品牌"])
    site_brand_top = site_brand[[c for c in top_brands_list if c in site_brand.columns]]
    fig_sb = px.imshow(site_brand_top, text_auto=True, aspect="auto",
                       color_continuous_scale=[[0, "#f0fdf4"], [1, COLORS["primary"]]],
                       template=T)
    fig_sb.update_layout(xaxis_title="品牌", yaxis_title="网站",
                          coloraxis_showscale=False, height=280, margin=dict(t=10, b=10))
    st.plotly_chart(fig_sb, use_container_width=True)
    st.caption("Elf Bar 在所有一次性渠道均有布局；烟油品牌（Vgod、Bar Juice 等）集中在特定网站。")

    st.markdown('<p class="section-title">品牌代表性口味（可选择）</p>', unsafe_allow_html=True)
    brand_sel = st.selectbox("选择品牌", fdf["品牌"].unique(), key="brand_flavor_sel")
    bf = fdf[fdf["品牌"] == brand_sel]["口味名称"].value_counts().head(10).reset_index()
    bf.columns = ["口味名称", "出现次数"]
    if not bf.empty:
        bf = bf.sort_values("出现次数", ascending=True)
        fig_bf = px.bar(bf, x="出现次数", y="口味名称", orientation="h",
                        text="出现次数",
                        color_discrete_sequence=[COLORS["primary"]], template=T)
        fig_bf.update_traces(textposition="outside")
        fig_bf.update_layout(height=380, margin=dict(t=10, b=10),
                              yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_bf, use_container_width=True)

    st.markdown('<p class="section-title">品牌来源国 × 口味分类偏好</p>', unsafe_allow_html=True)
    country_cat = fdf.groupby(["品牌来源国", "分类"]).size().reset_index(name="数量")
    fig_cc = px.bar(country_cat, x="品牌来源国", y="数量", color="分类",
                    barmode="stack", color_discrete_map=CAT_COLOR_MAP, template=T)
    fig_cc.update_layout(xaxis_title="", yaxis_title="产品数",
                          legend_title_text="分类", margin=dict(t=10, b=10))
    st.plotly_chart(fig_cc, use_container_width=True)
    st.caption("美国品牌（Vgod、Bar Juice 等）热销烟草口味更多，与中国品牌水果主导形成明显差异化。")


# ============================================================
# TAB 5  产品规格
# ============================================================
with tabs[5]:
    st.markdown("## 💉 产品规格分析")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<p class="section-title">尼古丁浓度分布（mg）</p>', unsafe_allow_html=True)
        nic_flat = [n for nl in fdf["尼古丁档位列表"] for n in nl]
        if nic_flat:
            nic_df = pd.DataFrame(Counter(nic_flat).items(), columns=["浓度(mg)", "出现次数"])
            nic_df = nic_df.sort_values("浓度(mg)")
            fig_nic = px.bar(nic_df, x="浓度(mg)", y="出现次数",
                             text="出现次数",
                             color="浓度(mg)",
                             color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                             template=T)
            fig_nic.update_traces(textposition="outside")
            fig_nic.update_layout(coloraxis_showscale=False,
                                   margin=dict(t=10, b=10),
                                   xaxis_title="尼古丁浓度(mg)", yaxis_title="出现次数")
            st.plotly_chart(fig_nic, use_container_width=True)
            top_nic = Counter(nic_flat).most_common(1)[0]
            st.caption(f"50mg 盐尼古丁占绝对主导（{top_nic[1]}次），是孟加拉一次性市场的标准配置；25mg 为第二梯队，多档浓度产品满足不同用户群。")

    with col_b:
        st.markdown('<p class="section-title">尼古丁浓度 × 产品类型分布</p>', unsafe_allow_html=True)
        nic_type_cross = []
        for _, row in fdf.iterrows():
            for n in row["尼古丁档位列表"]:
                nic_type_cross.append({"产品类型": row["产品类型"], "尼古丁浓度(mg)": n})
        ntc_df = pd.DataFrame(nic_type_cross)
        if not ntc_df.empty:
            ntc_agg = ntc_df.groupby(["产品类型", "尼古丁浓度(mg)"]).size().reset_index(name="出现次数")
            fig_ntc = px.bar(ntc_agg, x="尼古丁浓度(mg)", y="出现次数", color="产品类型",
                             barmode="group",
                             color_discrete_map={
                                 "Disposable Vape": COLORS["primary"],
                                 "E-liquid":        COLORS["red"],
                             },
                             template=T)
            fig_ntc.update_layout(xaxis_title="尼古丁浓度(mg)", yaxis_title="出现次数",
                                   legend_title_text="", margin=dict(t=10, b=10))
            st.plotly_chart(fig_ntc, use_container_width=True)
            st.caption("一次性产品集中在50mg高浓度；烟油覆盖25–50mg多档，满足不同用户需求。")

    st.markdown('<p class="section-title">一次性电子烟 Puffs 口数分布</p>', unsafe_allow_html=True)
    disp_df = fdf[fdf["puffs口数"].notna()].copy()
    if not disp_df.empty:
        puffs_cnt = disp_df["puffs口数"].value_counts().reset_index()
        puffs_cnt.columns = ["puffs口数", "产品数"]
        puffs_cnt = puffs_cnt.sort_values("puffs口数")
        puffs_cnt["puffs口数_标签"] = puffs_cnt["puffs口数"].apply(lambda x: f"{int(x):,}")
        fig_pf = px.bar(puffs_cnt, x="puffs口数_标签", y="产品数",
                        text="产品数",
                        color="产品数",
                        color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                        template=T)
        fig_pf.update_traces(textposition="outside")
        fig_pf.update_layout(coloraxis_showscale=False,
                              margin=dict(t=10, b=10),
                              xaxis_title="puffs 口数", yaxis_title="产品数")
        st.plotly_chart(fig_pf, use_container_width=True)
        top_puffs = puffs_cnt.sort_values("产品数", ascending=False).iloc[0]
        st.caption(f"{int(top_puffs['puffs口数']):,} puffs 是市场最主流规格（{top_puffs['产品数']}款），12K–13K 大口数产品占据一次性市场主导。")

    st.markdown('<p class="section-title">设备类别 × 尼古丁浓度关系</p>', unsafe_allow_html=True)
    dev_nic = []
    for _, row in fdf.iterrows():
        for n in row["尼古丁档位列表"]:
            dev_nic.append({"设备类别": row["设备类别"], "尼古丁浓度(mg)": n})
    dev_nic_df = pd.DataFrame(dev_nic)
    if not dev_nic_df.empty:
        fig_dn = px.box(dev_nic_df, x="设备类别", y="尼古丁浓度(mg)",
                        color="设备类别",
                        color_discrete_map={"一次性": COLORS["primary"], "烟油": COLORS["red"]},
                        points="all", template=T)
        fig_dn.update_layout(showlegend=False,
                              xaxis_title="", yaxis_title="尼古丁浓度(mg)",
                              margin=dict(t=10, b=10))
        st.plotly_chart(fig_dn, use_container_width=True)
        st.caption("一次性设备集中在 50mg 超高浓度，烟油则涵盖多个浓度档位（25–50mg），一次性产品的高浓度策略符合孟加拉重度用户偏好。")


# ============================================================
# TAB 6  价格分析
# ============================================================
with tabs[6]:
    st.markdown("## 💰 价格分析")

    price_df = fdf.dropna(subset=["价格_数值"])
    bp_price = bp.dropna(subset=["价格_数值"])

    if price_df.empty:
        st.info("暂无价格数据。")
    else:
        price_cols = st.columns(len(fdf["设备类别"].dropna().unique()))
        for col, dev in zip(price_cols, fdf["设备类别"].dropna().unique()):
            sub = price_df[price_df["设备类别"] == dev]["价格_数值"]
            if not sub.empty:
                mean_val = sub.mean()
                min_val = sub.min()
                max_val = sub.max()
                cny_mean = mean_val * EXCHANGE_RATE
                col.markdown(
                    f'<div class="metric-card"><div class="value">৳{mean_val:,.0f}</div>'
                    f'<div class="label">{dev} 均价<br>（৳{min_val:,.0f}—৳{max_val:,.0f}）<br>≈¥{cny_mean:,.1f}</div></div>',
                    unsafe_allow_html=True
                )
        st.markdown("<br>", unsafe_allow_html=True)

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown('<p class="section-title">各产品类型价格分布</p>', unsafe_allow_html=True)
            fig_pb = px.box(price_df, x="产品类型", y="价格_数值",
                            color="产品类型",
                            color_discrete_map={
                                "Disposable Vape": COLORS["primary"],
                                "E-liquid":        COLORS["red"],
                            },
                            points="all", template=T)
            fig_pb.update_layout(showlegend=False, xaxis_title="",
                                  yaxis_title="价格（৳）", margin=dict(t=10, b=10))
            st.plotly_chart(fig_pb, use_container_width=True)
            disp_avg = price_df[price_df["产品类型"] == "Disposable Vape"]["价格_数值"].mean()
            liq_avg = price_df[price_df["产品类型"] == "E-liquid"]["价格_数值"].mean()
            st.caption(f"一次性均价 ৳{disp_avg:,.0f}（≈¥{disp_avg*EXCHANGE_RATE:,.1f}），烟油均价 ৳{liq_avg:,.0f}（≈¥{liq_avg*EXCHANGE_RATE:,.1f}）；一次性产品价格带更集中，竞争更激烈。")

        with col_r:
            st.markdown('<p class="section-title">各品牌来源国均价对比</p>', unsafe_allow_html=True)
            cp = price_df.groupby("品牌来源国")["价格_数值"].mean().reset_index()
            cp = cp.sort_values("价格_数值", ascending=False)
            cp["display"] = cp["价格_数值"].apply(lambda x: f"৳{x:,.0f} (¥{x*EXCHANGE_RATE:,.1f})")
            fig_cp = px.bar(cp, x="品牌来源国", y="价格_数值",
                            text="display",
                            color="品牌来源国",
                            color_discrete_map=COUNTRY_COLOR_MAP,
                            template=T)
            fig_cp.update_traces(texttemplate="%{text}", textposition="outside")
            fig_cp.update_layout(xaxis_title="", yaxis_title="均价（৳）",
                                  showlegend=False, margin=dict(t=10, b=10))
            st.plotly_chart(fig_cp, use_container_width=True)
            top_cp = cp.iloc[0]
            st.caption(f"{top_cp['品牌来源国']}品牌均价最高（৳{top_cp['价格_数值']:,.0f}）；各国品牌价格差异不明显，消费者可选择空间较大。")

        st.markdown('<p class="section-title">Elf Bar vs Voltbar 价格对比（brand popular 数据）</p>',
                    unsafe_allow_html=True)
        if not bp_price.empty:
            brand_price_bp = bp_price.groupby("品牌")["价格_数值"].agg(["mean", "min", "max"]).reset_index()
            brand_price_bp.columns = ["品牌", "均价", "最低价", "最高价"]
            brand_price_bp["display"] = brand_price_bp["均价"].apply(lambda x: f"৳{x:,.0f} (¥{x*EXCHANGE_RATE:,.1f})")
            fig_bpb = px.bar(brand_price_bp, x="品牌", y="均价",
                             text="display",
                             color="品牌",
                             color_discrete_map={"Elf Bar": COLORS["elfbar"], "Voltbar": COLORS["voltbar"]},
                             error_y=brand_price_bp["最高价"] - brand_price_bp["均价"],
                             template=T)
            fig_bpb.update_traces(texttemplate="%{text}", textposition="outside")
            fig_bpb.update_layout(showlegend=False, xaxis_title="",
                                   yaxis_title="均价（৳）", margin=dict(t=10, b=10))
            st.plotly_chart(fig_bpb, use_container_width=True)
            eb_avg = bp_price[bp_price["品牌"] == "Elf Bar"]["价格_数值"].mean()
            vb_avg = bp_price[bp_price["品牌"] == "Voltbar"]["价格_数值"].mean()
            st.caption(
                f"Elf Bar 均价 ৳{eb_avg:,.0f}（≈¥{eb_avg*EXCHANGE_RATE:,.1f}） vs Voltbar 均价 ৳{vb_avg:,.0f}（≈¥{vb_avg*EXCHANGE_RATE:,.1f}）；"
                f"{'Elf Bar' if eb_avg > vb_avg else 'Voltbar'} 定价略高，两者价差不大，竞争主要在口味差异而非价格。"
            )

        st.markdown('<p class="section-title">各网站均价对比</p>', unsafe_allow_html=True)
        site_price = price_df.groupby("网站名称")["价格_数值"].mean().reset_index()
        site_price = site_price.sort_values("价格_数值", ascending=True)
        site_price["display"] = site_price["价格_数值"].apply(lambda x: f"৳{x:,.0f} (¥{x*EXCHANGE_RATE:,.1f})")
        fig_sp = px.bar(site_price, x="价格_数值", y="网站名称", orientation="h",
                        text="display",
                        color="价格_数值",
                        color_continuous_scale=[[0, "#a7f3d0"], [1, COLORS["primary"]]],
                        template=T)
        fig_sp.update_traces(texttemplate="%{text}", textposition="outside")
        fig_sp.update_layout(coloraxis_showscale=False,
                              xaxis_title="均价（৳）", yaxis_title="",
                              margin=dict(t=10, b=10),
                              yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_sp, use_container_width=True)
        st.caption("不同渠道定价差异明显，部分网站以低价引流，部分网站强调精品溢价策略。")


# ============================================================
# TAB 7  总结洞察 + 原始数据
# ============================================================
with tabs[7]:
    st.markdown("## 📝 市场总结与关键洞察")

    top_flavors_list = fdf["口味名称"].value_counts().head(3).index.tolist()
    top_tags_3 = [f"{t}（{c}）" for t, c in Counter(all_tags).most_common(3)]
    ice_pct_total = (fdf["含冰/薄荷"] == "是").sum() / len(fdf)
    china_pct = (fdf["品牌来源国"] == "中国").sum() / len(fdf)
    multi_pct_total = fdf["口味复杂度"].isin(["双重复合", "三重以上复合"]).sum() / len(fdf)

    st.markdown('<div class="danger-card">', unsafe_allow_html=True)
    st.markdown("### ⚖️ 1. 法规窗口期：机遇与风险并存")
    st.markdown("- 2025年12月全面禁令 → 2026年4月议会撤销，**形成短暂的政策窗口期**。")
    st.markdown("- 撤销决定在公共卫生界引发强烈反对，**后续再度收紧的风险不可忽视**。")
    st.markdown("- 建议密切关注政策动向，采取柔性进入策略，避免大规模库存押注。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🥭 2. 热带水果主导，Mango 是最强共识口味")
    st.markdown(f"- 水果口味占比 **{(fdf['分类']=='水果').sum()/len(fdf):.0%}**，是市场绝对主流。")
    st.markdown("- 全市场最热门口味：**VCT**和**Mango Lychee Bubblegum**。")
    st.markdown(f"- 最高频口味元素：**{' / '.join(top_tags_3)}**。")
    st.markdown("- Mango 在零售热销榜与 Daraz 电商平台均高频出现，是进入孟加拉市场的**必备口味**。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🧊 3. 冰感是重要加分项，复合口味成为主流")
    st.markdown(f"- 明确含冰/薄荷产品占比 **{ice_pct_total:.0%}**（实际比例可能更高，部分未标注）。")
    st.markdown(f"- 复合口味（双重+三重以上）占比 **{multi_pct_total:.0%}**，单一口味竞争力趋弱。")
    st.markdown("- **Strawberry 在 Daraz 搜索量遥遥领先（42次SKU），也是孟加拉的热门趋势。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🥇 4. Elf Bar + Voltbar 双雄格局")
    st.markdown("- Elf Bar RAYA D1（13K）与 Voltbar Cartridge（12K）是孟加拉搜索热度最高的品牌之二。")
    st.markdown("- **Voltbar 搜索热度年均更高（37.9 vs 24.4）**，但 Elf Bar 品牌认知度更广。")
    st.markdown("- 两品牌均以 **水果+糖果** 口味为主，**价格接近（约 ৳1,850–2,000）**，口味差异化是主战场。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 💉 5. 50mg 盐尼古丁是市场标准配置")
    st.markdown("- 50mg 盐尼古丁出现次数遥遥领先，是孟加拉一次性市场的行业标准。")
    st.markdown("- 烟油品类提供多档浓度（25–50mg），满足不同需求。")
    st.markdown("- 美国品牌（Vgod、Bar Juice）在烟油品类以烟草口味差异化，与中国品牌水果路线形成互补。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="warning-card">', unsafe_allow_html=True)
    st.markdown(f"### 🌍 6. 中国品牌主导，本土品牌缺席")
    st.markdown(f"- 中国品牌占全市场产品数 **{china_pct:.0%}**，处于绝对统治地位。")
    st.markdown("- 调研数据中**未发现本土孟加拉品牌**，市场基本依赖进口。")
    st.markdown("- 这意味着本土渠道合作伙伴和代理商在市场准入中扮演关键角色。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="action-box">', unsafe_allow_html=True)
    st.markdown("### 📢 7. 产品开发建议")
    st.markdown("- **必备口味**：Mango、Watermelon、Strawberry——三大跨平台共识口味，零售榜+Daraz双高频。")
    st.markdown("- **潜力蓝海**：Bubblegum 系列——在零售高频元素中意外出现较多。")
    st.markdown("- **冰感组合优先**：如 Mango Ice、Watermelon Ice、Strawberry Ice，冰感显著提升吸引力。")
    st.markdown("- **规格建议**：一次性电子烟以12K–13K puffs + 50mg 盐尼古丁；烟油以尼古丁含量25mg、35mg、50mg对标市场主流，降低用户决策门槛。")
    st.markdown("- **定价策略**：参考 ৳1,650–2,000（约 ¥92–112）价格带，对标 Elf Bar / Voltbar，走量先于溢价。")
    st.markdown('</div>', unsafe_allow_html=True)
