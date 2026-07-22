"""
Strafor Kalıp Kapasite Hesaplama — Streamlit Uygulaması (Gelişmiş Dashboard Versiyonu)
=========================================================
Master Liste + Öngörüler dosyaları yüklenince otomatik kapasite/doluluk hesabı yapar. 
Tüm hücrelerde "blok kesim" arar. Pasta grafikler ve detaylı aylık analizler sunar.
"""

import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Kalıp Kapasite Dashboard", layout="wide", page_icon="📊")

AYLAR = ["Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
FIRMALAR = ["BASAŞ", "MEFA"]

# --------------------------------------------------------------------------
# Yardımcı fonksiyonlar
# --------------------------------------------------------------------------

_TR_MAP = str.maketrans({
    "ı": "i", "İ": "i", "I": "i", "ş": "s", "Ş": "s", "ğ": "g", "Ğ": "g",
    "ü": "u", "Ü": "u", "ö": "o", "Ö": "o", "ç": "c", "Ç": "c",
})

def normalize_col(name: str) -> str:
    s = re.sub(r"\s+", " ", str(name)).strip().translate(_TR_MAP)
    return s.upper()

def find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    norm_map = {normalize_col(c): c for c in df.columns}
    for alias in aliases:
        a = normalize_col(alias)
        if a in norm_map:
            return norm_map[a]
        for norm_name, orig in norm_map.items():
            if a in norm_name:
                return orig
    return None

def parse_parts(value) -> list[float]:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return []
    if isinstance(value, (int, float)):
        return [float(value)]
    nums = re.findall(r"\d+(?:[.,]\d+)?", str(value))
    return [float(n.replace(",", ".")) for n in nums]

def normalize_firma(value) -> str:
    s = str(value).upper()
    if "BASAŞ" in s or "BASAS" in s:
        return "BASAŞ"
    if "MEFA" in s:
        return "MEFA"
    return str(value).strip()

def normalize_kod(value) -> str:
    s = re.sub(r"\D", "", str(value))
    if len(s) >= 6:
        return s[:6]
    return s.zfill(6) if s else ""

# --------------------------------------------------------------------------
# Dosya Okuma & Garantili Blok Kesim Bulucu
# --------------------------------------------------------------------------

def flag_blok_kesim(df: pd.DataFrame) -> pd.Series:
    """Verilen DataFrame'in TÜM sütunlarını tarar ve 'blok kesim' içeriyorsa True döner."""
    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        mask |= df[col].astype(str).str.contains(r"blok\s*kesim", case=False, na=False)
    return mask

def load_master(file) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(file)
    warnings = []

    col_kalip_no = find_col(df, ["KALIP NO", "KALIPNO"])
    col_tanim = find_col(df, ["KALIP_TANIMI", "KALIP TANIMI", "TANIM"])
    col_kalip_sayisi = find_col(df, ["KALIP SAYISI"])
    col_goz = find_col(df, ["GOZ ADEDI", "GÖZ ADEDİ"])
    col_firma = find_col(df, ["FIRMA", "FİRMA"])
    col_cevrim = find_col(df, ["KALIP ÇEVRİM", "KALIP CEVRIM", "ÇEVRİM"])
    col_dv = find_col(df, ["DURAN VARLIK", "DURAN VARLIK NO"])

    required = {"KALIP NO": col_kalip_no, "GOZ ADEDI": col_goz, "FIRMA": col_firma, "Kalıp Çevrim": col_cevrim}
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(f"Master Liste'de şu sütunlar bulunamadı: {missing}")

    out = pd.DataFrame({
        "kod": df[col_kalip_no].apply(normalize_kod),
        "tanim": df[col_tanim] if col_tanim else "",
        "kalip_sayisi": pd.to_numeric(df[col_kalip_sayisi], errors="coerce").fillna(1) if col_kalip_sayisi else 1,
        "goz_raw": df[col_goz],
        "firma": df[col_firma].apply(normalize_firma),
        "cevrim": pd.to_numeric(df[col_cevrim], errors="coerce"),
        "duran_varlik": df[col_dv] if col_dv else "",
        "is_blok_kesim": flag_blok_kesim(df)  # Tüm tabloyu tarar
    })

    out["parts"] = out["goz_raw"].apply(parse_parts)
    out["parca_sayisi"] = out["parts"].apply(len)
    out["ambiguous"] = out["parca_sayisi"] > 1
    out["parca_per_cevrim"] = out["parts"].apply(sum) * out["kalip_sayisi"]

    bad_cevrim = out["cevrim"].isna() | (out["cevrim"] <= 0)
    out["hiz"] = np.where(bad_cevrim, np.nan, out["parca_per_cevrim"] / out["cevrim"])
    out = out[out["kod"] != ""].copy()

    return out, warnings

def load_ongoru(file) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(file)
    warnings = []

    col_malzeme = find_col(df, ["MALZEME"])
    col_firma = find_col(df, ["FIRMA", "FİRMA"])
    col_tanim = find_col(df, ["TANIM", "TANIM"])

    if col_malzeme is None or col_firma is None:
        raise ValueError("Öngörüler dosyasında 'Malzeme' veya 'Firma' sütunu bulunamadı.")

    used = {col_malzeme, col_firma, col_tanim} - {None}
    candidate_cols = [c for c in df.columns if c not in used]
    ay_cols = [c for c in candidate_cols if re.fullmatch(r"\d{5,6}", str(c).strip()) or pd.api.types.is_numeric_dtype(df[c])][:6]

    out = pd.DataFrame({
        "kod": df[col_malzeme].apply(normalize_kod),
        "firma": df[col_firma].apply(normalize_firma),
        "tanim": df[col_tanim] if col_tanim else "",
        "is_blok_kesim": flag_blok_kesim(df) # Tüm tabloyu tarar
    })
    
    for i, ay in enumerate(AYLAR):
        out[ay] = pd.to_numeric(df[ay_cols[i]], errors="coerce").fillna(0) if i < len(ay_cols) else 0.0

    agg = {ay: "sum" for ay in AYLAR}
    agg["tanim"] = "first"
    agg["is_blok_kesim"] = "any"
    grouped = out.groupby(["kod", "firma"], as_index=False).agg(agg)
    return grouped, warnings

# --------------------------------------------------------------------------
# Kapasite hesabı
# --------------------------------------------------------------------------

def compute_capacity(master_df: pd.DataFrame, ongoru_df: pd.DataFrame, calendar_df: pd.DataFrame):
    grp = master_df.groupby(["firma", "kod"], as_index=False).agg(
        birlesik_hiz=("hiz", lambda s: np.nansum(s) if s.notna().any() else np.nan),
        fiziksel_kalip_sayisi=("duran_varlik", "nunique"),
        ambiguous=("ambiguous", "any"),
        tanim=("tanim", "first"),
        is_blok_kesim=("is_blok_kesim", "any")
    )

    merged = pd.merge(grp, ongoru_df, on=["firma", "kod"], how="outer", suffixes=("_master", "_ongoru"))
    if "tanim_master" in merged.columns or "tanim_ongoru" in merged.columns:
        merged["tanim"] = merged.get("tanim_master", pd.Series(index=merged.index, dtype=object)).fillna(merged.get("tanim_ongoru", ""))
        merged.drop(columns=[c for c in ["tanim_master", "tanim_ongoru"] if c in merged.columns], inplace=True)

    merged["is_blok_kesim"] = merged.get("is_blok_kesim_master", pd.Series(False)).fillna(False) | merged.get("is_blok_kesim_ongoru", pd.Series(False)).fillna(False)
    
    merged["kalip_var"] = merged["birlesik_hiz"].notna()
    for ay in AYLAR:
        if ay not in merged.columns: merged[ay] = 0.0
    merged[AYLAR] = merged[AYLAR].fillna(0.0)
    merged["talep_var"] = merged[AYLAR].sum(axis=1) > 0

    cal = calendar_df.set_index(["firma", "ay"])
    for ay in AYLAR:
        ihtiyac_col, kapasite_col, doluluk_col = f"{ay}_ihtiyac_saat", f"{ay}_kapasite_saat", f"{ay}_doluluk_%"
        merged[ihtiyac_col] = merged.apply(lambda r: r[ay] / (r["birlesik_hiz"] * 3600) if r["kalip_var"] and r["birlesik_hiz"] > 0 else np.nan, axis=1)
        
        def kapasite_getir(row):
            try: return float(cal.loc[(row["firma"], ay)]["is_gunu"]) * float(cal.loc[(row["firma"], ay)]["gunluk_saat"]) * float(cal.loc[(row["firma"], ay)]["verimlilik"]) / 100.0
            except: return np.nan
            
        merged[kapasite_col] = merged.apply(kapasite_getir, axis=1)
        merged[doluluk_col] = np.where(merged[kapasite_col].fillna(0) > 0, merged[ihtiyac_col] / merged[kapasite_col] * 100, np.nan)

    merged["max_doluluk"] = merged[[f"{ay}_doluluk_%" for ay in AYLAR]].max(axis=1).fillna(0)

    # Durum Belirleme
    merged["durum"] = "Normal"
    merged.loc[merged["talep_var"] & ~merged["kalip_var"], "durum"] = "Eşleşmedi (Kalıp Yok)"
    merged.loc[~merged["talep_var"] & merged["kalip_var"], "durum"] = "Talep Yok"
    
    merged.loc[merged["is_blok_kesim"] == True, "durum"] = "Blok Kesim"

    normal_mask = merged["durum"] == "Normal"
    merged.loc[normal_mask & (merged["max_doluluk"] > 100), "durum"] = "Kritik (>%100)"
    merged.loc[normal_mask & (merged["max_doluluk"] <= 100) & (merged["max_doluluk"] >= 90), "durum"] = "Uyarı (%90-100)"
    merged.loc[normal_mask & (merged["max_doluluk"] < 90) & (merged["max_doluluk"] >= 80), "durum"] = "Dikkat (%80-90)"

    return merged

def default_calendar() -> pd.DataFrame:
    rows = [{"firma": f, "ay": a, "is_gunu": 26, "gunluk_saat": 20, "verimlilik": 90} for f in FIRMALAR for a in AYLAR]
    return pd.DataFrame(rows)

# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.title("🏭 Kalıp Kapasite Dashboard")

with st.sidebar:
    st.header("1) Dosyalar")
    master_file = st.file_uploader("Master Liste (.xlsx)", type=["xlsx"])
    ongoru_file = st.file_uploader("Öngörüler (.xlsx)", type=["xlsx"])

    st.header("2) Çalışma Takvimi")
    if "calendar_df" not in st.session_state:
        st.session_state.calendar_df = default_calendar()
    st.session_state.calendar_df = st.data_editor(st.session_state.calendar_df, num_rows="fixed", width='stretch', key="cal_editor")

if master_file and ongoru_file:
    master_df, _ = load_master(master_file)
    ongoru_df, _ = load_ongoru(ongoru_file)
    result = compute_capacity(master_df, ongoru_df, st.session_state.calendar_df)

    # Renk Paleti Tanımlamaları
    COLOR_MAP = {
        "Kritik (>%100)": "#FF4B4B", "Uyarı (%90-100)": "#FFA500", "Dikkat (%80-90)": "#FFD700",
        "Normal": "#00CC96", "Eşleşmedi (Kalıp Yok)": "#636EFA", "Talep Yok": "#AB63FA", "Blok Kesim": "#00B5F7"
    }

    # Metrikler
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam Ürün Sayısı", len(result))
    c2.metric("⚠️ >%80 Riskli Kalıp", int((result["max_doluluk"] >= 80).sum()))
    c3.metric("🛑 Eşleşmeyen", int(result["durum"].str.contains("Eşleşmedi").sum()))
    c4.metric("🧊 Blok Kesim", int(result["is_blok_kesim"].sum()))

    st.markdown("---")

    # Sekmeler
    tab_dash, tab_aylik, tab_liste, tab_blok, tab_indir = st.tabs([
        "📊 Genel Özet", "🥧 Aylık Analiz", "📋 Tüm Liste & Kritikler", "🧊 Blok Kesim", "⬇️ İndir"
    ])

    doluluk_cols = [f"{ay}_doluluk_%" for ay in AYLAR]

    with tab_dash:
        colA, colB = st.columns(2)
        with colA:
            st.subheader("Sistemdeki Tüm Kalıpların Dağılımı")
            durum_counts = result["durum"].value_counts().reset_index()
            durum_counts.columns = ["Durum", "Adet"]
            fig_pie_genel = px.pie(durum_counts, names="Durum", values="Adet", color="Durum", 
                                   color_discrete_map=COLOR_MAP, hole=0.4)
            st.plotly_chart(fig_pie_genel, use_container_width=True)
            
        with colB:
            st.subheader("En Kritik 15 Kalıp (Maksimum Doluluğa Göre)")
            kritik = result[(result["max_doluluk"] >= 80) & (result["durum"] != "Blok Kesim")].nlargest(15, "max_doluluk")
            if not kritik.empty:
                fig_bar_genel = px.bar(kritik, x="kod", y="max_doluluk", color="durum", text="max_doluluk",
                                       color_discrete_map=COLOR_MAP, labels={"kod": "Kalıp", "max_doluluk": "Maks. Doluluk %"})
                fig_bar_genel.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                st.plotly_chart(fig_bar_genel, use_container_width=True)
            else:
                st.success("Sistemde %80 doluluğu aşan riskli kalıp bulunmuyor.")

    with tab_aylik:
        st.subheader("Aylara Göre Darboğaz ve Pasta Grafiği")
        secili_ay = st.selectbox("Analiz Edilecek Ayı Seçin:", AYLAR)
        
        ay_col = f"{secili_ay}_doluluk_%"
        ay_df = result[(result["kalip_var"]) & (result["talep_var"]) & (result["durum"] != "Blok Kesim")].copy()
        
        # O aya özel risk ataması
        def ay_risk(val):
            if pd.isna(val) or val == 0: return "Talep Yok"
            elif val > 100: return "Kritik (>%100)"
            elif val >= 90: return "Uyarı (%90-100)"
            elif val >= 80: return "Dikkat (%80-90)"
            else: return "Normal"
            
        ay_df["Ay_Durumu"] = ay_df[ay_col].apply(ay_risk)
        
        c_pie, c_bar = st.columns([1, 1.5])
        with c_pie:
            ay_counts = ay_df["Ay_Durumu"].value_counts().reset_index()
            ay_counts.columns = ["Durum", "Adet"]
            fig_ay_pie = px.pie(ay_counts, names="Durum", values="Adet", color="Durum", 
                                color_discrete_map=COLOR_MAP, title=f"{secili_ay} Ayı Durumu")
            st.plotly_chart(fig_ay_pie, use_container_width=True)
            
        with c_bar:
            ay_kritik = ay_df[ay_df[ay_col] >= 80].nlargest(15, ay_col)
            if not ay_kritik.empty:
                fig_ay_bar = px.bar(ay_kritik, x="kod", y=ay_col, color="Ay_Durumu",
                                    color_discrete_map=COLOR_MAP, title=f"{secili_ay} Ayı En Yoğun Kalıplar")
                fig_ay_bar.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Kapasite")
                st.plotly_chart(fig_ay_bar, use_container_width=True)
            else:
                st.info(f"{secili_ay} ayında kapasiteyi zorlayan kalıp bulunmamaktadır.")

    with tab_liste:
        st.subheader("Tüm Veri Tablosu")
        st.dataframe(result[["firma", "kod", "tanim", "durum"] + doluluk_cols], width='stretch',
                     column_config={col: st.column_config.NumberColumn(col, format="%.1f%%") for col in doluluk_cols})

    with tab_blok:
        st.subheader("🧊 Blok Kesim Ürünler")
        blok_df = result[result["is_blok_kesim"] == True]
        if not blok_df.empty:
            st.dataframe(blok_df[["firma", "kod", "tanim", "durum"] + AYLAR], width='stretch')
        else:
            st.success("Blok kesim içeren not bulunamadı.")

    with tab_indir:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Kapasite Analizi", index=False)
            if not blok_df.empty:
                blok_df.to_excel(writer, sheet_name="Blok Kesimler", index=False)
        st.download_button("Tüm Sonuçları Excel Olarak İndir", data=buffer.getvalue(),
                           file_name="Dashboard_Kapasite_Analizi.xlsx", mime="application/vnd.ms-excel")
else:
    st.info("Lütfen sol menüden Excel dosyalarınızı yükleyin.")
