"""
Strafor Kalıp Kapasite Hesaplama — Streamlit Uygulaması
=========================================================
Master Liste + Öngörüler dosyaları yüklenince otomatik kapasite/doluluk
hesabı yapar. Blok kesimleri özel olarak ayırır ve %80 üzerini kritik sayar.
"""

import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Strafor Kalıp Kapasite Hesabı", layout="wide")

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
    """Sütun isimlerinde esnek eşleştirme (boşluk/case farkına bakmadan)."""
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
    """'2+1+1+1' -> [2,1,1,1] ; 3 -> [3] ; None/NaN -> []"""
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
    """Malzeme/KALIP NO değerini 6 haneli kod string'ine indirger."""
    s = re.sub(r"\D", "", str(value))
    if len(s) >= 6:
        return s[:6]
    return s.zfill(6) if s else ""


# --------------------------------------------------------------------------
# Master Liste okuma
# --------------------------------------------------------------------------

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

    required = {
        "KALIP NO": col_kalip_no, "GOZ ADEDI": col_goz, "FIRMA": col_firma,
        "Kalıp Çevrim": col_cevrim, "Duran Varlık": col_dv,
    }
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(f"Master Liste'de şu sütunlar bulunamadı: {missing}")

    # Yeni Özellik: Herhangi bir sütunda "blok kesim" ibaresi aranıyor
    blok_kesim_mask = pd.Series(False, index=df.index)
    for col in df.columns:
        if df[col].dtype == object:
            blok_kesim_mask |= df[col].astype(str).str.contains(r"blok\s*kesim", case=False, na=False)

    out = pd.DataFrame({
        "kod": df[col_kalip_no].apply(normalize_kod),
        "tanim": df[col_tanim] if col_tanim else "",
        "kalip_sayisi": pd.to_numeric(df[col_kalip_sayisi], errors="coerce").fillna(1) if col_kalip_sayisi else 1,
        "goz_raw": df[col_goz],
        "firma": df[col_firma].apply(normalize_firma),
        "cevrim": pd.to_numeric(df[col_cevrim], errors="coerce"),
        "duran_varlik": df[col_dv],
        "is_blok_kesim": blok_kesim_mask
    })

    out["parts"] = out["goz_raw"].apply(parse_parts)
    out["parca_sayisi"] = out["parts"].apply(len)
    out["ambiguous"] = out["parca_sayisi"] > 1
    out["parca_per_cevrim"] = out["parts"].apply(sum) * out["kalip_sayisi"]

    bad_cevrim = out["cevrim"].isna() | (out["cevrim"] <= 0)
    if bad_cevrim.any():
        warnings.append(f"{int(bad_cevrim.sum())} satırda Kalıp Çevrim boş/sıfır — bu kalıpların hızı hesaplanamadı.")
    out["hiz"] = np.where(bad_cevrim, np.nan, out["parca_per_cevrim"] / out["cevrim"])

    empty_kod = out["kod"] == ""
    if empty_kod.any():
        warnings.append(f"{int(empty_kod.sum())} satırda KALIP NO okunamadı, bu satırlar hesaba katılmadı.")
    out = out[~empty_kod].copy()

    return out, warnings


# --------------------------------------------------------------------------
# Öngörüler okuma
# --------------------------------------------------------------------------

def load_ongoru(file) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(file)
    warnings = []

    col_malzeme = find_col(df, ["MALZEME"])
    col_firma = find_col(df, ["FIRMA", "FİRMA"])
    col_tanim = find_col(df, ["TANIM", "TANIM"])

    if col_malzeme is None or col_firma is None:
        raise ValueError("Öngörüler dosyasında 'Malzeme' veya 'Firma' sütunu bulunamadı.")

    # Ay sütunlarını bul
    used = {col_malzeme, col_firma, col_tanim} - {None}
    candidate_cols = [c for c in df.columns if c not in used]
    ay_cols = []
    for c in candidate_cols:
        cname = str(c).strip()
        if re.fullmatch(r"\d{5,6}", cname) or pd.api.types.is_numeric_dtype(df[c]):
            ay_cols.append(c)
    ay_cols = ay_cols[:6]
    if len(ay_cols) < 6:
        warnings.append(f"Sadece {len(ay_cols)} aylık sütun tespit edildi (6 bekleniyordu). Sütun isimlerini kontrol edin.")

    # Öngörüler dosyasında da "blok kesim" ibaresi aranıyor (garanti olsun diye)
    blok_kesim_mask = pd.Series(False, index=df.index)
    for col in df.columns:
        if df[col].dtype == object:
            blok_kesim_mask |= df[col].astype(str).str.contains(r"blok\s*kesim", case=False, na=False)

    out = pd.DataFrame({
        "kod": df[col_malzeme].apply(normalize_kod),
        "firma": df[col_firma].apply(normalize_firma),
        "tanim": df[col_tanim] if col_tanim else "",
        "is_blok_kesim": blok_kesim_mask
    })
    for i, ay in enumerate(AYLAR):
        if i < len(ay_cols):
            out[ay] = pd.to_numeric(df[ay_cols[i]], errors="coerce").fillna(0)
        else:
            out[ay] = 0.0

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
        merged["tanim"] = merged.get("tanim_master", pd.Series(index=merged.index, dtype=object)).fillna(
            merged.get("tanim_ongoru", "")
        )
        merged.drop(columns=[c for c in ["tanim_master", "tanim_ongoru"] if c in merged.columns], inplace=True)

    # Master ve Öngörüden gelen blok kesim bilgilerini birleştir
    merged["is_blok_kesim"] = merged.get("is_blok_kesim_master", pd.Series(False, index=merged.index)).fillna(False) | \
                              merged.get("is_blok_kesim_ongoru", pd.Series(False, index=merged.index)).fillna(False)
    merged.drop(columns=[c for c in ["is_blok_kesim_master", "is_blok_kesim_ongoru"] if c in merged.columns], inplace=True, errors="ignore")

    merged["kalip_var"] = merged["birlesik_hiz"].notna()
    for ay in AYLAR:
        if ay not in merged.columns:
            merged[ay] = 0.0
    merged[AYLAR] = merged[AYLAR].fillna(0.0)
    merged["talep_var"] = merged[AYLAR].sum(axis=1) > 0

    # Aylık İhtiyaç Saat / Kapasite Saat / Doluluk %
    cal = calendar_df.set_index(["firma", "ay"])
    for ay in AYLAR:
        ihtiyac_col = f"{ay}_ihtiyac_saat"
        kapasite_col = f"{ay}_kapasite_saat"
        doluluk_col = f"{ay}_doluluk_%"

        def ihtiyac_hesapla(row):
            if not row["kalip_var"] or row["birlesik_hiz"] <= 0:
                return np.nan
            return row[ay] / (row["birlesik_hiz"] * 3600)

        merged[ihtiyac_col] = merged.apply(ihtiyac_hesapla, axis=1)

        def kapasite_getir(row):
            try:
                c = cal.loc[(row["firma"], ay)]
                return float(c["is_gunu"]) * float(c["gunluk_saat"]) * float(c["verimlilik"]) / 100.0
            except KeyError:
                return np.nan

        merged[kapasite_col] = merged.apply(kapasite_getir, axis=1)
        merged[doluluk_col] = np.where(
            merged[kapasite_col].fillna(0) > 0,
            merged[ihtiyac_col] / merged[kapasite_col] * 100,
            np.nan,
        )

    # Maksimum doluluk oranını sütun olarak kaydet (filtreleme için)
    merged["max_doluluk"] = merged[[f"{ay}_doluluk_%" for ay in AYLAR]].max(axis=1).fillna(0)

    # --- DURUM BELİRLEME (YENİ ARALIKLARLA) ---
    merged["durum"] = "Normal"
    merged.loc[merged["talep_var"] & ~merged["kalip_var"], "durum"] = "EŞLEŞMEDİ: Talep var, kalıp yok"
    merged.loc[~merged["talep_var"] & merged["kalip_var"], "durum"] = "BİLGİ: Kalıp var, talep yok"
    
    # 1. Öncelik: Blok Kesim mi?
    merged.loc[merged["is_blok_kesim"] == True, "durum"] = "Blok Kesim (Kalıp Aranmaz)"

    # 2. Normal olanların kapasite aralıklarına göre güncellenmesi
    normal_mask = merged["durum"] == "Normal"
    merged.loc[normal_mask & (merged["max_doluluk"] > 100), "durum"] = "KRİTİK: >%100"
    merged.loc[normal_mask & (merged["max_doluluk"] <= 100) & (merged["max_doluluk"] >= 90), "durum"] = "UYARI: %90-100"
    merged.loc[normal_mask & (merged["max_doluluk"] < 90) & (merged["max_doluluk"] >= 80), "durum"] = "DİKKAT: %80-90"

    # Göz Adedi belirsizlik notu ekleme (Blok kesimlere eklemiyoruz)
    ambig_mask = merged["ambiguous"].fillna(False) & (merged["durum"] != "Blok Kesim (Kalıp Aranmaz)")
    merged.loc[ambig_mask, "durum"] = merged["durum"] + " | Göz Adedi belirsiz"

    return merged


def default_calendar() -> pd.DataFrame:
    rows = []
    for firma in FIRMALAR:
        for ay in AYLAR:
            rows.append({"firma": firma, "ay": ay, "is_gunu": 26, "gunluk_saat": 20, "verimlilik": 90})
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# Streamlit arayüzü
# --------------------------------------------------------------------------

st.title("🏭 Strafor Kalıp Kapasite Hesabı")
st.caption(
    "Master Liste ve Öngörüler dosyalarını yükleyin — sistem (Firma, Kod) "
    "bazında kapasite hesabı yapar. Blok kesimleri ayırır, %80 üzerini kategorize eder."
)

with st.sidebar:
    st.header("1) Dosyalar")
    master_file = st.file_uploader("Master Liste (.xlsx)", type=["xlsx"])
    ongoru_file = st.file_uploader("Öngörüler (.xlsx)", type=["xlsx"])

    st.header("2) Çalışma Takvimi (Firma × Ay)")
    st.caption("İş günü, günlük çalışma saati ve verimlilik (%) — gerekirse düzenleyin.")
    if "calendar_df" not in st.session_state:
        st.session_state.calendar_df = default_calendar()
    st.session_state.calendar_df = st.data_editor(
        st.session_state.calendar_df, num_rows="fixed", width='stretch', key="cal_editor"
    )

if master_file and ongoru_file:
    try:
        master_df, master_warnings = load_master(master_file)
        ongoru_df, ongoru_warnings = load_ongoru(ongoru_file)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    for w in master_warnings + ongoru_warnings:
        st.warning(w)

    result = compute_capacity(master_df, ongoru_df, st.session_state.calendar_df)

    # ---- Özet metrikler ----
    kritik_sayisi = (result["max_doluluk"] >= 80).sum()
    esles_yok = result["durum"].str.contains("EŞLEŞMEDİ").sum()
    blok_kesim_sayisi = result["is_blok_kesim"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam (Firma, Kod)", len(result))
    c2.metric("⚠️ >%80 Doluluk (Kritik/Uyarı)", int(kritik_sayisi))
    c3.metric("🛑 Eşleşmeyen (Talep var)", int(esles_yok))
    c4.metric("🧊 Blok Kesim Kalıplar", int(blok_kesim_sayisi))

    # YENİ EKLENEN SEKMELER (Isı haritası kaldırıldı, Blok Kesim eklendi)
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📋 Genel Tablo", "🚨 Kritik Kalıplar (>= %80)", "🔍 Eşleşmeyen/Belirsiz", "🧊 Blok Kesim", "⬇️ İndir"]
    )

    doluluk_cols = [f"{ay}_doluluk_%" for ay in AYLAR]

    with tab1:
        firma_secim = st.multiselect("Firma filtrele", FIRMALAR, default=FIRMALAR)
        gosterilecek = result[result["firma"].isin(firma_secim)]
        st.dataframe(
            gosterilecek[["firma", "kod", "tanim", "fiziksel_kalip_sayisi", "birlesik_hiz", "durum"] + doluluk_cols],
            width='stretch',
            column_config={
                col: st.column_config.NumberColumn(col, format="%.1f%%") for col in doluluk_cols
            },
        )

    with tab2:
        # %80 ve üzerini filtrele
        kritik = result[result["max_doluluk"] >= 80].sort_values(by="max_doluluk", ascending=False)
        if kritik.empty:
            st.success("Hiçbir kalıp %80 doluluğu geçmiyor.")
        else:
            st.markdown("**%80 ve üzeri doluluğa sahip kalıplar (En yüksek doluluğa göre sıralı):**")
            st.dataframe(
                kritik[["firma", "kod", "tanim", "durum"] + doluluk_cols],
                width='stretch',
                column_config={col: st.column_config.NumberColumn(col, format="%.1f%%") for col in doluluk_cols},
            )
            
            ay_secim = st.selectbox("Grafik için Ay Seçiniz:", AYLAR)
            kritik_ay = kritik[kritik[f"{ay_secim}_doluluk_%"] >= 80]
            
            if not kritik_ay.empty:
                top20 = kritik_ay.nlargest(20, f"{ay_secim}_doluluk_%")
                
                # Grafik Barlarının Rengini Aralıklarına Göre Belirleme
                def get_color(val):
                    if val > 100: return "> %100 (Kritik)"
                    elif val >= 90: return "%90 - %100 (Uyarı)"
                    else: return "%80 - %90 (Dikkat)"
                
                top20["Risk_Seviyesi"] = top20[f"{ay_secim}_doluluk_%"].apply(get_color)
                
                color_map = {
                    "> %100 (Kritik)": "red",
                    "%90 - %100 (Uyarı)": "orange",
                    "%80 - %90 (Dikkat)": "#FFD700" # Altın Sarısı
                }

                fig = px.bar(
                    top20, x="kod", y=f"{ay_secim}_doluluk_%", color="Risk_Seviyesi",
                    title=f"{ay_secim} Ayı — En Kritik Kalıplar (>%80)",
                    labels={f"{ay_secim}_doluluk_%": "Doluluk %", "kod": "Kalıp Kodu"},
                    color_discrete_map=color_map
                )
                fig.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="Kapasite Sınırı")
                fig.add_hline(y=80, line_dash="dash", line_color="orange", annotation_text="%80 Sınırı")
                st.plotly_chart(fig, width='stretch')
            else:
                st.info(f"{ay_secim} ayında %80'i aşan kalıp bulunmuyor.")

    with tab3:
        st.subheader("Talep var ama kalıp eşleşmedi")
        st.dataframe(result[result["durum"].str.contains("EŞLEŞMEDİ")][["firma", "kod", "tanim"] + AYLAR], width='stretch')
        
        st.subheader("Göz Adedi belirsiz (birden çok parça notu — teyit gerekli)")
        st.dataframe(master_df[master_df["ambiguous"]][["firma", "kod", "tanim", "goz_raw", "duran_varlik"]], width='stretch')
        
        st.subheader("Kalıp var, talep yok (bilgi amaçlı)")
        st.dataframe(result[result["durum"].str.contains("talep yok")][["firma", "kod", "tanim"]], width='stretch')

    with tab4:
        st.subheader("🧊 Blok Kesim Olarak İşaretlenenler")
        st.info("Aşağıdaki ürünler (Master ya da Öngörü listesinde 'blok kesim' notu içerdiği için) eşleşmeyenler listesinden ayrılarak buraya taşınmıştır.")
        blok_df = result[result["is_blok_kesim"] == True]
        if not blok_df.empty:
            st.dataframe(blok_df[["firma", "kod", "tanim", "durum"] + AYLAR], width='stretch')
        else:
            st.success("Sistemde 'blok kesim' notu içeren ürün bulunamadı.")

    with tab5:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Kapasite Hesabı", index=False)
            master_df[master_df["ambiguous"]].to_excel(writer, sheet_name="Belirsiz Göz Adedi", index=False)
            
            # Blok kesimleri ayrı bir excel sekmesi olarak da verelim
            if not blok_df.empty:
                blok_df.to_excel(writer, sheet_name="Blok Kesim Kalıplar", index=False)
                
        st.download_button(
            "Excel olarak indir", data=buffer.getvalue(),
            file_name="kapasite_hesabi_sonuc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Başlamak için sol menüden Master Liste ve Öngörüler dosyalarını yükleyin.")
