"""
Strafor Kalıp Kapasite Hesaplama — Streamlit Uygulaması (v5 - Sade Arayüz)
=========================================================
Hiç Excel/formül bilmeyen biri de kullanabilsin diye arayüz sadeleştirildi.
Hesap mantığı (v4) aynen korundu, sadece ekran tasarımı basitleştirildi.
"""

import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Kalıp Kapasite Kontrolü", layout="wide", page_icon="🏭")

AYLAR = ["Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
FIRMALAR = ["BASAŞ", "MEFA"]

# --------------------------------------------------------------------------
# Yardımcı fonksiyonlar (hesap mantığı — dokunulmadı)
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

def normalize_plaka(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "Bilinmiyor"
    s = re.sub(r"\s+", "", str(value)).upper().replace("X", "*")
    return s if s else "Bilinmiyor"


def normalize_kod(value) -> str:
    s = re.sub(r"\D", "", str(value))
    if len(s) >= 6:
        return s[:6]
    return s.zfill(6) if s else ""

# --------------------------------------------------------------------------
# Blok Kesim çıkarma
# --------------------------------------------------------------------------

def _extract_blok_kesim_from_sheet(df: pd.DataFrame) -> list[dict]:
    """Tek bir sayfada 'blok kesim' geçen satırları çıkarır."""
    malzeme_col = find_col(df, ["MALZEME"])
    firma_col = find_col(df, ["FIRMA", "FİRMA"])
    tanim_col = find_col(df, ["TANIM", "KALIP_TANIMI", "KALIP TANIMI"])
    if malzeme_col is None or firma_col is None:
        return []

    used = {malzeme_col, firma_col, tanim_col} - {None}
    candidate_cols = [c for c in df.columns if c not in used]
    ay_cols = [c for c in candidate_cols if re.fullmatch(r"\d{5,6}", str(c).strip()) or pd.api.types.is_numeric_dtype(df[c])][:6]

    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        mask |= df[col].astype(str).str.contains(r"blok\s*kesim", case=False, na=False)
    blok_df = df[mask]

    out = []
    for _, row in blok_df.iterrows():
        f = normalize_firma(row[firma_col])
        orijinal_kod = str(row[malzeme_col]).strip()
        k = normalize_kod(row[malzeme_col])
        tanim = row[tanim_col] if tanim_col else ""
        ay_degerleri = {AYLAR[i]: (row[ay_cols[i]] if i < len(ay_cols) else 0) for i in range(len(AYLAR))}
        out.append({
            "firma": f, "6_haneli_kod": k, "Orijinal_Malzeme_Kodu": orijinal_kod,
            "tanim": tanim, **ay_degerleri, "is_blok_kesim": True,
        })
    return out


def extract_blok_kesim_list(*files) -> pd.DataFrame:
    """Yüklenen tüm dosyaların TÜM sayfalarını tarayıp 'blok kesim' geçen satırları toplar."""
    all_rows = []
    for file in files:
        if file is None:
            continue
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                try:
                    df = pd.read_excel(file, sheet_name=sheet)
                    all_rows.extend(_extract_blok_kesim_from_sheet(df))
                except Exception:
                    continue
        except Exception:
            continue

    result = pd.DataFrame(all_rows)
    if not result.empty:
        result = result.drop_duplicates(subset=["firma", "6_haneli_kod", "Orijinal_Malzeme_Kodu"])
    return result

# --------------------------------------------------------------------------
# Dosya okuma
# --------------------------------------------------------------------------

def load_master(file) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(file, sheet_name=0)
    warnings = []

    col_kalip_no = find_col(df, ["KALIP NO", "KALIPNO"])
    col_tanim = find_col(df, ["KALIP_TANIMI", "KALIP TANIMI", "TANIM"])
    col_kalip_sayisi = find_col(df, ["KALIP SAYISI"])
    col_goz = find_col(df, ["GOZ ADEDI", "GÖZ ADEDİ"])
    col_firma = find_col(df, ["FIRMA", "FİRMA"])
    col_cevrim = find_col(df, ["KALIP ÇEVRİM", "KALIP CEVRIM", "ÇEVRİM"])
    col_dv = find_col(df, ["DURAN VARLIK", "DURAN VARLIK NO"])
    col_plaka = find_col(df, ["PLAKA EBATI", "PLAKA EBADI", "PLAKA"])

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
        "plaka": df[col_plaka].apply(normalize_plaka) if col_plaka else "Bilinmiyor",
    })

    out["parts"] = out["goz_raw"].apply(parse_parts)
    out["ambiguous"] = out["parts"].apply(len) > 1
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
    })
    for i, ay in enumerate(AYLAR):
        out[ay] = pd.to_numeric(df[ay_cols[i]], errors="coerce").fillna(0) if i < len(ay_cols) else 0.0

    agg = {ay: "sum" for ay in AYLAR}
    agg["tanim"] = "first"
    grouped = out.groupby(["kod", "firma"], as_index=False).agg(agg)
    return grouped, warnings

# --------------------------------------------------------------------------
# Kapasite hesabı (mantık aynı — dokunulmadı)
# --------------------------------------------------------------------------

def compute_capacity(master_df, ongoru_df, blok_kesim_df, calendar_df):
    grp = master_df.groupby(["firma", "kod"], as_index=False).agg(
        birlesik_hiz=("hiz", lambda s: np.nansum(s) if s.notna().any() else np.nan),
        fiziksel_kalip_sayisi=("duran_varlik", "nunique"),
        ambiguous=("ambiguous", "any"),
        tanim=("tanim", "first"),
    )

    merged = pd.merge(grp, ongoru_df, on=["firma", "kod"], how="outer", suffixes=("_master", "_ongoru"))
    if "tanim_master" in merged.columns or "tanim_ongoru" in merged.columns:
        merged["tanim"] = merged.get("tanim_master", pd.Series(index=merged.index, dtype=object)).fillna(merged.get("tanim_ongoru", ""))
        merged.drop(columns=[c for c in ["tanim_master", "tanim_ongoru"] if c in merged.columns], inplace=True)

    if not blok_kesim_df.empty:
        blok_6_li = blok_kesim_df[["firma", "6_haneli_kod", "is_blok_kesim"]].rename(columns={"6_haneli_kod": "kod"}).drop_duplicates()
        merged = pd.merge(merged, blok_6_li, on=["firma", "kod"], how="left")
    else:
        merged["is_blok_kesim"] = False
    merged["is_blok_kesim"] = merged["is_blok_kesim"].fillna(False)

    merged["kalip_var"] = merged["birlesik_hiz"].notna()
    for ay in AYLAR:
        if ay not in merged.columns:
            merged[ay] = 0.0
    merged[AYLAR] = merged[AYLAR].fillna(0.0)
    merged["talep_var"] = merged[AYLAR].sum(axis=1) > 0

    cal = calendar_df.set_index(["firma", "ay"])
    for ay in AYLAR:
        ihtiyac_col, kapasite_col, doluluk_col = f"{ay}_ihtiyac_saat", f"{ay}_kapasite_saat", f"{ay}_doluluk_%"
        merged[ihtiyac_col] = merged.apply(
            lambda r: r[ay] / (r["birlesik_hiz"] * 3600) if r["kalip_var"] and r["birlesik_hiz"] > 0 else np.nan, axis=1
        )

        def kapasite_getir(row):
            try:
                c = cal.loc[(row["firma"], ay)]
                return float(c["is_gunu"]) * float(c["gunluk_saat"]) * float(c["verimlilik"]) / 100.0
            except Exception:
                return np.nan

        merged[kapasite_col] = merged.apply(kapasite_getir, axis=1)
        merged[doluluk_col] = np.where(
            merged[kapasite_col].fillna(0) > 0, merged[ihtiyac_col] / merged[kapasite_col] * 100, np.nan
        )

    doluluk_cols = [f"{ay}_doluluk_%" for ay in AYLAR]
    merged["max_doluluk"] = merged[doluluk_cols].max(axis=1).fillna(0)
    merged["en_yogun_ay"] = merged[doluluk_cols].fillna(-1).idxmax(axis=1).str.replace("_doluluk_%", "", regex=False)

    merged["durum"] = "🟢 Normal"
    merged.loc[merged["talep_var"] & ~merged["kalip_var"], "durum"] = "🔵 Kalıp bulunamadı"
    merged.loc[~merged["talep_var"] & merged["kalip_var"], "durum"] = "⚪ Talep yok"
    merged.loc[merged["is_blok_kesim"] == True, "durum"] = "🧊 Blok kesim"

    normal_mask = merged["durum"] == "🟢 Normal"
    merged.loc[normal_mask & (merged["max_doluluk"] > 100), "durum"] = "🔴 Kapasite Aşıldı"
    merged.loc[normal_mask & (merged["max_doluluk"] <= 100) & (merged["max_doluluk"] >= 90), "durum"] = "🟠 Sınırda"
    merged.loc[normal_mask & (merged["max_doluluk"] < 90) & (merged["max_doluluk"] >= 80), "durum"] = "🟡 Dikkat"

    return merged


def default_calendar() -> pd.DataFrame:
    rows = [{"firma": f, "ay": a, "is_gunu": 26, "gunluk_saat": 20, "verimlilik": 90} for f in FIRMALAR for a in AYLAR]
    return pd.DataFrame(rows)


BUCKET_SIRASI = ["0-10%", "10-20%", "20-30%", "30-40%", "40-50%", "50-60%", "60-70%", "70-80%", "80-90%", "90-100%", "100%+"]
BUCKET_RENK = {
    "0-10%": "#B3E5B4", "10-20%": "#A6DDA8", "20-30%": "#99D59C", "30-40%": "#8CCD90",
    "40-50%": "#F7E27A", "50-60%": "#F5D25C", "60-70%": "#F2C23E", "70-80%": "#F0B220",
    "80-90%": "#FFA500", "90-100%": "#FF7B00", "100%+": "#FF4B4B",
}

def doluluk_bucket(val: float) -> str | None:
    if pd.isna(val):
        return None
    if val >= 100:
        return "100%+"
    idx = min(int(val // 10), 9)
    return BUCKET_SIRASI[idx]


# --------------------------------------------------------------------------
# ARAYÜZ — SADE VE HERKESİN ANLAYACAĞI ŞEKİLDE
# --------------------------------------------------------------------------

st.title("🏭 Kalıp Kapasite Kontrolü")
st.write("İki dosyayı yükle, hangi kalıpların yetişemeyeceğini otomatik gör.")

# ---- SIDEBAR: sadece 2 adım + isteğe bağlı ayarlar ----
with st.sidebar:
    st.subheader("1️⃣ Dosyaları Yükle")
    master_file = st.file_uploader("Kalıp Listesi (Master Liste)", type=["xlsx"])
    ongoru_file = st.file_uploader("Talep / Öngörü Listesi", type=["xlsx"])

    with st.expander("⚙️ Gelişmiş: Çalışma Takvimi"):
        st.caption("İş günü, günlük saat ve verimlilik — değiştirmek istemiyorsan dokunma, varsayılan değerler kullanılır.")
        if "calendar_df" not in st.session_state:
            st.session_state.calendar_df = default_calendar()
        st.session_state.calendar_df = st.data_editor(
            st.session_state.calendar_df, num_rows="fixed", width="stretch", key="cal_editor"
        )

if not (master_file and ongoru_file):
    st.info("⬅️ Başlamak için soldan iki dosyayı yükle.")
    st.stop()

# ---- Hesaplama ----
with st.spinner("Hesaplanıyor..."):
    blok_df_raw = extract_blok_kesim_list(master_file, ongoru_file)
    master_df, _ = load_master(master_file)
    ongoru_df, _ = load_ongoru(ongoru_file)
    result = compute_capacity(master_df, ongoru_df, blok_df_raw, st.session_state.calendar_df)

doluluk_cols = [f"{ay}_doluluk_%" for ay in AYLAR]

# ---- Hangi satırların doluluğu gerçekten hesaplanabiliyor? ----
# Blok kesim -> zaten kalıp aranmıyor. Kalıp bulunamadı -> hesaplanamaz (tedarikçi teyidi gerekir).
result["hesaplanabilir"] = result["kalip_var"] & result["talep_var"] & ~result["is_blok_kesim"]

hesaplanamayan_tedarikci = result[result["talep_var"] & ~result["kalip_var"]].copy()   # 🔵
hesaplanamayan_blok = result[result["is_blok_kesim"]].copy()                            # 🧊
hesaplanabilenler = result[result["hesaplanabilir"]].copy()
hesaplanabilenler["bucket"] = hesaplanabilenler["max_doluluk"].apply(doluluk_bucket)

n_toplam_urun = len(result)
n_hesaplanabilir = len(hesaplanabilenler)
n_80_plus = int((hesaplanabilenler["max_doluluk"] >= 80).sum())
n_100_plus = int((hesaplanabilenler["max_doluluk"] >= 100).sum())
n_tedarikci = len(hesaplanamayan_tedarikci)
n_blok = len(hesaplanamayan_blok)

# ---- ÖZET KARTLARI ----
st.subheader("Durum Özeti")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Toplam Ürün", n_toplam_urun)
c2.metric("🔴 %100 Üstü (Kapasite Aşıldı)", n_100_plus)
c3.metric("🟠 %80 ve Üstü", n_80_plus, help="Doluluk oranı %80'i geçen, yakından takip edilmesi gereken kalıplar")
c4.metric("❓ Hesaplanamayan", n_tedarikci + n_blok, help="Tedarikçi teyidi bekleyen + blok kesimden gelen ürünler")

if n_100_plus > 0:
    st.error(f"⚠️ **{n_100_plus} ürün, elindeki kalıp kapasitesini aşıyor.** '🚨 Acil' sekmesinden bak.")
else:
    st.success("✅ Şu an kapasitesini aşan ürün yok.")

st.divider()

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
    [
        "📊 Genel Dağılım", "⏱️ Toplam Süre", "🚨 Acil (%80+)", "🏢 Firma Bazlı En Dolu 10",
        "🏭 Makine Bazlı Üretim", "🔎 Tüm Kalıpları Ara", "❓ Hesaplanamayanlar", "⬇️ İndir",
    ]
)

# ---- TAB 1: Genel Dağılım (üstte genel pasta, altta AYRI bir bölümde aylık pasta grid) ----
with tab1:
    st.write(f"Doluluğu hesaplanabilen **{n_hesaplanabilir}** ürünün, 6 aylık dönemdeki EN YOĞUN ayına göre dağılımı:")
    bucket_counts = hesaplanabilenler["max_doluluk"].apply(doluluk_bucket).value_counts().reindex(BUCKET_SIRASI).fillna(0).astype(int)
    bucket_counts = bucket_counts[bucket_counts > 0].reset_index()
    bucket_counts.columns = ["Doluluk Aralığı", "Ürün Sayısı"]
    if not bucket_counts.empty:
        fig = px.pie(
            bucket_counts, names="Doluluk Aralığı", values="Ürün Sayısı",
            color="Doluluk Aralığı", color_discrete_map=BUCKET_RENK, hole=0.35,
            category_orders={"Doluluk Aralığı": BUCKET_SIRASI},
        )
        fig.update_traces(sort=False)
        st.plotly_chart(fig, width="stretch")

    st.caption(
        f"Not: {n_tedarikci + n_blok} ürün bu grafiğe dahil değil (❓ Hesaplanamayanlar sekmesine bak) "
        "çünkü ya kalıbı bulunamadı ya da blok kesimden üretiliyor."
    )

    st.divider()
    st.subheader("📅 Aylara Göre Dağılım")
    st.caption("Her ayın kendi doluluk dağılımı, ayrı ayrı:")
    for satir_baslangic in range(0, len(AYLAR), 3):
        cols = st.columns(3)
        for i, ay in enumerate(AYLAR[satir_baslangic:satir_baslangic + 3]):
            with cols[i]:
                ay_bucket = hesaplanabilenler[f"{ay}_doluluk_%"].apply(doluluk_bucket).value_counts().reindex(BUCKET_SIRASI).fillna(0).astype(int)
                ay_bucket = ay_bucket[ay_bucket > 0].reset_index()
                ay_bucket.columns = ["Doluluk Aralığı", "Ürün Sayısı"]
                if ay_bucket.empty:
                    st.info(f"{ay}: veri yok")
                    continue
                fig_ay = px.pie(
                    ay_bucket, names="Doluluk Aralığı", values="Ürün Sayısı",
                    color="Doluluk Aralığı", color_discrete_map=BUCKET_RENK, hole=0.35, title=ay,
                )
                fig_ay.update_traces(sort=False, textinfo="value")
                fig_ay.update_layout(showlegend=False, margin=dict(t=40, b=0, l=0, r=0), height=260)
                st.plotly_chart(fig_ay, width="stretch")
    st.caption("Renk skalası üstteki genel grafikle aynı: yeşil düşük doluluk, kırmızı %100 üstü.")

# ---- TAB 2: Toplam Süre (bar chart + hesaplama açıklaması) ----
with tab2:
    st.subheader("Aylık Toplam Süre İhtiyacı")
    st.write("Firma bazında, her ay toplamda kaç saat kalıp/makine zamanı gerektiği:")

    sure_satirlari = []
    for firma in FIRMALAR:
        alt = hesaplanabilenler[hesaplanabilenler["firma"] == firma]
        makine_sayisi = master_df[master_df["firma"] == firma]["plaka"].nunique()
        makine_sayisi = max(makine_sayisi, 1)
        for ay in AYLAR:
            kapasite_deger = alt[f"{ay}_kapasite_saat"].mean()
            kapasite_tek = 0 if pd.isna(kapasite_deger) else kapasite_deger
            sure_satirlari.append({
                "Ay": ay, "Firma": firma,
                "İhtiyaç Saat": alt[f"{ay}_ihtiyac_saat"].sum(),
                "Kapasite Saat": kapasite_tek * makine_sayisi,
                "Makine Sayısı": makine_sayisi,
            })
    sure_df = pd.DataFrame(sure_satirlari)
    st.caption(
        f"Kapasite Saat burada **makine sayısıyla çarpılmış** toplam kapasiteyi gösteriyor "
        f"(BASAŞ: {master_df[master_df['firma']=='BASAŞ']['plaka'].nunique()} makine, "
        f"MEFA: {master_df[master_df['firma']=='MEFA']['plaka'].nunique()} makine — "
        "'Plaka Ebatı' farklı olan her kalıp ayrı bir makine sayıldı)."
    )

    fig_sure = px.bar(
        sure_df, x="Ay", y="İhtiyaç Saat", color="Firma", barmode="group",
        category_orders={"Ay": AYLAR}, text_auto=".0f",
    )
    st.plotly_chart(fig_sure, width="stretch")

    with st.expander("ℹ️ Bu grafik nasıl hesaplanıyor?"):
        st.markdown(
            """
Her ürün için önce **birleşik üretim hızı** bulunur — o ürünü üretebilen tüm kalıpların
(Kalıp Adet ÷ Kalıp Çevrim) değerleri toplanarak elde edilir (çevrim süreleri farklı
olsa bile doğru sonuç verir):

`Birleşik Hız = Σ (Kalıp Adet / Kalıp Çevrim)`  → parça/saniye

Sonra o ayki talebi karşılamak için gereken süre hesaplanır:

`İhtiyaç Saat = Talep Adedi / (Birleşik Hız × 3600)`

Bu grafikteki her bar, **o firmanın o aydaki tüm ürünlerinin İhtiyaç Saat toplamıdır**:

`Toplam İhtiyaç Saat = Σ (her ürünün kendi İhtiyaç Saati)`
            """
        )

    st.divider()
    st.subheader("İhtiyaç Saat / Kapasite Saat Karşılaştırması")
    fig_kars = px.bar(
        sure_df.melt(id_vars=["Ay", "Firma"], value_vars=["İhtiyaç Saat", "Kapasite Saat"], var_name="Tür", value_name="Saat"),
        x="Ay", y="Saat", color="Tür", barmode="group", facet_col="Firma",
        category_orders={"Ay": AYLAR},
    )
    st.plotly_chart(fig_kars, width="stretch")
    st.caption("Kapasite Saat = (İş Günü × Günlük Saat × Verimlilik) × Makine Sayısı.")

# ---- TAB 3: Acil (%80 ve üzeri) ----
with tab3:

    st.write("Doluluk oranı %80 ve üzerinde olan, yakından takip edilmesi gereken ürünler:")
    acil = hesaplanabilenler[hesaplanabilenler["max_doluluk"] >= 80].sort_values("max_doluluk", ascending=False)
    if acil.empty:
        st.success("Şu an %80 üzerinde doluluğa sahip ürün yok. 🎉")
    else:
        st.dataframe(
            acil[["firma", "kod", "tanim", "en_yogun_ay", "max_doluluk"]].rename(columns={
                "firma": "Firma", "kod": "Kalıp Kodu", "tanim": "Ürün Adı",
                "en_yogun_ay": "En Yoğun Ay", "max_doluluk": "En Yüksek Doluluk %",
            }),
            width="stretch", hide_index=True,
            column_config={
                "En Yüksek Doluluk %": st.column_config.ProgressColumn(
                    "En Yüksek Doluluk %", format="%.0f%%", min_value=0, max_value=200
                )
            },
        )
        st.caption("İpucu: Doluluk %100'ü geçiyorsa, o kalıp elindeki süre içinde talebi karşılayamıyor demektir.")

# ---- TAB 4: Firma Bazlı En Dolu 10 ----
with tab4:
    st.write("Her firmanın en dolu (en yoğun) 10 kalıbı, ayrı ayrı:")
    col_basas, col_mefa = st.columns(2)
    for firma, kolon in zip(FIRMALAR, [col_basas, col_mefa]):
        with kolon:
            st.markdown(f"#### {firma}")
            top10 = hesaplanabilenler[hesaplanabilenler["firma"] == firma].nlargest(10, "max_doluluk")
            if top10.empty:
                st.info("Veri yok.")
                continue
            st.dataframe(
                top10.sort_values("max_doluluk", ascending=False)[["firma", "kod", "tanim", "en_yogun_ay", "max_doluluk"]].rename(columns={
                    "firma": "Firma", "kod": "Kalıp Kodu", "tanim": "Ürün Adı",
                    "en_yogun_ay": "Ay", "max_doluluk": "Doluluk %",
                }),
                width="stretch", hide_index=True,
                column_config={
                    "Doluluk %": st.column_config.ProgressColumn(
                        "Doluluk %", format="%.0f%%", min_value=0, max_value=max(200, int(top10["max_doluluk"].max()) + 20)
                    )
                },
            )

# ---- TAB 5: Makine Bazlı Üretim (Plaka Ebatı = Makine Tipi) ----
with tab5:
    st.write(
        "Her farklı **Plaka Ebatı**, bir makine tipini temsil eder. "
        "Aşağıda, her firmanın hangi makinede kaç farklı ürün ürettiği gösteriliyor."
    )
    col_b, col_m = st.columns(2)
    for firma, kolon in zip(FIRMALAR, [col_b, col_m]):
        with kolon:
            st.markdown(f"#### {firma}")
            alt_master = master_df[master_df["firma"] == firma]
            makine_sayim = alt_master.groupby("plaka")["kod"].nunique().sort_values(ascending=False).reset_index()
            makine_sayim.columns = ["Makine (Plaka Ebatı)", "Ürün Sayısı"]
            if makine_sayim.empty:
                st.info("Veri yok.")
                continue
            fig_makine = px.bar(
                makine_sayim, x="Makine (Plaka Ebatı)", y="Ürün Sayısı", text_auto=True,
            )
            st.plotly_chart(fig_makine, width="stretch")
    st.caption("Not: Plaka Ebatı bilgisi Master Liste'den geliyor. Aynı ölçüdeki farklı yazımlar (örn. '817*980' / '817x980') otomatik birleştirildi.")

# ---- TAB 6: Arama ----
with tab6:
    arama = st.text_input("🔎 Kalıp kodu veya ürün adı yaz", placeholder="örn: 484070 veya köşe takviye")
    firma_secim = st.multiselect("Firma", FIRMALAR, default=FIRMALAR)


    gosterilecek = hesaplanabilenler[hesaplanabilenler["firma"].isin(firma_secim)]
    if arama:
        arama_l = arama.lower()
        gosterilecek = gosterilecek[
            gosterilecek["kod"].astype(str).str.lower().str.contains(arama_l)
            | gosterilecek["tanim"].astype(str).str.lower().str.contains(arama_l)
        ]

    detay_goster = st.checkbox("Ay ay detayı göster (6 ay ayrı ayrı)")
    if detay_goster:
        kolonlar = ["firma", "kod", "tanim"] + doluluk_cols
        col_config = {col: st.column_config.NumberColumn(col.replace("_doluluk_%", ""), format="%.0f%%") for col in doluluk_cols}
    else:
        kolonlar = ["firma", "kod", "tanim", "en_yogun_ay", "max_doluluk"]
        col_config = {"max_doluluk": st.column_config.ProgressColumn("En Yüksek Doluluk %", format="%.0f%%", min_value=0, max_value=200)}

    st.dataframe(
        gosterilecek[kolonlar].rename(columns={
            "firma": "Firma", "kod": "Kalıp Kodu", "tanim": "Ürün Adı", "en_yogun_ay": "En Yoğun Ay",
        }),
        width="stretch", hide_index=True, column_config=col_config,
    )
    st.caption("Bu liste sadece doluluğu hesaplanabilen ürünleri gösterir. Diğerleri için '❓ Hesaplanamayanlar' sekmesine bak.")

# ---- TAB 7: Hesaplanamayanlar (tedarikçi teyidi + blok kesim, ayrı ayrı) ----
with tab7:
    st.write("Bu ürünlerin doluluk oranı hesaplanamıyor — nedenleri farklı olduğu için iki ayrı grupta gösteriliyor.")

    st.markdown(f"#### 🔵 Tedarikçi ile doğrulanması gerekenler ")
    st.caption("Bu ürünler için talep var ama Kalıp Listesi'nde eşleşen bir kalıp bulunamadı. Kalıp bilgisi tedarikçiden/kalıphaneden teyit edilmeli.")
    if hesaplanamayan_tedarikci.empty:
        st.success("Bu grupta ürün yok.")
    else:
        st.dataframe(
            hesaplanamayan_tedarikci[["firma", "kod", "tanim"] + AYLAR].rename(
                columns={"firma": "Firma", "kod": "Kod", "tanim": "Ürün Adı"}
            ),
            width="stretch", hide_index=True,
        )

    st.markdown(f"#### 🧊 Blok kesimden gelenler ")
    st.caption("Bu ürünler kalıptan değil, blok kesimden üretiliyor — kapasite hesabına hiç dahil edilmiyor.")
    if hesaplanamayan_blok.empty and blok_df_raw.empty:
        st.success("Bu grupta ürün yok.")
    elif not blok_df_raw.empty:
        st.dataframe(
            blok_df_raw[["firma", "Orijinal_Malzeme_Kodu", "tanim"]].rename(
                columns={"firma": "Firma", "Orijinal_Malzeme_Kodu": "Malzeme Kodu", "tanim": "Ürün Adı"}
            ),
            width="stretch", hide_index=True,
        )

# ---- TAB 8: İndir ----
with tab8:
    st.write("Hesaplanan tüm sonuçları Excel dosyası olarak indirebilirsin — her grup ayrı sayfada.")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        hesaplanabilenler.drop(columns=["bucket"]).to_excel(writer, sheet_name="Doluluk Hesaplanan", index=False)
        hesaplanamayan_tedarikci.to_excel(writer, sheet_name="Tedarikci Teyidi Gereken", index=False)
        if not blok_df_raw.empty:
            blok_df_raw.to_excel(writer, sheet_name="Blok Kesimler", index=False)
    st.download_button(
        "📥 Excel Olarak İndir", data=buffer.getvalue(),
        file_name="Kalip_Kapasite_Sonuclari.xlsx",
        mime="application/vnd.ms-excel",
    )
