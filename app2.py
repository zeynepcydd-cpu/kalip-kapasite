"""
Strafor Kalıp Kapasite Hesaplama — Streamlit Uygulaması
=========================================================
Master Liste + Öngörüler dosyaları yüklenince otomatik kapasite/doluluk
hesabı yapar, %100 üstü ve eşleşmeyen kritik kalıpları ayrı gösterir.

Çalıştırmak için:
    pip install streamlit pandas numpy plotly openpyxl
    streamlit run kapasite_app.py
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

    out = pd.DataFrame({
        "kod": df[col_kalip_no].apply(normalize_kod),
        "tanim": df[col_tanim] if col_tanim else "",
        "kalip_sayisi": pd.to_numeric(df[col_kalip_sayisi], errors="coerce").fillna(1) if col_kalip_sayisi else 1,
        "goz_raw": df[col_goz],
        "firma": df[col_firma].apply(normalize_firma),
        "cevrim": pd.to_numeric(df[col_cevrim], errors="coerce"),
        "duran_varlik": df[col_dv],
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

    # Ay sütunlarını bul: sayısal başlıklı (YYYYMM) ilk 6 sütunu al
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

    out = pd.DataFrame({
        "kod": df[col_malzeme].apply(normalize_kod),
        "firma": df[col_firma].apply(normalize_firma),
        "tanim": df[col_tanim] if col_tanim else "",
    })
    for i, ay in enumerate(AYLAR):
        if i < len(ay_cols):
            out[ay] = pd.to_numeric(df[ay_cols[i]], errors="coerce").fillna(0)
        else:
            out[ay] = 0.0

    # Aynı (firma, kod) için birden fazla malzeme satırı olabilir (yoğunluk
    # varyantları gibi) — talebi topla, tanımlardan ilkini örnek olarak tut.
    agg = {ay: "sum" for ay in AYLAR}
    agg["tanim"] = "first"
    grouped = out.groupby(["kod", "firma"], as_index=False).agg(agg)
    return grouped, warnings


# --------------------------------------------------------------------------
# Kapasite hesabı
# --------------------------------------------------------------------------

def compute_capacity(master_df: pd.DataFrame, ongoru_df: pd.DataFrame, calendar_df: pd.DataFrame):
    # (Firma, Kod) bazında birleşik hız + kaç fiziksel kalıp + belirsizlik bayrağı
    grp = master_df.groupby(["firma", "kod"], as_index=False).agg(
        birlesik_hiz=("hiz", lambda s: np.nansum(s) if s.notna().any() else np.nan),
        fiziksel_kalip_sayisi=("duran_varlik", "nunique"),
        ambiguous=("ambiguous", "any"),
        tanim=("tanim", "first"),
    )

    merged = pd.merge(grp, ongoru_df, on=["firma", "kod"], how="outer", suffixes=("_master", "_ongoru"))
    if "tanim_master" in merged.columns or "tanim_ongoru" in merged.columns:
        merged["tanim"] = merged.get("tanim_master", pd.Series(index=merged.index, dtype=object)).fillna(
            merged.get("tanim_ongoru", "")
        )
        merged.drop(columns=[c for c in ["tanim_master", "tanim_ongoru"] if c in merged.columns], inplace=True)

    merged["kalip_var"] = merged["birlesik_hiz"].notna()
    for ay in AYLAR:
        if ay not in merged.columns:
            merged[ay] = 0.0
    merged[AYLAR] = merged[AYLAR].fillna(0.0)
    merged["talep_var"] = merged[AYLAR].sum(axis=1) > 0

    merged["durum"] = "Normal"
    merged.loc[merged["talep_var"] & ~merged["kalip_var"], "durum"] = "KRİTİK: Talep var, kalıp eşleşmedi"
    merged.loc[~merged["talep_var"] & merged["kalip_var"], "durum"] = "Bilgi: Kalıp var, talep yok"
    merged.loc[merged["ambiguous"].fillna(False), "durum"] = merged["durum"] + " | Göz Adedi belirsiz (teyit gerekli)"

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

    merged.loc[merged["durum"] == "Normal", "durum"] = merged.loc[merged["durum"] == "Normal"].apply(
        lambda r: "KRİTİK: %100 üstü doluluk" if any(
            (r.get(f"{ay}_doluluk_%") or 0) > 100 for ay in AYLAR
        ) else "Normal", axis=1
    )

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
    "Master Liste ve Öngörüler dosyalarını yükleyin — sistem (Firma, Altılı Kod) "
    "bazında kalıp hızlarını birleştirip aylık ihtiyaç/kapasite/doluluk hesaplar."
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
    kritik_100 = result["durum"].str.contains("100 üstü").sum()
    esles_yok = result["durum"].str.contains("eşleşmedi").sum()
    belirsiz = result["durum"].str.contains("belirsiz").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam (Firma, Kod) grubu", len(result))
    c2.metric("⚠️ %100 üstü doluluk", int(kritik_100))
    c3.metric("🛑 Talep var, kalıp yok", int(esles_yok))
    c4.metric("❓ Göz Adedi belirsiz", int(belirsiz))

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📋 Genel Tablo", "🚨 Kritik Kalıplar", "🌡️ Isı Haritası", "🔍 Eşleşmeyen/Belirsiz", "⬇️ İndir"]
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
        kritik = result[result["durum"].str.contains("100 üstü")]
        if kritik.empty:
            st.success("Hiçbir kalıp %100 doluluğu geçmiyor.")
        else:
            st.dataframe(
                kritik[["firma", "kod", "tanim"] + doluluk_cols],
                width='stretch',
                column_config={col: st.column_config.NumberColumn(col, format="%.1f%%") for col in doluluk_cols},
            )
            ay_secim = st.selectbox("Ay seç (grafik için)", AYLAR)
            top20 = kritik.nlargest(20, f"{ay_secim}_doluluk_%")
            fig = px.bar(
                top20, x="kod", y=f"{ay_secim}_doluluk_%", color="firma",
                title=f"{ay_secim} — En Kritik 20 Kalıp", labels={f"{ay_secim}_doluluk_%": "Doluluk %"},
            )
            fig.add_hline(y=100, line_dash="dash", line_color="red")
            st.plotly_chart(fig, width='stretch')

    with tab3:
        heat_df = result.melt(
            id_vars=["firma", "kod"], value_vars=doluluk_cols, var_name="ay", value_name="doluluk"
        )
        heat_df["ay"] = heat_df["ay"].str.replace("_doluluk_%", "", regex=False)
        pivot = heat_df.pivot_table(index=["firma", "kod"], columns="ay", values="doluluk")
        pivot = pivot[[a for a in AYLAR if a in pivot.columns]]
        fig2 = px.imshow(
            pivot, aspect="auto", color_continuous_scale="RdYlGn_r", zmin=0, zmax=150,
            labels=dict(color="Doluluk %"),
        )
        st.plotly_chart(fig2, width='stretch')

    with tab4:
        st.subheader("Talep var ama kalıp eşleşmedi")
        st.dataframe(result[result["durum"].str.contains("eşleşmedi")][["firma", "kod", "tanim"] + AYLAR])
        st.subheader("Göz Adedi belirsiz (birden çok parça notu — teyit gerekli)")
        st.dataframe(
            master_df[master_df["ambiguous"]][["firma", "kod", "tanim", "goz_raw", "duran_varlik"]]
        )
        st.subheader("Kalıp var, talep yok (bilgi amaçlı)")
        st.dataframe(result[result["durum"].str.contains("Kalıp var, talep yok")][["firma", "kod", "tanim"]])

    with tab5:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Kapasite Hesabı", index=False)
            master_df[master_df["ambiguous"]].to_excel(writer, sheet_name="Belirsiz Göz Adedi", index=False)
        st.download_button(
            "Excel olarak indir", data=buffer.getvalue(),
            file_name="kapasite_hesabi_sonuc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Başlamak için sol menüden Master Liste ve Öngörüler dosyalarını yükleyin.")
