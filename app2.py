import streamlit as st
import pandas as pd
import plotly.express as px

# Sayfa ayarları
st.set_page_config(page_title="Kalıp Kapasite Analizi", layout="wide")
st.title("🏭 Kalıp Kapasite ve Darboğaz Analiz Aracı")

st.markdown("""
Bu araç, master kalıp listesi ve öngörülen üretim adetlerini karşılaştırarak 
kapasite aşımı riski taşıyan kalıpları tespit eder.
""")

# Sütun düzeni ile dosya yükleme alanları
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Master Liste Yükle (Excel/CSV)", type=["xlsx", "csv"])
with col2:
    forecast_file = st.file_uploader("Öngörü Listesi Yükle (Excel/CSV)", type=["xlsx", "csv"])

if master_file and forecast_file:
    try:
        # Dosyaları okuma
        df_master = pd.read_excel(master_file) if master_file.name.endswith('.xlsx') else pd.read_csv(master_file)
        df_forecast = pd.read_excel(forecast_file) if forecast_file.name.endswith('.xlsx') else pd.read_csv(forecast_file)
        
        st.success("Dosyalar başarıyla yüklendi!")

        # 1. Öngörü listesinde her bir kalıp için toplam üretilecek adetleri topla
        kalip_talepleri = df_forecast.groupby('Kalıp_Kodu')['Talep_Adedi'].sum().reset_index()

        # 2. Master liste ile birleştir (Merge)
        df_merged = pd.merge(kalip_talepleri, df_master, on='Kalıp_Kodu', how='left')

        # 3. Gerekli Zaman (Kapasite) Hesabı
        df_merged['Gerekli_Baski'] = df_merged['Talep_Adedi'] / df_merged['Kavite_Sayısı']
        df_merged['Gerekli_Sure_Saat'] = (df_merged['Gerekli_Baski'] * df_merged['Çevrim_Süresi_sn']) / 3600
        
        # 4. Doluluk Oranı (Kapasite Kullanım Oranı)
        df_merged['Doluluk_Orani_%'] = (df_merged['Gerekli_Sure_Saat'] / df_merged['Aylık_Kapasite_Saat']) * 100
        df_merged['Doluluk_Orani_%'] = df_merged['Doluluk_Orani_%'].round(2)
        
        # 5. Kritik Kalıpları Filtreleme (Örn: %90 ve üzeri)
        st.subheader("⚠️ Kapasite Aşımı veya Kritik Seviyedeki Kalıplar (>= %90)")
        kritik_kaliplar = df_merged[df_merged['Doluluk_Orani_%'] >= 90].sort_values(by='Doluluk_Orani_%', ascending=False)
        
        if not kritik_kaliplar.empty:
            # --- İNTERAKTİF GRAFİK ---
            fig = px.bar(
                kritik_kaliplar,
                x='Kalıp_Kodu',
                y='Doluluk_Orani_%',
                color='Doluluk_Orani_%',
                color_continuous_scale='Reds',
                title="Kritik Kalıpların Doluluk Oranları (%)",
                labels={'Kalıp_Kodu': 'Kalıp Kodu', 'Doluluk_Orani_%': 'Doluluk Oranı (%)'},
                text='Doluluk_Orani_%'
            )
            
            fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            
            # Y eksenini veriye göre dinamik ayarla ki grafik şık dursun
            max_oran = max(kritik_kaliplar['Doluluk_Orani_%'])
            fig.update_layout(yaxis=dict(range=[0, max_oran + (max_oran * 0.1)]))
            
            st.plotly_chart(fig, use_container_width=True)

            # --- DETAY TABLOSU ---
            st.markdown("### 📋 Detaylı Veri Tablosu")
            st.dataframe(
                kritik_kaliplar[['Kalıp_Kodu', 'Talep_Adedi', 'Aylık_Kapasite_Saat', 'Gerekli_Sure_Saat', 'Doluluk_Orani_%']],
                use_container_width=True
            )
        else:
            st.info("Kapasitesi %90'ı aşan kalıp bulunmamaktadır. Sistem rahat.")

        # Tüm tabloyu indirme butonu
        st.markdown("---")
        csv = df_merged.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Tüm Analiz Sonuçlarını İndir (CSV)",
            data=csv,
            file_name='kapasite_analiz_sonucu.csv',
            mime='text/csv',
        )

    except Exception as e:
        st.error(f"Veri işlenirken bir hata oluştu. Excel/CSV sütun isimlerinizin koddakilerle birebir aynı olduğundan emin olun. Hata detayı: {e}")