import os
import shutil
import difflib
import base64
from datetime import datetime
import pandas as pd
import yaml
import streamlit as st
import streamlit.components.v1 as components

from arsiv_yoneticisi import ArsivYoneticisi
from ai_servisi import AIServisi
from belge_ureticisi import BelgeUreticisi, para_formatla

st.set_page_config(page_title="Ödeme Yazısı Otomasyonu", layout="wide")

# İlgili alana yumuşak kaydırma (scroll) yaptıran yardımcı fonksiyon
def alana_kaydir(element_id: str):
    js = f"""
    <script>
        const el = window.parent.document.getElementById('{element_id}');
        if (el) {{
            el.scrollIntoView({{behavior: 'smooth', block: 'center'}});
        }}
    </script>
    """
    components.html(js, height=0)

# Yapılandırmaları yükle
@st.cache_resource
def ayarlari_yukle():
    with open("config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    with open("sablon_cumleleri.yaml", "r", encoding="utf-8") as f:
        sablonlar = yaml.safe_load(f)
    return config, sablonlar

config, sablon_cumleleri = ayarlari_yukle()

# Loglama fonksiyonu
def log_yaz(tur, firma, tutar, baz_alinan, skor, ai_kullanildi):
    log_yolu = "cikti/uretim_kayitlari.csv"
    kayit = {
        "tarih_saat": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "odeme_turu": tur,
        "firma": firma,
        "tutar": tutar,
        "baz_alinan_dosya": baz_alinan,
        "benzerlik_skoru": f"{skor:.2f}",
        "ai_kullanildi": "Evet" if ai_kullanildi else "Hayır"
    }
    df = pd.DataFrame([kayit])
    if not os.path.exists(log_yolu):
        df.to_csv(log_yolu, index=False, encoding="utf-8-sig")
    else:
        df.to_csv(log_yolu, mode="a", header=False, index=False, encoding="utf-8-sig")

# Metin farklarını renkli gösteren yardımcı fonksiyon
def farklari_renklendir(eski, yeni):
    matcher = difflib.SequenceMatcher(None, eski.split(), yeni.split())
    sonuc = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            sonuc.append(" ".join(yeni.split()[j1:j2]))
        elif tag in ['replace', 'insert']:
            sonuc.append(f"<span style='background-color: #d4edda; color: #155724; font-weight: bold;'>{' '.join(yeni.split()[j1:j2])}</span>")
    return " ".join(sonuc)

# Arşiv Yöneticisini Başlat
if "arsiv" not in st.session_state:
    st.session_state.arsiv = ArsivYoneticisi("arsiv")

arsiv = st.session_state.arsiv

st.title("📄 Resmi Ödeme Yazısı Hazırlama Sistemi")

# Üst Bilgi / Arşiv Yenile Butonu
col_info, col_btn = st.columns([5, 1])
with col_info:
    st.caption(f"Arşivde toplam {len(arsiv.belgeler)} adet taranmış ödeme yazısı mevcut.")
with col_btn:
    if st.button("🔄 Arşivi Yenile"):
        arsiv.indeksle()
        st.success("Arşiv güncellendi!")
        st.rerun()

if arsiv.okunamayanlar:
    with st.expander("⚠️ Okunamayan Dosyalar"):
        for o in arsiv.okunamayanlar:
            st.write(f"- **{o['dosya']}** ({o['tur']}): {o['hata']}")

st.divider()

# Form Alanları
col1, col2 = st.columns(2)

turler = arsiv.turleri_getir()
if not turler:
    st.warning("Arşivde henüz ödeme türü klasörü bulunamadı. Lütfen 'arsiv' klasörüne klasörlerinizi ekleyin.")
    st.stop()

with col1:
    secilen_tur = st.selectbox("1. Ödeme Türü", turler)
    
    # Firma Listesi
    try:
        df_firmalar = pd.read_excel("firmalar.xlsx")
        firma_listesi = df_firmalar["Firma_Kisa_Ad"].tolist()
        secilen_firma_kisa = st.selectbox("2. Firma Seçimi", firma_listesi)
        secilen_firma_satir = df_firmalar[df_firmalar["Firma_Kisa_Ad"] == secilen_firma_kisa].iloc[0]
        unvan_varsayilan = str(secilen_firma_satir.get("Firma_Unvan", ""))
        banka_varsayilan = str(secilen_firma_satir.get("Banka", ""))
        iban_varsayilan = str(secilen_firma_satir.get("IBAN", ""))
    except Exception as e:
        st.error(f"firmalar.xlsx okunamadı: {e}")
        st.stop()

    st.markdown('<div id="alan_firma_unvan"></div>', unsafe_allow_html=True)
    firma_unvan = st.text_input("Firma Tam Unvanı *", value=unvan_varsayilan)
    
    banka_bilgisi = st.text_input("Banka / Şube", value=banka_varsayilan)
    
    st.markdown('<div id="alan_iban"></div>', unsafe_allow_html=True)
    iban_bilgisi = st.text_input("IBAN Numarası *", value=iban_varsayilan)

    # İmzacılar
    try:
        df_imzacilar = pd.read_excel("imzacilar.xlsx")
        imzaci_secenekleri = [f"{row['Ad_Soyad']} - {row['Unvan']}" for _, row in df_imzacilar.iterrows()]
        
        st.markdown('<div id="alan_imzacilar"></div>', unsafe_allow_html=True)
        secilen_imzaci_metinleri = st.multiselect(
            "3. İmzalayacak Kişiler (1-3 Kişi) *", 
            imzaci_secenekleri, 
            default=[imzaci_secenekleri[0]] if imzaci_secenekleri else None
        )
        imzacilar = []
        for im in secilen_imzaci_metinleri:
            ad, unvan = im.split(" - ", 1)
            imzacilar.append({"ad_soyad": ad, "unvan": unvan})
    except Exception as e:
        st.error(f"imzacilar.xlsx okunamadı: {e}")
        st.stop()

with col2:
    col_ay, col_yil = st.columns(2)
    aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
    with col_ay:
        secilen_ay = st.selectbox("Ay", aylar, index=datetime.now().month - 1)
    with col_yil:
        secilen_yil = st.selectbox("Yıl", [2025, 2026, 2027], index=1)

    st.markdown('<div id="alan_tutar"></div>', unsafe_allow_html=True)
    tutar_str = st.text_input("KDV'siz Tutar (TL) *", value="10.000,00", help="Örnek format: 15.000,00")
    
    # Türkçe formatı sayıya çevirme
    try:
        temiz_tutar = tutar_str.replace(".", "").replace(",", ".")
        tutar_sayi = float(temiz_tutar)
    except:
        tutar_sayi = 0.0

    kdv_orani = st.selectbox("KDV Oranı (%)", config["finans"]["kdv_oranlari"], index=config["finans"]["kdv_oranlari"].index(config["finans"]["varsayilan_kdv"]))
    kdvli_tutar_sayi = tutar_sayi * (1 + (kdv_orani / 100))
    st.info(f"Hesaplanan KDV Dahil Tutar: **{para_formatla(kdvli_tutar_sayi)} TL**")

    # Ödeme Cümlesi Kalıbı
    sablon_anahtarlari = list(sablon_cumleleri.keys())
    sablon_adlari = [sablon_cumleleri[k]["ad"] for k in sablon_anahtarlari]
    secilen_sablon_ad = st.selectbox("Ödeme Şablonu", sablon_adlari)
    secilen_sablon_key = sablon_anahtarlari[sablon_adlari.index(secilen_sablon_ad)]

    # Bütçe Tablosu (Opsiyonel)
    with st.expander("📊 Bütçe Tablosu Ekle (Opsiyonel)"):
        tablo_aktif = st.checkbox("Bütçe tablosunu dahil et", value=False)
        toplam_butce = st.number_input("Toplam Bütçe (TL)", value=0.0, step=1000.0)
        kullanilan_butce = st.number_input("Daha Önce Kullanılan Bütçe (TL)", value=0.0, step=1000.0)

st.subheader("Hizmet ve Kapsam Tanımı")
onerilen_kapsamlar = arsiv.turdeki_kapsam_onerileri(secilen_tur)
if onerilen_kapsamlar:
    secilen_oneri = st.selectbox("Eski Kapsamlardan Seç veya Aşağıya Yaz", ["(Yeni Metin Gireceğim)"] + onerilen_kapsamlar)
    kapsam_taslak = "" if secilen_oneri == "(Yeni Metin Gireceğim)" else secilen_oneri
else:
    kapsam_taslak = ""

st.markdown('<div id="alan_kapsam"></div>', unsafe_allow_html=True)
kullanici_kapsam = st.text_area("Hizmet Açıklaması / Detayı *", value=kapsam_taslak, height=100, placeholder="Örnek: 15-16 Nisan tarihlerinde İTÜ Çekirdek etkinliği için afiş baskısı alındı.")

# Oturuma durum kaydetme
if "islenmis_konu" not in st.session_state:
    st.session_state.islenmis_konu = ""
if "islenmis_kapsam" not in st.session_state:
    st.session_state.islenmis_kapsam = ""
if "en_benzer_yazi" not in st.session_state:
    st.session_state.en_benzer_yazi = None
if "en_yuksek_skor" not in st.session_state:
    st.session_state.en_yuksek_skor = 0.0

if st.button("🔍 Yazıyı Hazırla ve Önizle", type="primary"):
    # Zorunlu alan doğrulamaları
    eksik = False
    
    if not firma_unvan.strip():
        st.error("⚠️ Lütfen 'Firma Tam Unvanı' alanını doldurunuz!")
        alana_kaydir("alan_firma_unvan")
        eksik = True
    elif not iban_bilgisi.strip():
        st.error("⚠️ Lütfen 'IBAN Numarası' alanını doldurunuz!")
        alana_kaydir("alan_iban")
        eksik = True
    elif len(imzacilar) == 0:
        st.error("⚠️ Lütfen en az 1 imzalayacak kişi seçiniz!")
        alana_kaydir("alan_imzacilar")
        eksik = True
    elif tutar_sayi <= 0:
        st.error("⚠️ Lütfen geçerli bir 'KDV'siz Tutar' giriniz (0'dan büyük olmalıdır)!")
        alana_kaydir("alan_tutar")
        eksik = True
    elif not kullanici_kapsam.strip():
        st.error("⚠️ Lütfen 'Hizmet Açıklaması / Detayı' alanını doldurunuz!")
        alana_kaydir("alan_kapsam")
        eksik = True

    if not eksik:
        with st.spinner("Arşiv taranıyor ve metin hazırlanıyor..."):
            benzerler = arsiv.benzerleri_bul(kullanici_kapsam, secilen_tur, limit=3)
            ai_kullanildi = False
            
            if benzerler and benzerler[0][1] >= 0.90:
                # Neredeyse birebir eşleşme
                st.session_state.en_benzer_yazi = benzerler[0][0]
                st.session_state.en_yuksek_skor = benzerler[0][1]
                st.session_state.islenmis_konu = st.session_state.en_benzer_yazi["konu"]
                st.session_state.islenmis_kapsam = kullanici_kapsam
            else:
                # Gemini ile kurumsal üsluba göre üret
                try:
                    ai = AIServisi(
                        model_name=config["gemini"]["model_name"],
                        temperature=config["gemini"]["temperature"]
                    )
                    konu, kapsam = ai.konu_ve_kapsam_uret(kullanici_kapsam, benzerler)
                    st.session_state.islenmis_konu = konu
                    st.session_state.islenmis_kapsam = kapsam
                    ai_kullanildi = True
                    if benzerler:
                        st.session_state.en_benzer_yazi = benzerler[0][0]
                        st.session_state.en_yuksek_skor = benzerler[0][1]
                    else:
                        st.session_state.en_benzer_yazi = None
                        st.session_state.en_yuksek_skor = 0.0
                except Exception as err:
                    st.error(f"Yapay zekâ servisi hatası: {err}")
                    st.stop()

# Önizleme Ekranı
if st.session_state.islenmis_kapsam:
    st.divider()
    st.header("📋 Önizleme ve Doğrulama")
    
    skor = st.session_state.en_yuksek_skor
    esik = config["arama"]["esik_degeri"]

    if skor < esik:
        st.warning(f"⚠️ Arşivde yakın örnek bulunamadı (Benzerlik Skoru: %{skor*100:.1f}). Kapsam yapay zekâ ile yazıldı, lütfen dikkatle kontrol edin.")
    else:
        st.success(f"✅ Baz Alınan Belge: **{st.session_state.en_benzer_yazi['dosya_adi']}** (Benzerlik: %{skor*100:.1f})")

    # Yan Yana Kıyaslama
    col_eski, col_yeni = st.columns(2)
    with col_eski:
        st.markdown("**Arşivdeki Orijinal Kapsam:**")
        eski_metin = st.session_state.en_benzer_yazi["kapsam"] if st.session_state.en_benzer_yazi else "(Benzer eski belge yok)"
        st.info(eski_metin)

    with col_yeni:
        st.markdown("**Yeni Kapsam (Farklar Vurgulu):**")
        if st.session_state.en_benzer_yazi:
            fark_html = farklari_renklendir(eski_metin, st.session_state.islenmis_kapsam)
            st.markdown(f"<div style='padding: 10px; border: 1px solid #ccc; border-radius: 5px;'>{fark_html}</div>", unsafe_allow_html=True)
        else:
            st.write(st.session_state.islenmis_kapsam)

    st.write("")
    duzenlenen_konu = st.text_input("Konu Satırı (Düzenlenebilir):", value=st.session_state.islenmis_konu)
    duzenlenen_kapsam = st.text_area("Kapsam Paragrafı (Düzenlenebilir):", value=st.session_state.islenmis_kapsam, height=120)

    # Ödeme cümlesini üret
    sablon_ham = sablon_cumleleri[secilen_sablon_key]["metin"]
    odeme_cumlesi_uretilen = sablon_ham.replace("{{ tutar_kdvsiz }}", para_formatla(tutar_sayi))\
                                        .replace("{{ tutar_kdvli }}", para_formatla(kdvli_tutar_sayi))\
                                        .replace("{{ firma_unvan }}", firma_unvan)

    st.markdown("**Oluşacak Ödeme Cümlesi:**")
    st.code(odeme_cumlesi_uretilen, language="text")

    # Onay kutusu ve doğrudan indirme
    onay = st.checkbox("Metinleri, tutarları ve hesaplamaları kontrol ettim, onaylıyorum.")
    arsive_ekle = st.checkbox("Bu yazıyı ilerideki aramalar için arşive de dahil et", value=False)

    if st.button("💾 Word Belgesini Oluştur ve İndir", disabled=not onay, type="primary"):
        belge_verisi = {
            "ay": secilen_ay,
            "yil": secilen_yil,
            "sayi_formati": config["belge"]["sayi_formati"],
            "hitap": config["belge"]["hitap"],
            "kapanis": config["belge"]["kapanis"],
            "konu": duzenlenen_konu,
            "kapsam": duzenlenen_kapsam,
            "odeme_cumlesi": odeme_cumlesi_uretilen,
            "firma_unvan": firma_unvan,
            "banka": banka_bilgisi,
            "iban": iban_bilgisi,
            "tutar_kdvsiz": para_formatla(tutar_sayi),
            "tutar_kdvli": para_formatla(kdvli_tutar_sayi),
            "tutar_kdvsiz_sayi": tutar_sayi,
            "imzacilar": imzacilar,
            "tablo_aktif": tablo_aktif,
            "butce_toplam": toplam_butce,
            "butce_kullanilan": kullanilan_butce,
            "butce_odenecek": tutar_sayi
        }

        dosya_adi = f"{secilen_tur}_{secilen_ay}_{secilen_yil}_Odeme_Yazisi.docx"
        cikti_tam_yol = os.path.join("cikti", dosya_adi)

        uretici = BelgeUreticisi("sablonlar/sablon.docx")
        basarili, mesaj = uretici.uret(belge_verisi, cikti_tam_yol)

        if basarili:
            st.success(f"🎉 Belge başarıyla oluşturuldu ve indirme başlatıldı: `{dosya_adi}`")
            log_yaz(secilen_tur, firma_unvan, tutar_sayi, 
                    st.session_state.en_benzer_yazi["dosya_adi"] if st.session_state.en_benzer_yazi else "-",
                    skor, True)
            
            if arsive_ekle:
                hedef_klasor = os.path.join("arsiv", secilen_tur)
                os.makedirs(hedef_klasor, exist_ok=True)
                shutil.copy(cikti_tam_yol, os.path.join(hedef_klasor, dosya_adi))
                st.info("Belge arşive kaydedildi.")

            # Dosyayı oku ve Base64 üzerinden tarayıcıya otomatik indirme yaptır
            with open(cikti_tam_yol, "rb") as fp:
                dosya_baytlari = fp.read()
                b64 = base64.b64encode(dosya_baytlari).decode()

            otomatik_indir_js = f"""
            <script>
                const link = window.parent.document.createElement('a');
                link.href = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}';
                link.download = '{dosya_adi}';
                window.parent.document.body.appendChild(link);
                link.click();
                window.parent.document.body.removeChild(link);
            </script>
            """
            components.html(otomatik_indir_js, height=0)

            # Tarayıcı popup/otomatik indirme kısıtlamalarına karşı yedek buton
            st.download_button(
                label="📥 Otomatik indirme başlamadıysa buraya tıklayın",
                data=dosya_baytlari,
                file_name=dosya_adi,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error(f"Hata: {mesaj}")