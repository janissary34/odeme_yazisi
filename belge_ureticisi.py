import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def para_formatla(deger):
    try:
        val = float(deger)
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(deger)


def hucre_kenarlik_ekle(hucre):
    tcPr = hucre._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'CCCCCC')
        tcBorders.append(border)
    tcPr.append(tcBorders)


class BelgeUreticisi:
    def __init__(self, sablon_yolu="sablonlar/sablon.docx"):
        self.sablon_yolu = sablon_yolu

    def uret(self, veri, cikti_yolu):
        try:
            # Şablonu aç: Header (Logo) ve Footer (Adres bilgisi) korunur
            if os.path.exists(self.sablon_yolu):
                doc = Document(self.sablon_yolu)
                # Gövdedeki eski taslak paragraflarını temizle
                for p in list(doc.paragraphs):
                    p._element.getparent().remove(p._element)
                for t in list(doc.tables):
                    t._element.getparent().remove(t._element)
            else:
                doc = Document()

            ana_punto = float(veri.get("punto", 14))
            kucuk_punto = max(8.5, ana_punto - 2.5)

            # Sayfa kenar boşlukları
            for s in doc.sections:
                s.top_margin = Inches(0.8)
                s.bottom_margin = Inches(0.8)
                s.left_margin = Inches(1.1)
                s.right_margin = Inches(1.1)

            # 1. TARİH (Sağa Yaslı)
            p_tarih = doc.add_paragraph()
            p_tarih.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_tarih.paragraph_format.space_before = Pt(0)
            p_tarih.paragraph_format.space_after = Pt(12)
            r_tarih = p_tarih.add_run(f"{veri.get('ay', '')} {veri.get('yil', '')}")
            r_tarih.font.name = "Calibri"
            r_tarih.font.size = Pt(ana_punto)

            # 2. SAYI & KONU
            ham_sayi = str(veri.get("sayi_formati", "T-YK")).replace("Sayı:", "").strip()
            p_sayi = doc.add_paragraph()
            p_sayi.paragraph_format.space_before = Pt(0)
            p_sayi.paragraph_format.space_after = Pt(2)
            r_sayi = p_sayi.add_run(f"Sayı: {ham_sayi}")
            r_sayi.font.name = "Calibri"
            r_sayi.font.size = Pt(ana_punto)

            ham_konu = str(veri.get("konu", "")).replace("Konu:", "").strip()
            p_konu = doc.add_paragraph()
            p_konu.paragraph_format.space_before = Pt(0)
            p_konu.paragraph_format.space_after = Pt(22)
            r_konu = p_konu.add_run(f"Konu: {ham_konu}")
            r_konu.font.name = "Calibri"
            r_konu.font.size = Pt(ana_punto)

            # 3. HİTAP (Ortalı)
            p_hitap = doc.add_paragraph()
            p_hitap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_hitap.paragraph_format.space_before = Pt(0)
            p_hitap.paragraph_format.space_after = Pt(20)
            r_hitap = p_hitap.add_run(veri.get("hitap", "ARI TEKNOKENT GENEL MÜDÜRLÜĞÜ'NE"))
            r_hitap.bold = True
            r_hitap.font.name = "Calibri"
            r_hitap.font.size = Pt(ana_punto)

            # 4. KAPSAM PARAGRAFI
            p_kapsam = doc.add_paragraph()
            p_kapsam.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_kapsam.paragraph_format.space_before = Pt(0)
            p_kapsam.paragraph_format.space_after = Pt(14)
            p_kapsam.paragraph_format.line_spacing = 1.15
            r_kapsam = p_kapsam.add_run(veri.get("kapsam", ""))
            r_kapsam.font.name = "Calibri"
            r_kapsam.font.size = Pt(ana_punto)

            # 5. ÖDEME CÜMLESİ
            p_odeme = doc.add_paragraph()
            p_odeme.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_odeme.paragraph_format.space_before = Pt(0)
            p_odeme.paragraph_format.space_after = Pt(14)
            p_odeme.paragraph_format.line_spacing = 1.15

            tutar_kdvsiz = str(veri.get("tutar_kdvsiz", ""))
            tutar_kdvli = str(veri.get("tutar_kdvli", ""))
            firma_unvan = str(veri.get("firma_unvan", "")).strip()

            r1 = p_odeme.add_run("Söz konusu hizmetin ")
            r_toplam = p_odeme.add_run("toplam")
            r_toplam.bold = True
            r2 = p_odeme.add_run(f" bedeli {tutar_kdvsiz} TL + KDV (")
            r_kdvli = p_odeme.add_run(f"{tutar_kdvli} TL")
            r_kdvli.bold = True
            r3 = p_odeme.add_run(") olup , bu bedelin ")
            r_firma = p_odeme.add_run(firma_unvan)
            r_firma.bold = True
            r4 = p_odeme.add_run("’ne ödenmesi için gereğini izninize arz ederim.")

            for r in [r1, r_toplam, r2, r_kdvli, r3, r_firma, r4]:
                r.font.name = "Calibri"
                r.font.size = Pt(ana_punto)

            # 6. DİNAMİK NOT SATIRI (Checkbox Seçimlerine Göre)
            ek_sozlesme = veri.get("ek_sozlesme", False)
            ek_komisyon = veri.get("ek_komisyon", False)
            not_metni = ""

            if ek_sozlesme and ek_komisyon:
                not_metni = "Not: Sözleşme/ek protokol ve komisyon kararı ile fiyat tekliflerine ekte yer verilmiştir."
            elif ek_sozlesme:
                not_metni = "Not: Sözleşme ve devamı olarak hazırlanan ek protokol belgesi ilaveten sunulmuştur."
            elif ek_komisyon:
                not_metni = "Not: Komisyon kararı ve fiyat tekliflerine ekte yer verilmiştir."

            if not_metni:
                p_not = doc.add_paragraph()
                p_not.paragraph_format.space_before = Pt(0)
                p_not.paragraph_format.space_after = Pt(16)
                r_not = p_not.add_run(not_metni)
                r_not.font.name = "Calibri"
                r_not.font.size = Pt(max(9.0, ana_punto - 1.5))

            # 7. SAYGILARIMLA
            p_saygi = doc.add_paragraph()
            p_saygi.paragraph_format.space_before = Pt(0)
            p_saygi.paragraph_format.space_after = Pt(14)
            r_saygi = p_saygi.add_run("Saygılarımla,")
            r_saygi.font.name = "Calibri"
            r_saygi.font.size = Pt(ana_punto)

            # 8. BÜTÇE TABLOSU (Aktifse)
            if veri.get("tablo_aktif", False):
                toplam_b = float(veri.get("butce_toplam", 0.0))
                kullanilan_b = float(veri.get("butce_kullanilan", 0.0))
                odenecek_b = float(veri.get("butce_odenecek", 0.0))
                kalan_b = toplam_b - kullanilan_b - odenecek_b

                tablo = doc.add_table(rows=2, cols=5)
                tablo.alignment = WD_TABLE_ALIGNMENT.CENTER

                basliklar = ["Bütçe Kaynağı", "Toplam Bakiye", "Kullanılan Bütçe", "Ödenecek Tutar", "Kalan Bakiye"]
                degerler = [
                    "Pazarlama PR Çekirdek-Ajans",
                    f"{para_formatla(toplam_b)}TL",
                    f"{para_formatla(kullanilan_b)}TL",
                    f"{para_formatla(odenecek_b)}TL",
                    f"{para_formatla(kalan_b)}TL"
                ]

                for i, b in enumerate(basliklar):
                    cell = tablo.cell(0, i)
                    cell.text = b
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell.paragraphs[0].runs[0].bold = True
                    cell.paragraphs[0].runs[0].font.name = "Calibri"
                    cell.paragraphs[0].runs[0].font.size = Pt(kucuk_punto)
                    hucre_kenarlik_ekle(cell)

                for i, d in enumerate(degerler):
                    cell = tablo.cell(1, i)
                    cell.text = d
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_run = cell.paragraphs[0].runs[0]
                    p_run.font.name = "Calibri"
                    p_run.font.size = Pt(kucuk_punto)
                    if i == 4:
                        p_run.bold = True
                    hucre_kenarlik_ekle(cell)

                p_bosluk = doc.add_paragraph()
                p_bosluk.paragraph_format.space_before = Pt(12)

            # 9. İMZACILAR ALANI
            imzacilar = veri.get("imzacilar", [])
            if imzacilar:
                p_imza_bosluk = doc.add_paragraph()
                p_imza_bosluk.paragraph_format.space_before = Pt(10)
                for imzaci in imzacilar:
                    p_imza = doc.add_paragraph()
                    p_imza.paragraph_format.space_before = Pt(0)
                    p_imza.paragraph_format.space_after = Pt(2)
                    r_imzaci_ad = p_imza.add_run(imzaci.get("ad_soyad", ""))
                    r_imzaci_ad.bold = True
                    r_imzaci_ad.font.name = "Calibri"
                    r_imzaci_ad.font.size = Pt(ana_punto)

                    p_unvan = doc.add_paragraph()
                    p_unvan.paragraph_format.space_before = Pt(0)
                    p_unvan.paragraph_format.space_after = Pt(14)
                    r_imzaci_unvan = p_unvan.add_run(imzaci.get("unvan", ""))
                    r_imzaci_unvan.font.name = "Calibri"
                    r_imzaci_unvan.font.size = Pt(max(9.0, ana_punto - 1.5))

            # 10. ALT BİLGİ: FİRMA, BANKA VE IBAN
            doc.add_paragraph().paragraph_format.space_after = Pt(20)

            p_alt_firma = doc.add_paragraph()
            p_alt_firma.paragraph_format.space_before = Pt(0)
            p_alt_firma.paragraph_format.space_after = Pt(2)
            r_alt_f = p_alt_firma.add_run(firma_unvan)
            r_alt_f.bold = True
            r_alt_f.font.name = "Calibri"
            r_alt_f.font.size = Pt(max(9.5, ana_punto - 1.0))

            banka = str(veri.get("banka", "")).strip()
            if banka and banka.lower() != "nan":
                p_banka = doc.add_paragraph()
                p_banka.paragraph_format.space_before = Pt(0)
                p_banka.paragraph_format.space_after = Pt(2)
                r_banka = p_banka.add_run(banka)
                r_banka.font.name = "Calibri"
                r_banka.font.size = Pt(kucuk_punto)

            iban = str(veri.get("iban", "")).strip()
            if iban and iban.lower() != "nan":
                p_iban = doc.add_paragraph()
                p_iban.paragraph_format.space_before = Pt(0)
                p_iban.paragraph_format.space_after = Pt(0)
                iban_metni = iban if iban.startswith("IBAN:") else f"IBAN: {iban}"
                r_iban = p_iban.add_run(iban_metni)
                r_iban.font.name = "Calibri"
                r_iban.font.size = Pt(kucuk_punto)

            os.makedirs(os.path.dirname(cikti_yolu), exist_ok=True)
            doc.save(cikti_yolu)
            return True, "Belge başarıyla oluşturuldu."

        except Exception as e:
            return False, f"Belge üretim hatası: {str(e)}"