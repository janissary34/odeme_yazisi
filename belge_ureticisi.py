import os
import re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def para_formatla(deger):
    """Sayısal tutarı Türkçe para formatına çevirir: 128000 -> 128.000,00"""
    try:
        val = float(deger)
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(deger)


def hucre_kenarlik_ekle(hucre):
    """Hücreye ince gri kurumsal kenarlık uygular."""
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
            # Şablon varsa logolu şablonu baz al, yoksa sıfırdan belge aç
            if os.path.exists(self.sablon_yolu):
                doc = Document(self.sablon_yolu)
            else:
                doc = Document()

            # Sayfa kenar boşlukları
            sections = doc.sections
            for s in sections:
                s.top_margin = Inches(0.8)
                s.bottom_margin = Inches(0.8)
                s.left_margin = Inches(1.0)
                s.right_margin = Inches(1.0)

            # 1. TARİH (Sağ Üst Köşe)
            p_tarih = doc.add_paragraph()
            p_tarih.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_tarih.paragraph_format.space_before = Pt(0)
            p_tarih.paragraph_format.space_after = Pt(8)
            run_tarih = p_tarih.add_run(f"{veri.get('ay', '')} {veri.get('yil', '')}")
            run_tarih.font.name = "Calibri"
            run_tarih.font.size = Pt(11)

            # 2. SAYI & KONU (Sol Blok)
            p_sayi = doc.add_paragraph()
            p_sayi.paragraph_format.space_before = Pt(0)
            p_sayi.paragraph_format.space_after = Pt(2)
            r_sayi = p_sayi.add_run(f"Sayı: {veri.get('sayi_formati', 'T-YK')}")
            r_sayi.font.name = "Calibri"
            r_sayi.font.size = Pt(11)

            p_konu = doc.add_paragraph()
            p_konu.paragraph_format.space_before = Pt(0)
            p_konu.paragraph_format.space_after = Pt(28)
            r_konu = p_konu.add_run(f"Konu: {veri.get('konu', '')}")
            r_konu.font.name = "Calibri"
            r_konu.font.size = Pt(11)

            # 3. HİTAP (Ortalı)
            p_hitap = doc.add_paragraph()
            p_hitap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_hitap.paragraph_format.space_before = Pt(0)
            p_hitap.paragraph_format.space_after = Pt(24)
            r_hitap = p_hitap.add_run(veri.get("hitap", "ARI TEKNOKENT GENEL MÜDÜRLÜĞÜ'NE"))
            r_hitap.bold = True
            r_hitap.font.name = "Calibri"
            r_hitap.font.size = Pt(11.5)

            # 4. KAPSAM PARAGRAFI (İki Yana Yaslı)
            p_kapsam = doc.add_paragraph()
            p_kapsam.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_kapsam.paragraph_format.space_before = Pt(0)
            p_kapsam.paragraph_format.space_after = Pt(14)
            p_kapsam.paragraph_format.line_spacing = 1.15
            r_kapsam = p_kapsam.add_run(veri.get("kapsam", ""))
            r_kapsam.font.name = "Calibri"
            r_kapsam.font.size = Pt(11)

            # 5. ÖDEME CÜMLESİ (Kritik Kelimeler Vurgulu / Bold)
            p_odeme = doc.add_paragraph()
            p_odeme.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_odeme.paragraph_format.space_before = Pt(0)
            p_odeme.paragraph_format.space_after = Pt(16)
            p_odeme.paragraph_format.line_spacing = 1.15

            tutar_kdvsiz = str(veri.get("tutar_kdvsiz", ""))
            tutar_kdvli = str(veri.get("tutar_kdvli", ""))
            firma_unvan = str(veri.get("firma_unvan", "")).strip()

            r1 = p_odeme.add_run("Söz konusu hizmetin ")
            r_toplam = p_odeme.add_run("toplam")
            r_toplam.bold = True
            r2 = p_odeme.add_run(f" bedeli olan {tutar_kdvsiz} TL + KDV (")
            r_kdvli = p_odeme.add_run(f"{tutar_kdvli} TL")
            r_kdvli.bold = True
            r3 = p_odeme.add_run(")'nin ")
            r_firma = p_odeme.add_run(firma_unvan)
            r_firma.bold = True
            r4 = p_odeme.add_run("’ne ödenmesi için gereğini izninize arz ederim.")

            for r in [r1, r_toplam, r2, r_kdvli, r3, r_firma, r4]:
                r.font.name = "Calibri"
                r.font.size = Pt(11)

            # 6. STANDART NOT
            p_not = doc.add_paragraph()
            p_not.paragraph_format.space_before = Pt(0)
            p_not.paragraph_format.space_after = Pt(16)
            r_not = p_not.add_run("Not: Sözleşme ve devamı olarak hazırlanan ek protokol belgesi ilaveten sunulmuştur.")
            r_not.font.name = "Calibri"
            r_not.font.size = Pt(10.5)

            # 7. SAYGILARIMLA
            p_saygi = doc.add_paragraph()
            p_saygi.paragraph_format.space_before = Pt(0)
            p_saygi.paragraph_format.space_after = Pt(10)
            r_saygi = p_saygi.add_run("Saygılarımla,")
            r_saygi.font.name = "Calibri"
            r_saygi.font.size = Pt(11)

            # 8. BÜTÇE TABLOSU (Varsa)
            if veri.get("tablo_aktif", False):
                toplam_b = float(veri.get("butce_toplam", 0.0))
                kullanilan_b = float(veri.get("butce_kullanilan", 0.0))
                odenecek_b = float(veri.get("butce_odenecek", 0.0))
                kalan_b = toplam_b - kullanilan_b - odenecek_b

                tablo = doc.add_table(rows=2, cols=5)
                tablo.alignment = WD_TABLE_ALIGNMENT.CENTER
                tablo.autofit = True

                basliklar = ["Bütçe Kaynağı", "Toplam Bakiye", "Kullanılan Bütçe", "Ödenecek Tutar", "Kalan Bakiye"]
                degerler = [
                    "Pazarlama PR Çekirdek-Ajans",
                    f"{para_formatla(toplam_b)}TL",
                    f"{para_formatla(kullanilan_b)}TL",
                    f"{para_formatla(odenecek_b)}TL",
                    f"{para_formatla(kalan_b)}TL"
                ]

                # Başlık satırı
                for i, b in enumerate(basliklar):
                    cell = tablo.cell(0, i)
                    cell.text = b
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell.paragraphs[0].runs[0].bold = True
                    cell.paragraphs[0].runs[0].font.name = "Calibri"
                    cell.paragraphs[0].runs[0].font.size = Pt(10)
                    hucre_kenarlik_ekle(cell)

                # Değer satırı
                for i, d in enumerate(degerler):
                    cell = tablo.cell(1, i)
                    cell.text = d
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_run = cell.paragraphs[0].runs[0]
                    p_run.font.name = "Calibri"
                    p_run.font.size = Pt(9.5)
                    if i == 4:  # Kalan Bakiye kalın olsun
                        p_run.bold = True
                    hucre_kenarlik_ekle(cell)

                # Tablodan sonra boşluk bırak
                p_bosluk = doc.add_paragraph()
                p_bosluk.paragraph_format.space_before = Pt(14)
                p_bosluk.paragraph_format.space_after = Pt(0)

            # 9. İMZACILAR ALANI (Tablonun altında veya Saygılarımla altında)
            imzacilar = veri.get("imzacilar", [])
            if imzacilar:
                p_imza_bosluk = doc.add_paragraph()
                p_imza_bosluk.paragraph_format.space_before = Pt(10)
                p_imza_bosluk.paragraph_format.space_after = Pt(0)

                for imzaci in imzacilar:
                    p_imza = doc.add_paragraph()
                    p_imza.paragraph_format.space_before = Pt(0)
                    p_imza.paragraph_format.space_after = Pt(2)
                    r_imzaci_ad = p_imza.add_run(imzaci.get("ad_soyad", ""))
                    r_imzaci_ad.bold = True
                    r_imzaci_ad.font.name = "Calibri"
                    r_imzaci_ad.font.size = Pt(11)

                    p_unvan = doc.add_paragraph()
                    p_unvan.paragraph_format.space_before = Pt(0)
                    p_unvan.paragraph_format.space_after = Pt(12)
                    r_imzaci_unvan = p_unvan.add_run(imzaci.get("unvan", ""))
                    r_imzaci_unvan.font.name = "Calibri"
                    r_imzaci_unvan.font.size = Pt(10)

            # 10. ALT BİLGİ (FİRMA, BANKA, IBAN)
            doc.add_paragraph().paragraph_format.space_after = Pt(18)

            p_alt_firma = doc.add_paragraph()
            p_alt_firma.paragraph_format.space_before = Pt(0)
            p_alt_firma.paragraph_format.space_after = Pt(2)
            r_alt_f = p_alt_firma.add_run(firma_unvan)
            r_alt_f.bold = True
            r_alt_f.font.name = "Calibri"
            r_alt_f.font.size = Pt(10.5)

            banka = str(veri.get("banka", "")).strip()
            if banka and banka.lower() != "nan":
                p_banka = doc.add_paragraph()
                p_banka.paragraph_format.space_before = Pt(0)
                p_banka.paragraph_format.space_after = Pt(2)
                r_banka = p_banka.add_run(banka)
                r_banka.font.name = "Calibri"
                r_banka.font.size = Pt(10)

            iban = str(veri.get("iban", "")).strip()
            if iban and iban.lower() != "nan":
                p_iban = doc.add_paragraph()
                p_iban.paragraph_format.space_before = Pt(0)
                p_iban.paragraph_format.space_after = Pt(0)
                r_iban = p_iban.add_run(f"IBAN: {iban}")
                r_iban.font.name = "Calibri"
                r_iban.font.size = Pt(10)

            # Çıktı klasörünü garantiye al ve kaydet
            os.makedirs(os.path.dirname(cikti_yolu), exist_ok=True)
            doc.save(cikti_yolu)
            return True, "Belge başarıyla oluşturuldu."

        except Exception as e:
            return False, f"Belge üretim hatası: {str(e)}"