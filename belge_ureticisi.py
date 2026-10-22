import os
import re
from typing import Dict, Any, List, Tuple
from docxtpl import DocxTemplate, RichText

def para_formatla(deger: float) -> str:
    """Sayıyı Türkçe para birimi formatına (ör. 77.500,00) çevirir."""
    metin = f"{deger:,.2f}"
    # Python varsayılanı 1,234.56 -> Türkçe 1.234,56
    metin = metin.replace(",", "X").replace(".", ",").replace("X", ".")
    return metin

class BelgeUreticisi:
    def __init__(self, sablon_yolu: str = "sablonlar/sablon.docx"):
        self.sablon_yolu = sablon_yolu

    def dogrula(self, veri: Dict[str, Any]) -> List[str]:
        """Tüm kritik kontrolleri yapar, hata varsa listeler."""
        hatalar = []

        if not veri.get("konu", "").strip():
            hatalar.append("Konu satırı boş olamaz.")
        if not veri.get("kapsam", "").strip():
            hatalar.append("Kapsam metni boş olamaz.")
        if not veri.get("tutar_kdvsiz_sayi") or veri["tutar_kdvsiz_sayi"] <= 0:
            hatalar.append("Geçerli bir KDV'siz tutar girilmelidir.")
        if not veri.get("imzacilar"):
            hatalar.append("En az bir imzacı seçilmelidir.")
        if not veri.get("firma_unvan", "").strip():
            hatalar.append("Firma unvanı boş olamaz.")

        # Tablo aktifse bütçe kontrolleri
        if veri.get("tablo_aktif"):
            toplam = veri.get("butce_toplam", 0.0)
            kullanilan = veri.get("butce_kullanilan", 0.0)
            odenecek = veri.get("butce_odenecek", 0.0)
            kalan = toplam - kullanilan - odenecek
            if kalan < 0:
                hatalar.append(f"Bütçe yetersiz! Kalan bakiye eksiye düşüyor ({para_formatla(kalan)} TL).")

        return hatalar

    def uret(self, veri: Dict[str, Any], cikti_yolu: str) -> Tuple[bool, str]:
        """Doğrulama başarılıysa Word dosyasını oluşturur."""
        hatalar = self.dogrula(veri)
        if hatalar:
            return False, "\n".join(hatalar)

        if not os.path.exists(self.sablon_yolu):
            return False, f"Şablon dosyası bulunamadı: {self.sablon_yolu}. Lütfen Aşama 5'i tamamlayın."

        doc = DocxTemplate(self.sablon_yolu)

        # Şablondaki kalın gösterimler için RichText nesneleri
        rt_firma = RichText(veri["firma_unvan"], bold=True)
        rt_kdvsiz = RichText(veri["tutar_kdvsiz"], bold=True)
        rt_kdvli = RichText(veri["tutar_kdvli"], bold=True)

        baglam = {
            "tarih": f"{veri['ay']} {veri['yil']}",
            "sayi": veri.get("sayi_formati", "Sayı: T-YK"),
            "hitap": veri.get("hitap", "ARI TEKNOKENT GENEL MÜDÜRLÜĞÜ’NE"),
            "konu": veri["konu"],
            "kapsam": veri["kapsam"],
            "odeme_cumlesi": veri["odeme_cumlesi"],
            "kapanis": veri.get("kapanis", "Saygılarımla,"),
            "imzacilar": veri["imzacilar"],
            "firma_unvan": rt_firma,
            "banka": veri.get("banka", ""),
            "iban": veri.get("iban", ""),
            "tablo_aktif": veri.get("tablo_aktif", False),
            "butce_toplam": para_formatla(veri.get("butce_toplam", 0.0)),
            "butce_kullanilan": para_formatla(veri.get("butce_kullanilan", 0.0)),
            "butce_odenecek": para_formatla(veri.get("butce_odenecek", 0.0)),
            "butce_kalan": para_formatla(veri.get("butce_toplam", 0.0) - veri.get("butce_kullanilan", 0.0) - veri.get("butce_odenecek", 0.0))
        }

        try:
            doc.render(baglam)
            doc.save(cikti_yolu)
            return True, "Belge başarıyla oluşturuldu."
        except Exception as e:
            return False, f"Şablon doldurulurken hata oluştu: {str(e)}"