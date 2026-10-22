import os
import re
import math
from collections import Counter
from typing import List, Dict, Any, Tuple
from docx import Document

def turkce_kucuk_harf(metin: str) -> str:
    """Türkçe karakterleri güvenli bir şekilde küçük harfe çevirir."""
    harita = {
        'İ': 'i', 'I': 'ı', 'Ş': 'ş', 'Ğ': 'ğ',
        'Ü': 'ü', 'Ö': 'ö', 'Ç': 'ç'
    }
    for buyuk, kucuk in harita.items():
        metin = metin.replace(buyuk, kucuk)
    return metin.lower()

def kelimelere_ayir(metin: str) -> List[str]:
    """Metni temizler ve 1-3 gram kelime listesine dönüştürür."""
    temiz = turkce_kucuk_harf(metin)
    temiz = re.sub(r'[^\w\s]', ' ', temiz)
    tokens = [w for w in temiz.split() if len(w) > 1]
    
    # 1'li ve 2'li kelime öbeklerini ekle (n-gram benzerliği)
    ngrams = list(tokens)
    for i in range(len(tokens) - 1):
        ngrams.append(f"{tokens[i]} {tokens[i+1]}")
    return ngrams

class SafTfidf:
    """Windows DLL engeline takılmayan saf Python TF-IDF ve Benzerlik motoru."""
    def __init__(self):
        self.idf = {}
        self.dokuman_sayisi = 0

    def fit(self, dokumanlar: List[List[str]]):
        self.dokuman_sayisi = len(dokumanlar)
        if self.dokuman_sayisi == 0:
            return
        df = Counter()
        for doc in dokumanlar:
            benzersiz_kelimeler = set(doc)
            for w in benzersiz_kelimeler:
                df[w] += 1
        
        for w, sayac in df.items():
            self.idf[w] = math.log((1 + self.dokuman_sayisi) / (1 + sayac)) + 1.0

    def vektor_olustur(self, tokens: List[str]) -> Dict[str, float]:
        tf = Counter(tokens)
        toplam = len(tokens) if tokens else 1
        vektor = {}
        for w, c in tf.items():
            if w in self.idf:
                vektor[w] = (c / toplam) * self.idf[w]
            else:
                vektor[w] = (c / toplam) * 1.0
        return vektor

    @staticmethod
    def kosinus_benzerligi(vec1: Dict[str, float], vec2: Dict[str, float]) -> float:
        kesisim = set(vec1.keys()) & set(vec2.keys())
        nokta_carpim = sum(vec1[w] * vec2[w] for w in kesisim)
        norm1 = math.sqrt(sum(v ** 2 for v in vec1.values()))
        norm2 = math.sqrt(sum(v ** 2 for v in vec2.values()))
        if norm1 == 0 or norm2 == 0:
            return 0.0
        return float(nokta_carpim / (norm1 * norm2))

class ArsivYoneticisi:
    def __init__(self, arsiv_dizini: str = "arsiv"):
        self.arsiv_dizini = arsiv_dizini
        self.belgeler: List[Dict[str, Any]] = []
        self.okunamayanlar: List[Dict[str, str]] = []
        self.motor = SafTfidf()
        self.belge_vektorleri: List[Dict[str, float]] = []
        self.indeksle()

    def turleri_getir(self) -> List[str]:
        """Arşivdeki klasör adlarını ödeme türleri olarak döner."""
        if not os.path.exists(self.arsiv_dizini):
            return []
        turler = [d for d in os.listdir(self.arsiv_dizini) 
                  if os.path.isdir(os.path.join(self.arsiv_dizini, d)) and not d.startswith('.')]
        return sorted(turler)

    def _docx_ayristir(self, dosya_yolu: str, tur: str) -> Dict[str, Any]:
        doc = Document(dosya_yolu)
        paragraflar = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        
        konu = ""
        kapsam_list = []
        odeme_cumlesi = ""
        hitap_bulundu = False
        kapanis_bulundu = False
        
        for p in paragraflar:
            p_temiz = p.replace(":", ": ").strip()
            if p_temiz.lower().startswith("konu"):
                konu = p.split(":", 1)[-1].strip() if ":" in p else p
                continue
                
            if "genel müdürlüğü" in p.lower() or "genel müdürüğü" in p.lower():
                hitap_bulundu = True
                continue
                
            if "saygılarımla" in p.lower():
                kapanis_bulundu = True
                continue
                
            if hitap_bulundu and not kapanis_bulundu:
                if any(x in p.lower() for x in ["toplam bedeli", "ödenmesi için", "onayınızı arz", "izinize arz", "kredi kartı"]):
                    odeme_cumlesi = p
                else:
                    kapsam_list.append(p)

        kapsam = " ".join(kapsam_list).strip()
        dosya_adi = os.path.basename(dosya_yolu)
        
        if not konu and not kapsam:
            raise ValueError(f"Belgeden anlamlı metin çıkarılamadı: {dosya_adi}")

        return {
            "dosya_adi": dosya_adi,
            "dosya_yolu": dosya_yolu,
            "tur": tur,
            "konu": konu,
            "kapsam": kapsam,
            "odeme_cumlesi": odeme_cumlesi,
            "ham_metin": "\n".join(paragraflar)
        }

    def indeksle(self):
        """Arşivdeki tüm docx dosyalarını tarar ve TF-IDF matrisini oluşturur."""
        self.belgeler = []
        self.okunamayanlar = []
        self.belge_vektorleri = []
        
        if not os.path.exists(self.arsiv_dizini):
            return

        for tur in self.turleri_getir():
            tur_yolu = os.path.join(self.arsiv_dizini, tur)
            for dosya in os.listdir(tur_yolu):
                if dosya.startswith("~$") or not dosya.lower().endswith(".docx"):
                    continue
                tam_yol = os.path.join(tur_yolu, dosya)
                try:
                    veri = self._docx_ayristir(tam_yol, tur)
                    self.belgeler.append(veri)
                except Exception as e:
                    self.okunamayanlar.append({"dosya": dosya, "tur": tur, "hata": str(e)})

        if self.belgeler:
            korpus_tokens = [kelimelere_ayir(b["kapsam"] + " " + b["konu"]) for b in self.belgeler]
            self.motor.fit(korpus_tokens)
            self.belge_vektorleri = [self.motor.vektor_olustur(t) for t in korpus_tokens]

    def benzerleri_bul(self, hedef_metin: str, tur: str, limit: int = 3) -> List[Tuple[Dict[str, Any], float]]:
        """Verilen metne ve türe en çok benzeyen eski yazıları döner."""
        if not self.belgeler or not hedef_metin.strip():
            return []

        tur_indeksleri = [i for i, b in enumerate(self.belgeler) if b["tur"].lower() == tur.lower()]
        if not tur_indeksleri:
            return []

        hedef_tokens = kelimelere_ayir(hedef_metin)
        hedef_vektor = self.motor.vektor_olustur(hedef_tokens)

        sonuclar = []
        for idx in tur_indeksleri:
            skor = self.motor.kosinus_benzerligi(hedef_vektor, self.belge_vektorleri[idx])
            sonuclar.append((self.belgeler[idx], skor))

        sonuclar.sort(key=lambda x: x[1], reverse=True)
        return sonuclar[:limit]

    def turdeki_kapsam_onerileri(self, tur: str) -> List[str]:
        """Kullanıcıya yazarken fikir vermesi için o türdeki eski kapsamları listeler."""
        oneriler = []
        for b in self.belgeler:
            if b["tur"].lower() == tur.lower() and b["kapsam"]:
                oneriler.append(b["kapsam"])
        return list(set(oneriler))