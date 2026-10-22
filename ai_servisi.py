import os
from typing import List, Dict, Any, Tuple
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

# Streamlit Cloud üzerinde secrets desteği
try:
    import streamlit as st
    STREAMLIT_SECRETS_KEY = st.secrets.get("GEMINI_API_KEY", None)
except Exception:
    STREAMLIT_SECRETS_KEY = None


class AIServisi:
    def __init__(self, model_name: str = "models/gemini-2.5-flash-lite", temperature: float = 0.2):
        api_key = os.getenv("GEMINI_API_KEY") or STREAMLIT_SECRETS_KEY
        if not api_key:
            raise ValueError("GEMINI_API_KEY bulunamadı! Lütfen ayarları kontrol edin.")
        
        self.client = genai.Client(api_key=api_key)
        self.temperature = temperature

    def konu_ve_kapsam_uret(
        self,
        kullanici_kapsam_girisi: str,
        ornekler: List[Tuple[Dict[str, Any], float]]
    ) -> Tuple[str, str]:
        ornek_metinler = ""
        for i, (b, skor) in enumerate(ornekler[:3], 1):
            ornek_metinler += (
                f"\n--- ÖRNEK {i} ---\n"
                f"Konu: {b.get('konu', '')}\n"
                f"Kapsam: {b.get('kapsam', '')}\n"
            )

        sistem_talimati = (
            "Sen resmi kurumsal ödeme yazıları hazırlayan uzman bir asistansın.\n"
            "GÖREVİN:\n"
            "1. Kullanıcının verdiği iş/hizmet açıklamasını incele.\n"
            "2. Verilen örneklerdeki kurumsal dil, ton ve kalıplara harfiyen sadık kalarak bir 'Konu' satırı ve 'Kapsam' paragrafı oluştur.\n"
            "3. Tutar, IBAN veya özel fiyat bilgisi ekleme.\n"
            "4. Yanıtı SADECE ve SADECE şu formatta ver, başka hiçbir selamlama veya açıklama yazma:\n"
            "KONU: [Buraya konu metni, sonu mutlaka Hk. ile bitsin]\n"
            "KAPSAM: [Buraya resmi kapsam paragrafı]"
        )

        kullanici_istegi = f"""
KURUMSAL ARŞİVİMİZDEN ÖRNEK YAZILAR:
{ornek_metinler if ornek_metinler else "Örnek bulunamadı, genel resmi Teknokent üslubunu kullan."}

YENİ HİZMET DETAYLARI:
{kullanici_kapsam_girisi}

Yukarıdaki detaylara göre sadece KONU ve KAPSAM alanlarını üret:
"""

        denenecekler = [
            "models/gemini-2.5-flash-lite",
            "models/gemini-flash-lite-latest",
            "models/gemini-3.1-flash-lite-preview",
            "models/gemini-2.5-flash"
        ]

        for m in denenecekler:
            try:
                response = self.client.models.generate_content(
                    model=m,
                    contents=kullanici_istegi,
                    config=types.GenerateContentConfig(
                        system_instruction=sistem_talimati,
                        temperature=self.temperature
                    )
                )
                
                yanit_metni = response.text.strip()
                konu = ""
                kapsam = ""
                
                for satir in yanit_metni.split("\n"):
                    satir_temiz = satir.strip()
                    if satir_temiz.startswith("KONU:"):
                        konu = satir_temiz.replace("KONU:", "").strip()
                    elif satir_temiz.startswith("KAPSAM:"):
                        kapsam = satir_temiz.replace("KAPSAM:", "").strip()

                return (konu or "Hizmet Bedeli Hk."), (kapsam or kullanici_kapsam_girisi)

            except Exception:
                continue

        return "Hizmet Bedeli Ödemesi Hk.", kullanici_kapsam_girisi