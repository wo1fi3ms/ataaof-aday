#!/usr/bin/env python3
"""
html_creator.py — Excel soru bankasından HTML quiz sayfası oluşturur.

Kullanım:
  python html_creator.py "Büyük Veri"
  python html_creator.py "Siber Güvenlik"
  python html_creator.py "Herhangi Ders" --excel dosya.xlsx --output cikti.html

Excel dosyaları ve HTML çıktıları ilgili sınıf/dönem klasöründe toplanır. Data/ ham kaynak klasörüdür.
"""

import sys, re, json, argparse, os, unicodedata
from pathlib import Path
import pandas as pd
from html import escape

# fmt: off
DERS_CONFIG = {
    # ── 1. Sınıf Güz ──────────────────────────────────────────────────────────
    "Atatürk İlkeleri ve İnkılap Tarihi I": {
        "excel":  "Data/Ataturk_Ilkeleri_I_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Ataturk_Ilkeleri_ve_Inkilap_Tarihi_I_Soru_Bankasi.html",
        "emoji":  "🏛️",
    },
    "Genel İktisat": {
        "excel":  "Data/Genel_Iktisat_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Genel_Iktisat_Soru_Bankasi.html",
        "emoji":  "📊",
    },
    "İşletme Yönetimi": {
        "excel":  "Data/Isletme_Yonetimi_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Isletme_Yonetimi_Soru_Bankasi.html",
        "emoji":  "🏢",
    },
    "Matematik I": {
        "excel":  "Data/Matematik_I_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Matematik_I_Soru_Bankasi.html",
        "emoji":  "📐",
    },
    "Türk Dili I": {
        "excel":  "Data/Turk_Dili_I_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Turk_Dili_I_Soru_Bankasi.html",
        "emoji":  "📝",
    },
    "Yabancı Dil I": {
        "excel":  "Data/Yabanci_Dil_I_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Yabanci_Dil_I_Soru_Bankasi.html",
        "emoji":  "🌐",
    },
    "Yönetim Bilişim Sistemleri": {
        "excel":  "Data/Yonetim_Bilisim_Sistemleri_Soru_Bankasi.xlsx",
        "output": "1_sinif/guz/Yonetim_Bilisim_Sistemleri_Soru_Bankasi.html",
        "emoji":  "💻",
    },
    # ── 1. Sınıf Bahar ────────────────────────────────────────────────────────
    "Afet ve Acil Durum Mevzuatı": {
        "excel":   "1_sinif/bahar/Afet_ve_Acil_Durum_Mevzuati_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Afet_ve_Acil_Durum_Mevzuati_Soru_Bankasi.html",
        "emoji":   "📜",
        "pdf_dir": "1_sinif/bahar/pdfs/Afet_ve_Acil_Durum_Mevzuati",
    },
    "Dosyalama ve Arşivleme": {
        "excel":   "1_sinif/bahar/Dosyalama_ve_Arsivleme_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Dosyalama_ve_Arsivleme_Soru_Bankasi.html",
        "emoji":   "🗂️",
        "pdf_dir": "1_sinif/bahar/pdfs/Dosyalama_ve_Arsivleme",
    },
    "İlk Yardım ve Acil Sağlık Hizmetleri": {
        "excel":   "1_sinif/bahar/Ilk_Yardim_ve_Acil_Saglik_Hizmetleri_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Ilk_Yardim_ve_Acil_Saglik_Hizmetleri_Soru_Bankasi.html",
        "emoji":   "🚑",
        "pdf_dir": "1_sinif/bahar/pdfs/Ilk_Yardim_ve_Acil_Saglik_Hizmetleri",
    },
    "Kimyasal Maddeler ve Tehlikeleri": {
        "excel":   "1_sinif/bahar/Kimyasal_Maddeler_ve_Tehlikeleri_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Kimyasal_Maddeler_ve_Tehlikeleri_Soru_Bankasi.html",
        "emoji":   "⚗️",
        "pdf_dir": "1_sinif/bahar/pdfs/Kimyasal_Maddeler_ve_Tehlikeleri",
    },
    "Lojistik Yönetimi": {
        "excel":   "1_sinif/bahar/Lojistik_Yonetimi_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Lojistik_Yonetimi_Soru_Bankasi.html",
        "emoji":   "🚚",
        "pdf_dir": "1_sinif/bahar/pdfs/Lojistik_Yonetimi",
    },
    "Temel Bilgi Teknolojileri II": {
        "excel":   "1_sinif/bahar/Temel_Bilgi_Teknolojileri_II_Soru_Bankasi.xlsx",
        "output":  "1_sinif/bahar/Temel_Bilgi_Teknolojileri_II_Soru_Bankasi.html",
        "emoji":   "💻",
        "pdf_dir": "1_sinif/bahar/pdfs/Temel_Bilgi_Teknolojileri_II",
    },
    "Algoritmalar ve Programlamaya Giriş": {
        "excel":  "Data/Algoritmalar_ve_Programlamaya_Giris_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Algoritmalar_ve_Programlamaya_Giris_Soru_Bankasi.html",
        "emoji":  "🔢",
    },
    "Atatürk İlkeleri ve İnkılap Tarihi II": {
        "excel":  "Data/Ataturk_Ilkeleri_II_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Ataturk_Ilkeleri_ve_Inkilap_Tarihi_II_Soru_Bankasi.html",
        "emoji":  "🏛️",
    },
    "Hukukun Temel Kavramları": {
        "excel":  "Data/Hukukun_Temel_Kavramlari_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Hukukun_Temel_Kavramlari_Soru_Bankasi.html",
        "emoji":  "⚖️",
    },
    "Matematik II": {
        "excel":  "Data/Matematik_II_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Matematik_II_Soru_Bankasi.html",
        "emoji":  "📐",
    },
    "Türk Dili II": {
        "excel":  "Data/Turk_Dili_II_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Turk_Dili_II_Soru_Bankasi.html",
        "emoji":  "📝",
    },
    "Yabancı Dil II": {
        "excel":  "Data/Yabanci_Dil_II_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Yabanci_Dil_II_Soru_Bankasi.html",
        "emoji":  "🌐",
    },
    "Yönetim ve Organizasyon": {
        "excel":  "Data/Yonetim_ve_Organizasyon_Soru_Bankasi.xlsx",
        "output": "1_sinif/bahar/Yonetim_ve_Organizasyon_Soru_Bankasi.html",
        "emoji":  "🏢",
    },
    # ── 2. Sınıf Güz ──────────────────────────────────────────────────────────
    "Muhasebe": {
        "excel":  "Data/Muhasebe_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Muhasebe_Soru_Bankasi.html",
        "emoji":  "💰",
    },
    "Nesneye Yönelik Programlama": {
        "excel":  "Data/Nesneye_Yonelik_Programlama_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Nesneye_Yonelik_Programlama_Soru_Bankasi.html",
        "emoji":  "🧩",
    },
    "Olasılık ve İstatistik": {
        "excel":  "Data/Olasilik_ve_Istatistik_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Olasilik_ve_Istatistik_Soru_Bankasi.html",
        "emoji":  "📊",
    },
    "Pazarlama İlkeleri": {
        "excel":  "Data/Pazarlama_Ilkeleri_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Pazarlama_Ilkeleri_Soru_Bankasi.html",
        "emoji":  "📣",
    },
    "Üretim Yönetimi": {
        "excel":  "Data/Uretim_Yonetimi_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Uretim_Yonetimi_Soru_Bankasi.html",
        "emoji":  "🏭",
    },
    "Veri Yapıları": {
        "excel":  "Data/Veri_Yapilari_Soru_Bankasi.xlsx",
        "output": "2_sinif/guz/Veri_Yapilari_Soru_Bankasi.html",
        "emoji":  "🗂️",
    },
    # ── 2. Sınıf Bahar ────────────────────────────────────────────────────────
    "Bilgisayar Organizasyonu": {
        "excel":  "Data/Bilgisayar_Organizasyonu_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Bilgisayar_Organizasyonu_Soru_Bankasi.html",
        "emoji":  "🖥️",
    },
    "İletişim Sosyolojisi": {
        "excel":  "2_sinif/bahar/Iletisim_Sosyolojisi_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Iletisim_Sosyolojisi_Soru_Bankasi.html",
        "emoji":  "🗣️",
        "pdfs": {
            "2023-2024 Bütünleme": "1EB0JauewmE0s_UWG-f2RLmghr2-bkxNt",
            "2023-2024 Final":     "1ZtSeq4TVnymKUlEDks249R-eg79zMmve",
            "2023-2024 Vize":      "1QQTUNl5_HpFyWKJezopMP5gRWuWfR60p",
            "2024-2025 Final":     "1ka-0E6g5w1lB9h8oLq29GPVR9ZrOiimx",
            "2024-2025 Vize":      "1g8icwutuZQhRugJhgJmWW6wsjQGS3UVz",
            "2024-2025 Bütünleme": "1gPyFeIDJrsiS2JZTTFz7QssRDYU6ceZQ",
            "2025-2026 Vize":      "1OwCObfpO5t_xHXdgdYLGh0BX0-BPzoBe",
            "2020-2021 Final":     "1B8G5igJ5E3I7XltssMhaPmJYKbWzKtAn",
            "2024 Bütünleme":      "1EB0JauewmE0s_UWG-f2RLmghr2-bkxNt",
            "2024 Final":          "1ZtSeq4TVnymKUlEDks249R-eg79zMmve",
            "2024 Vize":           "1QQTUNl5_HpFyWKJezopMP5gRWuWfR60p",
            "2025 Final":          "1ka-0E6g5w1lB9h8oLq29GPVR9ZrOiimx",
            "2025 Vize":           "1g8icwutuZQhRugJhgJmWW6wsjQGS3UVz",
            "2025 Bütünleme":      "1gPyFeIDJrsiS2JZTTFz7QssRDYU6ceZQ",
            "2026 Vize":           "1OwCObfpO5t_xHXdgdYLGh0BX0-BPzoBe",
            "2021 Final":          "1B8G5igJ5E3I7XltssMhaPmJYKbWzKtAn",
        },
    },
    "Maliyet Muhasebesi": {
        "excel":  "2_sinif/bahar/Maliyet_Muhasebesi_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Maliyet_Muhasebesi_Soru_Bankasi.html",
        "emoji":  "💰",
        "pdfs": {
            "2023-2024 Bütünleme": "1bt-3t3r0BCtQJOFB1tJfiE0zijteOxO5",
            "2023-2024 Vize":      "1_FE2FOGNMw9b6JUUSCyE2CMCa-j-5E7y",
            "2024-2025 Vize":      "1bYiCYPlH3THqZ9kDVzhlVNBxM9NEqBGc",
            "2024-2025 Final":     "1Wxij9KkTQvnIKGjK5H36lk0H-FWE6bHh",
            "2024-2025 Bütünleme": "1Z3gEOHQQFqzPSPgQTVR2l1D3vLIflxDJ",
            "2025-2026 Vize":      "1ccn_WK5rj-_TEA6oN_bBwdvcUbGGfBpc",
            "2015-2016 Vize":      "1KEeegg7JoSELo2Xw_xAmg0FICy8ek1k",
            "2016 Vize":           "1KEeegg7JoSELo2Xw_xAmg0FICy8ek1k",
            "2020-2021 Final":     "1qzH_Mph9YP2vRU0ZOFd_WlBo4q6SHHw5",
            "2021 Final":          "1qzH_Mph9YP2vRU0ZOFd_WlBo4q6SHHw5",
            "2022-2023 Vize":      "1TeGhsqGVstWGrWBj22x7IqO9Y38h4zZW",
            "2023 Vize":           "1TeGhsqGVstWGrWBj22x7IqO9Y38h4zZW",
            "2022-2023 Final":     "1O-366ddRxFlOeD_zRC4mvZq6FPnTRkcR",
            "2023 Final":          "1O-366ddRxFlOeD_zRC4mvZq6FPnTRkcR",
            "2023-2024 Final":     "1NF2JVS4J7To5My909KGAoHJtnq5noJwp",
            "2024 Final":          "1NF2JVS4J7To5My909KGAoHJtnq5noJwp",
            "2024 Bütünleme":      "1bt-3t3r0BCtQJOFB1tJfiE0zijteOxO5",
            "2024 Vize":           "1_FE2FOGNMw9b6JUUSCyE2CMCa-j-5E7y",
            "2025 Vize":           "1bYiCYPlH3THqZ9kDVzhlVNBxM9NEqBGc",
            "2025 Final":          "1Wxij9KkTQvnIKGjK5H36lk0H-FWE6bHh",
            "2025 Bütünleme":      "1Z3gEOHQQFqzPSPgQTVR2l1D3vLIflxDJ",
            "2026 Vize":           "1ccn_WK5rj-_TEA6oN_bBwdvcUbGGfBpc",
        },
    },
    "Ofis Uygulamaları": {
        "excel":  "Data/Ofis_Uygulamalari_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Ofis_Uygulamalari_Soru_Bankasi.html",
        "emoji":  "📋",
    },
    "Veri Tabanı Yönetim Sistemleri": {
        "excel":  "Data/Veri_Tabani_Yonetim_Sistemleri_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Veri_Tabani_Yonetim_Sistemleri_Soru_Bankasi.html",
        "emoji":  "🗄️",
    },
    "Web Programlama I": {
        "excel":  "2_sinif/bahar/Web_Programlama_I_Soru_Bankasi.xlsx",
        "output": "2_sinif/bahar/Web_Programlama_I_Soru_Bankasi.html",
        "emoji":  "🌐",
        "pdfs": {
            "2020-2021 Vize":      "1Go1OY3_nmVoigq8za4cdIq4hyFUmR-T8",
            "2022-2023 Bütünleme": "1vlvcAgIzhlJaQMYOJPZxIuHz-8rx41kz",
            "2022-2023 Final":     "1HUxv74jvThHhvZgpvZi0JT0s5KuLBO_l",
            "2022-2023 Vize":      "1b05wShHTX6eSDdOHADOBM1QPmsYAlcNL",
            "2023-2024 Bütünleme": "1ZlDGBVzvpBS3pum6qyal4cM0qRrNFj9E",
            "2023-2024 Final":     "1yEN7DejurlhAKE4XlugQS-mjfihp5Y3x",
            "2023-2024 Vize":      "1kcvLMbzXvPxAowJay2HTqhVyKV3voYel",
            "2024-2025 Bütünleme": "1lr0Cm4vuI4OxX1B0ru0vkM2-VHDTgZRm",
            "2024-2025 Final":     "1l27XTvP2r6MT90gO0BqPl72ELD5NtKMa",
            "2024-2025 Vize":      "1NUTrBDGedhleePDCqiuATUcKomLgdEH0",
            "2025-2026 Bütünleme": "1hg2fhFihhqf_kwFyr5t342mESrg64Id8",
            "2025-2026 Vize":      "1j86M96vstR5tNkuYCrKkWta987EyG9_q",
            "2021 Vize":           "1Go1OY3_nmVoigq8za4cdIq4hyFUmR-T8",
            "2023 Bütünleme":      "1vlvcAgIzhlJaQMYOJPZxIuHz-8rx41kz",
            "2023 Final":          "1HUxv74jvThHhvZgpvZi0JT0s5KuLBO_l",
            "2023 Vize":           "1b05wShHTX6eSDdOHADOBM1QPmsYAlcNL",
            "2024 Bütünleme":      "1ZlDGBVzvpBS3pum6qyal4cM0qRrNFj9E",
            "2024 Final":          "1yEN7DejurlhAKE4XlugQS-mjfihp5Y3x",
            "2024 Vize":           "1kcvLMbzXvPxAowJay2HTqhVyKV3voYel",
            "2025 Bütünleme":      "1lr0Cm4vuI4OxX1B0ru0vkM2-VHDTgZRm",
            "2025 Final":          "1l27XTvP2r6MT90gO0BqPl72ELD5NtKMa",
            "2025 Vize":           "1NUTrBDGedhleePDCqiuATUcKomLgdEH0",
            "2026 Bütünleme":      "1hg2fhFihhqf_kwFyr5t342mESrg64Id8",
            "2026 Vize":           "1j86M96vstR5tNkuYCrKkWta987EyG9_q",
        },
    },
    # ── 3. Sınıf Güz ──────────────────────────────────────────────────────────
    "E-Ticaret": {
        "excel":  "Data/E_Ticaret_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/E_Ticaret_Soru_Bankasi.html",
        "emoji":  "🛒",
    },
    "İşletim Sistemleri": {
        "excel":  "Data/Isletim_Sistemleri_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/Isletim_Sistemleri_Soru_Bankasi.html",
        "emoji":  "⚙️",
    },
    "İşletme Finansına Giriş": {
        "excel":  "Data/Isletme_Finansina_Giris_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/Isletme_Finansina_Giris_Soru_Bankasi.html",
        "emoji":  "💹",
    },
    "Mobil Programlama": {
        "excel":  "Data/Mobil_Programlama_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/Mobil_Programlama_Soru_Bankasi.html",
        "emoji":  "📱",
    },
    "Web Programlama II": {
        "excel":  "Data/Web_Programlama_II_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/Web_Programlama_II_Soru_Bankasi.html",
        "emoji":  "🌐",
    },
    "Yöneylem Araştırması": {
        "excel":  "Data/Yoneylem_Arastirmasi_Soru_Bankasi.xlsx",
        "output": "3_sinif/guz/Yoneylem_Arastirmasi_Soru_Bankasi.html",
        "emoji":  "📉",
    },
    # ── 3. Sınıf Bahar ────────────────────────────────────────────────────────
    "Benzetim": {
        "excel":  "3_sinif/bahar/Benzetim_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Benzetim_Soru_Bankasi.html",
        "emoji":  "🔬",
        "pdfs": {
            "2020-2021 Final":     "1Bbe2quHTKyhD2zgM2HdXCZKvh9BPXPFH",
            "2022-2023 Vize":      "1pY8o4VjWpixS1YrWwB3dyvcRJ0ic_oDd",
            "2023-2024 Bütünleme": "1XSaNqS6tEqaOjrIygauQMyNZVMKQRcme",
            "2023-2024 Final":     "1OHFsWBDN8SQIDTskfqQe50Z5nB_Jy_Ie",
            "2023-2024 Vize":      "1oY0prmMWz2DwTNs2-87xjQLQtstmzg9W",
            "2024-2025 Bütünleme": "18NFZZ-IcHUpHAYOP6HEcs7gnnrgBkZxk",
            "2024-2025 Final":     "1jFuIYpGn7c_fxCjTEFbMde5tKeD6OWMR",
            "2024-2025 Vize":      "1xLBatBe28aqRpIOLltRYKR1xG0z-nNBU",
            "2025-2026 Vize":      "1-md5W7-DUBIuJKWnClWIiA-Ae9oRyA2B",
            "2021 Final":          "1Bbe2quHTKyhD2zgM2HdXCZKvh9BPXPFH",
            "2023 Vize":           "1pY8o4VjWpixS1YrWwB3dyvcRJ0ic_oDd",
            "2024 Bütünleme":      "1XSaNqS6tEqaOjrIygauQMyNZVMKQRcme",
            "2024 Final":          "1OHFsWBDN8SQIDTskfqQe50Z5nB_Jy_Ie",
            "2024 Vize":           "1oY0prmMWz2DwTNs2-87xjQLQtstmzg9W",
            "2025 Bütünleme":      "18NFZZ-IcHUpHAYOP6HEcs7gnnrgBkZxk",
            "2025 Final":          "1jFuIYpGn7c_fxCjTEFbMde5tKeD6OWMR",
            "2025 Vize":           "1xLBatBe28aqRpIOLltRYKR1xG0z-nNBU",
            "2026 Vize":           "1-md5W7-DUBIuJKWnClWIiA-Ae9oRyA2B",
        },
    },
    "Bilgi Sistemleri Proje Yönetimi": {
        "excel":  "3_sinif/bahar/Bilgi_Sistemleri_Proje_Yonetimi_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Bilgi_Sistemleri_Proje_Yonetimi_Soru_Bankasi.html",
        "emoji":  "📋",
        "pdfs": {
            "2022-2023 Vize":      "15L_KJDTe9tdeOfPqaEJEaKHD1eiNwmU7",
            "2023-2024 Bütünleme": "19U2qI2dZurOsGK7C_ZDl2vzGhNWD2Kh-",
            "2023-2024 Final":     "1TVth_gvbFjvA1R0fSKc3RroAyPCJEI5Q",
            "2023-2024 Vize":      "1MW9O9UwlSjl-CAqr3oG0Op3Tdx8Bz85t",
            "2023-2024 Vize B":    "1MW9O9UwlSjl-CAqr3oG0Op3Tdx8Bz85t",
            "2024-2025 Bütünleme": "1Q_dtx9qyuqDGthMDrK6MJIjw1cGPcvCY",
            "2024-2025 Final":     "1-3lxpljCv_XqpUvdxLrjRpyF8aO8QPnk",
            "2024-2025 Final A":   "1HDBDe0pz4_syBdoCjJELsjpDciuB-Y1I",
            "2024-2025 Vize":      "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2024-2025 Vize A":    "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2024-2025 Vize B":    "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2025-2026 Vize":      "1Kcaa5QyskQPX1g94XaQ2fcsnOC8Fd5xx",
            "2023 Vize":           "15L_KJDTe9tdeOfPqaEJEaKHD1eiNwmU7",
            "2024 Bütünleme":      "19U2qI2dZurOsGK7C_ZDl2vzGhNWD2Kh-",
            "2024 Final":          "1TVth_gvbFjvA1R0fSKc3RroAyPCJEI5Q",
            "2024 Vize":           "1MW9O9UwlSjl-CAqr3oG0Op3Tdx8Bz85t",
            "2024 Vize B":         "1MW9O9UwlSjl-CAqr3oG0Op3Tdx8Bz85t",
            "2025 Bütünleme":      "1Q_dtx9qyuqDGthMDrK6MJIjw1cGPcvCY",
            "2025 Final":          "1-3lxpljCv_XqpUvdxLrjRpyF8aO8QPnk",
            "2025 Final A":        "1HDBDe0pz4_syBdoCjJELsjpDciuB-Y1I",
            "2025 Vize":           "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2025 Vize A":         "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2025 Vize B":         "1uayL56w6Ubj54TYXeKgRBSlEO6XnEMwj",
            "2026 Vize":           "1Kcaa5QyskQPX1g94XaQ2fcsnOC8Fd5xx",
        },
    },
    "Bilgisayar Ağları": {
        "excel":  "3_sinif/bahar/Bilgisayar_Aglari_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Bilgisayar_Aglari_Soru_Bankasi.html",
        "emoji":  "🌐",
        "pdfs": {
            "2022-2023 Vize":      "1nRElFX-l-l5FOnM8A-Aht-YHB5F-7TLY",
            "2023-2024 Bütünleme": "1L8FCG6qSHDxtazore_C4CLDpFQh1epCO",
            "2023-2024 Final":     "1IG3WsOMdC_tovdRCvR3l6rmeMabfAdDp",
            "2023-2024 Vize":      "11yMdJdxPEcEprWTEW6eKgRM7QCvD6EyC",
            "2024-2025 Bütünleme": "1eYDy772eFhzLzwUApgsbUn7f96eZUGNF",
            "2024-2025 Final":     "161Qt-Ez3BwIQi1u0RlaO3m0Zvem0Fg5z",
            "2024-2025 Vize":      "1azgP1OlTH8enSgJtrq1ss9JY052HRvtq",
            "2025-2026 Vize":      "1cn4f2yfEkh1nBTcW9_xIkPmsN6CkYuxl",
            "2025-2026 Vize A":    "1cn4f2yfEkh1nBTcW9_xIkPmsN6CkYuxl",
            "2025-2026 Vize B":    "1HD7DpipL8Iadl-OplJ6w8-dRyiAQTxG5",
            "2023 Vize":           "1nRElFX-l-l5FOnM8A-Aht-YHB5F-7TLY",
            "2024 Bütünleme":      "1L8FCG6qSHDxtazore_C4CLDpFQh1epCO",
            "2024 Final":          "1IG3WsOMdC_tovdRCvR3l6rmeMabfAdDp",
            "2024 Vize":           "11yMdJdxPEcEprWTEW6eKgRM7QCvD6EyC",
            "2025 Bütünleme":      "1eYDy772eFhzLzwUApgsbUn7f96eZUGNF",
            "2025 Final":          "161Qt-Ez3BwIQi1u0RlaO3m0Zvem0Fg5z",
            "2025 Vize":           "1azgP1OlTH8enSgJtrq1ss9JY052HRvtq",
            "2026 Vize":           "1cn4f2yfEkh1nBTcW9_xIkPmsN6CkYuxl",
            "2026 Vize A":         "1cn4f2yfEkh1nBTcW9_xIkPmsN6CkYuxl",
            "2026 Vize B":         "1HD7DpipL8Iadl-OplJ6w8-dRyiAQTxG5",
        },
    },
    "Bilimsel Araştırma Teknikleri": {
        "excel":  "3_sinif/bahar/Bilimsel_Arastirma_Teknikleri_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Bilimsel_Arastirma_Teknikleri_Soru_Bankasi.html",
        "emoji":  "🔍",
        "pdfs": {
            "2022-2023 Vize":      "1yuAVL5aYPu7fQZ_VdkAQcqhS-YQmd9vU",
            "2023-2024 Bütünleme": "1KUFTI9hL52_DChzs0vgmMf5_jTr0nMx6",
            "2023-2024 Final":     "1Wr0PQRUakq2Mnse7Ec4DgRHQGg9dzxL2",
            "2023-2024 Vize":      "1tjt7i8aHbTpxfmUSgJ1AhVUCbUbF06yk",
            "2024-2025 Bütünleme": "1adSnQghRUpBiKQ31LGzEZIg-NBzY5drq",
            "2024-2025 Final":     "1OcPKzbbAur_ZiAHLj72ao9aIQPasJ-9V",
            "2024-2025 Vize":      "12zXMy1cC-8eiejAd2FOKyBTa86PC2cxH",
            "2025-2026 Vize":      "1ObN9U_bH9C6TGvi2KRiJ8s4lVw9s-ux0",
            "2023 Vize":           "1yuAVL5aYPu7fQZ_VdkAQcqhS-YQmd9vU",
            "2024 Bütünleme":      "1KUFTI9hL52_DChzs0vgmMf5_jTr0nMx6",
            "2024 Final":          "1Wr0PQRUakq2Mnse7Ec4DgRHQGg9dzxL2",
            "2024 Vize":           "1tjt7i8aHbTpxfmUSgJ1AhVUCbUbF06yk",
            "2025 Bütünleme":      "1adSnQghRUpBiKQ31LGzEZIg-NBzY5drq",
            "2025 Final":          "1OcPKzbbAur_ZiAHLj72ao9aIQPasJ-9V",
            "2025 Vize":           "12zXMy1cC-8eiejAd2FOKyBTa86PC2cxH",
            "2026 Vize":           "1ObN9U_bH9C6TGvi2KRiJ8s4lVw9s-ux0",
        },
    },
    "Bilişim Hukuku": {
        "excel":  "3_sinif/bahar/Bilisim_Hukuku_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Bilisim_Hukuku_Soru_Bankasi.html",
        "emoji":  "⚖️",
        "pdfs": {
            "2022-2023 Bütünleme": "15_JbMLIUCN4gXnwCjIIkPn6z2pu4yjs3",
            "2022-2023 Final":     "1XbGodRnev3BfTFu-JBrCd_dPcqwm7Wdk",
            "2022-2023 Vize":      "1VsPvDnfdrtIHO96g0cRuFM7sNtDi1GQ3",
            "2023-2024 Bütünleme": "14WiAGi0xMLjF7LxEF2CMeWIrPnY_SYIp",
            "2023-2024 Final":     "11pPE1j96S7j4VwSDmbfeYVcDIKhhsuOq",
            "2023-2024 Üç Ders":   "1xS5Hx7_PatH3rB62SyODsZmKIWxiBVUc",
            "2023-2024 Vize":      "1sHZgi9ImhONAanQa6TajvW0kTJMdudbH",
            "2024-2025 Bütünleme": "1_gJQADrDH_qt0R-P0xspkw7-X8ZmvZAU",
            "2024-2025 Final":     "1hGyLMFZlRk8Ilz_egko_Q3l39l5Jdjau",
            "2024-2025 Vize":      "1UC4_68oE0sO669S1vyxTvuXjloX5HOPF",
            "2025-2026 Vize":      "1IckTEWwRzeuzi4eeT4v9KAini5BvuXMf",
            "2023 Bütünleme":      "15_JbMLIUCN4gXnwCjIIkPn6z2pu4yjs3",
            "2023 Final":          "1XbGodRnev3BfTFu-JBrCd_dPcqwm7Wdk",
            "2023 Vize":           "1VsPvDnfdrtIHO96g0cRuFM7sNtDi1GQ3",
            "2024 Final":          "11pPE1j96S7j4VwSDmbfeYVcDIKhhsuOq",
            "2024 Üç Ders":        "1xS5Hx7_PatH3rB62SyODsZmKIWxiBVUc",
            "2024 Vize":           "1sHZgi9ImhONAanQa6TajvW0kTJMdudbH",
            "2025 Bütünleme":      "1_gJQADrDH_qt0R-P0xspkw7-X8ZmvZAU",
            "2025 Final":          "1hGyLMFZlRk8Ilz_egko_Q3l39l5Jdjau",
            "2025 Vize":           "1UC4_68oE0sO669S1vyxTvuXjloX5HOPF",
            "2026 Vize":           "1IckTEWwRzeuzi4eeT4v9KAini5BvuXMf",
            "2024 Bütünleme":      "14WiAGi0xMLjF7LxEF2CMeWIrPnY_SYIp",
            "Üç Ders":             "1xS5Hx7_PatH3rB62SyODsZmKIWxiBVUc",
        },
    },
    "Büyük Veri": {
        "excel":  "3_sinif/bahar/Buyuk_Veri_Soru_Bankasi.xlsx",
        "output": "3_sinif/bahar/Buyuk_Veri_Soru_Bankasi.html",
        "emoji":  "📊",
        "pdfs": {
            "2022-2023 Vize":      "1SAAXGaNWurGXmBkJR-bLR_eZxy_2qVLF",
            "2023-2024 Bütünleme": "1oJZ18qhOK7D5gkRujYo1N0IwVTE3UJ76",
            "2023-2024 Final":     "1RkZOZRID8Xx5jZ0I59r7J_dgIzHiVgrb",
            "2023-2024 Vize":      "13yHRf3GWb5nK_JOpgNwUypkluOB7z3g7",
            "2024-2025 Bütünleme": "1TiUOvM57duhoC5dJUm1E7vPD71Y9PRlM",
            "2024-2025 Final":     "1cAv4S712nTRVhXsoOLwSOY5aSOSutLo4",
            "2024-2025 Vize":      "1sC9HiLwABziTa54Tapft78CxNB8e4U9M",
            "2025-2026 Vize":      "1JS6_kKLnidd6dsOLF3c3Pu1w8E6kWMGA",
            "2022-2023 Final":     ["1WqamU5-CqYeSLkuw3o1ZwG79iZFXn09s", "18HNvAWbofKRsMm3H9DHDLbi8Is65DgHH"],
            "2023 Final":          ["1WqamU5-CqYeSLkuw3o1ZwG79iZFXn09s", "18HNvAWbofKRsMm3H9DHDLbi8Is65DgHH"],
            "2023 Vize":           "1SAAXGaNWurGXmBkJR-bLR_eZxy_2qVLF",
            "2024 Bütünleme":      "1oJZ18qhOK7D5gkRujYo1N0IwVTE3UJ76",
            "2024 Final":          "1RkZOZRID8Xx5jZ0I59r7J_dgIzHiVgrb",
            "2024 Vize":           "13yHRf3GWb5nK_JOpgNwUypkluOB7z3g7",
            "2025 Bütünleme":      "1TiUOvM57duhoC5dJUm1E7vPD71Y9PRlM",
            "2025 Final":          "1cAv4S712nTRVhXsoOLwSOY5aSOSutLo4",
            "2025 Vize":           "1sC9HiLwABziTa54Tapft78CxNB8e4U9M",
            "2026 Vize":           "1JS6_kKLnidd6dsOLF3c3Pu1w8E6kWMGA",
        },
    },
    # ── 4. Sınıf Güz ──────────────────────────────────────────────────────────
    "Endüstri 4.0": {
        "excel":  "Data/Endustri_4_0_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/Endustri_4_0_Soru_Bankasi.html",
        "emoji":  "🏭",
    },
    "Karar Alma Teknikleri": {
        "excel":  "Data/Karar_Alma_Teknikleri_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/Karar_Alma_Teknikleri_Soru_Bankasi.html",
        "emoji":  "🎯",
    },
    "Kullanıcı Deneyimi Tasarımı": {
        "excel":  "Data/Kullanici_Deneyimi_Tasarimi_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/Kullanici_Deneyimi_Tasarimi_Soru_Bankasi.html",
        "emoji":  "🎨",
    },
    "Python Programlama": {
        "excel":  "Data/Python_Programlama_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/Python_Programlama_Soru_Bankasi.html",
        "emoji":  "🐍",
    },
    "R Programlama ile Veri Analizi": {
        "excel":  "Data/R_Programlama_ile_Veri_Analizi_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/R_Programlama_ile_Veri_Analizi_Soru_Bankasi.html",
        "emoji":  "📈",
    },
    "Yazılım Kalite ve Testi": {
        "excel":  "Data/Yazilim_Kalite_ve_Testi_Soru_Bankasi.xlsx",
        "output": "4_sinif/guz/Yazilim_Kalite_ve_Testi_Soru_Bankasi.html",
        "emoji":  "🧪",
    },
    # ── 4. Sınıf Bahar ────────────────────────────────────────────────────────
    "Bilişim Sistemleri Analiz ve Tasarımı": {
        "excel":  "Data/Bilisim_Sistemleri_Analiz_ve_Tasarimi_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Bilisim_Sistemleri_Analiz_ve_Tasarimi_Soru_Bankasi.html",
        "emoji":  "🔧",
    },
    "Kurumsal Bilgi Sistemleri": {
        "excel":  "Data/Kurumsal_Bilgi_Sistemleri_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Kurumsal_Bilgi_Sistemleri_Soru_Bankasi.html",
        "emoji":  "🏢",
    },
    "Siber Güvenlik": {
        "excel":  "Data/Siber_Guvenlik_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Siber_Guvenlik_Soru_Bankasi.html",
        "emoji":  "🔒",
    },
    "Sistem Yönetimi ve Bulut Bilişim": {
        "excel":  "Data/Sistem_Yonetimi_ve_Bulut_Bilisim_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Sistem_Yonetimi_ve_Bulut_Bilisim_Soru_Bankasi.html",
        "emoji":  "☁️",
    },
    "Veri Madenciliği": {
        "excel":  "Data/Veri_Madenciligi_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Veri_Madenciligi_Soru_Bankasi.html",
        "emoji":  "⛏️",
    },
    "Yapay Zeka ve Makine Öğrenmesi": {
        "excel":  "Data/Yapay_Zeka_ve_Makine_Ogrenmesi_Soru_Bankasi.xlsx",
        "output": "4_sinif/bahar/Yapay_Zeka_ve_Makine_Ogrenmesi_Soru_Bankasi.html",
        "emoji":  "🤖",
    },
}
# fmt: on

OPTION_LETTERS = ["A", "B", "C", "D", "E"]
OPTION_COLS    = ["A Şıkkı", "B Şıkkı", "C Şıkkı", "D Şıkkı", "E Şıkkı"]

# Standard tab order — only shown when the data actually contains that exam type
STANDARD_TABS = [
    ("Final",     "📙 Final"),
    ("Bütünleme", "📗 Bütünleme"),
    ("Yaz Okulu", "☀️ Yaz Okulu"),
    ("Üç Ders",   "📘 Üç Ders"),
    ("Vize",      "📝 Vize"),
]


# ─── PDF auto-discovery ───────────────────────────────────────────────────────

# ASCII filename part → Turkish label suffix
_PDF_SINAV_MAP = {
    "Butunleme": "Bütünleme",
    "Final":     "Final",
    "Yaz_Okulu": "Yaz Okulu",
}

def discover_pdfs(pdf_dir, output_path):
    """Scan pdf_dir for YYYY-YYYY_Sinav.pdf files, return {label: rel_path} dict.
    Paths are relative to the output HTML file's directory."""
    if not pdf_dir or not os.path.isdir(pdf_dir):
        return {}
    out_dir = str(Path(output_path).parent)
    result = {}
    for fname in sorted(os.listdir(pdf_dir)):
        if not fname.lower().endswith(".pdf"):
            continue
        m = re.match(r'^(\d{4}-\d{4})_(.+)\.pdf$', fname)
        if not m:
            continue
        year = m.group(1)
        sinav_key = m.group(2)
        sinav_tr = _PDF_SINAV_MAP.get(sinav_key, sinav_key.replace("_", " "))
        label = f"{year} {sinav_tr}"
        rel = os.path.relpath(os.path.join(pdf_dir, fname), out_dir)
        result[label] = rel
    return result


# ─── normalisation helpers ────────────────────────────────────────────────────

def _str(v):
    if isinstance(v, float) and v != v:
        return ""
    s = str(v).strip()
    return "" if s == "nan" else s

def clean_option(text, letter):
    s = _str(text)
    return re.sub(rf"^{letter}\)\s*", "", s, flags=re.IGNORECASE)

def clean_answer_letter(text):
    s = _str(text)
    m = re.match(r"^([A-Ea-e])\)", s)
    if m:
        return m.group(1).upper()
    if re.match(r"^[A-Ea-e]$", s):
        return s.upper()
    return s

def normalize_exams_str(text):
    s = _str(text)
    parts = re.split(r"\s*[,|]\s*", s)
    return " | ".join(p.strip() for p in parts if p.strip())

def parse_exams(exams_str):
    return [e.strip() for e in exams_str.split(" | ") if e.strip()]


# ─── question builder ─────────────────────────────────────────────────────────

def build_questions(df):
    questions = []
    for _, row in df.iterrows():
        q_text = _str(row.get("Soru", ""))
        if not q_text:
            continue

        options = []
        for letter, col in zip(OPTION_LETTERS, OPTION_COLS):
            opt = clean_option(row.get(col, ""), letter)
            if opt:
                options.append(f"{letter}) {opt}")

        ans_letter = clean_answer_letter(row.get("Doğru Cevap", ""))
        ans_text   = next((o for o in options if o.startswith(ans_letter + ")")), ans_letter)

        exams_str = normalize_exams_str(row.get("Sınav(lar)", ""))
        exams     = parse_exams(exams_str)
        exp       = _str(row.get("Açıklama", ""))

        questions.append(dict(q=q_text, options=options, ans=ans_text, exp=exp, exams=exams))

    # "N Kez Çıktı!" markers
    for q in questions:
        count = len([e for e in q["exams"] if "Kez" not in e])
        if count >= 2:
            q["exams"].append(f"{count} Kez Çıktı!")

    return questions


# ─── tab detection (standard order) ──────────────────────────────────────────

def detect_tabs(questions):
    all_exams = {e for q in questions for e in q["exams"] if "Kez Çıktı" not in e}
    return [
        (label, keyword)
        for keyword, label in STANDARD_TABS
        if any(keyword in e for e in all_exams)
    ]


# ─── HTML generation ──────────────────────────────────────────────────────────

CSS = """
:root{
  /* ATA-AÖF ana lacivert — UI chrome */
  --green:#0d1460;--green-dk:#080e46;--green-lt:#e8eaf6;
  /* Vurgu yeşili — geri bildirim / başarı */
  --accent:#00CB54;--accent-dk:#009e40;--accent-lt:#e6faf0;
  --gold:#e1b77e;
  --wine:#7b1d34;--wine-dk:#5c1526;
  --banko:#c0392b;--banko-dk:#922b21;
  --bg:#f7f7f8;--card-bg:#fff;--text:#202428;--border:#e4e4e7;
  --opt-bg:#fafafa;--opt-border:#e4e4e7;
  --muted:#606065;
}
body.dark-mode{
  /* UI chrome — okunabilir lacivert-lavanta */
  --green:#8899ee;--green-dk:#c0caff;--green-lt:#131c42;
  /* Geri bildirim yeşili */
  --accent:#00CB54;--accent-dk:#4ade80;--accent-lt:#012a14;
  --gold:#e1b77e;
  /* Daha ölçülü wine ve banko */
  --wine:#a05868;--wine-dk:#7d4052;
  --banko:#c06060;--banko-dk:#9a4848;
  --bg:#08091e;--card-bg:#0f1230;--text:#d8daec;--border:#252d60;
  --opt-bg:#111830;--opt-border:#252d60;
}
*{box-sizing:border-box;}
body{font-family:'Open Sans','Segoe UI',Tahoma,sans-serif;background:var(--bg);color:var(--text);margin:0;padding:0 0 48px;line-height:1.65;transition:background .25s,color .25s;}

/* ── topbar ── */
.site-topbar{position:sticky;top:0;z-index:200;background:var(--green-dk);display:flex;align-items:center;justify-content:space-between;padding:8px 22px;box-shadow:0 2px 8px rgba(0,0,0,.22);}
body.dark-mode .site-topbar{background:#06082a;}
.topbar-left{display:flex;align-items:center;gap:14px;}
.topbar-home{display:inline-flex;align-items:center;gap:5px;font-size:.78em;font-weight:700;color:rgba(255,255,255,.75);text-decoration:none;border:1px solid rgba(255,255,255,.25);border-radius:6px;padding:3px 10px;transition:all .2s;white-space:nowrap;}
.topbar-home:hover{background:rgba(255,255,255,.12);color:#fff;}
.topbar-uni{font-size:.75em;color:rgba(255,255,255,.45);letter-spacing:.03em;}
.topbar-right{display:flex;align-items:center;gap:12px;}
.topbar-course{font-size:.88em;font-weight:700;color:#fff;}
.dark-mode-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:8px;padding:5px 11px;font-size:1.05em;cursor:pointer;transition:all .2s;color:#fff;line-height:1;}
.dark-mode-btn:hover{background:rgba(255,255,255,.25);transform:scale(1.08);}

/* ── hero header ── */
.header{background:linear-gradient(160deg,var(--green-dk) 0%,var(--green) 100%);color:#fff;padding:32px 24px 28px;text-align:center;border-bottom:4px solid var(--gold);box-shadow:0 3px 14px rgba(0,0,0,.2);}
body.dark-mode .header{background:linear-gradient(160deg,#05071c 0%,#0a0e44 100%);}
.header-badge{display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);border-radius:20px;padding:4px 16px;font-size:.77em;font-weight:700;letter-spacing:.05em;color:rgba(255,255,255,.92);margin-bottom:14px;text-transform:uppercase;}
.header h1{margin:0 0 8px;font-size:1.9em;font-weight:700;letter-spacing:-.01em;}
.header-subtitle{margin:0;font-size:.93em;color:var(--gold);font-weight:600;letter-spacing:.03em;}

/* ── layout ── */
.container{max-width:1200px;margin:0 auto;padding:0 16px;}
.page-layout{display:flex;gap:22px;align-items:flex-start;margin-top:24px;}
#quiz-container{flex:1;min-width:0;display:grid;grid-template-columns:1fr 1fr;gap:18px;animation:fadeIn .3s ease-in-out;}
#quiz-container>*{min-width:0;}
@media(max-width:900px){#quiz-container{grid-template-columns:1fr;}}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:translateY(0);}}

/* ── sidebar ── */
.sidebar{width:192px;flex-shrink:0;position:sticky;top:48px;display:flex;flex-direction:column;gap:8px;background:var(--card-bg);padding:14px 12px;border-radius:10px;border:1px solid var(--border);box-shadow:0 2px 10px rgba(0,0,0,.07);}
.sidebar-label{font-size:.68em;font-weight:700;text-transform:uppercase;color:#aaa;letter-spacing:.07em;margin-top:2px;}
.sidebar-home{display:flex;align-items:center;justify-content:center;gap:6px;padding:9px 11px;background:var(--card-bg);color:var(--green-dk);border:2px solid var(--green);border-radius:7px;font-size:.82em;font-weight:700;text-decoration:none;transition:all .2s;}
.sidebar-home:hover{background:var(--green-lt);}
.tabs{display:flex;flex-direction:column;gap:5px;}
.tab-btn{background:var(--card-bg);color:var(--green-dk);border:2px solid var(--green);padding:8px 11px;font-size:.82em;font-weight:700;font-family:inherit;border-radius:7px;cursor:pointer;transition:all .2s;text-align:left;width:100%;}
.tab-btn:hover{background:var(--green-lt);}
.tab-btn.active{background:var(--green);color:#fff;border-color:var(--green);box-shadow:0 3px 8px rgba(13,20,96,.28);}
.tab-btn.tab-banko{background:#fde8e8;color:var(--banko);border-color:var(--banko);}
.tab-btn.tab-banko:hover,.tab-btn.tab-banko.active{background:var(--banko);color:#fff;}
body.dark-mode .tab-btn.tab-banko{background:#1e0808;color:#e08080;border-color:#7a3030;}
body.dark-mode .tab-btn.tab-banko:hover,body.dark-mode .tab-btn.tab-banko.active{background:#7a3030;color:#fff;}
.toggle-all-btn{background:var(--wine);color:#fff;border:none;padding:10px 11px;font-size:.84em;font-weight:700;font-family:inherit;border-radius:7px;cursor:pointer;box-shadow:0 3px 8px rgba(123,29,52,.28);transition:all .2s;width:100%;text-align:center;}
.toggle-all-btn:hover{background:var(--wine-dk);transform:translateY(-1px);}
.stats-bar{text-align:center;font-size:.77em;color:#666;font-weight:600;}
.stats-bar>span{display:inline-block;background:var(--card-bg);border:1px solid var(--border);border-radius:12px;padding:2px 9px;}
.stats-short{display:none;}
.score-panel{display:flex;flex-direction:column;gap:6px;margin-top:2px;}
.score-item{display:flex;align-items:center;justify-content:center;gap:6px;font-size:.91em;font-weight:700;border-radius:8px;padding:7px 10px;}
.score-item span{font-size:1.3em;font-weight:800;min-width:24px;text-align:center;}
.score-correct-item{background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;}
.score-wrong-item{background:#fde8e8;color:#c0392b;border:1px solid #f87171;}
body.dark-mode .score-correct-item{background:#012a14;color:#4ade80;border-color:#025a28;}
body.dark-mode .score-wrong-item{background:#2a0808;color:#f08888;border-color:#5a1818;}
.sidebar-meta{display:flex;flex-direction:column;gap:6px;margin-top:2px;}
.search-box{width:100%;padding:8px 10px;border:1.5px solid var(--border);border-radius:7px;font-size:.83em;font-family:inherit;background:var(--card-bg);color:var(--text);outline:none;transition:border-color .2s;box-sizing:border-box;}
.search-box:focus{border-color:var(--accent);}
.search-box::placeholder{color:#aaa;}
.tab-btn.tab-review{color:#92400e;border-color:#d97706;background:#fff8ec;position:relative;-webkit-user-select:none;user-select:none;-webkit-touch-callout:none;}
.tab-btn.tab-review:hover{background:#fde68a;color:#78350f;}
.tab-btn.tab-review.active{background:#d97706!important;color:#fff!important;border-color:#d97706!important;box-shadow:0 3px 8px rgba(217,119,6,.35)!important;}
body.dark-mode .tab-btn.tab-review{color:#c8963a;border-color:#7a5010;background:#1a1000;}
body.dark-mode .tab-btn.tab-review:hover{background:#2a1a00;color:#e0b050;}
body.dark-mode .tab-btn.tab-review.active{background:#7a5010!important;color:#fff!important;border-color:#7a5010!important;box-shadow:0 3px 8px rgba(120,80,16,.4)!important;}
.review-info-btn{position:absolute;top:4px;right:5px;width:15px;height:15px;border-radius:50%;border:1.5px solid currentColor;background:none;font-size:.6em;font-weight:900;line-height:13px;text-align:center;cursor:pointer;padding:0;color:inherit;opacity:.65;font-family:inherit;display:none;}
.review-tooltip{display:none;position:absolute;top:calc(100% + 6px);right:0;width:220px;background:var(--card-bg);border:1px solid var(--border);border-radius:8px;padding:9px 11px;font-size:.75em;font-weight:400;color:var(--text);box-shadow:0 4px 14px rgba(0,0,0,.15);z-index:200;text-align:left;line-height:1.5;pointer-events:none;}
@media(min-width:721px){
  .review-info-btn{display:block;}
  .review-info-btn:hover+.review-tooltip,.review-info-btn:focus+.review-tooltip{display:block;}
}
.progress-wrap{display:flex;flex-direction:column;gap:4px;margin-top:2px;}
.progress-label{display:flex;justify-content:space-between;align-items:center;font-size:.75em;font-weight:700;color:var(--text);}
.progress-reset{background:none;border:none;font-size:.7em;color:#aaa;cursor:pointer;padding:0;font-family:inherit;transition:color .2s;}
.progress-reset:hover{color:var(--banko);}
.progress-bar-bg{height:7px;background:var(--border);border-radius:4px;overflow:hidden;}
.progress-bar-fill{height:100%;background:var(--accent);border-radius:4px;transition:width .4s ease;}
.q-prev{font-size:.8em;font-weight:700;padding:2px 7px;border-radius:4px;line-height:1.4;}
.q-prev-correct{background:#d1fae5;color:#065f46;}
.q-prev-wrong{background:#fde8e8;color:#c0392b;}
body.dark-mode .q-prev-correct{background:#012a14;color:#4ade80;border:1px solid #025a28;}
body.dark-mode .q-prev-wrong{background:#2a0808;color:#f08888;border:1px solid #5a1818;}

/* ── mobile sidebar ── */
@media(max-width:720px){
  .page-layout{flex-direction:column;margin-top:8px;}
  .header{display:none;}
  .sidebar{position:fixed;top:48px;left:0;right:0;z-index:100;flex-direction:column;gap:5px;padding:8px 12px;border-radius:0;border:none;border-bottom:2px solid var(--border);box-shadow:0 2px 8px rgba(0,0,0,.13);width:100vw!important;max-width:100vw;overflow-x:hidden;box-sizing:border-box;}
  .sidebar-home,.sidebar-label{display:none;}
  .tabs{flex-direction:row;flex-wrap:nowrap;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;gap:5px;}
  .tabs::-webkit-scrollbar{display:none;}
  .tab-btn{white-space:nowrap;width:auto;flex-shrink:0;padding:7px 12px;font-size:.8em;}
  .sidebar-meta{flex-direction:row;align-items:center;gap:6px;flex-wrap:nowrap;margin-top:0;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;}
  .sidebar-meta::-webkit-scrollbar{display:none;}
  .stats-bar{flex-shrink:0;font-size:.73em;white-space:nowrap;text-align:left;}
  .stats-bar>span{padding:2px 7px;}
  .score-panel{flex-direction:row;gap:4px;margin-top:0;flex-shrink:0;}
  .score-item{padding:4px 6px;font-size:.74em;gap:3px;white-space:nowrap;}
  .score-item span{font-size:1em;min-width:14px;}
  .stats-full{display:none;}
  .stats-short{display:inline;}
  .toggle-all-btn{width:100%;padding:7px 10px;font-size:.8em;}
  .sidebar.scroll-collapsed{flex-direction:row;align-items:center;gap:6px;padding:5px 12px;}
  .sidebar.scroll-collapsed .tab-btn:not(.active){display:none;}
  .sidebar.scroll-collapsed nav.tabs{flex-shrink:0;overflow:visible;}
  .sidebar.scroll-collapsed .sidebar-meta{flex:1;min-width:0;gap:5px;overflow-x:auto;}
  .sidebar.scroll-collapsed .score-item{padding:3px 5px;}
  .sidebar.scroll-collapsed .progress-wrap{flex-shrink:0;}
  .sidebar.scroll-collapsed .toggle-all-btn{display:none;}
  .sidebar.scroll-collapsed .search-box{display:none;}
  .sidebar.scroll-collapsed #tag-filter-indicator{display:none;}
  .sidebar.scroll-collapsed .prog-label{display:none;}
  .sidebar.scroll-collapsed .score-label{display:none;}
  .sidebar.scroll-collapsed .stats-full{display:none;}
  .sidebar.scroll-collapsed .stats-short{display:inline;}
  #quiz-container{grid-template-columns:1fr;}
}

/* ── question cards ── */
.question-card{background:var(--card-bg);border-radius:9px;box-shadow:0 2px 10px rgba(0,0,0,.06);padding:22px 24px;border-left:5px solid var(--green);transition:box-shadow .2s;display:flex;flex-direction:column;}
.question-card:hover{box-shadow:0 5px 18px rgba(13,20,96,.12);}
.question-card .btn-toggle{margin-top:auto;}
.q-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:13px;border-bottom:1px solid var(--border);padding-bottom:10px;gap:10px;}
.q-meta{display:flex;flex-direction:column;gap:3px;}
.q-num{color:var(--green-dk);font-size:.86em;font-weight:700;text-transform:uppercase;letter-spacing:.05em;}
.exam-tags{display:flex;flex-wrap:wrap;gap:5px;justify-content:flex-end;}
.exam-tag{padding:2px 8px;border-radius:4px;font-size:.73em;font-weight:700;letter-spacing:.02em;}
.uc-tag{background:#ede9fe;color:#4c1d95;}
.vize-tag{background:#dbeafe;color:#1e3a8a;}
.fin-tag{background:#d1fae5;color:#065f46;}
.but-tag{background:#fef3c7;color:#7a5800;border:1px solid var(--gold);}
.yaz-tag{background:#ffedd5;color:#9a3412;border:1px solid #fb923c;}
.freq-tag{background:var(--banko);color:#fff;}
body.dark-mode .uc-tag{background:#1e1540;color:#9474cc;}
body.dark-mode .vize-tag{background:#0e1e3c;color:#5a8fcc;}
body.dark-mode .fin-tag{background:#112c1c;color:#5aaa7a;}
body.dark-mode .but-tag{background:#28210a;color:#b88c30;}
body.dark-mode .yaz-tag{background:#281808;color:#b87030;}
body.dark-mode .freq-tag{background:#2c1010;color:#b85858;}
.exam-tag.tag-clickable{cursor:pointer;transition:transform .15s,opacity .15s;}
.exam-tag.tag-clickable:hover{transform:scale(1.08);opacity:.75;}
.question-card.tag-hidden{display:none;}
.sidebar.tag-filter-active .tabs,.sidebar.tag-filter-active .sidebar-home,.sidebar.tag-filter-active .sidebar-label,.sidebar.tag-filter-active .toggle-all-btn{opacity:.28;pointer-events:none;user-select:none;}
.sidebar.tag-filter-active #tag-filter-indicator{opacity:1!important;pointer-events:auto!important;}
#tag-filter-indicator{display:none;border-radius:7px;padding:8px 10px;font-size:.82em;font-weight:700;align-items:center;gap:6px;}
#tag-filter-indicator.visible{display:flex;}
#tag-filter-indicator .tfi-label{flex:1;line-height:1.3;}
#tag-filter-indicator .tfi-close{background:none;border:none;cursor:pointer;font-size:1.1em;padding:0 2px;font-weight:900;opacity:.65;line-height:1;font-family:inherit;}
#tag-filter-indicator .tfi-close:hover{opacity:1;}
#tag-filter-indicator.tfi-fin{background:#d1fae5;color:#065f46;}
#tag-filter-indicator.tfi-vize{background:#dbeafe;color:#1e3a8a;}
#tag-filter-indicator.tfi-but{background:#fef3c7;color:#7a5800;border:1px solid var(--gold);}
#tag-filter-indicator.tfi-uc{background:#ede9fe;color:#4c1d95;}
#tag-filter-indicator.tfi-yaz{background:#ffedd5;color:#9a3412;border:1px solid #fb923c;}
#tag-filter-indicator.tfi-other{background:var(--bg);color:#444;border:1px solid var(--border);}
body.dark-mode #tag-filter-indicator.tfi-fin{background:#112c1c;color:#5aaa7a;border:1px solid #1e4830;}
body.dark-mode #tag-filter-indicator.tfi-vize{background:#0e1e3c;color:#5a8fcc;border:1px solid #1a3460;}
body.dark-mode #tag-filter-indicator.tfi-but{background:#28210a;color:#b88c30;border:1px solid #483808;}
body.dark-mode #tag-filter-indicator.tfi-uc{background:#1e1540;color:#9474cc;border:1px solid #362060;}
body.dark-mode #tag-filter-indicator.tfi-yaz{background:#281808;color:#b87030;border:1px solid #483018;}
body.dark-mode #tag-filter-indicator.tfi-other{color:var(--text);}
.question-text{font-size:1.05em;font-weight:600;color:var(--text);margin:0 0 14px;white-space:pre-line;line-height:1.55;}
.neg-hl{text-decoration:underline;text-decoration-style:wavy;text-decoration-color:var(--banko);text-underline-offset:3px;font-weight:700;}
.q-table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;margin:8px 0 12px;}
.q-table{border-collapse:collapse!important;width:100%!important;margin:0!important;font-size:.88em;font-weight:400;white-space:normal!important;table-layout:auto;word-break:break-word;}
.q-table th,.q-table td{border:1px solid var(--border)!important;padding:5px 9px!important;vertical-align:middle!important;}
.q-table th{background:var(--green-lt)!important;color:var(--green-dk)!important;font-weight:700!important;text-align:center!important;}
.q-table td{text-align:left;}
.q-table td.num{text-align:right;font-variant-numeric:tabular-nums;font-family:'Courier New',monospace;font-size:.95em;}
.q-table tr:nth-child(even) td{background:var(--opt-bg);}
.q-table tfoot td,.q-table tr:last-child td[style*="font-weight"]{font-weight:700;background:var(--green-lt);}
body.dark-mode .q-table th{background:#1a2060!important;color:#c0caff!important;}
body.dark-mode .options li.wrong-highlight{background:#2a0808;border-color:#7a3030;color:#f08888;}
body.dark-mode .options li.correct-highlight{background:#012a14;border-color:#025a28;color:#4ade80;}
@media(max-width:600px){.q-table{font-size:.78em!important;}.q-table th,.q-table td{padding:4px 6px!important;}}
.options{list-style:none;padding:0;margin:0 0 16px;display:flex;flex-direction:column;gap:7px;}
.options li{background:var(--opt-bg);padding:9px 14px;border-radius:7px;border:1px solid var(--opt-border);font-size:.96em;transition:background .22s,border-color .22s;cursor:pointer;user-select:none;color:var(--text);}
.options li.correct-highlight{background:var(--accent-lt);border-color:var(--accent);font-weight:700;color:var(--accent-dk);}
.options li.wrong-highlight{background:#fde8e8;border-color:#e74c3c;font-weight:700;color:#c0392b;}
.btn-toggle{background:var(--green);color:#fff;border:none;padding:10px 0;width:100%;font-size:.96em;font-weight:700;font-family:inherit;border-radius:7px;cursor:pointer;transition:background .2s;}
.btn-toggle:hover{background:var(--green-dk);}
.answer-area{display:none;margin-top:13px;padding:14px 16px;background:var(--accent-lt);border-left:4px solid var(--accent);border-radius:0 7px 7px 0;}
.correct-ans{font-size:1.04em;color:var(--accent-dk);margin:0 0 7px;font-weight:700;}
.exp-text{margin:0;color:var(--text);line-height:1.55;font-size:.92em;}

/* ── PDF button & modal ── */
.pdf-open-btn{display:none;background:var(--green);color:#fff;border:none;padding:9px 11px;font-size:.82em;font-weight:700;font-family:inherit;border-radius:7px;cursor:pointer;width:100%;text-align:center;transition:background .2s;}
.pdf-open-btn:hover{background:var(--green-dk);}
.pdf-open-btn.visible{display:block;}
.pdf-modal{display:none;position:fixed;inset:0;z-index:1000;background:rgba(0,0,0,.65);padding:16px;align-items:stretch;}
.pdf-modal.open{display:flex;flex-direction:column;}
.pdf-modal-inner{background:#fff;border-radius:10px;overflow:hidden;display:flex;flex-direction:column;flex:1;}
.pdf-modal-header{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:var(--green-dk);color:#fff;font-weight:700;font-size:.9em;gap:8px;}
.pdf-modal-title{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.pdf-modal-close{background:none;border:none;color:#fff;font-size:1.5em;cursor:pointer;line-height:1;padding:0 4px;opacity:.8;}
.pdf-modal-close:hover{opacity:1;}
.pdf-nav-btn{background:rgba(255,255,255,.2);border:none;color:#fff;font-size:1.1em;cursor:pointer;border-radius:5px;padding:2px 10px;line-height:1.6;transition:background .15s;}
.pdf-nav-btn:hover{background:rgba(255,255,255,.35);}
.pdf-nav-btn:disabled{opacity:.3;cursor:default;}
.pdf-nav-counter{font-size:.8em;opacity:.85;white-space:nowrap;}
#pdfModalFrames{display:flex;flex-direction:column;flex:1;min-height:0;}
#pdfModalFrames iframe{flex:1;width:100%;border:none;min-height:0;}

/* ── footer ── */
.page-footer{text-align:center;margin-top:44px;padding:16px;font-size:.77em;color:#888;border-top:2px solid var(--border);}
.page-footer strong{color:var(--green-dk);}
"""

JS = """
let allAnswersVisible=false,scoreCorrect=0,scoreWrong=0;
function updateScore(){document.getElementById('score-correct').textContent=scoreCorrect;document.getElementById('score-wrong').textContent=scoreWrong;}
function resetScore(){scoreCorrect=0;scoreWrong=0;updateScore();}

function shortenLabel(e){
  return e
    .replace(/(\d{4})-(\d{4})/g,(_,a,b)=>a.slice(2)+'-'+b.slice(2))
    .replace(/^(\d{4}) /,(_,y)=>y.slice(2)+' ')
    .replace('Bütünleme','Büt');
}
function tagCssClass(e){
  if(e.includes('Kez Çıktı')) return 'freq-tag';
  if(e.includes('Üç Ders'))   return 'uc-tag';
  if(e.includes('Yaz Okulu')) return 'yaz-tag';
  if(e.includes('Vize'))      return 'vize-tag';
  if(e.includes('Final'))     return 'fin-tag';
  if(e.includes('Bütünleme')) return 'but-tag';
  return '';
}
function tagTfiClass(e){
  if(e.includes('Üç Ders'))   return 'tfi-uc';
  if(e.includes('Yaz Okulu')) return 'tfi-yaz';
  if(e.includes('Vize'))      return 'tfi-vize';
  if(e.includes('Final'))     return 'tfi-fin';
  if(e.includes('Bütünleme')) return 'tfi-but';
  return 'tfi-other';
}
function highlightNeg(text){
  var L='[a-zA-ZçğışöüÇĞİIŞÖÜ]';
  return text.replace(
    new RegExp(
      // Çok kelimeli kalıplar önce
      'doğru\\\\s+değil'+L+'*'
      +'|yanlış\\\\s+değil'+L+'*'
      +'|uygun\\\\s+değil'+L+'*'
      +'|gerekli\\\\s+değil'+L+'*'
      +'|mümkün\\\\s+değil'+L+'*'
      +'|'+L+'+mamas'+L+'\\\\s+gerek'+L+'*'
      +'|'+L+'+memes'+L+'\\\\s+gerek'+L+'*'
      +'|değil\\\\s+midir?'
      +'|değil\\\\s+mi(?!'+L+')'
      // Tek kelime negatifleri
      +'|değildir'
      +'|yanlıştır'
      +'|yoktur(?!'+L+')'
      +'|olamaz(?!'+L+')'
      +'|olmaz(?!'+L+')'
      // -mamaktadır/-memektedir (süregelen olumsuz)
      +'|'+L+'+mamaktadır'+L+'*'
      +'|'+L+'+memektedir'+L+'*'
      // -mamalıdır/-memelidir (zorunluluk olumsuz)
      +'|'+L+'+mamalıdır'
      +'|'+L+'+memelidir'
      +'|'+L+'+mamalı(?!'+L+')'
      +'|'+L+'+memeli(?!'+L+')'
      // -maz/-mez geniş (tüm Türkçe geniş zaman olumsuz)
      +'|'+L+'+maz(?!'+L+')'
      +'|'+L+'+mez(?!'+L+')',
      'gi'
    ),
    function(m){return '<u class="neg-hl">'+m+'</u>';}
  );
}
function getProgress(){try{return JSON.parse(localStorage.getItem(PROGRESS_KEY)||'{}');}catch(e){return {};}}
function saveProgress(qid,result){
  var p=getProgress();p[String(qid)]=result;
  localStorage.setItem(PROGRESS_KEY,JSON.stringify(p));
  updateProgressBar();
  var icon=document.getElementById('prev-'+qid);
  if(icon){icon.className='q-prev '+(result==='correct'?'q-prev-correct':'q-prev-wrong');icon.textContent=result==='correct'?'✓ Doğru':'✗ Yanlış';}
}
function updateProgressBar(){
  var p=getProgress();var total=sorular.length;
  var done=Object.keys(p).filter(function(k){return parseInt(k)<total;}).length;
  var el=document.getElementById('prog-done');if(el)el.textContent=done;
  var et=document.getElementById('prog-total');if(et)et.textContent=total;
  var fill=document.getElementById('prog-fill');if(fill)fill.style.width=(total?(done/total*100):0)+'%';
}
function resetProgress(){
  if(!confirm('Bu ders için tüm ilerleme sıfırlansın mı?'))return;
  localStorage.removeItem(PROGRESS_KEY);
  updateProgressBar();
  var activeTab=document.querySelector('.tab-btn.active');
  if(activeTab)activeTab.click();
}
function updateStatsBar(){
  searchQuestions(document.getElementById('searchBox')?document.getElementById('searchBox').value:'');
}
function applyTagFilter(fullLabel,tfiClass){
  document.querySelectorAll('.question-card').forEach(card=>{
    const exams=JSON.parse(card.dataset.exams||'[]');
    card.classList.toggle('tag-hidden',!exams.includes(fullLabel));
  });
  updateStatsBar();
  const ind=document.getElementById('tag-filter-indicator');
  ind.className='visible '+tfiClass;
  ind.querySelector('.tfi-label').textContent=shortenLabel(fullLabel)+' Görüntüleniyor';
  document.querySelector('.sidebar').classList.add('tag-filter-active');
  // PDF button
  const pdfBtn=document.getElementById('pdfOpenBtn');
  if(pdfBtn){
    const val=pdfMap[fullLabel];
    if(val){
      const ids=Array.isArray(val)?val:[val];
      pdfBtn.dataset.pdfIds=JSON.stringify(ids);
      pdfBtn.dataset.pdfLabels=JSON.stringify([fullLabel]);
      pdfBtn.dataset.pdfLabel=shortenLabel(fullLabel);
      pdfBtn.classList.add('visible');
    } else pdfBtn.classList.remove('visible');
  }
}
function clearTagFilter(){
  document.querySelectorAll('.question-card').forEach(c=>c.classList.remove('tag-hidden'));
  updateStatsBar();
  document.getElementById('tag-filter-indicator').className='';
  document.querySelector('.sidebar').classList.remove('tag-filter-active');
  const pdfBtn=document.getElementById('pdfOpenBtn');
  if(pdfBtn)pdfBtn.classList.remove('visible');
}
let _pdfIds=[],_pdfLabels=[],_pdfIdx=0;
function _pdfLoad(idx){
  const frames=document.getElementById('pdfModalFrames');
  frames.innerHTML='';
  const f=document.createElement('iframe');
  const _src=_pdfIds[idx];
  // Local path (contains '/' or ends with '.pdf') or legacy Google Drive ID
  f.src=(_src.includes('/')||_src.endsWith('.pdf'))?_src:'https://drive.google.com/file/d/'+_src+'/preview';
  f.allowFullscreen=true;
  frames.appendChild(f);
  const counter=document.getElementById('pdfCounter');
  if(_pdfIds.length>1){counter.textContent=(_pdfIdx+1)+' / '+_pdfIds.length;}
  else{counter.textContent='';}
  // Modal başlığını her PDF'e göre güncelle
  const lbl=_pdfLabels[idx]?shortenLabel(_pdfLabels[idx]):null;
  document.getElementById('pdfModalTitle').textContent=(lbl||document.getElementById('pdfOpenBtn').dataset.pdfLabel||'Sınav')+' Sınav Kağıdı';
  document.getElementById('pdfPrev').disabled=idx===0;
  document.getElementById('pdfNext').disabled=idx===_pdfIds.length-1;
}
function pdfNav(dir){
  const next=_pdfIdx+dir;
  if(next<0||next>=_pdfIds.length)return;
  _pdfIdx=next;_pdfLoad(_pdfIdx);
}
function openPdfModal(){
  const btn=document.getElementById('pdfOpenBtn');
  if(!btn)return;
  _pdfIds=JSON.parse(btn.dataset.pdfIds||'[]');
  _pdfLabels=JSON.parse(btn.dataset.pdfLabels||'[]');
  if(!_pdfIds.length)return;
  _pdfIdx=0;
  document.getElementById('pdfPrev').style.display=_pdfIds.length>1?'':'none';
  document.getElementById('pdfNext').style.display=_pdfIds.length>1?'':'none';
  _pdfLoad(0);
  document.getElementById('pdfModal').classList.add('open');
  document.body.style.overflow='hidden';
}
function closePdfModal(){
  document.getElementById('pdfModal').classList.remove('open');
  document.getElementById('pdfModalFrames').innerHTML='';
  document.body.style.overflow='';
}

function renderQuestions(filter){
  const container=document.getElementById('quiz-container');
  container.innerHTML='';resetScore();
  const examCount=s=>s.exams.filter(e=>!e.includes('Kez')).length;
  let filtered;
  if(filter==='all') filtered=[...sorular];
  else if(filter==='banko') filtered=sorular.filter(s=>examCount(s)>=2);
  else if(filter==='review'){const p=getProgress();filtered=sorular.filter(s=>p[String(s.id)]!=='correct');}
  else filtered=sorular.filter(s=>s.exams.some(e=>e.includes(filter)));
  filtered.sort((a,b)=>examCount(b)-examCount(a));
  if(filtered.length===0){container.innerHTML='<p style="text-align:center;color:#888;padding:30px;">Bu kategoride soru bulunamadı.</p>';document.getElementById('stats-bar').innerHTML=`<span>0<span class="stats-full"> soru gösteriliyor</span><span class="stats-short"> Soru</span></span>`;return;}
  filtered.forEach((soru,idx)=>{
    const card=document.createElement('div');card.className='question-card';
    card.dataset.exams=JSON.stringify(soru.exams);
    const tagsHTML=soru.exams.map(e=>{
      const css=tagCssClass(e);
      const isFreq=e.includes('Kez Çıktı');
      const clickable=isFreq?'':' tag-clickable';
      const data=isFreq?'':` data-full="${e.replace(/"/g,'&quot;')}" data-tfi="${tagTfiClass(e)}"`;
      return `<span class="exam-tag${css?' '+css:''}${clickable}"${data}>${shortenLabel(e)}</span>`;
    }).join('');
    const ansLetter=soru.ans?soru.ans.trim().charAt(0):'';
    const optionsHTML=soru.options.map(o=>{
      const isCorrect=ansLetter&&o.trim().startsWith(ansLetter+')');
      return `<li${isCorrect?' data-correct="true"':''} onclick="selectOption(this)">${o}</li>`;
    }).join('');
    const answerId=`answer-${idx}-${filter}`;
    card.dataset.qid=soru.id;
    const prevResult=getProgress()[String(soru.id)];
    const prevHTML=prevResult?`<span class="q-prev ${prevResult==='correct'?'q-prev-correct':'q-prev-wrong'}" id="prev-${soru.id}">${prevResult==='correct'?'✓ Doğru':'✗ Yanlış'}</span>`:`<span class="q-prev" id="prev-${soru.id}"></span>`;
    card.innerHTML=`
      <div class="q-header">
        <div class="q-meta"><span class="q-num">Soru ${idx+1}</span>${prevHTML}</div>
        <div class="exam-tags">${tagsHTML}</div>
      </div>
      <div class="question-text">${highlightNeg(soru.q)}</div>
      <ul class="options">${optionsHTML}</ul>
      <button class="btn-toggle" onclick="toggleAnswer('${answerId}',this)">Cevabı Göster</button>
      <div class="answer-area" id="${answerId}" style="${allAnswersVisible?'display:block':''}">
        <p class="correct-ans">✔ Doğru Cevap: ${soru.ans}</p>
        <p class="exp-text">${soru.exp}</p>
      </div>`;
    if(allAnswersVisible){const cl=card.querySelector('.options li[data-correct="true"]');if(cl)cl.classList.add('correct-highlight');card.querySelector('.btn-toggle').textContent='Cevabı Gizle';}
    card.querySelectorAll('.exam-tag.tag-clickable').forEach(tag=>{
      tag.addEventListener('click',()=>applyTagFilter(tag.dataset.full,tag.dataset.tfi));
    });
    card.querySelectorAll('.question-text table').forEach(t=>{t.classList.add('q-table');if(!t.parentNode.classList.contains('q-table-wrap')){var w=document.createElement('div');w.className='q-table-wrap';t.parentNode.insertBefore(w,t);w.appendChild(t);}});
    container.appendChild(card);
  });
  searchQuestions(document.getElementById('searchBox')?document.getElementById('searchBox').value:'');
}
function searchQuestions(q){
  q=(q||'').trim().toLowerCase();
  var cards=document.querySelectorAll('#quiz-container .question-card');
  var visible=0;
  cards.forEach(function(card){
    if(card.classList.contains('tag-hidden')){card.style.display='none';return;}
    var text=(card.querySelector('.question-text')||{}).textContent||'';
    var opts=Array.from(card.querySelectorAll('.options li')).map(function(li){return li.textContent;}).join(' ');
    var match=!q||text.toLowerCase().includes(q)||opts.toLowerCase().includes(q);
    card.style.display=match?'':'none';
    if(match)visible++;
  });
  document.getElementById('stats-bar').innerHTML='<span>'+visible+'<span class="stats-full"> soru gösteriliyor</span><span class="stats-short"> Soru</span></span>';
}
function setCorrectHighlight(card,visible){card.querySelectorAll('.options li[data-correct="true"]').forEach(li=>li.classList.toggle('correct-highlight',visible));}
function selectOption(li){
  const card=li.closest('.question-card');if(card.dataset.answered)return;
  card.dataset.answered='true';const isCorrect=li.dataset.correct==='true';
  if(isCorrect){li.classList.add('correct-highlight');card.dataset.answeredResult='correct';scoreCorrect++;}
  else{li.classList.add('wrong-highlight');card.querySelectorAll('.options li[data-correct="true"]').forEach(l=>l.classList.add('correct-highlight'));card.dataset.answeredResult='wrong';scoreWrong++;}
  saveProgress(card.dataset.qid,isCorrect?'correct':'wrong');
  updateScore();
  const area=card.querySelector('.answer-area');
  if(area&&area.style.display!=='block'){area.style.display='block';const btn=card.querySelector('.btn-toggle');if(btn)btn.textContent='Cevabı Gizle';}
}
function toggleAnswer(id,btn){
  const area=document.getElementById(id);const card=btn.closest('.question-card');
  if(area.style.display==='none'||area.style.display===''){
    area.style.display='block';btn.textContent='Cevabı Gizle';setCorrectHighlight(card,true);
  }else{
    area.style.display='none';btn.textContent='Cevabı Göster';setCorrectHighlight(card,false);
    card.querySelectorAll('.options li.wrong-highlight').forEach(li=>li.classList.remove('wrong-highlight'));
    if(card.dataset.answeredResult==='correct')scoreCorrect=Math.max(0,scoreCorrect-1);
    else if(card.dataset.answeredResult==='wrong')scoreWrong=Math.max(0,scoreWrong-1);
    delete card.dataset.answered;delete card.dataset.answeredResult;updateScore();
  }
}
function toggleAllAnswers(){
  allAnswersVisible=!allAnswersVisible;
  document.getElementById('toggleAllBtn').textContent=allAnswersVisible?'Tüm Cevapları Gizle':'Tüm Cevapları Göster';
  document.querySelectorAll('.answer-area').forEach(a=>a.style.display=allAnswersVisible?'block':'none');
  document.querySelectorAll('.btn-toggle').forEach(b=>b.textContent=allAnswersVisible?'Cevabı Gizle':'Cevabı Göster');
  document.querySelectorAll('.question-card').forEach(c=>setCorrectHighlight(c,allAnswersVisible));
}
function filterQuestions(filter,clickedBtn){
  clearTagFilter();
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  clickedBtn.classList.add('active');allAnswersVisible=false;
  document.getElementById('toggleAllBtn').textContent='Tüm Cevapları Göster';
  renderQuestions(filter);
  // Sekme filtresi için PDF butonunu güncelle
  const pdfBtn=document.getElementById('pdfOpenBtn');
  if(pdfBtn){
    const TAB_FILTERS=['Final','Bütünleme','Yaz Okulu','Vize','Üç Ders'];
    const isTabFilter=TAB_FILTERS.some(k=>filter===k||filter.includes(k));
    if(isTabFilter){
      const tabIds=[],tabLabels=[];
      // pdfMap anahtarlarını filtre anahtar kelimesine göre filtrele, kronolojik sırala
      Object.keys(pdfMap).sort().forEach(label=>{
        if(label.includes(filter)){
          const val=pdfMap[label];
          const ids=Array.isArray(val)?val:[val];
          ids.forEach(id=>{tabIds.push(id);tabLabels.push(label);});
        }
      });
      if(tabIds.length){
        pdfBtn.dataset.pdfIds=JSON.stringify(tabIds);
        pdfBtn.dataset.pdfLabels=JSON.stringify(tabLabels);
        pdfBtn.dataset.pdfLabel=filter;
        pdfBtn.classList.add('visible');
      } else {
        pdfBtn.classList.remove('visible');
      }
    } else {
      pdfBtn.classList.remove('visible');
    }
  }
}
renderQuestions('all');
updateProgressBar();
(function(){
  var _tb=document.querySelector('.site-topbar');
  var _sb=document.querySelector('.sidebar');
  var _qc=document.getElementById('quiz-container');
  function syncPad(){
    if(!_sb||!_qc)return;
    if(window.innerWidth<=720){
      var tbH=_tb?_tb.offsetHeight:48;
      _sb.style.top=tbH+'px';
      _qc.style.paddingTop=_sb.offsetHeight+'px';
    } else {
      _sb.style.top='';
      _qc.style.paddingTop='';
    }
  }
  syncPad();
  window.addEventListener('resize',syncPad,{passive:true});
  var mq=window.matchMedia('(max-width:720px)');
  var ticking=false;
  function onScroll(){
    if(ticking)return;
    ticking=true;
    requestAnimationFrame(function(){
      if(_sb)_sb.classList.toggle('scroll-collapsed',window.scrollY>90);
      ticking=false;
    });
  }
  function setup(){
    if(mq.matches){window.addEventListener('scroll',onScroll,{passive:true});}
    else{window.removeEventListener('scroll',onScroll);if(_sb)_sb.classList.remove('scroll-collapsed');syncPad();}
  }
  if(mq.addEventListener)mq.addEventListener('change',setup);
  else mq.addListener(setup);
  setup();
})();

/* ── Dark Mode ── */
(function(){
  var toggle=document.getElementById('darkModeToggle');
  if(!toggle)return;
  var KEY='ataaof-aday-dark-mode';
  var prefersDark=window.matchMedia('(prefers-color-scheme: dark)').matches;
  var saved=localStorage.getItem(KEY);
  var isDark=saved!==null?saved==='true':prefersDark;
  if(isDark)document.body.classList.add('dark-mode');
  toggle.textContent=isDark?'☀️':'🌙';
  toggle.addEventListener('click',function(){
    document.body.classList.toggle('dark-mode');
    var nowDark=document.body.classList.contains('dark-mode');
    localStorage.setItem(KEY,String(nowDark));
    toggle.textContent=nowDark?'☀️':'🌙';
  });
})();
// Tekrar Et toggle
function toggleReview(btn){
  var isActive=btn.classList.contains('active');
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  allAnswersVisible=false;
  document.getElementById('toggleAllBtn').textContent='Tüm Cevapları Göster';
  clearTagFilter();
  if(isActive){
    var allBtn=document.querySelector('.tab-btn:first-child');
    if(allBtn)allBtn.classList.add('active');
    renderQuestions('all');
  } else {
    btn.classList.add('active');
    renderQuestions('review');
  }
}
"""


def generate_html(course, emoji, questions, tabs, home_url="index.html", pdfs=None):
    total    = len(questions)
    pdfs     = pdfs or {}
    pdf_map_js = json.dumps(pdfs, ensure_ascii=False)

    # HTML-escape text fields so raw tags like <footer> don't break innerHTML rendering
    def _he(s):
        return escape(s, quote=False) if isinstance(s, str) else s

    sorular_js = ",\n".join(
        f"{{id:{i},q:{json.dumps(q['q'],ensure_ascii=False)},"
        f"options:{json.dumps([_he(o) for o in q['options']],ensure_ascii=False)},"
        f"ans:{json.dumps(_he(q['ans']),ensure_ascii=False)},"
        f"exp:{json.dumps(_he(q['exp']),ensure_ascii=False)},"
        f"exams:{json.dumps(q['exams'],ensure_ascii=False)}}}"
        for i, q in enumerate(questions)
    )
    progress_key = json.dumps(f"auzef-progress-{course}")

    tab_btns = "\n".join(
        f'        <button class="tab-btn" onclick="filterQuestions(\'{fval}\',this)">{label}</button>'
        for label, fval in tabs
    )

    type_labels = [lbl for lbl, _ in tabs]
    subtitle = " · ".join(["Eksiksiz Çıkmış Sorular"] + type_labels) if type_labels else "Çıkmış Sınav Soruları"

    ec = escape(course)
    es = escape(subtitle)

    return f"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0">
<title>{ec} – Soru Havuzu</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>

<div class="site-topbar">
  <div class="topbar-left">
    <a href="{home_url}" class="topbar-home">← Ana Sayfa</a>
    <span class="topbar-uni">İstanbul Üniversitesi · AUZEF · Yönetim Bilişim Sistemleri</span>
  </div>
  <div class="topbar-right">
    <span class="topbar-course">{emoji} {ec}</span>
    <button id="darkModeToggle" class="dark-mode-btn" aria-label="Dark Mode" title="Karanlık Mod">🌙</button>
  </div>
</div>

<div class="header">
  <div class="header-badge">{emoji} AUZEF Soru Havuzu</div>
  <h1>{ec}</h1>
  <p class="header-subtitle">{es}</p>
</div>

<div class="container">
  <div class="page-layout">
    <div id="quiz-container"></div>

    <aside class="sidebar">
      <a href="{home_url}" class="sidebar-home">⌂ Ana Sayfa</a>

      <span class="sidebar-label">Filtrele</span>
      <nav class="tabs">
        <button class="tab-btn active" onclick="filterQuestions('all',this)">Tüm Sorular</button>
{tab_btns}
        <button class="tab-btn tab-banko" onclick="filterQuestions('banko',this)">🔥 Çok Çıkanlar</button>
        <button class="tab-btn tab-review" onclick="toggleReview(this)">🔄 Devam Et
          <span class="review-info-btn" onclick="event.stopPropagation()" tabindex="0">ℹ</span>
          <div class="review-tooltip">Daha önce doğru cevapladığın sorular çıkarılarak teste devam eder.</div>
        </button>
      </nav>

      <input class="search-box" id="searchBox" type="search" placeholder="🔍 Sorularda ara..." oninput="searchQuestions(this.value)">

      <div id="tag-filter-indicator">
        <span class="tfi-label"></span>
        <button class="tfi-close" onclick="clearTagFilter()">✕</button>
      </div>
      <button id="pdfOpenBtn" class="pdf-open-btn" onclick="openPdfModal()">📄 Sınav PDF Göster</button>

      <div class="progress-wrap">
        <div class="progress-label">
          <span><span id="prog-done">0</span> / <span id="prog-total">0</span><span class="prog-label"> tamamlandı</span></span>
          <button class="progress-reset" onclick="resetProgress()" title="İlerlemeyi sıfırla">sıfırla</button>
        </div>
        <div class="progress-bar-bg"><div class="progress-bar-fill" id="prog-fill" style="width:0%"></div></div>
      </div>

      <div class="sidebar-meta">
        <div class="stats-bar" id="stats-bar"></div>
        <div class="score-panel">
          <div class="score-item score-correct-item">✓ <span id="score-correct">0</span><span class="score-label"> Doğru</span></div>
          <div class="score-item score-wrong-item">✗ <span id="score-wrong">0</span><span class="score-label"> Yanlış</span></div>
        </div>
      </div>
      <button class="toggle-all-btn" id="toggleAllBtn" onclick="toggleAllAnswers()">Tüm Cevapları Göster</button>
    </aside>
  </div>

  <div class="page-footer">
    <strong>{ec}</strong> · Çıkmış Sorular Soru Havuzu · {total} Soru
  </div>
</div>

<div class="pdf-modal" id="pdfModal">
  <div class="pdf-modal-inner">
    <div class="pdf-modal-header">
      <span class="pdf-modal-title" id="pdfModalTitle">Sınav Kağıdı</span>
      <button class="pdf-nav-btn" id="pdfPrev" onclick="pdfNav(-1)">&#8249;</button>
      <span class="pdf-nav-counter" id="pdfCounter"></span>
      <button class="pdf-nav-btn" id="pdfNext" onclick="pdfNav(1)">&#8250;</button>
      <button class="pdf-modal-close" onclick="closePdfModal()">✕</button>
    </div>
    <div id="pdfModalFrames"></div>
  </div>
</div>

<script>
const sorular=[
{sorular_js}
];
const pdfMap={pdf_map_js};
const PROGRESS_KEY={progress_key};
{JS}
</script>
</body>
</html>"""


# ─── Excel standardisation ────────────────────────────────────────────────────

def standardize_excel(path):
    wb      = pd.ExcelFile(path)
    changed = False
    sheets  = {}

    for sheet in wb.sheet_names:
        df = wb.parse(sheet)

        if sheet == "Renk Açıklaması":
            if "Unnamed: 0" in df.columns:
                COLOR_MAP = {"mavi":"Mavi","yeşil":"Yeşil","sarı":"Sarı",
                             "turuncu":"Turuncu","kırmızı":"Kırmızı","mor":"Mor"}
                renk_col, acik_col = [], []
                for a in df["Açıklama"].dropna().tolist():
                    a_str = str(a)
                    color = next((v for k,v in COLOR_MAP.items() if k in a_str.lower()), "")
                    clean = re.sub(r'\s*\(.*?\)\s*','',a_str).strip()
                    clean = re.sub(r'\s+soruları?$','',clean).strip()
                    renk_col.append(color); acik_col.append(clean)
                df = pd.DataFrame({"Renk": renk_col, "Açıklama": acik_col})
                changed = True

        elif "Soru" in df.columns:
            # Rename old-format columns (A→"A Şıkkı", Sınavlar→"Sınav(lar)")
            _alias = {old: new for old, new in
                      [("A","A Şıkkı"),("B","B Şıkkı"),("C","C Şıkkı"),("D","D Şıkkı"),("E","E Şıkkı"),
                       ("Sınavlar","Sınav(lar)")]
                      if old in df.columns and new not in df.columns}
            if _alias:
                df = df.rename(columns=_alias)
                changed = True

            for letter, col in zip(OPTION_LETTERS, OPTION_COLS):
                if col in df.columns:
                    orig = df[col].copy()
                    df[col] = df[col].apply(
                        lambda v: clean_option(v, letter)
                        if not (isinstance(v, float) and v != v) else v)
                    if not orig.equals(df[col]):
                        changed = True

            if "Doğru Cevap" in df.columns:
                orig = df["Doğru Cevap"].copy()
                df["Doğru Cevap"] = df["Doğru Cevap"].apply(
                    lambda v: clean_answer_letter(v)
                    if not (isinstance(v, float) and v != v) else v)
                if not orig.equals(df["Doğru Cevap"]):
                    changed = True

            if "Sınav(lar)" in df.columns:
                orig = df["Sınav(lar)"].copy()
                df["Sınav(lar)"] = df["Sınav(lar)"].apply(
                    lambda v: normalize_exams_str(v)
                    if not (isinstance(v, float) and v != v) else v)
                if not orig.equals(df["Sınav(lar)"]):
                    changed = True

        sheets[sheet] = df

    if changed:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sheet, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  ✓ Excel standardize edildi: {path}")
    else:
        print(f"  ✓ Excel zaten standart: {path}")
    return changed


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Excel → HTML quiz sayfası üretici")
    parser.add_argument("ders",          help='Ders adı. Örn: "Kamuoyu Araştırmaları"')
    parser.add_argument("--excel",       help="Excel dosya yolu (varsayılan: otomatik)")
    parser.add_argument("--output",      help="Çıktı HTML yolu (varsayılan: otomatik)")
    parser.add_argument("--no-standardize", action="store_true",
                        help="Excel'i standartlaştırmayı atla")
    args = parser.parse_args()

    course      = args.ders
    config      = DERS_CONFIG.get(course, {})
    excel_path  = args.excel  or config.get("excel")
    output_path = args.output or config.get("output")
    emoji       = config.get("emoji", "📚")
    pdfs        = config.get("pdfs", {})
    pdf_dir     = config.get("pdf_dir")
    if pdf_dir and not pdfs:
        pdfs = discover_pdfs(pdf_dir, output_path)

    if not excel_path:
        print(f"HATA: '{course}' için Excel bulunamadı. --excel ile belirtin.")
        sys.exit(1)
    if not output_path:
        safe        = re.sub(r"[^\w]", "_", course)
        output_path = f"{safe}_Soru_Bankasi.html"

    if not args.no_standardize:
        print("[1/3] Excel standardize ediliyor...")
        standardize_excel(excel_path)

    print(f"[2/3] Sorular okunuyor: {excel_path}")
    df        = pd.read_excel(excel_path, sheet_name="Tüm Sorular")
    questions = build_questions(df)
    tabs      = detect_tabs(questions)
    print(f"  {len(questions)} soru  |  sekmeler: {[lbl for lbl,_ in tabs]}")

    depth    = len(Path(output_path).parts) - 1
    home_url = "../" * depth + "index.html"

    print(f"[3/3] HTML oluşturuluyor: {output_path}  (ana sayfa: {home_url})")
    html = generate_html(course, emoji, questions, tabs, home_url=home_url, pdfs=pdfs)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✓ Tamamlandı → {output_path}")


if __name__ == "__main__":
    main()
