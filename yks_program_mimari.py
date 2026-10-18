import streamlit as st
import openpyxl
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import PatternFill, Font, Alignment
from copy import copy
from io import BytesIO
import datetime
import math
import re

try:
    import yt_dlp
    YTDLP_AVAILABLE = True
except ImportError:
    YTDLP_AVAILABLE = False

st.set_page_config(page_title="Akıllı YKS Program Mimarı v2", page_icon="🗓️", layout="wide")

# --- TÜRKÇE KARAKTER DUYARLI KÜÇÜK HARF DÖNÜŞTÜRÜCÜ ---
def tr_lower(text):
    if not text:
        return ""
    return str(text).replace('İ', 'i').replace('I', 'ı').lower().strip()

def normalize_subject_name(raw_name):
    words = str(raw_name).strip().split()
    out = []
    for w in words:
        up = w.replace('i', 'İ').replace('ı', 'I').upper()
        if up in ["TYT", "AYT", "YKS", "MSÜ", "DYK"]:
            out.append(up)
        else:
            first_char = w[0].replace('i', 'İ').replace('ı', 'I').upper()
            rest = tr_lower(w[1:])
            out.append(first_char + rest)
    return " ".join(out)

# --- TAKVİM VE RESMİ TATİLLER ---
TR_MONTHS = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan",
    5: "Mayıs", 6: "Haziran", 7: "Temmuz", 8: "Ağustos",
    9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

FIXED_HOLIDAYS = {
    (1, 1): "1 Ocak Yılbaşı Tatili",
    (4, 23): "23 Nisan Ulusal Egemenlik ve Çocuk Bayramı",
    (5, 1): "1 Mayıs Emek ve Dayanışma Günü",
    (5, 19): "19 Mayıs Atatürk'ü Anma, Gençlik ve Spor Bayramı",
    (7, 15): "15 Temmuz Demokrasi ve Millî Birlik Günü",
    (8, 30): "30 Ağustos Zafer Bayramı",
    (10, 29): "29 Ekim Cumhuriyet Bayramı Tatili"
}

IGNORE_KEYWORDS = [
    "(zorunlu)", "öğle arası", "akşam yemeği", "ara öğün", "uyku vakti",
    "kahvaltı", "serbest", "dyk kursu", "deneme", "çarşı izni",
    "okuldaki derslerin", "dershanedeki derslerin", "kitap okuma",
    "sorularını hocalara sor", "problem + paragraf", "eksik kapatma", "dershane"
]

DEFAULT_COLORS = {
    "fizik": "#53D2FF",
    "kimya": "#92D050",
    "biyoloji": "#FFFF00",
    "matematik": "#FFC000",
    "geometri": "#8EA9DB",
    "türkçe": "#EEF7BB",
    "dil bilgisi": "#EEF7BB",
    "tarih": "#F4B084",
    "coğrafya": "#A9D08E",
    "felsefe": "#D9E1F2"
}

# =====================================================================
# 1. GÖMÜLÜ MEB MÜFREDAT KONU VE AĞIRLIK HAVUZU (OYNATMA LİSTESİZ MOD İÇİN)
# =====================================================================
MEB_CURRICULUM = {
    "TYT Fizik": [
        ("Fizik Bilimine Giriş", 1), ("Madde ve Özellikleri (Özkütle, Dayanıklılık, Adezyon-Kohezyon)", 2),
        ("Düzgün Doğrusal Hareket ve İvme", 2), ("Kuvvet ve Newton'ın Hareket Yasaları", 2),
        ("İş, Güç ve Enerji", 2), ("Isı, Sıcaklık ve Genleşme", 3), ("Elektrostatik", 2),
        ("Elektrik Akımı, Direnç, Üreteç ve Lambalar", 3), ("Manyetizma", 1),
        ("Katı, Sıvı, Gaz ve Akışkan Basıncı", 2), ("Kaldırma Kuvveti", 2),
        ("Dalgalar (Temel Kavramlar, Yay, Su, Ses ve Deprem)", 3),
        ("Optik - 1 (Aydınlanma, Gölge, Düzlem ve Küresel Aynalar)", 3),
        ("Optik - 2 (Kırılma, Mercekler, Prizma ve Renkler)", 3)
    ],
    "TYT Kimya": [
        ("Kimya Bilimi (Simya, Disiplinler ve Sembolik Dil)", 2), ("Atom Modelleri ve Atomun Yapısı", 2),
        ("Periyodik Sistem ve Periyodik Özellikler", 2), ("Kimyasal Türler Arası Güçlü Etkileşimler", 2),
        ("Zayıf Etkileşimler ve Fiziksel-Kimyasal Değişimler", 2), ("Maddenin Halleri (Katı, Sıvı, Gaz ve Hal Değişimi)", 3),
        ("Doğa ve Kimya", 1), ("Kimyanın Temel Kanunları", 2), ("Mol Kavramı", 2),
        ("Kimyasal Tepkime Türleri ve Hesaplamalar", 3), ("Karışımlar, Derişim ve Ayırma Teknikleri", 3),
        ("Asitler, Bazlar ve Tuzlar", 3), ("Kimya Her Yerde", 1)
    ],
    "TYT Biyoloji": [
        ("Canlıların Ortak Özellikleri ve İnorganik Bileşikler", 2),
        ("Organik Bileşikler (Karbohidrat, Lipit, Protein, Vitamin)", 2),
        ("Enzimler, Hormonlar, Nükleik Asitler (DNA-RNA) ve ATP", 2),
        ("Hücre Zarı ve Madde Geçişleri", 2), ("Hücre Organelleri ve Çekirdek", 2),
        ("Canlıların Sınıflandırılması (Bakteri, Arke, Protista, Bitki, Mantar, Hayvan, Virüs)", 3),
        ("Hücre Bölünmeleri (Mitoz ve Eşeysiz Üreme)", 2), ("Mayoz Bölünme ve Eşeyli Üreme", 2),
        ("Kalıtım (Mendel, Kan Grupları, Eşeye Bağlı Kalıtım ve Soy Ağaçları)", 4),
        ("Ekosistem Ekolojisi, Madde Döngüleri ve Güncel Çevre Sorunları", 2)
    ],
    "TYT Matematik": [
        ("Temel Kavramlar ve Sayı Kümeleri", 2), ("Tek-Çift, Ardışık ve Asal Sayılar, Faktöriyel", 2),
        ("Sayı Basamakları ve Çözümleme", 2), ("Bölme ve Bölünebilme", 2), ("EBOB - EKOK", 2),
        ("Rasyonel ve Ondalık Sayılar", 2), ("1. Dereceden Denklemler ve Basit Eşitsizlikler", 2),
        ("Mutlak Değer", 2), ("Üslü Sayılar", 2), ("Köklü Sayılar", 2), ("Çarpanlara Ayırma", 2),
        ("Oran - Orantı", 2), ("Sayı ve Kesir Problemleri", 4), ("Yaş ve İşçi Problemleri", 2),
        ("Hareket (Hız) Problemleri", 3), ("Yüzde, Kâr-Zarar ve Karışım Problemleri", 3),
        ("Grafik ve Rutin Olmayan Problemler", 2), ("Mantık", 1), ("Kümeler ve Kartezyen Çarpım", 2),
        ("Fonksiyonlar", 4), ("Polinomlar ve 2. Dereceden Denklemler", 3),
        ("Permütasyon, Kombinasyon ve Binom", 3), ("Olasılık ve Veri-İstatistik", 3)
    ],
    "TYT Geometri": [
        ("Doğruda ve Üçgende Açılar", 2), ("Dik Üçgen, Pisagor ve Öklid", 2),
        ("İkizkenar ve Eşkenar Üçgen", 2), ("Üçgende Açıortay ve Kenarortay", 2),
        ("Üçgende Eşlik ve Benzerlik", 3), ("Üçgende Alan ve Açı-Kenar Bağıntıları", 3),
        ("Çokgenler ve Genel Dörtgenler", 2), ("Paralelkenar ve Eşkenar Dörtgen", 2),
        ("Dikdörtgen, Kare ve Deltoid", 3), ("Yamuk", 2), ("Çemberde Açı ve Uzunluk", 3),
        ("Dairede Çevre ve Alan", 2), ("Analitik Geometri (Nokta ve Doğru Analitiği)", 3),
        ("Katı Cisimler (Prizma, Küp, Silindir, Piramit, Koni, Küre)", 3)
    ],
    "TYT Türkçe": [
        ("Sözcükte ve Söz Öbeklerinde Anlam", 2), ("Cümlede Anlam ve Kavramlar", 2),
        ("Paragrafta Yapı, Ana Düşünce ve Yardımcı Düşünce", 3), ("Ses Bilgisi", 1),
        ("Sözcükte Yapı ve Ekler", 2), ("Sözcük Türleri (İsim, Sıfat, Zamir, Zarf, Edat-Bağlaç)", 3),
        ("Fiiller, Ek Fiil ve Fiilimsiler", 2), ("Fiilde Çatı", 1), ("Cümlenin Ögeleri", 2),
        ("Cümle Türleri", 1), ("Yazım Kuralları", 2), ("Noktalama İşaretleri", 2), ("Anlatım Bozuklukları", 1)
    ],
    "TYT Tarih": [
        ("Tarih ve Zaman", 1), ("İnsanlığın İlk Dönemleri ve Orta Çağ'da Dünya", 2),
        ("İlk ve Orta Çağlarda Türk Dünyası", 2), ("İslam Medeniyetinin Doğuşu ve İlk Türk İslam Devletleri", 2),
        ("Yerleşme ve Devletleşme Sürecinde Selçuklu Türkiyesi", 1),
        ("Beylikten Devlete Osmanlı Siyaseti ve Dünya Gücü Osmanlı", 2),
        ("Osmanlı Kültür ve Medeniyeti", 1), ("Değişen Dünya Dengeleri, Duraklama ve Gerileme Dönemi", 2),
        ("Osmanlı Dağılma Dönemi ve Islahatlar", 2), ("XX. Yüzyıl Başlarında Osmanlı ve I. Dünya Savaşı", 2),
        ("Milli Mücadele Hazırlık Dönemi (Genelgeler ve Kongreler)", 2),
        ("I. TBMM Dönemi ve Kurtuluş Savaşı Cepheler/Antlaşmalar", 2),
        ("Atatürkçülük, Atatürk İlkeleri ve Türk İnkılabı", 2), ("Atatürk Dönemi Türk Dış Politikası", 1)
    ],
    "TYT Coğrafya": [
        ("Doğa ve İnsan, Dünya'nın Şekli ve Hareketleri", 2), ("Coğrafi Konum ve Harita Bilgisi", 2),
        ("Atmosfer, Sıcaklık, Basınç ve Rüzgarlar", 2), ("Nem, Yağış, Büyük İklim Tipleri ve Türkiye İklimi", 3),
        ("İç Kuvvetler, Kayaçlar ve Tektonik Oluşum", 2), ("Dış Kuvvetler ve Türkiye'nin Yer Şekilleri", 3),
        ("Dünya'da ve Türkiye'de Su, Toprak ve Bitki Varlığı", 3),
        ("Nüfus, Nüfus Piramitleri, Göçler ve Yerleşme", 3),
        ("Ekonomik Faaliyetler, Bölge Türleri ve Ulaşım Yolları", 2), ("Çevre ve Toplum, Doğal Afetler", 2)
    ],
    "AYT Matematik": [
        ("Polinomlar ve Çarpanlara Ayırma", 2), ("2. Dereceden Denklemler, Karmaşık Sayılar ve Parabol", 4),
        ("2. Dereceden Eşitsizlikler", 2), ("Fonksiyonlarda Uygulamalar ve Dönüşümler", 3),
        ("Trigonometri (11. ve 12. Sınıf Tamamı)", 5), ("Logaritma", 3), ("Diziler (Aritmetik ve Geometrik)", 3),
        ("Limit ve Süreklilik", 3), ("Türev ve Türev Uygulamaları", 6), ("İntegral ve Alan Hesabı", 6)
    ],
    "AYT Fizik": [
        ("Vektörler ve Bağıl Hareket", 2), ("Newton'ın Hareket Yasaları", 2),
        ("Bir ve İki Boyutta Sabit İvmeli Hareket (Atışlar)", 3), ("İş, Güç, Enerji ve İtme-Momentum", 4),
        ("Tork, Denge ve Kütle Merkezi", 3), ("Basit Makineler", 2),
        ("Elektriksel Kuvvet, Alan, Potansiyel ve Paralel Levhalar, Sığaçlar", 4),
        ("Manyetizma ve Elektromanyetik İndüksiyon, Alternatif Akım", 4),
        ("Düzgün Çembersel Hareket, Dönerek Öteleme ve Açısal Momentum", 3),
        ("Kütle Çekim, Kepler ve Basit Harmonik Hareket", 3),
        ("Dalga Mekaniği (Su, Ses, Işıkta Girişim, Kırınım, Doppler) ve EM Dalgalar", 3),
        ("Atom Fiziğine Giriş, Radyoaktivite ve Modern Fizik (Özel Görelilik, Fotoelektrik, Compton)", 4),
        ("Modern Fiziğin Teknolojideki Uygulamaları", 2)
    ],
    "AYT Kimya": [
        ("Modern Atom Teorisi ve Periyodik Özellikler", 2), ("Gazlar", 3), ("Sıvı Çözeltiler ve Çözünürlük", 3),
        ("Kimyasal Tepkimelerde Enerji (Entalpi)", 2), ("Kimyasal Tepkimelerde Hız", 2),
        ("Kimyasal Tepkimelerde Denge", 3), ("Sulu Çözelti Dengeleri (Asit-Baz ve Kçç)", 4),
        ("Kimya ve Elektrik (Redoks, Piller, Elektroliz)", 4), ("Karbon Kimyasına Giriş", 2),
        ("Organik Bileşikler (Hidrokarbonlar ve Fonksiyonel Gruplar)", 5)
    ],
    "AYT Biyoloji": [
        ("Sinir Sistemi ve Endokrin Sistem, Duyu Organları", 4), ("Destek ve Hareket Sistemi", 2),
        ("Sindirim Sistemi", 2), ("Dolaşım ve Bağışıklık Sistemi", 3), ("Solunum Sistemi", 2),
        ("Üriner Sistem ve Üreme Sistemi", 3), ("Komünite ve Popülasyon Ekolojisi", 2),
        ("Genden Proteine (Nükleik Asitler, Protein Sentezi, Biyoteknoloji)", 3),
        ("Canlılarda Enerji Dönüşümleri (Fotosentez, Kemosentez ve Hücresel Solunum)", 4),
        ("Bitki Biyolojisi", 4), ("Canlılar ve Çevre", 1)
    ]
}

# =====================================================================
# 2. GÖMÜLÜ POPÜLER HOCA KAMPLARI (TEK TIKLA YÜKLEME İÇİN)
# =====================================================================
BUILTIN_CAMPS = {
    "Özcan Aykın - 55 Günde TYT Fizik": [
        "1. Gün: Fizik Bilimine Giriş - 1 & 2 (Alt Dallar, Büyüklükler)",
        "1. Gün: Fizik Bilimine Giriş Soru Çözümü + Soru Bankası",
        "2. Gün: Madde ve Özellikleri - 1 & 2 (Kütle, Hacim, Özkütle)",
        "3. ve 4. Gün: Karışımların Özkütlesi ve Dayanıklılık",
        "5. Gün: Madde ve Özellikleri - 5 (Adezyon ve Kohezyon)",
        "3. ve 5. Gün: Madde ve Özellikleri Soru Çözümü - 1 & 2",
        "6. Gün: Hareket - 1 & 2 (Temel Kavramlar, Düzgün Doğrusal Hareket)",
        "7. Gün: Hareket - 3 (İvme)",
        "7. Gün: Hareket Soru Çözümü - 1 & 2 + Soru Bankası",
        "8. Gün: Kuvvet - 1 & 2 (Temel Kuvvetler, Newton Yasaları)",
        "9. Gün: Kuvvet - 3 & 4 (Newton Yasaları, Sürtünme Kuvveti)",
        "10. Gün: Kuvvet Soru Çözümü - 1 & 2 + Soru Bankası",
        "11. Gün: İş, Güç, Enerji - 1 & 2 (İş Kavramı, Mekanik Enerji)",
        "12. Gün: İş, Güç, Enerji - 3 & 4 (Enerji Korunumu, Güç, Verim)",
        "13. Gün: İş, Güç, Enerji Soru Çözümü - 1 & 2 + Soru Bankası",
        "14. Gün: Isı ve Sıcaklık - 1 & 2 (İç Enerji, Sıcaklık Ölçekleri)",
        "15. Gün: Isı ve Sıcaklık - 3 & 4 (Öz Isı, Isı Sığası, Hal Değişimi)",
        "16. Gün: Isı ve Sıcaklık - 5 & 6 (Isıl Denge, Enerji İletim Yolları)",
        "18. Gün: Genleşme - 1 & 2",
        "17. ve 18. Gün: Isı-Sıcaklık Soru Çözümü 1-2 ve Genleşme Soru Çözümü",
        "19. Gün: Elektrostatik - 1 & 2 (Elektrik Yükleri, Elektriklenme)",
        "20. Gün: Elektrostatik - 3 & 4 (Elektroskop, Topraklama)",
        "21. Gün: Elektrostatik - 5 & 6 (Coulomb Kuvveti, Elektrik Alan)",
        "22. Gün: Elektrostatik Soru Çözümü - 1 & 2",
        "23. Gün: Elektrik Akımı - 1 & 2 (Akım, Potansiyel Fark, Direnç)",
        "24. Gün: Elektrik Devreleri - 1 & 2 (Ohm Yasası, Dirençler)",
        "25. Gün: Elektrik Devreleri - 3 (Üreteçler) + Soru Çözümü 1-2",
        "26. Gün: Elektrik Devreleri - 4 (Elektriksel Güç ve Lambalar) + Soru Çözümü",
        "27. Gün: Manyetizma - 1 & 2 (Mıknatıs, Akım ve Manyetik Alan)",
        "28. Gün: Manyetizma - 3 & 4 (Elektromıknatıs, Dünya'nın Manyetik Alanı)",
        "29. Gün: Manyetizma Soru Çözümü - 1 & 2 + Soru Bankası",
        "30. Gün: Basınç - 1 (Katı Basıncı) + Soru Çözümü",
        "31. Gün: Basınç - 2, 3 & 4 (Sıvı Basıncı 1, 2 ve 3)",
        "32. Gün: Basınç - 5 & 6 (Açık Hava ve Kapalı Kap Gaz Basıncı)",
        "33. Gün: Basınç - 7 & 8 (Basınç Ölçen Aletler, Akışkan Basıncı)",
        "34. Gün: Durgun Sıvı, Akışkan ve Gaz Basıncı Soru Çözümü (1-2)",
        "35. Gün: Kaldırma Kuvveti - 1 & 2",
        "36. Gün: Kaldırma Kuvveti Soru Çözümü - 1 & 2 + Soru Bankası",
        "37. Gün: Optik: Aydınlanma - 1 & 2 (Işık Şiddeti, Akı, Aydınlanma)",
        "38. Gün: Optik: Gölge - 1 & 2 + Aydınlanma/Gölge Soru Çözümü",
        "39. Gün: Optik: Yansıma Kanunları ve Düzlem Ayna - 1",
        "40. Gün: Optik: Düzlem Ayna - 2 (Görüş Alanı) + Soru Çözümü",
        "41. Gün: Küresel Aynalar - 1 & 2 (Çukur ve Tümsek Özel Işınlar)",
        "42. Gün: Küresel Aynalar - 3 & 4 (Görüntü) + 43. Gün: Soru Çözümü",
        "44. Gün: Işığın Kırılması (Snell Yasası, Sınır Açısı, Tam Yansıma)",
        "45. Gün: Küresel Yüzeylerden Geçiş, Görünür Uzaklık ve Soru Çözümü",
        "46. Gün: Mercekler, Odak Uzaklığı ve Özel Işınlar",
        "47. Gün: İnce ve Kalın Kenarlı Merceklerde Görüntü + Soru Çözümü",
        "48. Gün: Optik: Prizmalar ve Renkler",
        "49. Gün: Prizma ve Renkler Soru Çözümü (Optik Bitti)",
        "50. Gün: Dalgalar - 1 & 2 (Temel Kavramlar ve Sınıflandırma)",
        "51. ve 52. Gün: Yay Dalgaları - 1, 2 & 3 + Soru Çözümü",
        "53. ve 54. Gün: Su Dalgaları - 1, 2, 3 & 4 (Yansıma, Hız ve Kırılma)",
        "55. Gün: Ses ve Deprem Dalgaları + Soru Çözümü (Kamp Bitti)"
    ],
    "Görkem Şahin - 47 Günde TYT Kimya": [
        "Giriş + 1) Simyadan Kimyaya", "2) Kimyanın ve Kimyacıların Çalışması",
        "3) Kimyanın Sembolik Dili", "4) Kimya Uygulamalarında İş Sağlığı ve Güvenliği",
        "5) Atom Modelleri", "6) Atomun Yapısı", "7) Periyodik Sistem",
        "8) Elementlerin Sınıflandırılması", "9) Periyodik Özellikler 1", "10) Periyodik Özellikler 2",
        "11) Kimyasal Türler Arası Etkileşimler", "12) Güçlü Etkileşimler - 1 (İyonik Bağ)",
        "13) Güçlü Etkileşimler - 2 (İyonik Bağ)", "14) Kovalent Bağların Oluşumu",
        "15) Kovalent Bileşiklerin Adlandırılması ve Metalik Bağ", "16) Zayıf Etkileşimler",
        "17) Fiziksel ve Kimyasal Değişimler", "18) Maddenin Halleri", "19) Katılar",
        "20) Sıvılar", "21) Gazlar", "22) Doğa ve Kimya", "23) Kütlenin Korunumu Kanunu",
        "24) Sabit Oranlar Kanunu", "25) Katlı Oranlar Kanunu", "26) Mol Kavramı 1",
        "27) Mol Kavramı 2", "28) Mol Kavramı Soru Çözümü", "29) Kimyasal Tepkime Türleri",
        "30) Kimyasal Hesaplamalar 1", "31) Kimyasal Hesaplamalar 2", "32) Kimyasal Hesaplamalar 3",
        "33) Kimyasal Hesaplamalar Soru Çözümü", "34) Karışımların Sınıflandırılması 1",
        "35) Karışımların Sınıflandırılması 2", "36) Çözünme Süreci", "37) Çözeltilerde Derişim",
        "38) Koligatif Özellikler", "39) Ayırma ve Saflaştırma Teknikleri", "40) Asitler ve Bazlar",
        "41) Asitler ve Bazların Tepkimeleri 1", "42) Asitler ve Bazların Tepkimeleri 2",
        "43) Hayatımızdaki Asitler - Bazlar ve Tuzlar", "44) Kimya Her Yerde - 1 & 2 (Kamp Bitti)"
    ],
    "Biosem (Zeduva) - 39 Günde TYT Biyoloji": [
        "1. Gün: Canlıların Genel Özellikleri", "2. Gün: İnorganik Bileşikler",
        "3. Gün: Karbohidratlar", "4. Gün: Lipitler (Yağlar)", "5. Gün: Proteinler",
        "6. Gün: Hormonlar ve Vitaminler", "7. Gün: Enzimler", "8. Gün: Enzimleri Etkileyen Faktörler",
        "9. Gün: Nükleik Asitler ve DNA", "10. Gün: RNA ve ATP Metabolizma",
        "11. Gün: Hücre ve Hücre Zarı", "12. Gün: Pasif Taşıma",
        "13. Gün: Aktif Taşıma, Endositoz ve Ekzositoz", "14. Gün: Ribozom ve Sentrozom",
        "15. Gün: Tek Zarlı Organeller", "16. Gün: Çift Zarlı Organeller",
        "17. Gün: Çekirdek ve Hücre İskeleti", "18. Gün: Bilimsel Yöntem ve Hücre Genel Tekrar",
        "19. Gün: Ekoloji İle İlgili Kavramlar", "20. Gün: Ekosistemi Oluşturan Bileşenler",
        "21. Gün: Besin Zinciri ve Besin Piramidi", "22. Gün: Madde Döngüleri ve Güncel Çevre Sorunları",
        "23. Gün: İkili Adlandırma", "24. Gün: Bakteriler Alemi", "25. Gün: Arkeler Alemi",
        "26. Gün: Protista, Bitki ve Mantar Alemi", "27. Gün: Omurgasız Hayvanlar",
        "28. Gün: Omurgalı Hayvanlar ve Virüsler", "29. Gün: Bölünme İle İlgili Kavramlar",
        "30. Gün: Mitoz Bölünme", "31. Gün: Mayoz ve Eşeysiz Üreme", "32. Gün: Eşeyli Üreme ve Bölünmeler Tekrar",
        "33. Gün: Kalıtım Kavramları ve Gamet Bulma", "34. Gün: Çaprazlamalar",
        "35. Gün: Eş Baskınlık ve Çok Alellik", "36. Gün: Kan Grupları",
        "37. Gün: Eşeye Bağlı Kalıtım", "38. Gün: Otozomal Kalıtım",
        "39. Gün: X ve Y'ye Bağlı Kalıtım, Soy Ağaçları ve Genel Tekrar (Kamp Bitti)"
    ],
    "Kenan Kara - 50 Günde TYT-AYT Geometri": [
        "1. Gün: Temel Kavramlar (1-2)", "2. Gün: Üçgende Açı (1-2-3)", "3. Gün: Ek Çizimler (1-2-3)",
        "4. Gün: Dik Üçgen (1-2)", "5. Gün: Dik Üçgen (3-4)", "6. Gün: İkizkenar Üçgen (1-2)",
        "7. Gün: Eşkenar Üçgen (1-2)", "8. Gün: Açıortay (1-2-3)", "9. Gün: Kenarortay (1-2)",
        "10. Gün: Üçgende Eşlik (1-2)", "11. Gün: Üçgende Benzerlik (1-2-3)", "12. Gün: Üçgende Benzerlik (4-5)",
        "13. Gün: Üçgende Alan (1-2)", "14. Gün: Üçgende Alan (3-4)", "15. Gün: Açı - Kenar Bağıntıları (1-2-3)",
        "16. Gün: Üçgende Merkezler", "17. Gün: Dörtgen (1-2)", "18. Gün: Paralelkenar (1-2)",
        "19. Gün: Paralelkenar (3-4)", "20. Gün: Eşkenar Dörtgen", "21. Gün: Deltoid",
        "22. Gün: Dikdörtgen (1-2-3)", "23. Gün: Kare (1-2-3)", "24. Gün: Yamuk (1-2-3)",
        "25. Gün: Yamuk (4-5)", "26. Gün: Çokgenler (1-2)", "27. Gün: Çokgenler (3-4)",
        "28. Gün: Çember ve Çemberde Açı (1-2)", "29. Gün: Çemberde Açı (3-4)",
        "30. Gün: Çemberde Uzunluk (1-2)", "31. Gün: Çemberde Uzunluk (3-4)", "32. Gün: Çemberin Çevresi",
        "33. Gün: Dairede Alan (1-2)", "34. Gün: Dairede Alan (3-4)", "35. Gün: Noktanın Analitiği (1-2)",
        "36. Gün: Noktanın Analitiği (3-4)", "37. Gün: Doğru Analitiği (1-2-3-4)", "38. Gün: Doğru Analitiği (5-6)",
        "39. Gün: Doğru Analitiği (7-8)", "40. Gün: Simetri (1-2)", "41. Gün: Öteleme ve Dönme-Döndürme",
        "42. Gün: Çember Analitiği (1-2)", "43. Gün: Temel Diklik Teoremi", "44. Gün: Dik Prizmalar (1-2)",
        "45. Gün: Dik Prizmalar (3-4)", "46. Gün: Küp (1-2-3)", "47. Gün: Silindir (1-2)",
        "48. Gün: Piramit (1-2)", "49. Gün: Koni (1-2)", "50. Gün: Küre (1-2) (Kamp Bitti)"
    ],
    "Sadettin Akyayla - 29 Günde TYT Tarih": [
        "1. Gün: Tarih ve Zaman", "2. Gün: İnsanlığın İlk Dönemleri",
        "3. Gün: İlk ve Orta Çağlarda Türk Dünyası - I", "4. Gün: İlk ve Orta Çağlarda Türk Dünyası - II",
        "5. Gün: İlk ve Orta Çağlarda Türk Dünyası - III", "6. Gün: İslam Medeniyetinin Doğuşu",
        "7. Gün: Türklerin İslamiyet'i Kabulü - I", "8. Gün: Türklerin İslamiyet'i Kabulü - II",
        "9. Gün: Orta Çağ'da Dünya", "10. Gün: Yerleşme ve Devletleşme Sürecinde Selçuklu Türkiyesi",
        "11. Gün: Beylikten Devlete Osmanlı Siyaseti", "12. Gün: Dünya Gücü Osmanlı",
        "13. Gün: Osmanlı Devleti Kültür ve Uygarlığı", "14. Gün: Yeni Çağ'da Avrupa",
        "15. Gün: Osmanlı Devleti Duraklama Dönemi", "16. Gün: Osmanlı Devleti Gerileme Dönemi",
        "17. Gün: Yakın Çağ'da Avrupa", "18. Gün: Osmanlı Devleti Dağılma Dönemi - I",
        "19. Gün: Osmanlı Devleti Dağılma Dönemi - II", "20. Gün: XX. Yüzyıl Başlarında Osmanlı - I",
        "21. Gün: XX. Yüzyıl Başlarında Osmanlı - II", "22. Gün: XX. Yüzyıl Başlarında Osmanlı - III",
        "23. Gün: İşgallerin Başlaması ve Millî Mücadele'ye Hazırlık",
        "24. Gün: BMM'nin Açılması ve Sevr Barış Antlaşması", "25. Gün: Millî Mücadele Muharebeler Dönemi",
        "26. Gün: Mudanya ve Lozan Barış Antlaşması", "27. Gün: Atatürk Dönemi İç Politika ve İlkeler",
        "28. Gün: Atatürkçülük ve Türk İnkılabı", "29. Gün: Atatürk Dönemi Türk Dış Politikası (Bitti)"
    ],
    "Zeduva - 39 Günde TYT Coğrafya": [
        "1. Gün: Doğa ve İnsan - Coğrafyanın Bölümleri", "2. Gün: Coğrafyanın Tarihi Gelişimi",
        "3. Gün: Atmosferin Katmanları ve Özellikleri", "4. Gün: Sıcaklık", "5. Gün: Basınç ve Rüzgarlar",
        "6. Gün: Nem ve Yağış", "7. Gün: Büyük İklim Tipleri", "8. Gün: Türkiye İklimi",
        "9. Gün: Dünya'nın Tektonik Oluşumu", "10. Gün: İç Kuvvetler", "11. Gün: Kayaçlar",
        "12. Gün: Türkiye'de İç Kuvvetler", "13. Gün: Dış Kuvvetler: Akarsular",
        "14. Gün: Dış Kuvvetler: Karstik Şekiller", "15. Gün: Rüzgar ve Buzul",
        "16. Gün: Dalga ve Akıntılar - Kıyı Tipleri", "17. Gün: Türkiye'de Dış Kuvvetler",
        "18. Gün: Türkiye'nin Yer Şekilleri", "19. Gün: Dünya'da Sular", "20. Gün: Türkiye'de Sular",
        "21. Gün: Dünya'da Topraklar", "22. Gün: Türkiye'de Topraklar", "23. Gün: Dünya'da Bitkiler",
        "24. Gün: Türkiye'de Bitkiler", "25. Gün: Dünya'da Yerleşmeler", "26. Gün: Türkiye'de Yerleşmeler",
        "27. Gün: Nüfusun Önemi ve Özellikleri", "28. Gün: Dünya'da Nüfusun Dağılışı",
        "29. Gün: Nüfus Piramitleri", "30. Gün: Türkiye'de Nüfus", "31. Gün: Dünya ve Türkiye'de Göçler",
        "32. Gün: Ekonomik Faaliyetler", "33. Gün: Bölge ve Bölge Türleri", "34. Gün: Ulaşım",
        "35. Gün: Çevre ve Toplum", "36. Gün: Doğal Afetler", "37. Gün: Dünya'nın Şekli ve Hareketleri",
        "38. Gün: Coğrafi Konum", "39. Gün: Harita Bilgisi (Kamp Bitti)"
    ],
    "Emek Hoca - 10 Günde Dil Bilgisi": [
        "1. Gün: Ses Bilgisi", "2. Gün: Yapı Bilgisi", "3. Gün: Sözcük Türleri - 1",
        "4. Gün: Sözcük Türleri - 2", "5. Gün: Sözcük Türleri - 3", "6. Gün: Fiiller ve Fiilde Yapı",
        "7. Gün: Ek Fiil ve Fiilimsiler", "8. Gün: Fiilde Çatı", "9. Gün: Cümlenin Ögeleri",
        "10. Gün: Cümle Türleri", "Yazım Kuralları Genel Tekrar", "Noktalama İşaretleri",
        "3 Saatte Tüm Dil Bilgisi Genel Tekrar (Kamp Bitti)"
    ]
}

# =====================================================================
# 3. DİNAMİK ŞABLON ANALİZ MOTORU (HER ÖĞRENCİNİN SAAT VE SATIR YAPISINA UYUMLU)
# =====================================================================
def analyze_template_structure(ws):
    """
    Öğrencinin şablonunda kaç satır olduğunu, başlıkların hangi satırda başladığını
    ve 'Başlangıç Zamanı:' hücresinin koordinatını dinamik olarak tespit eder.
    """
    first_non_empty = None
    last_non_empty = 1
    week_info_row = None
    header_row = None
    start_time_coord = (3, 6)  # Varsayılan F3

    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, 11)]
        if any(v is not None for v in row_vals):
            if first_non_empty is None:
                first_non_empty = r
            last_non_empty = r
        
        for c_idx, v in enumerate(row_vals, 1):
            if v is not None:
                s = tr_lower(v)
                if header_row is None and s in ['pzt', 'pazartesi']:
                    header_row = r
                if week_info_row is None and 'hafta' in s:
                    week_info_row = r
                if 'başlangıç zamanı' in s or 'baslangic zamani' in s:
                    start_time_coord = (r, c_idx + 1)

    stride = last_non_empty if (first_non_empty and first_non_empty > 1) else (last_non_empty + 1)
    return {
        "first_row": first_non_empty or 2,
        "last_row": last_non_empty,
        "week_info_row": week_info_row or 3,
        "header_row": header_row or 4,
        "stride": stride,
        "start_time_coord": start_time_coord
    }

def get_cell_color_key(cell):
    if not cell.fill or not cell.fill.fill_type:
        return None
    fg = cell.fill.fgColor
    if not fg:
        return None
    if fg.type == 'rgb' and fg.rgb and fg.rgb != '00000000':
        return f"rgb:{fg.rgb}"
    elif fg.type == 'theme':
        return f"theme:{fg.theme}:{round(fg.tint or 0.0, 2)}"
    return None

def is_distributable_subject(val_str):
    if not val_str:
        return False
    low = tr_lower(val_str)
    for kw in IGNORE_KEYWORDS:
        if kw in low:
            if low in ["tyt türkçe soru çözümü", "problem denemeleri"]:
                return True
            return False
    return True

def parse_template_subjects(ws, struct):
    """
    Şablondaki dersleri, hangi gün kaç etüt olduklarını (örn. Salı 3 etüt, Cumartesi 2 etüt)
    ve altlarındaki aynı renkli 'Eksik Kapatma' hücrelerini otomatik algılar.
    """
    subjects = {}
    header_row = struct["header_row"]
    last_row = struct["last_row"]

    for c in range(4, 11):  # Pzt (4) .. Paz (10)
        r = header_row + 1
        while r <= last_row:
            val = ws.cell(r, c).value
            val_str = str(val).strip() if val is not None else ""
            if is_distributable_subject(val_str):
                norm_name = normalize_subject_name(val_str)
                subj_color = get_cell_color_key(ws.cell(r, c))
                rows_in_group = [r]
                
                next_r = r + 1
                while next_r <= last_row:
                    next_val = ws.cell(next_r, c).value
                    next_str = str(next_val).strip() if next_val is not None else ""
                    next_color = get_cell_color_key(ws.cell(next_r, c))
                    
                    if is_distributable_subject(next_str) and normalize_subject_name(next_str) == norm_name:
                        rows_in_group.append(next_r)
                        next_r += 1
                    elif "eksik kapatma" in tr_lower(next_str) and subj_color and next_color == subj_color:
                        rows_in_group.append(next_r)
                        next_r += 1
                    else:
                        break
                subjects.setdefault(norm_name, []).append((c, rows_in_group))
                r = next_r
            else:
                r += 1
    return subjects

@st.cache_data(show_spinner=False)
def fetch_youtube_playlist(url):
    if not YTDLP_AVAILABLE:
        return []
    ydl_opts = {'extract_flat': True, 'quiet': True, 'no_warnings': True}
    videos = []
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)
        if 'entries' in info:
            for idx, entry in enumerate(info['entries'], 1):
                if entry and entry.get('title'):
                    title = entry.get('title')
                    clean_title = title.split('|')[0].strip() if '|' in title else title.strip()
                    duration_min = round((entry.get('duration') or 1800) / 60.0, 1)
                    videos.append({
                        "index": idx,
                        "title": clean_title,
                        "duration": max(duration_min, 5.0)
                    })
    return videos

def build_items_from_meb(curriculum_list, study_style, total_active_days):
    """
    Oynatma listesi kullanmayan (konuyu bilen ve soru çözecek olan veya kitaptan çalışacak)
    öğrenciler için MEB müfredatını konu ağırlıklarına göre alt görevlere böler.
    """
    items = []
    for topic, weight in curriculum_list:
        if study_style == "Sadece Soru Çözümü & Branş Denemesi":
            if weight >= 3:
                items.append({"title": f"{topic} - Soru Bankası Çözümü (Test 1-3)", "weight": weight * 0.5})
                items.append({"title": f"{topic} - ÖSYM Tarzı Soru Tarama", "weight": weight * 0.5})
            else:
                items.append({"title": f"{topic} - Soru Bankası Çözümü", "weight": float(weight)})
        else:
            # Kitaptan Konu Çalışması + Soru Çözümü
            items.append({"title": f"{topic} - Konu Çalışması ve Özet", "weight": weight * 0.5})
            items.append({"title": f"{topic} - Soru Bankası Çözümü", "weight": weight * 0.5})
    return items

def distribute_items_smart(items, active_day_groups, mode, fallback_label="Genel Tekrar ve Branş Denemesi"):
    """
    active_day_groups: [(week_idx, col_idx, [row1, row2, ...]), ...]
    Günlerin etüt sayısına (örn. Salı 3 etüt, Cumartesi 2 etüt) duyarlı ağırlıklı dağıtım yapar.
    """
    total_days = len(active_day_groups)
    if total_days == 0 or not items:
        return {}

    cell_assignments = {}
    total_slot_capacity = sum(len(rows) for (_, _, rows) in active_day_groups)
    total_weight = sum(it['weight'] for it in items)

    if "Fen / Kamp Modu" in mode:
        idx = 0
        n = len(items)
        rem_capacity = total_slot_capacity
        
        for d_idx, (w_idx, c_idx, rows) in enumerate(active_day_groups):
            rem_days = total_days - d_idx
            day_cap = len(rows)
            
            if idx >= n:
                for r_idx in rows:
                    cell_assignments[(w_idx, r_idx, c_idx)] = fallback_label
                continue

            # Bu günün etüt kapasitesine (2 ders veya 3 ders olmasına) göre hedef ağırlık
            rem_weight = sum(it['weight'] for it in items[idx:])
            target_day_w = rem_weight * (day_cap / max(1, rem_capacity))
            rem_capacity -= day_cap

            day_items = []
            curr_w = 0.0
            while idx < n:
                day_items.append(items[idx])
                curr_w += items[idx]['weight']
                idx += 1
                if (n - idx) < (rem_days - 1):
                    break
                if curr_w >= target_day_w and len(day_items) >= max(1, day_cap - 1):
                    break

            if day_cap == 1:
                cell_assignments[(w_idx, rows[0], c_idx)] = " + ".join(x['title'] for x in day_items)
            elif day_cap == 2:
                if len(day_items) == 1:
                    cell_assignments[(w_idx, rows[0], c_idx)] = day_items[0]['title']
                    cell_assignments[(w_idx, rows[1], c_idx)] = f"{day_items[0]['title']} Soru Çözümü"
                else:
                    mid = math.ceil(len(day_items) / 2)
                    cell_assignments[(w_idx, rows[0], c_idx)] = " + ".join(x['title'] for x in day_items[:mid])
                    cell_assignments[(w_idx, rows[1], c_idx)] = " + ".join(x['title'] for x in day_items[mid:])
            else:
                soru_items = [x for x in day_items if any(k in tr_lower(x['title']) for k in ['soru', 'test', 'efişinado', 'tekrar', 'deneme', 'tarama'])]
                konu_items = [x for x in day_items if x not in soru_items]
                if not konu_items:
                    konu_items = day_items[:-1] if len(day_items) > 1 else day_items
                    soru_items = day_items[-1:] if len(day_items) > 1 else []

                if len(konu_items) == 1:
                    cell_assignments[(w_idx, rows[0], c_idx)] = konu_items[0]['title']
                    cell_assignments[(w_idx, rows[1], c_idx)] = f"{konu_items[0]['title']} (Devam)"
                else:
                    mid = math.ceil(len(konu_items) / 2)
                    cell_assignments[(w_idx, rows[0], c_idx)] = " + ".join(x['title'] for x in konu_items[:mid])
                    cell_assignments[(w_idx, rows[1], c_idx)] = " + ".join(x['title'] for x in konu_items[mid:])

                soru_text = " + ".join(x['title'] for x in soru_items) if soru_items else f"{konu_items[-1]['title']} Soru Bankası Çözümü"
                for r_idx in rows[2:]:
                    cell_assignments[(w_idx, r_idx, c_idx)] = soru_text

    elif "Blok Tekrarı" in mode:
        idx = 0
        n = len(items)
        rem_capacity = total_slot_capacity
        for d_idx, (w_idx, c_idx, rows) in enumerate(active_day_groups):
            rem_days = total_days - d_idx
            day_cap = len(rows)
            if idx >= n:
                for r_idx in rows:
                    cell_assignments[(w_idx, r_idx, c_idx)] = fallback_label
                continue
            rem_weight = sum(it['weight'] for it in items[idx:])
            target_day_w = rem_weight * (day_cap / max(1, rem_capacity))
            rem_capacity -= day_cap

            day_items = []
            curr_w = 0.0
            while idx < n:
                day_items.append(items[idx])
                curr_w += items[idx]['weight']
                idx += 1
                if (n - idx) < (rem_days - 1):
                    break
                if curr_w >= target_day_w:
                    break
            combined = " + ".join(x['title'] for x in day_items)
            for r_idx in rows:
                cell_assignments[(w_idx, r_idx, c_idx)] = combined

    else:
        all_slots = [(w_idx, r_idx, c_idx) for (w_idx, c_idx, rows) in active_day_groups for r_idx in rows]
        total_slots = len(all_slots)
        n = len(items)
        if n <= total_slots:
            extra_slots = total_slots - n
            step = total_slots / (extra_slots + 1) if extra_slots > 0 else total_slots + 1
            extra_positions = set(int(round(step * i)) - 1 for i in range(1, extra_slots + 1))
            idx = 0
            last_title = "Ünite"
            for s_i, slot_key in enumerate(all_slots):
                if s_i in extra_positions or idx >= n:
                    cell_assignments[slot_key] = fallback_label if idx >= n else f"{last_title} Soru Bankası Çözümü"
                else:
                    cell_assignments[slot_key] = items[idx]['title']
                    last_title = items[idx]['title']
                    idx += 1
        else:
            target_w = total_weight / total_slots
            idx = 0
            for s_i, slot_key in enumerate(all_slots):
                rem_slots = total_slots - s_i
                slot_items = []
                curr_w = 0.0
                while idx < n:
                    slot_items.append(items[idx])
                    curr_w += items[idx]['weight']
                    idx += 1
                    if (n - idx) < (rem_slots - 1):
                        break
                    if curr_w >= target_w:
                        break
                cell_assignments[slot_key] = " + ".join(x['title'] for x in slot_items)

    return cell_assignments

def format_week_date_range(monday_date):
    sunday_date = monday_date + datetime.timedelta(days=6)
    m1 = TR_MONTHS[monday_date.month]
    m2 = TR_MONTHS[sunday_date.month]
    if monday_date.month == sunday_date.month:
        return f"({monday_date.day} - {sunday_date.day} {m1})"
    return f"({monday_date.day} {m1} - {sunday_date.day} {m2})"

def generate_multiweek_excel(template_bytes, num_weeks, start_date, holidays_dict, subject_configs):
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    ws_template = wb.active
    struct = analyze_template_structure(ws_template)
    
    last_row = struct["last_row"]
    header_row = struct["header_row"]
    week_info_row = struct["week_info_row"]
    stride = struct["stride"]
    st_row, st_col = struct["start_time_coord"]
    st_col_letter = openpyxl.utils.get_column_letter(st_col)

    ws_prog = wb.copy_worksheet(ws_template)
    ws_prog.title = "Program"
    
    if ws_prog.max_row > last_row:
        ws_prog.delete_rows(last_row + 1, ws_prog.max_row - last_row)
    ws_prog.tables.clear()
    ws_prog.row_breaks.brk.clear()

    start_monday = start_date - datetime.timedelta(days=start_date.weekday())
    template_subjects = parse_template_subjects(ws_template, struct)

    all_cell_assignments = {}
    custom_fills = {}

    for subj_name, day_groups in template_subjects.items():
        cfg = subject_configs.get(subj_name, {})
        if not cfg.get("enabled", False) or not cfg.get("items"):
            continue

        if cfg.get("use_custom_color") and cfg.get("color_hex"):
            hex_clean = cfg["color_hex"].lstrip('#').upper()
            custom_fills[subj_name] = PatternFill(start_color=f"FF{hex_clean}", end_color=f"FF{hex_clean}", fill_type="solid")

        active_groups = []
        for w in range(1, num_weeks + 1):
            week_monday = start_monday + datetime.timedelta(days=(w - 1) * 7)
            for (col_idx, rows) in day_groups:
                cell_date = week_monday + datetime.timedelta(days=(col_idx - 4))
                if cell_date < start_date or cell_date in holidays_dict:
                    continue
                active_groups.append((w, col_idx, rows))

        items = cfg["items"]
        start_idx = max(1, cfg.get("start_item", 1)) - 1
        end_idx = min(len(items), cfg.get("end_item", len(items)))
        selected_items = items[start_idx:end_idx]

        fallback_lbl = f"{subj_name} Branş Denemesi + Analizi"
        subj_assignments = distribute_items_smart(selected_items, active_groups, cfg.get("mode", ""), fallback_lbl)
        all_cell_assignments.update(subj_assignments)
        
        if subj_name in custom_fills:
            for (_, r_list) in day_groups:
                for r_i in r_list:
                    for c_i in range(4, 11):
                        if (c_i, r_list) in day_groups:
                            custom_fills[(r_i, c_i)] = custom_fills[subj_name]

    for w in range(1, num_weeks + 1):
        offset = (w - 1) * stride
        week_monday = start_monday + datetime.timedelta(days=(w - 1) * 7)
        
        for r in range(1, last_row + 1):
            target_r = offset + r
            orig_h = ws_template.row_dimensions[r].height
            ws_prog.row_dimensions[target_r].height = 40.15 if (orig_h is None or orig_h < 35) else orig_h
            
            for c in range(1, 11):
                src_cell = ws_template.cell(r, c)
                dst_cell = ws_prog.cell(target_r, c)
                
                val = src_cell.value
                if isinstance(val, str) and val.startswith('='):
                    # Şablondaki tüm satır referanslarını (B5, C12 vb.) o haftanın satırına kaydır
                    val = re.sub(
                        r'\b([A-J])(\d+)\b',
                        lambda m: f"{m.group(1)}{int(m.group(2)) + offset}",
                        val
                    )
                dst_cell.value = val

                if src_cell.has_style:
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format

                if r > header_row and c >= 4:
                    dst_cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                    dst_cell.font = Font(name='Verdana', size=10, bold=src_cell.font.bold if src_cell.font else False)

        # Hafta numarası ve tarih aralığını yaz
        ws_prog.cell(offset + week_info_row, 2).value = f"{w}. Hafta:"
        ws_prog.cell(offset + week_info_row, 3).value = format_week_date_range(week_monday)

        for c in range(4, 11):
            cell_date = week_monday + datetime.timedelta(days=(c - 4))
            for r in range(header_row + 1, last_row + 1):
                target_r = offset + r
                dst_cell = ws_prog.cell(target_r, c)
                
                if (r, c) in custom_fills:
                    dst_cell.fill = copy(custom_fills[(r, c)])

                if (w, r, c) in all_cell_assignments:
                    dst_cell.value = all_cell_assignments[(w, r, c)]
                else:
                    orig_val = str(ws_template.cell(r, c).value or "").strip()
                    if is_distributable_subject(orig_val) or "eksik kapatma" in tr_lower(orig_val):
                        if cell_date < start_date:
                            dst_cell.value = f"Program Başlangıcı: {start_date.day} {TR_MONTHS[start_date.month]}"
                        elif cell_date in holidays_dict:
                            dst_cell.value = holidays_dict[cell_date]

        tab_ref = f"B{offset + header_row}:J{offset + last_row}"
        tbl = Table(displayName=f"ProgramTablo{w}", ref=tab_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="Program", showFirstColumn=True, showLastColumn=False,
            showRowStripes=True, showColumnStripes=True
        )
        ws_prog.add_table(tbl)

        # Haftalar arası boş satır yüksekliği ve Sayfa Sonu (RowBreak)
        if w < num_weeks:
            sep_row = offset + stride + (1 if struct["first_row"] > 1 else 0)
            ws_prog.row_dimensions[sep_row].height = 40.15
            ws_prog.row_breaks.append(Break(id=sep_row, man=True))

    col_widths = {'A': 3.23, 'B': 15.70, 'C': 13.0, 'D': 32.0, 'E': 32.0, 'F': 32.0, 'G': 32.0, 'H': 32.0, 'I': 32.0, 'J': 32.0}
    for col_letter, width in col_widths.items():
        ws_prog.column_dimensions[col_letter].width = width

    ws_prog.page_setup.orientation = ws_prog.ORIENTATION_LANDSCAPE
    ws_prog.page_setup.paperSize = ws_prog.PAPERSIZE_A4
    ws_prog.sheet_properties.pageSetUpPr.fitToPage = True
    ws_prog.page_setup.fitToWidth = 1
    ws_prog.page_setup.fitToHeight = 0

    wb.defined_names['BaşlangıçSaati'] = DefinedName(name='BaşlangıçSaati', attr_text=f"Program!${st_col_letter}${st_row}")

    if len(wb.sheetnames) > 1 and 'Genel Şablon' in wb.sheetnames:
        wb.remove(wb['Genel Şablon'])

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =====================================================================
# STREAMLIT KULLANICI ARAYÜZÜ
# =====================================================================
st.title("🎓 Akıllı YKS Çalışma Programı Mimarı v2.0")
st.markdown(
    "Hangi öğrencinin şablonunu yüklerseniz yükleyin (Pansiyonlu, Sabah 06:00 etütlü, Dershaneli vb.) "
    "**saatleri, satır sayısını, renkleri ve sayfa aralıklarını otomatik tanır.** "
    "İster YouTube kampı, ister **oynatma listesiz MEB Konu / Soru Çözümü** dağıtın."
)

with st.sidebar:
    st.header("1. Öğrenci Şablonu ve Takvim")
    uploaded_template = st.file_uploader("Öğrencinin Şablon Excel'i (.xlsx)", type=["xlsx"])
    
    c_d1, c_d2 = st.columns(2)
    with c_d1:
        start_date = st.date_input("Başlangıç Tarihi", value=datetime.date(2026, 10, 19))
    with c_d2:
        num_weeks = st.number_input("Kaç Haftalık?", min_value=1, max_value=52, value=12)

    st.subheader("🎉 Tatil Ayarları")
    use_official_holidays = st.checkbox("Resmi Tatilleri Otomatik Atla (29 Ekim, 1 Ocak vb.)", value=True)
    custom_holidays_input = st.date_input("Ekstra Tatil Günleri (İsteğe Bağlı)", value=[])

holidays_dict = {}
if use_official_holidays:
    for d_off in range(num_weeks * 7 + 14):
        d = start_date - datetime.timedelta(days=7) + datetime.timedelta(days=d_off)
        if (d.month, d.day) in FIXED_HOLIDAYS:
            holidays_dict[d] = FIXED_HOLIDAYS[(d.month, d.day)]

if isinstance(custom_holidays_input, (list, tuple)):
    for cd in custom_holidays_input:
        holidays_dict[cd] = "Özel İzin / Tatil Günü"

if uploaded_template is not None:
    template_bytes = uploaded_template.read()
    wb_temp = openpyxl.load_workbook(BytesIO(template_bytes))
    ws_temp = wb_temp.active
    struct = analyze_template_structure(ws_temp)
    detected_subjects = parse_template_subjects(ws_temp, struct)

    st.success(
        f"✅ Şablon Analizi Tamamlandı (Toplam {struct['last_row']} satır/hafta) | "
        f"Tespit Edilen Dersler: **{', '.join(f'{k} ({len(v)} gün/hafta)' for k, v in detected_subjects.items())}**"
    )

    st.header("2. Ders Kaynakları (Kamp / Soru Bankası / Müfredat)")
    tabs = st.tabs(list(detected_subjects.keys()))
    subject_configs = {}

    for tab, (subj_name, day_groups) in zip(tabs, detected_subjects.items()):
        with tab:
            low_s = tr_lower(subj_name)
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                enabled = st.checkbox(f"'{subj_name}' dersini programa dağıt", value=True, key=f"en_{subj_name}")
            with col2:
                def_mode_idx = 0 if any(x in low_s for x in ["fizik", "kimya", "biyoloji"]) else (1 if "matematik" in low_s else 2)
                mode = st.selectbox(
                    "Etüt Yerleşim Modu",
                    [
                        "Fen / Kamp Modu (İlk Etütler Konu + Son Etüt Soru Çözümü)",
                        "Blok Tekrarı Modu (Aynı Gün Alt Alta Etütlerde Aynı Başlık)",
                        "Ardışık Mod (Her Etüde Sırayla Görev + Araya Soru Çözümü)"
                    ],
                    index=def_mode_idx,
                    key=f"mode_{subj_name}"
                )
            with col3:
                def_color = next((v for k, v in DEFAULT_COLORS.items() if k in low_s), "#53D2FF")
                use_custom_color = st.checkbox("Rengi Değiştir", value=False, key=f"usecol_{subj_name}")
                color_hex = st.color_picker("Renk", value=def_color, key=f"col_{subj_name}", disabled=not use_custom_color)

            source_type = st.radio(
                "Öğrenci Bu Dersi Nasıl Çalışacak?",
                [
                    "🔥 Hazır Hoca Kampından Seç (Özcan Aykın, Görkem Şahin, Biosem, Kenan Kara vb.)",
                    "📝 Konuyu Biliyor -> Sadece Soru Çözümü & Branş Denemesi Dağıt (MEB Müfredatı)",
                    "📖 Kitaptan / Fasikülden Konu + Soru Dağıt (Oynatma Listesiz MEB Müfredatı)",
                    "🔗 YouTube Oynatma Listesi Linki Yapıştır (Otomatik Çek)",
                    "✍️ Özel Konu / Video Listesi Yapıştır"
                ],
                key=f"srctype_{subj_name}"
            )

            items = []
            if "Hazır Hoca Kampından" in source_type:
                # Ders adına uygun kampı varsayılan seç
                camp_names = list(BUILTIN_CAMPS.keys())
                def_camp_idx = next((i for i, cn in enumerate(camp_names) if any(w in tr_lower(cn) for w in low_s.split() if len(w) > 3)), 0)
                selected_camp = st.selectbox("Hazır Kamp Seçin:", camp_names, index=def_camp_idx, key=f"camp_{subj_name}")
                for line in BUILTIN_CAMPS[selected_camp]:
                    items.append({"title": line, "weight": 1.0})

            elif "MEB Müfredatı" in source_type:
                meb_keys = list(MEB_CURRICULUM.keys())
                def_meb_idx = next((i for i, mk in enumerate(meb_keys) if tr_lower(mk) in low_s or any(w in tr_lower(mk) for w in low_s.split() if len(w) > 3)), 0)
                selected_meb = st.selectbox("Müfredat Seçin:", meb_keys, index=def_meb_idx, key=f"meb_{subj_name}")
                style = "Sadece Soru Çözümü & Branş Denemesi" if "Sadece Soru" in source_type else "Kitaptan Konu + Soru"
                items = build_items_from_meb(MEB_CURRICULUM[selected_meb], style, len(day_groups) * num_weeks)

            elif "YouTube" in source_type:
                yt_url = st.text_input("YouTube Oynatma Listesi URL'si:", key=f"yt_{subj_name}")
                if yt_url.strip() and YTDLP_AVAILABLE:
                    with st.spinner("Videolar ve süre ağırlıkları çekiliyor..."):
                        fetched = fetch_youtube_playlist(yt_url.strip())
                        for v in fetched:
                            items.append({"title": f"{v['index']}. {v['title']}", "weight": v['duration']})

            else:
                raw_text = st.text_area("Her satıra 1 başlık yazın (Ağırlık için satır sonuna '| 2' ekleyebilirsiniz):", height=150, key=f"txt_{subj_name}")
                if raw_text.strip():
                    for line in raw_text.strip().split('\n'):
                        if line.strip():
                            if '|' in line and line.rsplit('|', 1)[1].strip().replace('.', '').isdigit():
                                t_p, w_p = line.rsplit('|', 1)
                                items.append({"title": t_p.strip(), "weight": float(w_p.strip())})
                            else:
                                items.append({"title": line.strip(), "weight": 1.0})

            start_item, end_item = 1, max(1, len(items))
            if items:
                sc1, sc2 = st.columns(2)
                with sc1:
                    start_item = st.number_input("Başlangıç Sırası (Kaldığı Konu/Video):", min_value=1, max_value=len(items), value=1, key=f"st_{subj_name}")
                    st.caption(f"▶️ Başlangıç: **{items[start_item - 1]['title']}**")
                with sc2:
                    end_item = st.number_input("Bitiş Sırası (Hedeflenen Konu/Video):", min_value=1, max_value=len(items), value=len(items), key=f"end_{subj_name}")
                    st.caption(f"⏹️ Bitiş: **{items[end_item - 1]['title']}**")

            subject_configs[subj_name] = {
                "enabled": enabled, "mode": mode, "use_custom_color": use_custom_color,
                "color_hex": color_hex, "items": items, "start_item": start_item, "end_item": end_item
            }

    st.divider()
    if st.button("🚀 Öğrenciye Özel Çok Haftalık Excel Programını Üret", type="primary", use_container_width=True):
        with st.spinner("Öğrencinin saat şablonu, tatiller ve konu ağırlıkları işleniyor..."):
            excel_data = generate_multiweek_excel(template_bytes, num_weeks, start_date, holidays_dict, subject_configs)
        st.success("🎉 Program hazır! Öğrencinin özel saatleri, renkleri ve sayfa sonları birebir korundu.")
        st.download_button(
            label="📥 Excel Dosyasını İndir (.xlsx)",
            data=excel_data,
            file_name=f"Calisma_Programi_{num_weeks}_Hafta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
