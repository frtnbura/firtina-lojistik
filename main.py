import os
import json
import base64
import requests
import threading
from io import BytesIO
import openpyxl
from openpyxl import load_workbook
from flask import Flask
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, WebAppInfo
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ================= AYARLAR =================
TELEGRAM_TOKEN = "8047383930:AAH3O_2GC9x-iERPfr7FqdiX8zzwMhSnqVA"
GITHUB_TOKEN = "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"  # GitHub'dan aldığınız ghp_ ile başlayan tokenı buraya yazın
GITHUB_REPO = "frtnbura/firtina-lojistik"
WEBAPP_URL = "https://frtnbura.github.io/firtina-lojistik/index.html"
# ============================================

# Render'ın port beklemesini sağlayan arka plan web sunucusu
web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Fırtına Lojistik 7/24 Bulut Sistemi Aktif!"

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    web_app.run(host='0.0.0.0', port=port)

def github_excel_guncelle(dosya_adi, basliklar, yeni_satir):
    """GitHub üzerindeki Excel dosyasını günceller veya yoksa sıfırdan oluşturup satırı ekler."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{dosya_adi}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    r = requests.get(url, headers=headers)
    wb = None
    sha = None

    if r.status_code == 200:
        try:
            file_data = r.json()
            sha = file_data.get("sha")
            raw_content = file_data.get("content", "")
            if raw_content:
                content = base64.b64decode(raw_content)
                wb = load_workbook(BytesIO(content))
                ws = wb.active
        except Exception:
            wb = None

    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Kayitlar"
        ws.append(basliklar)

    ws.append(yeni_satir)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    encoded_content = base64.b64encode(out.read()).decode("utf-8")
    payload = {
        "message": f"Mobil Veri Kaydı: {yeni_satir[0]}",
        "content": encoded_content
    }
    if sha:
        payload["sha"] = sha
        
    requests.put(url, headers=headers, json=payload)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [[KeyboardButton(text="⚡ FIRTINA LOJİSTİK YÖNETİM", web_app=WebAppInfo(url=WEBAPP_URL))]]
    await update.message.reply_text(
        "🚛 *Fırtına Lojistik 7/24 Bulut Sistemi Aktif!*\n\nAşağıdaki butona tıklayarak sefer, masraf, yakıt, yevmiye ve veresiye kayıtlarınızı girebilirsiniz.",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True),
        parse_mode="Markdown"
    )

async def veri_yakala(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        veri = json.loads(update.effective_message.web_app_data.data)
        modul = veri.get("modul")
        
        # 1. NAKLİYE SEFERİ
        if modul == "nakliye_sefer":
            basliklar = ["Tarih", "Plaka", "Şoför", "Tonaj", "Birim Fiyat", "Yakıt", "Yükleme", "Harcırah", "Net Kar"]
            satir = [
                veri.get("tarih"), veri.get("plaka"), veri.get("sofor"), 
                veri.get("tonaj"), veri.get("fiyat"), veri.get("yakit"), 
                veri.get("yukleme"), veri.get("harcirah"), veri.get("net")
            ]
            github_excel_guncelle("nakliye_kamyon_hesap_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"✅ *Sefer Kaydedildi!*\n📅 `{veri.get('tarih')}` | 🚛 `{veri.get('plaka')}`\n⚖️ *{veri.get('tonaj')} Ton* | 💵 *Net: {veri.get('net')} TL*", parse_mode="Markdown")

        # 2. DİKİLİ SEFERİ
        elif modul == "dikili_sefer":
            basliklar = ["Tarih", "Bölge", "İstif", "Tür", "Tonaj", "Fiyat", "Tutar", "Plaka"]
            satir = [
                veri.get("tarih"), veri.get("bolge"), veri.get("istif"), 
                veri.get("tur"), veri.get("tonaj"), veri.get("fiyat"), 
                veri.get("tutar"), veri.get("plaka")
            ]
            github_excel_guncelle("dikili_sefer_hesap_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"🌲 *Dikili Seferi Kaydedildi!*\n📍 `{veri.get('bolge')}` | 🪵 `{veri.get('tur')}`\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        # 3. DİKİLİ VERESİYE & TAHSİLAT
        elif modul == "dikili_veresiye":
            basliklar = ["Tarih", "Müşteri", "İşlem Türü", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("musteri"), veri.get("islem_turu"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle("dikili_veresiye_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"📑 *Veresiye/Tahsilat Kaydedildi!*\n👤 `{veri.get('musteri')}` | 💵 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        # 4. ELEMAN YEVMİYE & AVANS (Kategoriye Göre Dosyaya Ayrılır)
        elif modul == "yevmiye":
            kategori = veri.get("kategori", "diger").lower()
            dosya_haritasi = {
                "yuklemeci": "yevmiye_yuklemeciler.xlsx",
                "sofor": "yevmiye_soforler.xlsx",
                "kesimci": "yevmiye_kesimciler.xlsx",
                "diger": "yevmiye_diger.xlsx"
            }
            dosya_adi = dosya_haritasi.get(kategori, "yevmiye_diger.xlsx")
            basliklar = ["Tarih", "İsim", "Kategori", "İşlem Türü", "Tutar / Yevmiye", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("isim"), veri.get("kategori"), veri.get("islem_turu"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle(dosya_adi, basliklar, satir)
            await update.message.reply_text(f"👷 *Yevmiye/Avans Kaydedildi!*\n👤 `{veri.get('isim')}` ({veri.get('kategori')})\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        # 5. YAKIT KAYDI
        elif modul == "yakit":
            basliklar = ["Tarih", "Plaka", "Litre", "Birim Fiyat", "Toplam Tutar", "Kilometre", "İstasyon"]
            satir = [
                veri.get("tarih"), veri.get("plaka"), veri.get("litre"), 
                veri.get("birim_fiyat"), veri.get("tutar"), veri.get("km"), veri.get("istasyon")
            ]
            github_excel_guncelle("yakit_kayitlari.xlsx", basliklar, satir)
            await update.message.reply_text(f"⛽ *Yakıt Kaydedildi!*\n🚛 `{veri.get('plaka')}` | ⛽ `{veri.get('litre')} Lt`\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        # 6. SABİT GİDERLER & MAAŞLAR
        elif modul == "sabit_gider":
            basliklar = ["Tarih", "Gider Türü", "İlgili Kişi / Plaka", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("gider_turu"), veri.get("ilgili"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle("sabit_giderler_maaslar.xlsx", basliklar, satir)
            await update.message.reply_text(f"🏢 *Sabit Gider Kaydedildi!*\n📋 `{veri.get('gider_turu')}` | 💵 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        # 7. ARIZA & SANAYİ & BAKIM
        elif modul == "ariza_bakim":
            basliklar = ["Tarih", "Plaka", "Yapılan İşlem / Parça", "Usta / Servis", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("plaka"), veri.get("islem"), veri.get("servis"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle("ariza_sanayi_bakim.xlsx", basliklar, satir)
            await update.message.reply_text(f"🔧 *Arıza/Bakım Masrafı Kaydedildi!*\n🚛 `{veri.get('plaka')}` | 🛠️ `{veri.get('islem')}`\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

    except Exception as e:
        await update.message.reply_text(f"❌ *Kayıt işlenirken hata oluştu:* `{str(e)}`", parse_mode="Markdown")

if __name__ == "__main__":
    threading.Thread(target=run_flask, daemon=True).start()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, veri_yakala))
    print("Fırtına Lojistik Telegram Köprüsü Başlatıldı...")
    app.run_polling()
