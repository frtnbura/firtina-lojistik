import os
import json
import requests
import threading
from flask import Flask
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, WebAppInfo
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ================= AYARLAR =================
TELEGRAM_TOKEN = "8047383930:AAH3O_2GC9x-iERPfr7FqdiX8zzwMhSnqVA"
# Yukarıda kopyaladığınız Google Apps Script URL'sini buraya yapıştırın:
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbxZ7-UaIdc9iMzW6N4svrHKXol7E7R37RicIhLyKOaPs2KNIOlOJkZ31v1eHLxmtITy5A/exec"
WEBAPP_URL = "https://frtnbura.github.io/firtina-lojistik/index.html"
# ============================================

web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Fırtına Lojistik 7/24 Bulut Sistemi Aktif!"

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    web_app.run(host='0.0.0.0', port=port)

def google_tabloya_kaydet(modul, basliklar, satir):
    payload = {
        "modul": modul,
        "basliklar": basliklar,
        "satir": satir
    }
    r = requests.post(GOOGLE_SHEET_URL, json=payload)
    if r.status_code != 200:
        raise Exception(f"Google Bağlantı Hatası: {r.text}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [[KeyboardButton(text="⚡ FIRTINA LOJİSTİK YÖNETİM", web_app=WebAppInfo(url=WEBAPP_URL))]]
    await update.message.reply_text(
        "🚛 *Fırtına Lojistik 7/24 Bulut Sistemi Aktif!*\n\nAşağıdaki butona tıklayarak sefer, masraf, yakıt ve veresiye kayıtlarınızı girebilirsiniz.",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True),
        parse_mode="Markdown"
    )

async def veri_yakala(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        raw_data = update.effective_message.web_app_data.data
        veri = json.loads(raw_data)
        modul = veri.get("modul")
        
        if modul == "nakliye_sefer":
            basliklar = ["Tarih", "Plaka", "Şoför", "Tonaj", "Birim Fiyat", "Yakıt", "Yükleme", "Harcırah", "Net Kar"]
            satir = [veri.get("tarih"), veri.get("plaka"), veri.get("sofor"), veri.get("tonaj"), veri.get("fiyat"), veri.get("yakit"), veri.get("yukleme"), veri.get("harcirah"), veri.get("net")]
            google_tabloya_kaydet("Nakliye Seferleri", basliklar, satir)
            await update.message.reply_text(f"✅ *Sefer Tabloya Kaydedildi!*\n📅 `{veri.get('tarih')}` | 🚛 `{veri.get('plaka')}`\n⚖️ *{veri.get('tonaj')} Ton* | 💵 *Net: {veri.get('net')} TL*", parse_mode="Markdown")

        elif modul == "dikili_sefer":
            basliklar = ["Tarih", "Bölge", "İstif", "Tür", "Tonaj", "Fiyat", "Tutar", "Plaka"]
            satir = [veri.get("tarih"), veri.get("bolge"), veri.get("istif"), veri.get("tur"), veri.get("tonaj"), veri.get("fiyat"), veri.get("tutar"), veri.get("plaka")]
            google_tabloya_kaydet("Dikili Seferleri", basliklar, satir)
            await update.message.reply_text(f"🌲 *Dikili Seferi Kaydedildi!*\n📍 `{veri.get('bolge')}` | 🪵 `{veri.get('tur')}`\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        elif modul == "dikili_veresiye":
            basliklar = ["Tarih", "Müşteri", "İşlem Türü", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("musteri"), veri.get("islem_turu"), veri.get("tutar"), veri.get("aciklama")]
            google_tabloya_kaydet("Dikili Veresiye", basliklar, satir)
            await update.message.reply_text(f"📑 *Veresiye/Tahsilat Kaydedildi!*\n👤 `{veri.get('musteri')}` | 💵 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        elif modul == "masraf":
            basliklar = ["Tarih", "Kategori", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("kategori"), veri.get("tutar"), veri.get("aciklama")]
            google_tabloya_kaydet("Genel Masraflar", basliklar, satir)
            await update.message.reply_text(f"🛠️ *Masraf Kaydedildi!*\n🏷️ `{veri.get('kategori')}` | 💸 *{veri.get('tutar')} TL*", parse_mode="Markdown")

    except Exception as e:
        await update.message.reply_text(f"❌ *Kayıt işlenirken hata oluştu:*\n`{str(e)}`", parse_mode="Markdown")

if __name__ == "__main__":
    threading.Thread(target=run_flask, daemon=True).start()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, veri_yakala))
    print("Fırtına Lojistik Telegram Köprüsü Başlatıldı...")
    app.run_polling()
