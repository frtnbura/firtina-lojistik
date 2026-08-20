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
GITHUB_TOKEN = "ghp_y9JPiEEPtvozMKPfmNcaV41WE7BBnL4RTwtw"
GITHUB_REPO = "frtnbura/firtina-lojistik"
WEBAPP_URL = "https://frtnbura.github.io/firtina-lojistik/index.html"
# ============================================

web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Fırtına Lojistik 7/24 Bulut Sistemi Aktif!"

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    web_app.run(host='0.0.0.0', port=port)

def github_excel_guncelle(dosya_adi, basliklar, yeni_satir):
    """GitHub API üzerinden Excel dosyasını kontrol eder, yoksa oluşturur, varsa yeni satırı ekler."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{dosya_adi}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }
    
    # 1. Dosya var mı kontrol et
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

    # Yeni satırı ekle
    ws.append(yeni_satir)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    # Base64 formatına çevir ve GitHub'a yükle
    encoded_content = base64.b64encode(out.read()).decode("utf-8")
    payload = {
        "message": f"Mobil Veri Kaydı: {yeni_satir[0]}",
        "content": encoded_content,
        "branch": "main"
    }
    if sha:
        payload["sha"] = sha
        
    put_response = requests.put(url, headers=headers, json=payload)
    if put_response.status_code not in [200, 201]:
        raise Exception(f"GitHub API Hatası ({put_response.status_code}): {put_response.text}")

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
            satir = [
                veri.get("tarih"), veri.get("plaka"), veri.get("sofor"), 
                veri.get("tonaj"), veri.get("fiyat"), veri.get("yakit"), 
                veri.get("yukleme"), veri.get("harcirah"), veri.get("net")
            ]
            github_excel_guncelle("nakliye_kamyon_hesap_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"✅ *Sefer GitHub'a Yazıldı!*\n📅 `{veri.get('tarih')}` | 🚛 `{veri.get('plaka')}`\n⚖️ *{veri.get('tonaj')} Ton* | 💵 *Net: {veri.get('net')} TL*", parse_mode="Markdown")

        elif modul == "dikili_sefer":
            basliklar = ["Tarih", "Bölge", "İstif", "Tür", "Tonaj", "Fiyat", "Tutar", "Plaka"]
            satir = [
                veri.get("tarih"), veri.get("bolge"), veri.get("istif"), 
                veri.get("tur"), veri.get("tonaj"), veri.get("fiyat"), 
                veri.get("tutar"), veri.get("plaka")
            ]
            github_excel_guncelle("dikili_sefer_hesap_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"🌲 *Dikili Seferi Kaydedildi!*\n📍 `{veri.get('bolge')}` | 🪵 `{veri.get('tur')}`\n💰 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        elif modul == "dikili_veresiye":
            basliklar = ["Tarih", "Müşteri", "İşlem Türü", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("musteri"), veri.get("islem_turu"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle("dikili_veresiye_takip.xlsx", basliklar, satir)
            await update.message.reply_text(f"📑 *Veresiye/Tahsilat Kaydedildi!*\n👤 `{veri.get('musteri')}` | 💵 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        elif modul == "masraf":
            basliklar = ["Tarih", "Kategori", "Tutar", "Açıklama"]
            satir = [veri.get("tarih"), veri.get("kategori"), veri.get("tutar"), veri.get("aciklama")]
            github_excel_guncelle("genel_masraflar.xlsx", basliklar, satir)
            await update.message.reply_text(f"🛠️ *Masraf Kaydedildi!*\n🏷️ `{veri.get('kategori')}` | 💸 *{veri.get('tutar')} TL*", parse_mode="Markdown")

        else:
            await update.message.reply_text(f"⚠️ Bilinmeyen modül türü: `{modul}`")

    except Exception as e:
        await update.message.reply_text(f"❌ *Kayıt işlenirken hata oluştu:*\n`{str(e)}`", parse_mode="Markdown")

if __name__ == "__main__":
    threading.Thread(target=run_flask, daemon=True).start()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, veri_yakala))
    print("Fırtına Lojistik Telegram Köprüsü Başlatıldı...")
    app.run_polling()
