import telebot
from telebot import types
import openpyxl
import json
import os
import re
import logging
import threading
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.pdfgen import canvas

# ---------------------------------------------------------------------------
# LOGGING — konsolga va bot.log fayliga yoziladi (xatolarni topish osonlashadi)
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("nortix_bot")

# BotFather'dan olgan tokeningiz (Railway'ning Variables bo'limidan olinadi)
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    log.critical("BOT_TOKEN topilmadi! Railway -> Variables bo'limiga BOT_TOKEN qo'shing.")
    raise SystemExit("BOT_TOKEN environment o'zgaruvchisi berilmagan.")

# Baza.xlsx faylining nomi
EXCEL_FILE = "Baza.xlsx"

bot = telebot.TeleBot(BOT_TOKEN)

# Excel fayl bilan ishlashda bir vaqtda faqat bitta oqim (thread) yozishi uchun qulf.
# Bot ko'p oqimli (threaded) rejimda ishlagani sabab, bu qulf bo'lmasa ombor sonini
# bir vaqtda ikki kishi o'zgartirsa, ma'lumot noto'g'ri yozilib qolishi mumkin.
EXCEL_LOCK = threading.Lock()

# Admin ID(lar). Bir nechta adminni vergul bilan ADMIN_IDS Railway o'zgaruvchisiga yozish mumkin
# masalan: ADMIN_IDS=199728470,123456789
ADMIN_ID = 199728470
_qoshimcha_adminlar = os.getenv("ADMIN_IDS", "")
ADMIN_IDLAR = {ADMIN_ID}
for _id in _qoshimcha_adminlar.split(","):
    _id = _id.strip()
    if _id.isdigit():
        ADMIN_IDLAR.add(int(_id))


def admin_mi(user_id):
    """Foydalanuvchi admin/menejerlardan biri ekanini tekshiradi."""
    try:
        return int(user_id) in ADMIN_IDLAR
    except (TypeError, ValueError):
        return False

# "ZAKAZ NORTIX" guruhi ID raqami (yangi buyurtmalar shu yerda tasdiqlanadi/bekor qilinadi)
# Guruhga botni admin qilib qo'shing, guruh ichida /groupid deb yozing — bot sizga ID'ni yuboradi.
# Keyin shu qatordagi 0 o'rniga o'sha ID'ni (masalan -1001234567890) yozing.
ZAKAZ_GRUPPA_ID = -5166542981

# "Nortix sklad" guruhi ID raqami (tasdiqlangan buyurtma PDF shu yerga yuboriladi, sklad shu yerda ishlaydi)
# Xuddi yuqoridagidek, botni shu guruhga ham admin qilib qo'shing va /groupid orqali ID oling.
NORTIX_SKLAD_GRUPPA_ID = -5345356975

# Supergroup ID'lari odatda -100 bilan boshlanadi va 13 xonali bo'ladi (masalan -1001234567890).
# Agar shu formatga mos kelmasa, ehtimol /groupid orqali olingan ID noto'g'ri kiritilgan.
for _nomi, _qiymati in (("ZAKAZ_GRUPPA_ID", ZAKAZ_GRUPPA_ID), ("NORTIX_SKLAD_GRUPPA_ID", NORTIX_SKLAD_GRUPPA_ID)):
    if not str(_qiymati).startswith("-100"):
        logging.getLogger("nortix_bot").warning(
            "%s = %s standart supergroup ID formatiga (-100...) mos kelmayapti — "
            "guruhga xabar yuborilmasligi mumkin, /groupid bilan qayta tekshiring.",
            _nomi, _qiymati,
        )

# Holat xotiralari
yangilash_holati = {}
rasm_kutilayotganlar = {}
aksiya_kutilayotganlar = {}
buyurtma_holati = {}
qidiruv_kutayotganlar = set()
savat = {}
kutilayotgan_buyurtmalar = {}
guruh_buyurtmalari = {}
buyurtma_id_hisoblagich = {"son": 0}


def _excel_atomik_saqlash(workbook):
    """Excel faylni avval vaqtinchalik faylga saqlab, keyin asl faylga almashtiradi.
    Shu tarzda, saqlash paytida bot yiqilib qolsa ham, Baza.xlsx buzilib qolmaydi."""
    vaqtinchalik = EXCEL_FILE + ".tmp"
    workbook.save(vaqtinchalik)
    os.replace(vaqtinchalik, EXCEL_FILE)


def ombor_malumotlarini_oqish():
    with EXCEL_LOCK:
        workbook = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        sheet = workbook.active

        malumotlar = {}
        for row in sheet.iter_rows(min_row=1, values_only=True):
            # Ustunlar tartibi: A=Brendlar, B=Kategoriya, C=Model, D=SONI, E=Narx, F=Rasm, G=Aksiya
            brend = row[0]
            kategoriya, model, soni = row[1], row[2], row[3]
            narxi = row[4] if len(row) > 4 else None
            rasm = row[5] if len(row) > 5 else None
            aksiya = row[6] if len(row) > 6 else None

            if kategoriya is None or model is None:
                continue
            kategoriya = str(kategoriya).strip()
            model = str(model).strip()
            brend = str(brend).strip() if brend is not None else ""
            aksiya = str(aksiya).strip() if aksiya is not None else ""

            if kategoriya.lower() == "kategoriya" or model.lower() == "model":
                continue

            soni = soni if soni is not None else 0
            # Narx validatsiyasi: Excel'da matn/bo'sh kiritilgan bo'lsa, dastur qulamasin.
            try:
                narxi = float(narxi) if narxi is not None else 0
            except (TypeError, ValueError):
                log.warning("Narx noto'g'ri formatda: kategoriya=%s model=%s narx=%r — 0 sifatida olindi.", kategoriya, model, narxi)
                narxi = 0

            if kategoriya not in malumotlar:
                malumotlar[kategoriya] = []
            malumotlar[kategoriya].append((model, soni, narxi, rasm, brend, aksiya))

        return malumotlar


def ombor_sonini_yangilash(kategoriya, model, yangi_soni):
    with EXCEL_LOCK:
        try:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
        except Exception:
            log.exception("Excel faylni ochishda xatolik (ombor_sonini_yangilash)")
            return False
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=1):
            row_kategoriya = row[1].value
            row_model = row[2].value
            if row_kategoriya is None or row_model is None:
                continue
            if str(row_kategoriya).strip() == kategoriya and str(row_model).strip() == model:
                row[3].value = yangi_soni
                _excel_atomik_saqlash(workbook)
                return True

        return False


def ombor_rasmini_yangilash(kategoriya, model, rasm_manzili):
    with EXCEL_LOCK:
        try:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
        except Exception:
            log.exception("Excel faylni ochishda xatolik (ombor_rasmini_yangilash)")
            return False
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=1):
            row_kategoriya = row[1].value
            row_model = row[2].value
            if row_kategoriya is None or row_model is None:
                continue
            if str(row_kategoriya).strip() == kategoriya and str(row_model).strip() == model:
                # sheet.cell(...) ustundagi ma'lumot mavjudligidan qat'i nazar ishlaydi
                # (avvalgi `len(row) > 5` tekshiruvi ba'zi Excel fayllarda doim False bo'lib,
                # rasm hech qachon saqlanmasligiga sabab bo'lardi).
                sheet.cell(row=row[0].row, column=6).value = rasm_manzili
                _excel_atomik_saqlash(workbook)
                return True

        return False


# ==========================================
# HOLATNI SAQLASH (savat, buyurtma jarayoni, tasdiq kutayotgan buyurtmalar)
# Bot qayta ishga tushganda (Railway deploy/restart) bu ma'lumotlar yo'qolib
# ketmasligi uchun har necha soniyada avtomatik diskka yoziladi.
# ==========================================
HOLAT_FAYLI = "bot_holati.json"
HOLAT_LOCK = threading.Lock()


def holatni_saqlash():
    with HOLAT_LOCK:
        try:
            data = {
                "savat": {str(k): v for k, v in savat.items()},
                "buyurtma_holati": {str(k): v for k, v in buyurtma_holati.items()},
                "kutilayotgan_buyurtmalar": {str(k): v for k, v in kutilayotgan_buyurtmalar.items()},
                "guruh_buyurtmalari": {str(k): v for k, v in guruh_buyurtmalari.items()},
                "kutilayotgan_royxatlar": {str(k): v for k, v in kutilayotgan_royxatlar.items()},
            }
            vaqtinchalik = HOLAT_FAYLI + ".tmp"
            with open(vaqtinchalik, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            os.replace(vaqtinchalik, HOLAT_FAYLI)
        except Exception:
            log.exception("Holatni saqlashda xatolik")


def holatni_yuklash():
    if not os.path.exists(HOLAT_FAYLI):
        return
    try:
        with open(HOLAT_FAYLI, "r", encoding="utf-8") as f:
            data = json.load(f)

        savat.update({int(k): v for k, v in data.get("savat", {}).items()})
        buyurtma_holati.update({int(k): v for k, v in data.get("buyurtma_holati", {}).items()})
        kutilayotgan_buyurtmalar.update({int(k): v for k, v in data.get("kutilayotgan_buyurtmalar", {}).items()})
        guruh_buyurtmalari.update({int(k): v for k, v in data.get("guruh_buyurtmalari", {}).items()})
        kutilayotgan_royxatlar.update({int(k): v for k, v in data.get("kutilayotgan_royxatlar", {}).items()})

        log.info(
            "Oldingi holat tiklandi: savat=%s, jarayondagi_buyurtma=%s, "
            "tasdiq_kutayotgan_buyurtma=%s, sklad_buyurtmalari=%s, tasdiq_kutayotgan_royxat=%s",
            len(savat), len(buyurtma_holati), len(kutilayotgan_buyurtmalar), len(guruh_buyurtmalari), len(kutilayotgan_royxatlar),
        )
    except Exception:
        log.exception("Holatni yuklashda xatolik")


def avtomatik_saqlash_oqimi():
    """Fon oqimida har 5 soniyada holatni diskka yozib turadi."""
    import time
    while True:
        time.sleep(5)
        holatni_saqlash()

def ombor_aksiyasini_yangilash(kategoriya, model, aksiya_matni):
    with EXCEL_LOCK:
        try:
            workbook = openpyxl.load_workbook(EXCEL_FILE)
        except Exception:
            log.exception("Excel faylni ochishda xatolik (ombor_aksiyasini_yangilash)")
            return False
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=1):
            row_kategoriya = row[1].value
            row_model = row[2].value
            if row_kategoriya is None or row_model is None:
                continue
            if str(row_kategoriya).strip() == kategoriya and str(row_model).strip() == model:
                sheet.cell(row=row[0].row, column=7).value = aksiya_matni
                _excel_atomik_saqlash(workbook)
                return True

        return False


BUYURTMALAR_FAYLI = "buyurtmalar.json"


def buyurtma_raqami(buyurtma_id):
    return f"INL-{buyurtma_id:07d}"


def hisoblagichni_tiklash():
    """Bot qayta ishga tushganda buyurtma raqami eski buyurtmalar bilan
    to'qnashmasligi uchun hisoblagichni buyurtmalar.json'dagi eng katta
    buyurtma_id'dan davom ettiradi."""
    eng_katta = 0
    try:
        if os.path.exists(BUYURTMALAR_FAYLI):
            with open(BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
                barcha = json.load(f)
            for b in barcha:
                bid = b.get("buyurtma_id") or 0
                if isinstance(bid, int) and bid > eng_katta:
                    eng_katta = bid
    except Exception:
        log.exception("Buyurtma hisoblagichini tiklashda xatolik")
    buyurtma_id_hisoblagich["son"] = eng_katta
    log.info("Buyurtma raqami hisoblagichi %s dan davom etadi.", eng_katta + 1)


NORTIX_KOK = colors.HexColor("#1877D6")
OMBOR_NOMI = "Xorazm ombori"  # Yuboruvchi ombor nomi — kerak bo'lsa shu qatorni o'zgartiring


def buyurtma_pdf_yaratish(buyurtma_id, buyurtma, tasdiqlagan_shaxs=None):
    """Buyurtma uchun 'Yuk xati' PDF hujjatini yaratadi va fayl yo'lini qaytaradi."""
    fayl_nomi = f"/tmp/{buyurtma_raqami(buyurtma_id)}.pdf"

    hujjat = SimpleDocTemplate(
        fayl_nomi, pagesize=A4,
        leftMargin=20 * mm, rightMargin=20 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm
    )
    elementlar = []

    logo_stil = ParagraphStyle("logo", fontName="Helvetica-Bold", fontSize=26, textColor=NORTIX_KOK)
    sarlavha_stil = ParagraphStyle("sarlavha", fontName="Helvetica-Bold", fontSize=20, textColor=NORTIX_KOK, alignment=TA_RIGHT)
    label_stil = ParagraphStyle("label", fontName="Helvetica", fontSize=9, textColor=colors.grey)
    qora_qalin_stil = ParagraphStyle("qoraQalin", fontName="Helvetica-Bold", fontSize=11)
    matn_stil = ParagraphStyle("matn", fontName="Helvetica", fontSize=10, leading=14)

    sarlavha_jadval = Table(
        [[Paragraph("Nortix", logo_stil), Paragraph("YUK XATI", sarlavha_stil)]],
        colWidths=[85 * mm, 85 * mm]
    )
    sarlavha_jadval.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elementlar.append(sarlavha_jadval)
    elementlar.append(Spacer(1, 6))
    elementlar.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#DDDDDD")))
    elementlar.append(Spacer(1, 12))

    qabul_qiluvchi = [
        Paragraph("Qabul qiluvchi (do'kon):", label_stil),
        Paragraph(buyurtma["manzil"], qora_qalin_stil),
        Paragraph(f"Egasi: {buyurtma['ism']}", matn_stil),
        Paragraph(f"Tel: {buyurtma['telefon']}", matn_stil),
    ]

    info_jadval = Table(
        [[
            [
                Paragraph("Faktura raqami:", label_stil),
                Paragraph(buyurtma_raqami(buyurtma_id), qora_qalin_stil),
                Spacer(1, 4),
                Paragraph("Sana:", label_stil),
                Paragraph(datetime.now().strftime("%d.%m.%Y %H:%M"), matn_stil),
            ],
            [
                Paragraph("Yuboruvchi:", label_stil),
                Paragraph(OMBOR_NOMI, qora_qalin_stil),
            ],
            qabul_qiluvchi,
        ]],
        colWidths=[55 * mm, 55 * mm, 60 * mm]
    )
    info_jadval.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elementlar.append(info_jadval)
    elementlar.append(Spacer(1, 18))

    jadval_malumot = [["Mahsulot nomi", "Miqdori", "Narx", "Summa"]]
    for item in buyurtma["itemlar"]:
        summa = item["son"] * item["narxi"]
        jadval_malumot.append([
            f"{item['model']} ({item['kategoriya']})",
            str(item["son"]),
            f"{item['narxi']:,.0f}",
            f"{summa:,.0f}",
        ])
    jadval_malumot.append(["JAMI", str(sum(i["son"] for i in buyurtma["itemlar"])), "", f"{buyurtma['jami_summa']:,.0f}"])

    mahsulot_jadval = Table(jadval_malumot, colWidths=[80 * mm, 25 * mm, 25 * mm, 40 * mm])
    mahsulot_jadval.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NORTIX_KOK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F5F7FA")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E3F0FD")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, -1), (0, -1), NORTIX_KOK),
        ("TEXTCOLOR", (-1, -1), (-1, -1), NORTIX_KOK),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E0E0")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elementlar.append(mahsulot_jadval)
    elementlar.append(Spacer(1, 20))

    tasdiqlovchi_matni = f"Tasdiqladi: Savdo bo'limi ({tasdiqlagan_shaxs})" if tasdiqlagan_shaxs else "Tasdiqladi: Savdo bo'limi"
    elementlar.append(Paragraph(tasdiqlovchi_matni, matn_stil))
    elementlar.append(Spacer(1, 30))

    imzo_jadval = Table(
        [
            ["_______________________", "_______________________", "_______________________"],
            ["Sotuvchi", "Omborchi", "Qabul qildi"],
        ],
        colWidths=[55 * mm, 55 * mm, 55 * mm]
    )
    imzo_jadval.setStyle(TableStyle([
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#BBBBBB")),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 1), (-1, 1), 2),
    ]))
    elementlar.append(imzo_jadval)
    elementlar.append(Spacer(1, 40))

    elementlar.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#DDDDDD")))
    elementlar.append(Spacer(1, 6))
    footer_jadval = Table(
        [[Paragraph("nortix.uz", matn_stil), Paragraph("Nortix", ParagraphStyle("footerLogo", fontName="Helvetica-Bold", fontSize=14, textColor=NORTIX_KOK, alignment=TA_RIGHT))]],
        colWidths=[85 * mm, 85 * mm]
    )
    footer_jadval.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elementlar.append(footer_jadval)

    hujjat.build(elementlar)
    return fayl_nomi


def buyurtmani_saqlash(chat_id, itemlar, ism, telefon, manzil, jami_summa, holat="tasdiqlangan", buyurtma_id=None):
    try:
        if os.path.exists(BUYURTMALAR_FAYLI):
            with open(BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
                barcha = json.load(f)
        else:
            barcha = []
    except Exception:
        barcha = []

    barcha.append({
        "buyurtma_id": buyurtma_id,
        "chat_id": chat_id,
        "sana": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "itemlar": itemlar,
        "ism": ism,
        "telefon": telefon,
        "manzil": manzil,
        "jami_summa": jami_summa,
        "holat": holat,
    })

    with open(BUYURTMALAR_FAYLI, "w", encoding="utf-8") as f:
        json.dump(barcha, f, ensure_ascii=False, indent=2)


def mijoz_buyurtmalarini_olish(chat_id):
    try:
        with open(BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
            barcha = json.load(f)
    except Exception:
        return []

    mijoznikilar = [b for b in barcha if b.get("chat_id") == chat_id]
    return list(reversed(mijoznikilar))


def barcha_buyurtmalarni_indeks_bilan_olish():
    try:
        with open(BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
            barcha = json.load(f)
    except Exception:
        return []

    return list(reversed(list(enumerate(barcha))))


KONTAKTLAR = [
    ("Muhiddin", "+998 91 999 40 30"),
    ("Durdivoy", "+998 91 422 27 72"),
    ("Murod", "+998 90 436 63 63"),
    ("Zafar", "+998 97 033 39 39"),
]
KONTAKT_MANZIL = "Toshkent city, Boulevar"
ISH_VAQTI = (
    "Dushanba: 9:30-18:30\n"
    "Seshanba: 9:30-18:30\n"
    "Chorshanba: 9:30-18:30\n"
    "Payshanba: 9:30-18:30\n"
    "Juma: 9:30-18:30\n"
    "Shanba: 9:30-18:30\n"
    "Yakshanba: Dam olish kuni"
)


def bosh_menyu_yaratish(user_id=None):
    """Asosiy pastki menyu tugmalari (Brend qo'shildi)"""
    menyu = types.ReplyKeyboardMarkup(resize_keyboard=True)
    menyu.add(types.KeyboardButton("🏷 Brendlar"), types.KeyboardButton("🔍 Qidiruv"))
    menyu.add(types.KeyboardButton("🛒 Savatim"), types.KeyboardButton("🧾 Buyurtmalarim"))
    menyu.add(types.KeyboardButton("☎️ Biz bilan bog'lanish"), types.KeyboardButton("📍 Manzilimiz"))
    menyu.add(types.KeyboardButton("ℹ️ Bot haqida"))

    if admin_mi(user_id):
        menyu.add(types.KeyboardButton("📷 Mahsulot rasmi"), types.KeyboardButton("🎉 Aksiya qo'shish"))

    return menyu


def orqaga_menyu_yaratish():
    menyu = types.ReplyKeyboardMarkup(resize_keyboard=True)
    menyu.add(types.KeyboardButton("🔙 Orqaga"), types.KeyboardButton("🏠 Bosh menyu"))
    return menyu



# ============================================================
# FAQAT ICHKI BOT: MENEJER + RAHBAR
# Mijozlar uchun registratsiya/katalog/buyurtma funksiyasi yo'q.
# ============================================================

MENEDJERLAR_FAYLI = "menedjerlar.json"
DOKONLAR_FAYLI = "dokonlar.json"
SAVDOLAR_FAYLI = "savdolar.json"
PLANLAR_FAYLI = "planlar.json"
RASXODLAR_FAYLI = "rasxodlar.json"

menedjer_holati = {}
dokon_holati = {}
savdo_holati = {}
rasxod_holati = {}

def json_yukla(fayl, default):
    try:
        if os.path.exists(fayl):
            with open(fayl, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        log.exception("%s o'qishda xatolik", fayl)
    return default

def json_saqlash(fayl, data):
    try:
        tmp = fayl + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, fayl)
        return True
    except Exception:
        log.exception("%s saqlashda xatolik", fayl)
        return False

menedjerlar = json_yukla(MENEDJERLAR_FAYLI, {})
dokonlar = json_yukla(DOKONLAR_FAYLI, {})
savdolar = json_yukla(SAVDOLAR_FAYLI, [])
planlar = json_yukla(PLANLAR_FAYLI, {})
rasxodlar = json_yukla(RASXODLAR_FAYLI, [])

# Eski holat saqlash funksiyasi bilan moslik uchun (mijoz registratsiyasi ishlatilmaydi)
kutilayotgan_royxatlar = {}

def menejer_ol(user_id):
    return menedjerlar.get(str(user_id))

def menejer_tasdiqlangan(user_id):
    m = menejer_ol(user_id)
    return bool(m and m.get("status") == "tasdiqlangan")

def rahbar_mi(user_id):
    # Hozircha ADMIN_IDS rahbar hisoblanadi.
    # Keyin alohida Rahbarlar ro'yxatiga o'tkazamiz.
    return admin_mi(user_id)

def menejer_ruxsat(user_id):
    return admin_mi(user_id) or menejer_tasdiqlangan(user_id)

def menejer_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add(types.KeyboardButton("🏪 Mening do'konlarim"))
    kb.add(types.KeyboardButton("🛒 Savdo kiritish"), types.KeyboardButton("📊 Mening savdom"))
    kb.add(types.KeyboardButton("🎯 Mening planim"), types.KeyboardButton("💰 Qarzdorlik"))
    kb.add(types.KeyboardButton("💸 Rasxod"), types.KeyboardButton("📦 Buyurtmalar"))
    kb.add(types.KeyboardButton("👤 Profilim"))
    return kb

def rahbar_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add(types.KeyboardButton("👨‍💼 Menejerlar"), types.KeyboardButton("🏪 Do'konlar"))
    kb.add(types.KeyboardButton("📊 Umumiy savdo"), types.KeyboardButton("🏆 Menejerlar reytingi"))
    kb.add(types.KeyboardButton("🎯 Planlar"), types.KeyboardButton("💰 Qarzdorlik"))
    kb.add(types.KeyboardButton("💸 Rasxodlar"))
    kb.add(types.KeyboardButton("📈 Savdo analitikasi"))
    kb.add(types.KeyboardButton("📷 Mahsulot rasmi"), types.KeyboardButton("🎉 Aksiya qo'shish"))
    kb.add(types.KeyboardButton("🔄 Ombor sonini yangilash"))
    return kb

def menejer_royxatdan_otishni_boshlash(chat_id):
    uid = chat_id
    mavjud = menejer_ol(uid)
    if mavjud:
        status = mavjud.get("status")
        if status == "tasdiqlangan":
            bot.send_message(chat_id, "✅ Sizning menejer profilingiz tasdiqlangan.", reply_markup=menejer_menu())
        elif status == "kutilmoqda":
            bot.send_message(chat_id, "⏳ Arizangiz admin tasdig'ini kutmoqda.")
        else:
            menedjer_holati[uid] = {"bosqich": "ism"}
            bot.send_message(chat_id, "Qayta ariza uchun 👤 Ism-familiyangizni kiriting:",
                             reply_markup=types.ReplyKeyboardRemove())
        return
    menedjer_holati[uid] = {"bosqich": "ism"}
    bot.send_message(chat_id,
        "👨‍💼 Menejer sifatida ro'yxatdan o'tish.\n\n"
        "👤 Ism-familiyangizni kiriting:",
        reply_markup=types.ReplyKeyboardRemove())

def manager_approval_kb(uid):
    kb = types.InlineKeyboardMarkup()
    kb.add(
        types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"men_ok:{uid}"),
        types.InlineKeyboardButton("❌ Rad etish", callback_data=f"men_rad:{uid}")
    )
    return kb

@bot.message_handler(commands=["manager", "menejer"])
def manager_command(message):
    menejer_royxatdan_otishni_boshlash(message.chat.id)

@bot.message_handler(func=lambda m: m.text == "👨‍💼 Menejer")
def manager_button(message):
    menejer_royxatdan_otishni_boshlash(message.chat.id)

@bot.message_handler(content_types=["text", "contact"],
                     func=lambda m: m.from_user.id in menedjer_holati)
def manager_registration(message):
    uid = message.from_user.id
    state = menedjer_holati.get(uid)
    if not state:
        return
    text = (message.text or "").strip()

    if state["bosqich"] == "ism":
        if len(text) < 2 or text.startswith("/"):
            bot.send_message(message.chat.id, "Iltimos, ism-familiyangizni to'g'ri kiriting.")
            return
        state["ism"] = text
        state["bosqich"] = "telefon"
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        kb.add(types.KeyboardButton("📱 Telefon raqamni yuborish", request_contact=True))
        bot.send_message(message.chat.id, "📱 Telefon raqamingizni yuboring:", reply_markup=kb)
        return

    if state["bosqich"] == "telefon":
        telefon = ""
        if message.content_type == "contact" and message.contact:
            telefon = message.contact.phone_number
        else:
            telefon = text
        telefon = telefon.replace(" ", "").replace("-", "")
        if not re.match(r"^\+?\d{9,15}$", telefon):
            bot.send_message(message.chat.id, "Telefon raqami noto'g'ri. Masalan: +998901234567")
            return

        menedjerlar[str(uid)] = {
            "telegram_id": uid,
            "username": message.from_user.username or "",
            "ism": state["ism"],
            "telefon": telefon,
            "status": "kutilmoqda",
            "sana": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        json_saqlash(MENEDJERLAR_FAYLI, menedjerlar)
        menedjer_holati.pop(uid, None)

        bot.send_message(message.chat.id,
            "✅ Menejerlik arizangiz yuborildi.\n\n"
            "⏳ Admin tasdiqlashini kuting.",
            reply_markup=types.ReplyKeyboardRemove())

        for admin_id in ADMIN_IDLAR:
            try:
                bot.send_message(admin_id,
                    f"🆕 YANGI MENEJER ARIZASI\n\n"
                    f"👤 {state['ism']}\n"
                    f"📱 {telefon}\n"
                    f"🆔 {uid}\n"
                    f"👤 @{message.from_user.username or 'username yo‘q'}",
                    reply_markup=manager_approval_kb(uid))
            except Exception:
                log.exception("Admin %s ga menejer arizasini yuborishda xato", admin_id)
        return

@bot.callback_query_handler(func=lambda c: c.data.startswith("men_ok:"))
def manager_approve(call):
    if not admin_mi(call.from_user.id):
        bot.answer_callback_query(call.id, "Ruxsat yo'q", show_alert=True)
        return
    uid = int(call.data.split(":", 1)[1])
    m = menedjer_ol(uid)
    if not m:
        bot.answer_callback_query(call.id, "Menejer topilmadi.", show_alert=True)
        return
    m["status"] = "tasdiqlangan"
    m["tasdiqlagan_admin"] = call.from_user.id
    m["tasdiqlangan_sana"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    json_saqlash(MENEDJERLAR_FAYLI, menedjerlar)
    bot.answer_callback_query(call.id, "Tasdiqlandi")
    try:
        bot.edit_message_text(call.message.text + "\n\n✅ TASDIQLANDI",
                              call.message.chat.id, call.message.message_id)
    except Exception:
        pass
    try:
        bot.send_message(uid,
            "🎉 Tabriklaymiz! Menejerlik profilingiz tasdiqlandi.\n\n"
            "Endi /start orqali menejer panelidan foydalanishingiz mumkin.",
            reply_markup=menejer_menu())
    except Exception:
        pass

@bot.callback_query_handler(func=lambda c: c.data.startswith("men_rad:"))
def manager_reject(call):
    if not admin_mi(call.from_user.id):
        bot.answer_callback_query(call.id, "Ruxsat yo'q", show_alert=True)
        return
    uid = int(call.data.split(":", 1)[1])
    m = menejer_ol(uid)
    if not m:
        bot.answer_callback_query(call.id, "Menejer topilmadi.", show_alert=True)
        return
    m["status"] = "rad_etilgan"
    m["rad_etgan_admin"] = call.from_user.id
    m["rad_sana"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    json_saqlash(MENEDJERLAR_FAYLI, menedjerlar)
    bot.answer_callback_query(call.id, "Rad etildi")
    try:
        bot.edit_message_text(call.message.text + "\n\n❌ RAD ETILDI",
                              call.message.chat.id, call.message.message_id)
    except Exception:
        pass
    try:
        bot.send_message(uid, "❌ Menejerlik arizangiz rad etildi. Qayta topshirish uchun /manager buyrug'ini bosing.")
    except Exception:
        pass

@bot.message_handler(commands=["start"])
def start_handler_manager(message):
    uid = message.from_user.id
    if admin_mi(uid):
        bot.send_message(message.chat.id, "👑 Rahbar paneli", reply_markup=rahbar_menu())
        return
    if menejer_tasdiqlangan(uid):
        bot.send_message(message.chat.id, "👨‍💼 Menejer paneli", reply_markup=menejer_menu())
        return
    menejer_royxatdan_otishni_boshlash(message.chat.id)

@bot.message_handler(func=lambda m: m.text == "👤 Profilim")
def manager_profile(message):
    m = menejer_ol(message.from_user.id)
    if not m:
        menejer_royxatdan_otishni_boshlash(message.chat.id); return
    bot.send_message(message.chat.id,
        f"👤 <b>{m.get('ism','')}</b>\n"
        f"📱 {m.get('telefon','')}\n"
        f"📊 Status: {m.get('status','')}",
        parse_mode="HTML")

def my_store_ids(uid):
    return [sid for sid, d in dokonlar.items() if str(d.get("manager_id")) == str(uid)]

@bot.message_handler(func=lambda m: m.text == "🏪 Mening do'konlarim")
def my_stores(message):
    uid = message.from_user.id
    if not menejer_tasdiqlangan(uid):
        bot.send_message(message.chat.id, "❌ Siz tasdiqlangan menejer emassiz."); return
    ids = my_store_ids(uid)
    if not ids:
        bot.send_message(message.chat.id, "🏪 Sizga hali do'kon biriktirilmagan.")
        return
    text = "🏪 <b>Mening do'konlarim</b>\n\n"
    for sid in ids:
        d = dokonlar[sid]
        text += f"• <b>{d.get('nomi','')}</b>\n  📍 {d.get('manzil','')}\n  📱 {d.get('telefon','')}\n\n"
    bot.send_message(message.chat.id, text, parse_mode="HTML")

# ---------------- DO'KON BOSHQARUVI (RAHBAR) ----------------
@bot.message_handler(func=lambda m: m.text == "🏪 Do'konlar" and rahbar_mi(m.from_user.id))
def stores_admin(message):
    kb = types.InlineKeyboardMarkup()
    kb.add(types.InlineKeyboardButton("➕ Do'kon qo'shish", callback_data="dokon_add"))
    kb.add(types.InlineKeyboardButton("📋 Do'konlar ro'yxati", callback_data="dokon_list"))
    bot.send_message(message.chat.id, "🏪 Do'konlar boshqaruvi:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data == "dokon_add")
def store_add_start(call):
    if not rahbar_mi(call.from_user.id): return
    dokon_holati[call.from_user.id] = {"bosqich":"nomi"}
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id, "🏪 Do'kon nomini kiriting:", reply_markup=types.ReplyKeyboardRemove())

@bot.callback_query_handler(func=lambda c: c.data == "dokon_list")
def store_list(call):
    if not rahbar_mi(call.from_user.id): return
    bot.answer_callback_query(call.id)
    if not dokonlar:
        bot.send_message(call.message.chat.id, "Hozircha do'konlar yo'q."); return
    for sid, d in dokonlar.items():
        manager = menejer_ol(d.get("manager_id"))
        mgr = manager.get("ism") if manager else "Biriktirilmagan"
        kb = types.InlineKeyboardMarkup()
        kb.add(types.InlineKeyboardButton("👨‍💼 Menejer biriktirish", callback_data=f"dokon_mgr:{sid}"))
        bot.send_message(call.message.chat.id,
            f"🏪 <b>{d.get('nomi','')}</b>\n📍 {d.get('manzil','')}\n"
            f"📱 {d.get('telefon','')}\n👨‍💼 Menejer: {mgr}",
            parse_mode="HTML", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("dokon_mgr:"))
def store_assign_start(call):
    if not rahbar_mi(call.from_user.id): return
    sid = call.data.split(":",1)[1]
    d = dokonlar.get(sid)
    if not d: return
    kb = types.InlineKeyboardMarkup()
    for uid, m in menedjerlar.items():
        if m.get("status") == "tasdiqlangan":
            kb.add(types.InlineKeyboardButton(m.get("ism","Noma'lum"), callback_data=f"assign:{sid}:{uid}"))
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id, f"🏪 {d.get('nomi')} uchun menejerni tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("assign:"))
def store_assign(call):
    if not rahbar_mi(call.from_user.id): return
    _, sid, uid = call.data.split(":",2)
    if sid not in dokonlar or not menejer_tasdiqlangan(uid):
        bot.answer_callback_query(call.id, "Ma'lumot topilmadi", show_alert=True); return
    dokonlar[sid]["manager_id"] = int(uid)
    json_saqlash(DOKONLAR_FAYLI, dokonlar)
    bot.answer_callback_query(call.id, "Biriktirildi")
    bot.send_message(call.message.chat.id, f"✅ {dokonlar[sid]['nomi']} do'koni menejerga biriktirildi.")
    try:
        bot.send_message(int(uid), f"🏪 Sizga yangi do'kon biriktirildi: {dokonlar[sid]['nomi']}")
    except Exception: pass

@bot.message_handler(func=lambda m: m.from_user.id in dokon_holati and dokon_holati[m.from_user.id].get("bosqich")=="plan" and rahbar_mi(m.from_user.id))
def plan_value_received(message):
    uid=message.from_user.id; st=dokon_holati.get(uid)
    raw=(message.text or "").replace(" ","").replace(",","").replace("$","")
    try: value=float(raw)
    except:
        bot.send_message(message.chat.id,"❌ Plan summasini raqam bilan kiriting."); return
    if value<0:
        bot.send_message(message.chat.id,"❌ Plan manfiy bo'lishi mumkin emas."); return
    mid=str(st["manager_id"]); oy=joriy_oy()
    planlar.setdefault(mid,{})[oy]=value
    json_saqlash(PLANLAR_FAYLI,planlar)
    dokon_holati.pop(uid,None)
    m=menejer_ol(mid)
    bot.send_message(message.chat.id,
        f"✅ {m.get('ism')} uchun {oy} oyi plani ${value:,.0f} qilib saqlandi.",
        reply_markup=rahbar_menu())
    try:
        bot.send_message(int(mid),f"🎯 Sizga {oy} oyi uchun yangi plan belgilandi: ${value:,.0f}")
    except Exception: pass


@bot.message_handler(func=lambda m: m.text == "💸 Rasxodlar" and rahbar_mi(m.from_user.id))
def expense_admin(message):
    if not rasxodlar:
        bot.send_message(message.chat.id, "💸 Hozircha rasxodlar kiritilmagan.")
        return
    jami = sum(float(x.get("summa", 0)) for x in rasxodlar if x.get("status") != "rad_etilgan")
    text = f"💸 <b>Rasxodlar</b>\n\n💰 Jami: <b>{jami:,.0f} so'm</b>\n\n"
    for x in reversed(rasxodlar[-30:]):
        status = x.get("status", "tasdiqlangan")
        belgi = {"kutilmoqda":"🕓", "tasdiqlangan":"✅", "rad_etilgan":"❌"}.get(status, "•")
        text += (f"{belgi} <b>{x.get('kategoriya','Boshqa')}</b> — {float(x.get('summa',0)):,.0f} so'm\n"
                  f"👤 {x.get("manager_name","Noma'lum")} | {x.get('sana','')}\n")
        if x.get("izoh"):
            text += f"📝 {x['izoh']}\n"
        text += "\n"
    bot.send_message(message.chat.id, text, parse_mode="HTML")

@bot.message_handler(func=lambda m: m.text == "📈 Savdo analitikasi" and rahbar_mi(m.from_user.id))
def analytics_admin(message):
    by_cat={}
    by_model={}
    for x in savdolar:
        by_cat[x["kategoriya"]]=by_cat.get(x["kategoriya"],0)+x.get("summa",0)
        by_model[x["model"]]=by_model.get(x["model"],0)+x.get("summa",0)
    text="📈 <b>Savdo analitikasi</b>\n\n<b>Kategoriya:</b>\n"
    for k,v in sorted(by_cat.items(),key=lambda z:z[1],reverse=True): text+=f"• {k}: ${v:,.0f}\n"
    text+="\n<b>Model:</b>\n"
    for k,v in sorted(by_model.items(),key=lambda z:z[1],reverse=True)[:20]: text+=f"• {k}: ${v:,.0f}\n"
    bot.send_message(message.chat.id,text,parse_mode="HTML")

@bot.message_handler(func=lambda m: m.text == "🔄 Ombor sonini yangilash" and rahbar_mi(m.from_user.id))
def stock_update_alias(message):
    # Eski /yangilash funksiyasidan foydalanish uchun
    bot.send_message(message.chat.id,"Ombor sonini yangilash uchun /yangilash buyrug'ini bosing.")

def bosh_menyu_yaratish(user_id=None):
    return rahbar_menu() if rahbar_mi(user_id) else menejer_menu()

def orqaga_menyu_yaratish():
    kb=types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add(types.KeyboardButton("🔙 Orqaga"), types.KeyboardButton("🏠 Bosh menyu"))
    return kb
# ADMIN RASM BO'LIMI FUNKSIYASI
def mahsulot_rasmi_menyu_chiqarish(chat_id, user_id):
    if not admin_mi(user_id):
        return

    try:
        malumotlar = ombor_malumotlarini_oqish()
    except Exception as e:
        bot.send_message(chat_id, f"Xatolik yuz berdi: {e}")
        return

    if not malumotlar:
        bot.send_message(chat_id, "Omborda mahsulot yo'q.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for kategoriya in malumotlar.keys():
        keyboard.add(types.InlineKeyboardButton(
            text=f"📁 {kategoriya}",
            callback_data=f"imgkat:{kategoriya}"
        ))

    keyboard.add(types.InlineKeyboardButton("🔙 Orqaga", callback_data="back_to_main_admin"))

    bot.send_message(
        chat_id,
        "📷 <b>Rasm qo'shish bo'limi:</b>\nQaysi kategoriyadagi mahsulotga rasm biriktirasiz?",
        parse_mode="HTML",
        reply_markup=keyboard
    )
    bot.send_message(
        chat_id,
        "Ortga qaytish uchun '🔙 Orqaga' tugmasini bosing:",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.text == "📷 Mahsulot rasmi")
def mahsulot_rasmi_boshlash(message):
    mahsulot_rasmi_menyu_chiqarish(message.chat.id, message.from_user.id)


@bot.callback_query_handler(func=lambda call: call.data.startswith("imgkat:"))
def rasm_kategoriya_tanlandi(call):
    bot.answer_callback_query(call.id)
    if not admin_mi(call.from_user.id):
        return

    kategoriya = call.data.split("imgkat:", 1)[1]
    malumotlar = ombor_malumotlarini_oqish()
    mahsulotlar = malumotlar.get(kategoriya, [])

    keyboard = types.InlineKeyboardMarkup()
    for item in mahsulotlar:
        model = item[0]
        rasm_bor = "🖼 " if (len(item) > 3 and item[3]) else ""
        keyboard.add(types.InlineKeyboardButton(
            text=f"{rasm_bor}{model}",
            callback_data=f"imgmod:{kategoriya}|{model}"
        ))

    keyboard.add(types.InlineKeyboardButton("🔙 Kategoriyalarga qaytish", callback_data="back_to_img_kat"))

    bot.send_message(
        call.message.chat.id,
        f"«<b>{kategoriya}</b>» kategoriyasidan modelni tanlang:",
        parse_mode="HTML",
        reply_markup=keyboard
    )


@bot.callback_query_handler(func=lambda call: call.data == "back_to_img_kat")
def back_to_img_kat_callback(call):
    bot.answer_callback_query(call.id)
    mahsulot_rasmi_menyu_chiqarish(call.message.chat.id, call.from_user.id)


@bot.callback_query_handler(func=lambda call: call.data == "back_to_main_admin")
def back_to_main_admin_callback(call):
    bot.answer_callback_query(call.id)
    rasm_kutilayotganlar.pop(call.from_user.id, None)
    aksiya_kutilayotganlar.pop(call.from_user.id, None)
    bot.send_message(
        call.message.chat.id,
        "🏠 Asosiy menyuga qaytdingiz.",
        reply_markup=bosh_menyu_yaratish(call.from_user.id)
    )


@bot.callback_query_handler(func=lambda call: call.data.startswith("imgmod:"))
def rasm_model_tanlandi(call):
    bot.answer_callback_query(call.id)
    if not admin_mi(call.from_user.id):
        return

    kategoriya, model = call.data.split("imgmod:", 1)[1].split("|", 1)

    rasm_kutilayotganlar[call.from_user.id] = {
        "kategoriya": kategoriya,
        "model": model,
    }

    bot.send_message(
        call.message.chat.id,
        f"«<b>{model}</b>» uchun rasmni yuboring (galereyadan) yoki rasm havolasini (link) yozing:\n\n"
        f"<i>Bekor qilish uchun '🔙 Orqaga' tugmasini bosing.</i>",
        parse_mode="HTML",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(content_types=['photo', 'text'], func=lambda message: message.from_user.id in rasm_kutilayotganlar)
def rasm_yoki_link_qabul_qilish(message):
    if message.text in ["🔙 Orqaga", "🏠 Bosh menyu"]:
        return

    holat = rasm_kutilayotganlar.get(message.from_user.id)
    if not holat:
        return

    kategoriya = holat["kategoriya"]
    model = holat["model"]

    rasm_manzili = None

    if message.photo:
        rasm_manzili = message.photo[-1].file_id
    elif message.text and message.text.startswith("http"):
        rasm_manzili = message.text.strip()
    else:
        bot.send_message(
            message.chat.id,
            "Iltimos, rasm yuboring yoki https://... bilan boshlanuvchi to'g'ri link kiriting.\n"
            "Bekor qilish uchun '🔙 Orqaga' tugmasini bosing.",
            reply_markup=orqaga_menyu_yaratish()
        )
        return

    muvaffaqiyatli = ombor_rasmini_yangilash(kategoriya, model, rasm_manzili)

    del rasm_kutilayotganlar[message.from_user.id]

    if muvaffaqiyatli:
        bot.send_message(
            message.chat.id,
            f"✅ <b>{model}</b> uchun rasm muvaffaqiyatli saqlandi!",
            parse_mode="HTML",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )
    else:
        bot.send_message(
            message.chat.id,
            "❌ Xatolik: ushbu model Excel'dan topilmadi.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )


def aksiya_menyu_chiqarish(chat_id, user_id):
    if not admin_mi(user_id):
        return

    try:
        malumotlar = ombor_malumotlarini_oqish()
    except Exception as e:
        bot.send_message(chat_id, f"Xatolik yuz berdi: {e}")
        return

    if not malumotlar:
        bot.send_message(chat_id, "Omborda mahsulot yo'q.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for kategoriya in malumotlar.keys():
        keyboard.add(types.InlineKeyboardButton(
            text=f"📁 {kategoriya}",
            callback_data=f"aksiyakat:{kategoriya}"
        ))

    keyboard.add(types.InlineKeyboardButton("🔙 Orqaga", callback_data="back_to_main_admin"))

    bot.send_message(
        chat_id,
        "🎉 <b>Aksiya bo'limi:</b>\nQaysi kategoriyadagi mahsulotga aksiya biriktirasiz?",
        parse_mode="HTML",
        reply_markup=keyboard
    )
    bot.send_message(
        chat_id,
        "Ortga qaytish uchun '🔙 Orqaga' tugmasini bosing:",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.text == "🎉 Aksiya qo'shish")
def aksiya_boshlash(message):
    aksiya_menyu_chiqarish(message.chat.id, message.from_user.id)


@bot.callback_query_handler(func=lambda call: call.data.startswith("aksiyakat:"))
def aksiya_kategoriya_tanlandi(call):
    bot.answer_callback_query(call.id)
    if not admin_mi(call.from_user.id):
        return

    kategoriya = call.data.split("aksiyakat:", 1)[1]
    malumotlar = ombor_malumotlarini_oqish()
    mahsulotlar = malumotlar.get(kategoriya, [])

    keyboard = types.InlineKeyboardMarkup()
    for item in mahsulotlar:
        model = item[0]
        aksiya_bor = "🎉 " if (len(item) > 5 and item[5]) else ""
        keyboard.add(types.InlineKeyboardButton(
            text=f"{aksiya_bor}{model}",
            callback_data=f"aksiyamod:{kategoriya}|{model}"
        ))

    keyboard.add(types.InlineKeyboardButton("🔙 Kategoriyalarga qaytish", callback_data="back_to_aksiya_kat"))

    bot.send_message(
        call.message.chat.id,
        f"«<b>{kategoriya}</b>» kategoriyasidan modelni tanlang:",
        parse_mode="HTML",
        reply_markup=keyboard
    )


@bot.callback_query_handler(func=lambda call: call.data == "back_to_aksiya_kat")
def back_to_aksiya_kat_callback(call):
    bot.answer_callback_query(call.id)
    aksiya_menyu_chiqarish(call.message.chat.id, call.from_user.id)


@bot.callback_query_handler(func=lambda call: call.data.startswith("aksiyamod:"))
def aksiya_model_tanlandi(call):
    bot.answer_callback_query(call.id)
    if not admin_mi(call.from_user.id):
        return

    kategoriya, model = call.data.split("aksiyamod:", 1)[1].split("|", 1)

    aksiya_kutilayotganlar[call.from_user.id] = {
        "kategoriya": kategoriya,
        "model": model,
    }

    bot.send_message(
        call.message.chat.id,
        f"«<b>{model}</b>» uchun aksiya matnini yozing (masalan: \"20% chegirma, 10-avgustgacha\"):\n\n"
        f"<i>Aksiyani olib tashlash uchun \"-\" belgisini yuboring.\n"
        f"Bekor qilish uchun '🔙 Orqaga' tugmasini bosing.</i>",
        parse_mode="HTML",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.from_user.id in aksiya_kutilayotganlar)
def aksiya_matnini_qabul_qilish(message):
    if message.text in ["🔙 Orqaga", "🏠 Bosh menyu"]:
        return

    holat = aksiya_kutilayotganlar.get(message.from_user.id)
    if not holat:
        return

    kategoriya = holat["kategoriya"]
    model = holat["model"]

    matn = (message.text or "").strip()
    if not matn:
        bot.send_message(message.chat.id, "Iltimos, aksiya matnini yozing.")
        return

    aksiya_matni = "" if matn == "-" else matn
    muvaffaqiyatli = ombor_aksiyasini_yangilash(kategoriya, model, aksiya_matni)

    del aksiya_kutilayotganlar[message.from_user.id]

    if muvaffaqiyatli:
        if aksiya_matni:
            xabar = f"✅ <b>{model}</b> uchun aksiya saqlandi:\n{aksiya_matni}"
        else:
            xabar = f"✅ <b>{model}</b> uchun aksiya olib tashlandi."
        bot.send_message(
            message.chat.id, xabar, parse_mode="HTML",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )
    else:
        bot.send_message(
            message.chat.id,
            "❌ Xatolik: ushbu model Excel'dan topilmadi.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )


@bot.callback_query_handler(func=lambda call: call.data.startswith("aksiya_korish:"))
def aksiya_korish(call):
    kategoriya, model = call.data.split("aksiya_korish:", 1)[1].split("|", 1)

    try:
        malumotlar = ombor_malumotlarini_oqish()
        mahsulotlar = {item[0]: item for item in malumotlar.get(kategoriya, [])}
        item = mahsulotlar.get(model)
        aksiya_matni = item[5] if item and len(item) > 5 and item[5] else "Bu mahsulot uchun aksiya topilmadi."
    except Exception:
        aksiya_matni = "Xatolik yuz berdi."

    bot.answer_callback_query(call.id, text=f"🎉 {model}\n\n{aksiya_matni}", show_alert=True)


@bot.message_handler(func=lambda message: message.text == "🔍 Qidiruv")
def qidiruv_boshlash(message):
    qidiruv_kutayotganlar.add(message.chat.id)
    bot.send_message(
        message.chat.id,
        "Model nomini (yoki uning bir qismini) yozing:",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.chat.id in qidiruv_kutayotganlar)
def qidiruv_natijasi(message):
    qidiruv_kutayotganlar.discard(message.chat.id)

    soz = (message.text or "").strip().lower()
    if not soz:
        bot.send_message(message.chat.id, "Iltimos, qidiruv uchun matn yozing.")
        return

    try:
        malumotlar = ombor_malumotlarini_oqish()
    except Exception as e:
        bot.send_message(message.chat.id, f"Xatolik yuz berdi: {e}")
        return

    topilganlar = []
    for kategoriya, mahsulotlar in malumotlar.items():
        for item in mahsulotlar:
            model, soni = item[0], item[1]
            narxi = item[2] if len(item) > 2 and item[2] else 0
            aksiya_bor = "🎉 " if (len(item) > 5 and item[5]) else ""
            if soz in model.lower() and soni and soni > 0:
                topilganlar.append((kategoriya, model, soni, narxi, aksiya_bor))

    if not topilganlar:
        bot.send_message(
            message.chat.id,
            "Hech narsa topilmadi. Boshqa so'z bilan urinib ko'ring.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )
        return

    keyboard = types.InlineKeyboardMarkup()
    for kategoriya, model, soni, narxi, aksiya_bor in topilganlar[:30]:
        keyboard.add(types.InlineKeyboardButton(
            text=f"{aksiya_bor}{model} — ${narxi:,.2f} ({kategoriya})",
            callback_data=f"buy:{kategoriya}|{model}"
        ))

    bot.send_message(
        message.chat.id,
        f"Topildi: {len(topilganlar)} ta natija",
        reply_markup=keyboard
    )


def savat_matni_yaratish(chat_id):
    itemlar = savat.get(chat_id, [])
    if not itemlar:
        return "Savatingiz bo'sh.", 0

    matn = "🛒 Savatingiz:\n\n"
    jami = 0
    for item in itemlar:
        summa = item["narxi"] * item["son"]
        jami += summa
        matn += f"• {item['model']} — {item['son']} dona x ${item['narxi']:,.2f} = ${summa:,.2f}\n"

    matn += f"\n💰 Jami: ${jami:,.2f}"
    return matn, jami


def savat_tugmalari(chat_id):
    keyboard = types.InlineKeyboardMarkup()
    itemlar = savat.get(chat_id, [])

    for idx, item in enumerate(itemlar):
        keyboard.add(types.InlineKeyboardButton(
            text=f"❌ O'chirish: {item['model']}",
            callback_data=f"del_cart:{idx}"
        ))

    keyboard.add(types.InlineKeyboardButton("➕ Yana mahsulot qo'shish", callback_data="yana_mahsulot"))
    keyboard.add(types.InlineKeyboardButton("🧾 Buyurtmani rasmiylashtirish", callback_data="savat_yakunlash"))
    keyboard.add(types.InlineKeyboardButton("🗑 Savatni bo'shatish", callback_data="savat_tozalash"))
    return keyboard


@bot.message_handler(func=lambda message: message.text == "🛒 Savatim")
def menyu_savat(message):
    matn, jami = savat_matni_yaratish(message.chat.id)
    if not savat.get(message.chat.id):
        bot.send_message(message.chat.id, matn)
    else:
        bot.send_message(message.chat.id, matn, reply_markup=savat_tugmalari(message.chat.id))


@bot.callback_query_handler(func=lambda call: call.data.startswith("del_cart:"))
def savatdan_ochirish(call):
    bot.answer_callback_query(call.id)
    idx = int(call.data.split("del_cart:", 1)[1])
    chat_id = call.message.chat.id

    itemlar = savat.get(chat_id, [])
    if 0 <= idx < len(itemlar):
        ochirilgan = itemlar.pop(idx)
        bot.send_message(chat_id, f"❌ «{ochirilgan['model']}» savatdan o'chirildi.")

    matn, jami = savat_matni_yaratish(chat_id)
    if not itemlar:
        bot.send_message(chat_id, "Savatingiz bo'sh qoldi.")
    else:
        bot.send_message(chat_id, matn, reply_markup=savat_tugmalari(chat_id))


@bot.callback_query_handler(func=lambda call: call.data == "yana_mahsulot")
def yana_mahsulot_qoshish(call):
    bot.answer_callback_query(call.id)
    kategoriyalarni_korsatish(call.message.chat.id)


@bot.callback_query_handler(func=lambda call: call.data == "back_to_categories")
def back_to_categories_callback(call):
    bot.answer_callback_query(call.id)
    kategoriyalarni_korsatish(call.message.chat.id)


@bot.callback_query_handler(func=lambda call: call.data == "savat_tozalash")
def savatni_tozalash(call):
    bot.answer_callback_query(call.id)
    savat.pop(call.message.chat.id, None)
    bot.send_message(call.message.chat.id, "🗑 Savatingiz tozalandi.")


HOLAT_MATNLARI = {
    "tasdiqlangan": "✅ Tasdiqlangan",
    "bekor_qilingan": "❌ Bekor qilingan",
}


@bot.message_handler(func=lambda message: message.text == "🧾 Buyurtmalarim")
def menyu_buyurtmalarim(message):
    if admin_mi(message.from_user.id):
        buyurtmalar = barcha_buyurtmalarni_indeks_bilan_olish()

        if not buyurtmalar:
            bot.send_message(message.chat.id, "Hozircha hech qanday zakaz yo'q.")
            return

        for idx, buyurtma in buyurtmalar[:20]:
            holat_matni = HOLAT_MATNLARI.get(buyurtma.get("holat", "tasdiqlangan"), "🕓 Kutilmoqda")
            matn = f"🧾 Zakaz #{buyurtma.get('buyurtma_id', '-')}  •  {buyurtma['sana']}  •  {holat_matni}\n\n"
            for item in buyurtma["itemlar"]:
                matn += f"• {item['model']} — {item['son']} dona x ${item['narxi']:,.2f}\n"
            matn += f"\n💰 Jami: ${buyurtma['jami_summa']:,.2f}"
            matn += f"\n👤 {buyurtma['ism']}  📞 {buyurtma['telefon']}"
            matn += f"\n📍 Manzil: {buyurtma['manzil']}"

            ochirish_klaviatura = types.InlineKeyboardMarkup()
            ochirish_klaviatura.add(types.InlineKeyboardButton(
                "🗑 Tarixdan o'chirish", callback_data=f"buyurtma_ochir:{idx}"
            ))

            bot.send_message(message.chat.id, matn, reply_markup=ochirish_klaviatura)
        return

    buyurtmalar = mijoz_buyurtmalarini_olish(message.chat.id)

    if not buyurtmalar:
        bot.send_message(message.chat.id, "Sizda hali buyurtmalar yo'q.")
        return

    for buyurtma in buyurtmalar[:10]:
        holat_matni = HOLAT_MATNLARI.get(buyurtma.get("holat", "tasdiqlangan"), "🕓 Kutilmoqda")
        matn = f"🧾 {buyurtma['sana']}  •  {holat_matni}\n\n"
        for item in buyurtma["itemlar"]:
            matn += f"• {item['model']} — {item['son']} dona x ${item['narxi']:,.2f}\n"
        matn += f"\n💰 Jami: ${buyurtma['jami_summa']:,.2f}"
        matn += f"\n📍 Manzil: {buyurtma['manzil']}"
        bot.send_message(message.chat.id, matn)


@bot.message_handler(func=lambda message: message.text == "☎️ Biz bilan bog'lanish")
def menyu_kontakt(message):
    kontaktlar_matni = "\n".join(f"👤 {ism}: {tel}" for ism, tel in KONTAKTLAR)
    bot.send_message(
        message.chat.id,
        f"👥 Menedjerlar:\n\n"
        f"{kontaktlar_matni}\n\n"
        f"🕐 Ish vaqti:\n{ISH_VAQTI}"
    )


@bot.message_handler(func=lambda message: message.text == "📍 Manzilimiz")
def menyu_manzil(message):
    bot.send_message(message.chat.id, f"📍 Manzil: {KONTAKT_MANZIL}")


@bot.message_handler(func=lambda message: message.text == "ℹ️ Bot haqida")
def menyu_haqida(message):
    bot.send_message(
        message.chat.id,
        "Bu bot orqali siz ombordagi mahsulotlar qoldig'ini ko'rishingiz "
        "va zakaz berishingiz mumkin."
    )


def kategoriyalarni_korsatish(chat_id):
    try:
        malumotlar = ombor_malumotlarini_oqish()
    except FileNotFoundError:
        bot.send_message(chat_id, "Xatolik: Baza.xlsx fayli topilmadi.")
        return
    except Exception as e:
        bot.send_message(chat_id, f"Xatolik yuz berdi: {e}")
        return

    if not malumotlar:
        bot.send_message(chat_id, "Omborda hozircha mahsulot yo'q.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for kategoriya in malumotlar.keys():
        keyboard.add(types.InlineKeyboardButton(
            text=kategoriya,
            callback_data=f"kat:{kategoriya}"
        ))

    bot.send_message(chat_id, "Kategoriyani tanlang:", reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: call.data.startswith("kat:"))
def kategoriya_tanlandi(call):
    bot.answer_callback_query(call.id)
    kategoriya = call.data.split("kat:", 1)[1]

    try:
        malumotlar = ombor_malumotlarini_oqish()
    except Exception as e:
        bot.send_message(call.message.chat.id, f"Xatolik yuz berdi: {e}")
        return

    mahsulotlar = malumotlar.get(kategoriya, [])
    mahsulotlar = [item for item in mahsulotlar if item[1] and item[1] > 0]

    if not mahsulotlar:
        bot.send_message(call.message.chat.id, "Bu kategoriyada hozircha omborda mahsulot yo'q.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for item in mahsulotlar:
        model, soni = item[0], item[1]
        narxi = item[2] if len(item) > 2 and item[2] else 0
        aksiya_bor = "🎉 " if (len(item) > 5 and item[5]) else ""
        keyboard.add(types.InlineKeyboardButton(
            text=f"{aksiya_bor}{model} — ${narxi:,.2f}",
            callback_data=f"buy:{kategoriya}|{model}"
        ))

    keyboard.add(types.InlineKeyboardButton("🔙 Kategoriyalarga qaytish", callback_data="back_to_categories"))

    bot.send_message(
        call.message.chat.id,
        f"📦 {kategoriya}\n\nZakaz bermoqchi bo'lgan modelni tanlang:",
        reply_markup=keyboard
    )


@bot.callback_query_handler(func=lambda call: call.data.startswith("buy:"))
def zakaz_boshlash(call):
    bot.answer_callback_query(call.id)

    kategoriya, model = call.data.split("buy:", 1)[1].split("|", 1)

    malumotlar = ombor_malumotlarini_oqish()
    mahsulotlar = {
        item[0]: (item[1], item[2], item[3] if len(item) > 3 else None, item[5] if len(item) > 5 else "")
        for item in malumotlar.get(kategoriya, [])
    }
    mavjud_son, narxi, rasm, aksiya_matni = mahsulotlar.get(model, (0, 0, None, ""))

    buyurtma_holati[call.message.chat.id] = {
        "bosqich": "son",
        "kategoriya": kategoriya,
        "model": model,
        "telegram_user_id": call.from_user.id,
        "telegram_username": call.from_user.username,
    }

    matn = f"«{model}»\n💰 Narxi: ${narxi:,.2f}\n📊 Omborda: {mavjud_son} dona bor."
    if aksiya_matni:
        matn += "\n🎉 Aksiyada!"
    matn += "\n\nNechta dona kerak? Raqam bilan yozing:"

    aksiya_klaviatura = None
    if aksiya_matni:
        aksiya_klaviatura = types.InlineKeyboardMarkup()
        aksiya_klaviatura.add(types.InlineKeyboardButton(
            "🎉 Aksiya haqida", callback_data=f"aksiya_korish:{kategoriya}|{model}"
        ))

    if rasm and str(rasm).strip():
        try:
            bot.send_photo(
                call.message.chat.id,
                photo=str(rasm).strip(),
                caption=matn,
                reply_markup=orqaga_menyu_yaratish()
            )
            if aksiya_klaviatura:
                bot.send_message(call.message.chat.id, "👇", reply_markup=aksiya_klaviatura)
            return
        except Exception:
            pass

    bot.send_message(
        call.message.chat.id,
        matn,
        reply_markup=orqaga_menyu_yaratish()
    )
    if aksiya_klaviatura:
        bot.send_message(call.message.chat.id, "👇", reply_markup=aksiya_klaviatura)


@bot.callback_query_handler(func=lambda call: call.data == "savat_yakunlash")
def savatni_yakunlash_boshlash(call):
    bot.answer_callback_query(call.id)

    if not savat.get(call.message.chat.id):
        bot.send_message(call.message.chat.id, "Savatingiz bo'sh.")
        return

    buyurtma_holati[call.message.chat.id] = {
        "bosqich": "ism",
        "telegram_user_id": call.from_user.id,
        "telegram_username": call.from_user.username,
    }

    bot.send_message(
        call.message.chat.id,
        "Ismingizni yozing:",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.chat.id in buyurtma_holati)
def zakaz_bosqichlari(message):
    holat = buyurtma_holati.get(message.chat.id)
    if not holat:
        return

    matn = message.text.strip() if message.text else ""

    if holat["bosqich"] == "son":
        if not matn.isdigit() or int(matn) <= 0:
            bot.send_message(message.chat.id, "Iltimos, faqat musbat butun son yozing (masalan: 3).")
            return

        so_ralgan_son = int(matn)

        malumotlar = ombor_malumotlarini_oqish()
        mahsulotlar = {item[0]: (item[1], item[2]) for item in malumotlar.get(holat["kategoriya"], [])}
        mavjud_son, narxi = mahsulotlar.get(holat["model"], (0, 0))

        if so_ralgan_son > mavjud_son:
            bot.send_message(
                message.chat.id,
                f"Kechirasiz, omborda faqat {mavjud_son} dona bor. Kamroq son kiriting:"
            )
            return

        savat.setdefault(message.chat.id, []).append({
            "kategoriya": holat["kategoriya"],
            "model": holat["model"],
            "son": so_ralgan_son,
            "narxi": narxi,
        })

        del buyurtma_holati[message.chat.id]

        matn_savat, jami = savat_matni_yaratish(message.chat.id)
        bot.send_message(
            message.chat.id,
            f"✅ Savatga qo'shildi!\n\n{matn_savat}",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )
        bot.send_message(
            message.chat.id,
            "Tanlang:",
            reply_markup=savat_tugmalari(message.chat.id)
        )
        return

    elif holat["bosqich"] == "ism":
        if not matn:
            bot.send_message(message.chat.id, "Iltimos, ismingizni yozing:")
            return
        holat["ism"] = matn
        holat["bosqich"] = "telefon"
        bot.send_message(
            message.chat.id,
            "Telefon raqamingizni yozing (masalan: +998901234567):",
            reply_markup=orqaga_menyu_yaratish()
        )

    elif holat["bosqich"] == "telefon":
        if not matn:
            bot.send_message(message.chat.id, "Iltimos, telefon raqamingizni yozing:")
            return
        holat["telefon"] = matn
        holat["bosqich"] = "manzil"
        bot.send_message(
            message.chat.id,
            "Yetkazib berish manzilini yozing:",
            reply_markup=orqaga_menyu_yaratish()
        )

    elif holat["bosqich"] == "manzil":
        if not matn:
            bot.send_message(message.chat.id, "Iltimos, manzilni yozing:")
            return
        holat["manzil"] = matn

        itemlar = savat.get(message.chat.id, [])
        if not itemlar:
            bot.send_message(
                message.chat.id,
                "Savatingiz bo'sh qoldi, buyurtma bekor qilindi.",
                reply_markup=bosh_menyu_yaratish(message.from_user.id)
            )
            del buyurtma_holati[message.chat.id]
            return

        jami_summa = 0
        mahsulotlar_matni = ""
        for item in itemlar:
            summa = item["narxi"] * item["son"]
            jami_summa += summa
            mahsulotlar_matni += (
                f"📦 {item['model']} ({item['kategoriya']})\n"
                f"   {item['son']} dona x ${item['narxi']:,.2f} = ${summa:,.2f}\n"
            )

        buyurtma_id_hisoblagich["son"] += 1
        buyurtma_id = buyurtma_id_hisoblagich["son"]

        username_qismi = f"@{holat['telegram_username']}" if holat["telegram_username"] else "yo'q"

        kutilayotgan_buyurtmalar[buyurtma_id] = {
            "chat_id": message.chat.id,
            "itemlar": itemlar,
            "ism": holat["ism"],
            "telefon": holat["telefon"],
            "manzil": holat["manzil"],
            "jami_summa": jami_summa,
        }
        holatni_saqlash()  # darhol diskka yozamiz — yangi buyurtma yo'qolmasligi kerak

        admin_xabari = (
            f"🆕 Yangi zakaz #{buyurtma_id} — tasdiq kutilmoqda\n\n"
            f"{mahsulotlar_matni}\n"
            f"💰 Jami: ${jami_summa:,.2f}\n\n"
            f"👤 Ism: {holat['ism']}\n"
            f"📞 Telefon: {holat['telefon']}\n"
            f"📍 Manzil: {holat['manzil']}\n"
            f"💬 Telegram: {username_qismi}"
        )

        tasdiq_klaviatura = types.InlineKeyboardMarkup()
        tasdiq_klaviatura.add(
            types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"admin_ok:{buyurtma_id}"),
            types.InlineKeyboardButton("❌ Bekor qilish", callback_data=f"admin_bekor:{buyurtma_id}"),
        )

        bot.send_message(ZAKAZ_GRUPPA_ID, admin_xabari, reply_markup=tasdiq_klaviatura)

        bot.send_message(
            message.chat.id,
            "🕓 Buyurtmangiz qabul qilindi va hozir tekshirilmoqda. Tasdiqlangach sizga xabar beramiz.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )

        savat.pop(message.chat.id, None)
        del buyurtma_holati[message.chat.id]


@bot.callback_query_handler(func=lambda call: call.data.startswith("admin_ok:"))
def admin_zakazni_tasdiqlash(call):
    bot.answer_callback_query(call.id)
    if str(call.message.chat.id) != str(ZAKAZ_GRUPPA_ID):
        return

    buyurtma_id = int(call.data.split("admin_ok:", 1)[1])
    buyurtma = kutilayotgan_buyurtmalar.pop(buyurtma_id, None)
    holatni_saqlash()
    if not buyurtma:
        bot.send_message(call.message.chat.id, "Bu buyurtma topilmadi — avval tasdiqlangan yoki bekor qilingan bo'lishi mumkin.")
        return

    malumotlar = ombor_malumotlarini_oqish()
    ombor_matni = ""
    for item in buyurtma["itemlar"]:
        mahsulotlar_dict = {i[0]: (i[1], i[2]) for i in malumotlar.get(item["kategoriya"], [])}
        mavjud_son, _ = mahsulotlar_dict.get(item["model"], (0, 0))
        yangi_son = max(0, mavjud_son - item["son"])
        ombor_sonini_yangilash(item["kategoriya"], item["model"], yangi_son)
        ombor_matni += f"📦 {item['model']} — omborda qoldi: {yangi_son} dona\n"

    buyurtmani_saqlash(
        buyurtma["chat_id"], buyurtma["itemlar"], buyurtma["ism"],
        buyurtma["telefon"], buyurtma["manzil"], buyurtma["jami_summa"],
        holat="tasdiqlangan", buyurtma_id=buyurtma_id
    )

    bot.send_message(
        buyurtma["chat_id"],
        "✅ Zakazingiz tasdiqlandi! Tez orada siz bilan bog'lanamiz."
    )

    tasdiqlagan_shaxs = call.from_user.first_name or call.from_user.username or "Nomaʼlum"

    try:
        bot.edit_message_text(
            call.message.text + f"\n\n✅ TASDIQLANDI ({tasdiqlagan_shaxs})",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            reply_markup=None
        )
    except Exception:
        pass

    bot.send_message(
        call.message.chat.id,
        f"✅ Zakaz #{buyurtma_id} tasdiqlandi va omborda ayirildi.\n"
        f"Tasdiqladi: {tasdiqlagan_shaxs}\n\n{ombor_matni}"
    )

    if NORTIX_SKLAD_GRUPPA_ID:
        pdf_yolu = buyurtma_pdf_yaratish(buyurtma_id, buyurtma, tasdiqlagan_shaxs)

        guruh_buyurtmalari[buyurtma_id] = {
            "manzil": buyurtma["manzil"],
            "ism": buyurtma["ism"],
        }

        guruh_izohi = (
            f"🛒 Yangi buyurtma!\n\n"
            f"{buyurtma_raqami(buyurtma_id)}\n"
            f"Buyurtma sanasi: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n\n"
            f"🏪 Do'kon: {buyurtma['manzil']}\n"
            f"👤 Buyurtma berdi: {buyurtma['ism']}\n\n"
            f"❗️ Iltimos buyurtma bilan tanishib chiqib buyurtma statusini belgilang 👇"
        )

        status_klaviatura = types.InlineKeyboardMarkup()
        status_klaviatura.add(
            types.InlineKeyboardButton("📦 Yuk ortilmoqda", callback_data=f"status_yuk:{buyurtma_id}")
        )
        status_klaviatura.add(
            types.InlineKeyboardButton("🚚 Yo'lga chiqdi", callback_data=f"status_yolga:{buyurtma_id}")
        )

        try:
            with open(pdf_yolu, "rb") as pdf_fayl:
                bot.send_document(NORTIX_SKLAD_GRUPPA_ID, pdf_fayl, caption=guruh_izohi, reply_markup=status_klaviatura)
        except Exception:
            bot.send_message(call.message.chat.id, "⚠️ Zakaz Nortix sklad guruhga yuborilmadi — NORTIX_SKLAD_GRUPPA_ID to'g'ri sozlanganini tekshiring.")
        finally:
            if os.path.exists(pdf_yolu):
                os.remove(pdf_yolu)


@bot.callback_query_handler(func=lambda call: call.data.startswith("status_yuk:"))
def status_yuk_belgilash(call):
    bot.answer_callback_query(call.id, "Belgilandi!")

    buyurtma_id = int(call.data.split("status_yuk:", 1)[1])
    yangi_izoh = call.message.caption + "\n\n✅ 📦 Yuk ortilmoqda - statusi belgilandi."

    yangi_klaviatura = types.InlineKeyboardMarkup()
    yangi_klaviatura.add(
        types.InlineKeyboardButton("🚚 Yo'lga chiqdi", callback_data=f"status_yolga:{buyurtma_id}")
    )

    try:
        bot.edit_message_caption(
            caption=yangi_izoh,
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            reply_markup=yangi_klaviatura
        )
    except Exception:
        pass

    if ZAKAZ_GRUPPA_ID:
        bot.send_message(
            ZAKAZ_GRUPPA_ID,
            f"📦 {buyurtma_raqami(buyurtma_id)} sonli buyurtma: Yuk ortilmoqda"
        )


@bot.callback_query_handler(func=lambda call: call.data.startswith("status_yolga:"))
def status_yolga_belgilash(call):
    bot.answer_callback_query(call.id, "Belgilandi!")

    buyurtma_id = int(call.data.split("status_yolga:", 1)[1])
    yangi_izoh = call.message.caption + "\n\n✅ 🚚 Yo'lga chiqdi — statusi belgilandi!"

    try:
        bot.edit_message_caption(
            caption=yangi_izoh,
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            reply_markup=None
        )
    except Exception:
        pass

    guruh_buyurtmalari.pop(buyurtma_id, None)

    if ZAKAZ_GRUPPA_ID:
        bot.send_message(
            ZAKAZ_GRUPPA_ID,
            f"🚚 {buyurtma_raqami(buyurtma_id)} sonli buyurtma: Yuk chiqib ketdi ✅"
        )


@bot.callback_query_handler(func=lambda call: call.data.startswith("admin_bekor:"))
def admin_zakazni_bekor_qilish(call):
    bot.answer_callback_query(call.id)
    if str(call.message.chat.id) != str(ZAKAZ_GRUPPA_ID):
        return

    buyurtma_id = int(call.data.split("admin_bekor:", 1)[1])
    buyurtma = kutilayotgan_buyurtmalar.pop(buyurtma_id, None)
    holatni_saqlash()
    if not buyurtma:
        bot.send_message(call.message.chat.id, "Bu buyurtma topilmadi — avval tasdiqlangan yoki bekor qilingan bo'lishi mumkin.")
        return

    buyurtmani_saqlash(
        buyurtma["chat_id"], buyurtma["itemlar"], buyurtma["ism"],
        buyurtma["telefon"], buyurtma["manzil"], buyurtma["jami_summa"],
        holat="bekor_qilingan", buyurtma_id=buyurtma_id
    )

    bot.send_message(
        buyurtma["chat_id"],
        "❌ Afsuski, buyurtmangiz bekor qilindi. Batafsil ma'lumot uchun biz bilan bog'laning."
    )

    bekor_qilgan_shaxs = call.from_user.first_name or call.from_user.username or "Nomaʼlum"

    try:
        bot.edit_message_text(
            call.message.text + f"\n\n❌ BEKOR QILINDI ({bekor_qilgan_shaxs})",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            reply_markup=None
        )
    except Exception:
        pass

    bot.send_message(
        call.message.chat.id,
        f"❌ Zakaz #{buyurtma_id} bekor qilindi. Ombor o'zgartirilmadi.\nBekor qildi: {bekor_qilgan_shaxs}"
    )


@bot.callback_query_handler(func=lambda call: call.data.startswith("buyurtma_ochir:"))
def buyurtmani_tarixdan_ochirish(call):
    bot.answer_callback_query(call.id)
    if not admin_mi(call.from_user.id):
        return

    idx = int(call.data.split("buyurtma_ochir:", 1)[1])

    try:
        with open(BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
            barcha = json.load(f)
    except Exception:
        barcha = []

    if 0 <= idx < len(barcha):
        ochirilgan = barcha.pop(idx)

        tiklangan_matn = ""
        if ochirilgan.get("holat", "tasdiqlangan") == "tasdiqlangan":
            try:
                malumotlar = ombor_malumotlarini_oqish()
            except Exception:
                malumotlar = {}

            for item in ochirilgan.get("itemlar", []):
                mahsulotlar_dict = {i[0]: i[1] for i in malumotlar.get(item["kategoriya"], [])}
                hozirgi_son = mahsulotlar_dict.get(item["model"], 0) or 0
                yangi_son = hozirgi_son + item["son"]
                ombor_sonini_yangilash(item["kategoriya"], item["model"], yangi_son)
                tiklangan_matn += f"📦 {item['model']} — omborga qaytarildi: +{item['son']} (endi {yangi_son} dona)\n"

        with open(BUYURTMALAR_FAYLI, "w", encoding="utf-8") as f:
            json.dump(barcha, f, ensure_ascii=False, indent=2)

        xabar = call.message.text + "\n\n🗑 TARIXDAN O'CHIRILDI"
        if tiklangan_matn:
            xabar += f"\n\n🔄 Ombor tiklandi:\n{tiklangan_matn}"

        try:
            bot.edit_message_text(
                xabar,
                chat_id=call.message.chat.id,
                message_id=call.message.message_id,
                reply_markup=None
            )
        except Exception:
            pass
    else:
        bot.send_message(call.message.chat.id, "Bu yozuv topilmadi — allaqachon o'chirilgan bo'lishi mumkin.")


@bot.message_handler(content_types=['document'])
def yangi_fayl_qabul_qilish(message):
    if not admin_mi(message.from_user.id):
        return

    fayl_nomi = message.document.file_name or ""
    if not fayl_nomi.lower().endswith(".xlsx"):
        bot.send_message(message.chat.id, "Faqat .xlsx fayl yuboring.")
        return

    try:
        fayl_info = bot.get_file(message.document.file_id)
        yuklab_olingan = bot.download_file(fayl_info.file_path)

        vaqtinchalik = EXCEL_FILE + ".yangi_tmp"
        with open(vaqtinchalik, "wb") as f:
            f.write(yuklab_olingan)

        # Yuklangan fayl haqiqatan ham to'g'ri Excel ekanini tekshiramiz —
        # buzilgan/noto'g'ri fayl bo'lsa, eski Baza.xlsx saqlanib qoladi.
        try:
            tekshiruv = openpyxl.load_workbook(vaqtinchalik, data_only=True)
            tekshiruv.close()
        except Exception:
            os.remove(vaqtinchalik)
            bot.send_message(message.chat.id, "❌ Fayl buzilgan yoki noto'g'ri formatda — eski ma'lumotlar o'zgarmadi.")
            return

        with EXCEL_LOCK:
            # Eski faylni zaxiraga olib qo'yamiz — noto'g'ri fayl kirib ketsa, tiklash mumkin bo'lsin.
            if os.path.exists(EXCEL_FILE):
                zaxira_nomi = f"Baza_zaxira_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                try:
                    import shutil
                    shutil.copy(EXCEL_FILE, zaxira_nomi)
                    log.info("Eski Excel fayl zaxiraga olindi: %s", zaxira_nomi)
                except Exception:
                    log.exception("Excel zaxira nusxa olishda xatolik")
            os.replace(vaqtinchalik, EXCEL_FILE)

        malumotlar = ombor_malumotlarini_oqish()
        jami_mahsulot = sum(len(v) for v in malumotlar.values())

        log.info("Admin %s tomonidan Excel fayl yangilandi: %s kategoriya, %s model.", message.from_user.id, len(malumotlar), jami_mahsulot)
        bot.send_message(
            message.chat.id,
            f"✅ Ombor fayli yangilandi!\n"
            f"Jami {len(malumotlar)} ta kategoriya, {jami_mahsulot} ta model topildi."
        )
    except Exception as e:
        log.exception("Yangi Excel fayl qabul qilishda xatolik")
        bot.send_message(message.chat.id, f"Xatolik yuz berdi: {e}")


@bot.message_handler(commands=['yangilash'])
def yangilash_boshlash(message):
    if not admin_mi(message.from_user.id):
        bot.send_message(message.chat.id, "Kechirasiz, bu buyruq faqat admin uchun.")
        return

    try:
        malumotlar = ombor_malumotlarini_oqish()
    except Exception as e:
        bot.send_message(message.chat.id, f"Xatolik yuz berdi: {e}")
        return

    if not malumotlar:
        bot.send_message(message.chat.id, "Omborda hozircha mahsulot yo'q.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for kategoriya in malumotlar.keys():
        keyboard.add(types.InlineKeyboardButton(
            text=kategoriya,
            callback_data=f"yangkat:{kategoriya}"
        ))

    bot.send_message(message.chat.id, "Qaysi kategoriyadagi mahsulotni yangilaysiz?", reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: call.data.startswith("yangkat:"))
def yangilash_kategoriya_tanlandi(call):
    bot.answer_callback_query(call.id)

    if not admin_mi(call.from_user.id):
        return

    kategoriya = call.data.split("yangkat:", 1)[1]
    malumotlar = ombor_malumotlarini_oqish()
    mahsulotlar = malumotlar.get(kategoriya, [])

    if not mahsulotlar:
        bot.send_message(call.message.chat.id, "Bu kategoriyada mahsulot topilmadi.")
        return

    keyboard = types.InlineKeyboardMarkup()
    for item in mahsulotlar:
        model, soni = item[0], item[1]
        keyboard.add(types.InlineKeyboardButton(
            text=f"{model} (hozir: {soni} dona)",
            callback_data=f"yangmodel:{kategoriya}|{model}"
        ))

    bot.send_message(call.message.chat.id, "Qaysi modelni yangilaysiz?", reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call: call.data.startswith("yangmodel:"))
def yangilash_model_tanlandi(call):
    bot.answer_callback_query(call.id)

    if not admin_mi(call.from_user.id):
        return

    kategoriya, model = call.data.split("yangmodel:", 1)[1].split("|", 1)

    yangilash_holati[call.from_user.id] = {
        "kategoriya": kategoriya,
        "model": model,
    }

    bot.send_message(
        call.message.chat.id,
        f"«{model}» uchun yangi sonni raqam bilan yuboring (masalan: 15):",
        reply_markup=orqaga_menyu_yaratish()
    )


@bot.message_handler(func=lambda message: message.from_user.id in yangilash_holati)
def yangi_sonni_qabul_qilish(message):
    holat = yangilash_holati.get(message.from_user.id)
    if not holat:
        return

    matn = message.text.strip()
    if not matn.isdigit():
        bot.send_message(message.chat.id, "Iltimos, faqat butun son yuboring (masalan: 15).")
        return

    yangi_soni = int(matn)
    kategoriya = holat["kategoriya"]
    model = holat["model"]

    muvaffaqiyatli = ombor_sonini_yangilash(kategoriya, model, yangi_soni)

    del yangilash_holati[message.from_user.id]

    if muvaffaqiyatli:
        bot.send_message(
            message.chat.id,
            f"✅ Yangilandi: {model} — endi {yangi_soni} dona.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )
    else:
        bot.send_message(
            message.chat.id,
            "Xatolik: bu model Excel'da topilmadi.",
            reply_markup=bosh_menyu_yaratish(message.from_user.id)
        )



if __name__ == "__main__":
    hisoblagichni_tiklash()
    threading.Thread(target=avtomatik_saqlash_oqimi, daemon=True).start()
    log.info("Menejer/Rahbar bot ishga tushdi...")
    while True:
        try:
            bot.infinity_polling(timeout=30, long_polling_timeout=30)
        except Exception:
            log.exception("Polling to'xtadi, 5 soniyadan keyin qayta urinamiz...")
            import time
            time.sleep(5)
@bot.message_handler(content_types=["text"], func=lambda m: m.from_user.id in dokon_holati and rahbar_mi(m.from_user.id))
def store_registration(message):
    uid = message.from_user.id
    st = dokon_holati.get(uid)
    if not st: return
    text = (message.text or "").strip()
    if text in ("🏠 Bosh menyu","🔙 Orqaga"):
        dokon_holati.pop(uid,None); bot.send_message(message.chat.id,"🏠",reply_markup=rahbar_menu()); return
    if st["bosqich"] == "nomi":
        if len(text)<2: bot.send_message(message.chat.id,"Do'kon nomini kiriting."); return
        st["nomi"]=text; st["bosqich"]="telefon"
        bot.send_message(message.chat.id,"📱 Do'kon telefon raqamini kiriting:")
    elif st["bosqich"] == "telefon":
        st["telefon"]=text; st["bosqich"]="manzil"
        bot.send_message(message.chat.id,"📍 Do'kon manzilini kiriting:")
    elif st["bosqich"] == "manzil":
        st["manzil"]=text
        sid = "DOK-" + str(len(dokonlar)+1).zfill(4)
        while sid in dokonlar:
            sid = "DOK-" + str(int(sid.split("-")[1])+1).zfill(4)
        dokonlar[sid] = {
            "id":sid, "nomi":st["nomi"], "telefon":st["telefon"],
            "manzil":st["manzil"], "manager_id":None,
            "sana":datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        json_saqlash(DOKONLAR_FAYLI,dokonlar)
        dokon_holati.pop(uid,None)
        bot.send_message(message.chat.id,f"✅ Do'kon qo'shildi: {st['nomi']}",reply_markup=rahbar_menu())

# ---------------- SAVDO KIRITISH ----------------
@bot.message_handler(func=lambda m: m.text == "🛒 Savdo kiritish")
def sale_start(message):
    uid=message.from_user.id
    if not menejer_tasdiqlangan(uid):
        bot.send_message(message.chat.id,"❌ Siz tasdiqlangan menejer emassiz."); return
    ids=my_store_ids(uid)
    if not ids:
        bot.send_message(message.chat.id,"❌ Avval sizga do'kon biriktirilishi kerak."); return
    kb=types.InlineKeyboardMarkup()
    for sid in ids:
        kb.add(types.InlineKeyboardButton(dokonlar[sid]["nomi"],callback_data=f"sale_store:{sid}"))
    bot.send_message(message.chat.id,"🏪 Savdo qaysi do'kon uchun?",reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("sale_store:"))
def sale_store(call):
    uid=call.from_user.id; sid=call.data.split(":",1)[1]
    if sid not in my_store_ids(uid):
        bot.answer_callback_query(call.id,"Bu do'kon sizga biriktirilmagan.",show_alert=True); return
    savdo_holati[uid]={"bosqich":"kategoriya","dokon_id":sid}
    bot.answer_callback_query(call.id)
    try: data=ombor_malumotlarini_oqish()
    except Exception: data={}
    kb=types.InlineKeyboardMarkup()
    for k in data: kb.add(types.InlineKeyboardButton(k,callback_data=f"sale_cat:{k}"))
    bot.send_message(call.message.chat.id,"📦 Kategoriyani tanlang:",reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("sale_cat:"))
def sale_category(call):
    uid=call.from_user.id; st=savdo_holati.get(uid)
    if not st: return
    cat=call.data.split(":",1)[1]; st["kategoriya"]=cat; st["bosqich"]="model"
    data=ombor_malumotlarini_oqish()
    kb=types.InlineKeyboardMarkup()
    for item in data.get(cat,[]):
        kb.add(types.InlineKeyboardButton(f"{item[0]} | {item[1]} dona",callback_data=f"sale_model:{cat}|{item[0]}"))
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,"Modelni tanlang:",reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("sale_model:"))
def sale_model(call):
    uid=call.from_user.id; st=savdo_holati.get(uid)
    if not st: return
    cat,model=call.data.split(":",1)[1].split("|",1)
    st.update({"kategoriya":cat,"model":model,"bosqich":"son"})
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,"🔢 Nechta dona sotildi?")

@bot.message_handler(func=lambda m: m.from_user.id in savdo_holati and savdo_holati[m.from_user.id].get("bosqich")=="son")
def sale_quantity(message):
    uid=message.from_user.id; st=savdo_holati[uid]
    if not (message.text or "").isdigit() or int(message.text)<=0:
        bot.send_message(message.chat.id,"Musbat son kiriting."); return
    st["son"]=int(message.text); st["bosqich"]="narx"
    data=ombor_malumotlarini_oqish()
    item=next((x for x in data.get(st["kategoriya"],[]) if x[0]==st["model"]),None)
    default_price=item[2] if item else 0
    bot.send_message(message.chat.id,f"💵 Sotuv narxini kiriting.\nTavsiya narx: ${default_price:,.0f}")

@bot.message_handler(func=lambda m: m.from_user.id in savdo_holati and savdo_holati[m.from_user.id].get("bosqich")=="narx")
def sale_price(message):
    uid=message.from_user.id; st=savdo_holati[uid]
    raw=(message.text or "").replace(" ","").replace(",","")
    try: price=float(raw)
    except:
        bot.send_message(message.chat.id,"Narxni raqam bilan kiriting."); return
    item={
        "id":len(savdolar)+1,
        "sana":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "manager_id":uid,
        "dokon_id":st["dokon_id"],
        "kategoriya":st["kategoriya"],
        "model":st["model"],
        "son":st["son"],
        "narx":price,
        "summa":st["son"]*price
    }
    savdolar.append(item); json_saqlash(SAVDOLAR_FAYLI,savdolar)
    savdo_holati.pop(uid,None)
    bot.send_message(message.chat.id,
        f"✅ Savdo saqlandi!\n\n🏪 {dokonlar[item['dokon_id']]['nomi']}\n"
        f"📦 {item['model']} — {item['son']} dona\n💰 ${item['summa']:,.0f}",
        reply_markup=menejer_menu())

@bot.message_handler(func=lambda m: m.text == "📊 Mening savdom")
def my_sales(message):
    uid=message.from_user.id
    if not menejer_tasdiqlangan(uid): return
    rows=[x for x in savdolar if str(x.get("manager_id"))==str(uid)]
    total=sum(x.get("summa",0) for x in rows)
    qty=sum(x.get("son",0) for x in rows)
    bot.send_message(message.chat.id,f"📊 <b>Mening savdom</b>\n\n📦 Mahsulot: {qty} dona\n💰 Jami savdo: ${total:,.0f}\n🧾 Savdo soni: {len(rows)}",parse_mode="HTML")

def joriy_oy():
    return datetime.now().strftime("%Y-%m")

def menejer_oylik_savdosi(uid, oy=None):
    oy = oy or joriy_oy()
    total = 0
    for x in savdolar:
        if str(x.get("manager_id")) != str(uid):
            continue
        sana = str(x.get("sana", ""))
        if sana[:7] == oy:
            total += float(x.get("summa", 0) or 0)
    return total

def plan_olish(uid, oy=None):
    oy = oy or joriy_oy()
    return float(planlar.get(str(uid), {}).get(oy, 0) or 0)

@bot.message_handler(func=lambda m: m.text == "🎯 Mening planim")
def my_plan(message):
    uid=message.from_user.id
    if not menejer_tasdiqlangan(uid):
        bot.send_message(message.chat.id,"❌ Siz tasdiqlangan menejer emassiz."); return
    oy=joriy_oy()
    plan=plan_olish(uid,oy)
    savdo=menejer_oylik_savdosi(uid,oy)
    foiz=(savdo/plan*100) if plan>0 else 0
    qolgan=max(plan-savdo,0)
    bot.send_message(message.chat.id,
        f"🎯 <b>Joriy oy plani — {oy}</b>\n\n"
        f"🎯 Plan: ${plan:,.0f}\n"
        f"📊 Savdo: ${savdo:,.0f}\n"
        f"📈 Bajarilish: {foiz:.1f}%\n"
        f"💰 Qolgan: ${qolgan:,.0f}",
        parse_mode="HTML")


# ============================================================
# 5-BOSQICH: QARZDORLIK TIZIMI
# Har bir do'kon bo'yicha qarz va to'lovlar alohida tarix sifatida saqlanadi.
# Menejer faqat o'ziga biriktirilgan do'konlarni ko'radi.
# Rahbar barcha do'konlar va menejerlar bo'yicha umumiy qarzdorlikni ko'radi.
# ============================================================

QARZDORLIK_FAYLI = "qarzdorlik.json"
qarzdorlik = json_yukla(QARZDORLIK_FAYLI, [])
qarz_holati = {}

def qarz_saqlash():
    json_saqlash(QARZDORLIK_FAYLI, qarzdorlik)

def qarz_balansi(store_id):
    """Do'konning joriy qarzi = qarz yozuvlari - to'lovlar."""
    balans = 0.0
    for x in qarzdorlik:
        if str(x.get("store_id")) != str(store_id):
            continue
        summa = float(x.get("summa", 0) or 0)
        if x.get("tur") == "qarz":
            balans += summa
        elif x.get("tur") == "tolov":
            balans -= summa
    return balans

def qarz_qoshish(store_id, manager_id, summa, tur, izoh=""):
    qarzdorlik.append({
        "id": len(qarzdorlik) + 1,
        "store_id": str(store_id),
        "manager_id": int(manager_id),
        "tur": tur,
        "summa": float(summa),
        "izoh": izoh,
        "sana": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })
    qarz_saqlash()

def qarz_store_nomi(store_id):
    d = dokonlar.get(str(store_id), {})
    return d.get("nomi", f"Do'kon #{store_id}")

def qarz_manager_store_ids(uid):
    return my_store_ids(uid)

def qarz_menu_keyboard(store_id):
    kb = types.InlineKeyboardMarkup()
    kb.add(
        types.InlineKeyboardButton("➕ Qarz qo'shish", callback_data=f"qarz_add:{store_id}"),
        types.InlineKeyboardButton("💵 To'lov kiritish", callback_data=f"tolov_add:{store_id}")
    )
    kb.add(types.InlineKeyboardButton("📜 Tarix", callback_data=f"qarz_tarix:{store_id}"))
    return kb

def qarz_dokonlar_xabari(chat_id, uid, rahbar=False):
    ids = list(dokonlar.keys()) if rahbar else my_store_ids(uid)
    if not ids:
        bot.send_message(chat_id, "🏪 Hozircha do'konlar mavjud emas.")
        return

    jami = 0
    matn = "💰 <b>Qarzdorlik</b>\n\n"
    for sid in ids:
        balans = qarz_balansi(sid)
        jami += balans
        d = dokonlar.get(str(sid), {})
        nomi = d.get("nomi", f"Do'kon #{sid}")
        belgi = "🔴" if balans > 0 else ("🟢" if balans == 0 else "🔵")
        matn += f"{belgi} <b>{nomi}</b>\n   Qarzdorlik: ${balans:,.0f}\n\n"

    matn += f"<b>Jami: ${jami:,.0f}</b>"
    bot.send_message(chat_id, matn, parse_mode="HTML")

    # Menejer uchun do'konni tanlash; rahbar uchun ham barcha do'konlar.
    kb = types.InlineKeyboardMarkup()
    for sid in ids:
        d = dokonlar.get(str(sid), {})
        kb.add(types.InlineKeyboardButton(
            f"{d.get('nomi') or ('Dokon #' + str(sid))} — ${qarz_balansi(sid):,.0f}",
            callback_data=f"qarz_store:{sid}"
        ))
    bot.send_message(chat_id, "Do'konni tanlang:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "💰 Qarzdorlik" and menejer_tasdiqlangan(m.from_user.id))
def my_debt(message):
    qarz_dokonlar_xabari(message.chat.id, message.from_user.id, rahbar=False)

@bot.message_handler(func=lambda m: m.text == "💰 Qarzdorlik" and rahbar_mi(m.from_user.id))
def debt_admin(message):
    qarz_dokonlar_xabari(message.chat.id, message.from_user.id, rahbar=True)

@bot.callback_query_handler(func=lambda c: c.data.startswith("qarz_store:"))
def qarz_store_selected(call):
    uid = call.from_user.id
    sid = call.data.split(":", 1)[1]
    if rahbar_mi(uid):
        ruxsat = sid in dokonlar
    else:
        ruxsat = menejer_tasdiqlangan(uid) and sid in my_store_ids(uid)
    if not ruxsat:
        bot.answer_callback_query(call.id, "Bu do'konga ruxsatingiz yo'q.", show_alert=True)
        return
    bot.answer_callback_query(call.id)
    nomi = qarz_store_nomi(sid)
    balans = qarz_balansi(sid)
    bot.send_message(
        call.message.chat.id,
        f"🏪 <b>{nomi}</b>\n\n💰 Joriy qarz: <b>${balans:,.0f}</b>",
        parse_mode="HTML",
        reply_markup=qarz_menu_keyboard(sid)
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("qarz_add:"))
def qarz_add_start(call):
    sid = call.data.split(":", 1)[1]
    uid = call.from_user.id
    if not ((rahbar_mi(uid) and sid in dokonlar) or
            (menejer_tasdiqlangan(uid) and sid in my_store_ids(uid))):
        bot.answer_callback_query(call.id, "Ruxsat yo'q.", show_alert=True)
        return
    qarz_holati[uid] = {"bosqich": "summa", "store_id": sid, "tur": "qarz"}
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,
        f"➕ <b>{qarz_store_nomi(sid)}</b>\n\n"
        "Qarz summasini kiriting.\nMasalan: 15000000",
        parse_mode="HTML", reply_markup=orqaga_menyu_yaratish())

@bot.callback_query_handler(func=lambda c: c.data.startswith("tolov_add:"))
def tolov_add_start(call):
    sid = call.data.split(":", 1)[1]
    uid = call.from_user.id
    if not ((rahbar_mi(uid) and sid in dokonlar) or
            (menejer_tasdiqlangan(uid) and sid in my_store_ids(uid))):
        bot.answer_callback_query(call.id, "Ruxsat yo'q.", show_alert=True)
        return
    qarz_holati[uid] = {"bosqich": "summa", "store_id": sid, "tur": "tolov"}
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id,
        f"💵 <b>{qarz_store_nomi(sid)}</b>\n\n"
        "To'lov summasini kiriting.\nMasalan: 5000000",
        parse_mode="HTML", reply_markup=orqaga_menyu_yaratish())

@bot.callback_query_handler(func=lambda c: c.data.startswith("qarz_tarix:"))
def qarz_history(call):
    sid = call.data.split(":", 1)[1]
    uid = call.from_user.id
    if not ((rahbar_mi(uid) and sid in dokonlar) or
            (menejer_tasdiqlangan(uid) and sid in my_store_ids(uid))):
        bot.answer_callback_query(call.id, "Ruxsat yo'q.", show_alert=True)
        return
    bot.answer_callback_query(call.id)
    yozuvlar = [x for x in qarzdorlik if str(x.get("store_id")) == str(sid)]
    if not yozuvlar:
        bot.send_message(call.message.chat.id, "📜 Bu do'kon bo'yicha tarix hali yo'q.")
        return
    matn = f"📜 <b>{qarz_store_nomi(sid)} — qarz tarixi</b>\n\n"
    for x in reversed(yozuvlar[-30:]):
        belgi = "➕ Qarz" if x.get("tur") == "qarz" else "➖ To'lov"
        matn += f"{belgi}: ${float(x.get('summa',0)):,.0f}\n"
        if x.get("izoh"):
            matn += f"   📝 {x['izoh']}\n"
        matn += f"   🕐 {x.get('sana','')}\n\n"
    matn += f"💰 <b>Qoldiq qarz: ${qarz_balansi(sid):,.0f}</b>"
    bot.send_message(call.message.chat.id, matn, parse_mode="HTML")

@bot.message_handler(func=lambda m: m.from_user.id in qarz_holati)
def qarz_input(message):
    uid = message.from_user.id
    st = qarz_holati.get(uid)
    if not st:
        return
    if st.get("bosqich") == "izoh":
        izoh = (message.text or "").strip()
        if izoh == "-":
            izoh = ""
        qarz_qoshish(st["store_id"], uid, st["summa"], st["tur"], izoh)
        qarz_holati.pop(uid, None)
        tur_matni = "➕ Qarz" if st["tur"] == "qarz" else "➖ To'lov"
        bot.send_message(message.chat.id,
            f"✅ {tur_matni} saqlandi: ${st['summa']:,.0f}\n"
            f"🏪 {qarz_store_nomi(st['store_id'])}\n"
            f"💰 Yangi qoldiq: ${qarz_balansi(st['store_id']):,.0f}",
            reply_markup=bosh_menyu_yaratish(uid))
        return

    raw = (message.text or "").strip().replace(" ", "").replace(",", "").replace("$", "")
    try:
        summa = float(raw)
    except ValueError:
        bot.send_message(message.chat.id, "❌ Summani faqat raqam bilan kiriting. Masalan: 15000000")
        return
    if summa <= 0:
        bot.send_message(message.chat.id, "❌ Summa 0 dan katta bo'lishi kerak.")
        return

    sid = st["store_id"]
    # To'lov mavjud qarzdan oshib ketmasin.
    if st["tur"] == "tolov":
        balans = qarz_balansi(sid)
        if balans <= 0:
            bot.send_message(message.chat.id, "🟢 Bu do'konning hozir qarzi yo'q.")
            qarz_holati.pop(uid, None)
            return
        if summa > balans:
            bot.send_message(message.chat.id,
                f"❌ To'lov qarzdan katta bo'lishi mumkin emas.\n"
                f"Joriy qarz: ${balans:,.0f}")
            return

    qarz_holati[uid] = {
        "bosqich": "izoh",
        "store_id": sid,
        "tur": st["tur"],
        "summa": summa
    }
    bot.send_message(message.chat.id,
        "📝 Izoh yozing yoki '-' yuboring:",
        reply_markup=orqaga_menyu_yaratish())


RASXOD_KATEGORIYALARI = [
    "🚗 Yo'l", "⛽ Benzin", "🍽 Ovqat", "🏨 Mehmonxona",
    "📱 Telefon", "📢 Reklama", "🚚 Yetkazib berish", "📦 Boshqa"
]

def rasxod_miqdori(manager_id=None, statuslar=None):
    natija = 0.0
    for x in rasxodlar:
        if manager_id is not None and str(x.get("manager_id")) != str(manager_id):
            continue
        if statuslar is not None and x.get("status") not in statuslar:
            continue
        natija += float(x.get("summa", 0) or 0)
    return natija

@bot.message_handler(func=lambda m: m.text == "💸 Rasxod" and menejer_tasdiqlangan(m.from_user.id))
def my_expense(message):
    uid = message.from_user.id
    shaxsiy = [x for x in rasxodlar if str(x.get("manager_id")) == str(uid)]
    jami = rasxod_miqqori = rasxod_miqdori(uid, {"kutilmoqda", "tasdiqlangan"})
    text = f"💸 <b>Mening rasxodlarim</b>\n\n💰 Jami: <b>{jami:,.0f} so'm</b>\n"
    if shaxsiy:
        text += "\n📜 Oxirgi rasxodlar:\n"
        for x in reversed(shaxsiy[-10:]):
            belgi = {"kutilmoqda":"🕓", "tasdiqlangan":"✅", "rad_etilgan":"❌"}.get(x.get("status"), "•")
            text += f"{belgi} {x.get('kategoriya','Boshqa')} — {float(x.get('summa',0)):,.0f} so'm\n"
    kb = types.InlineKeyboardMarkup()
    kb.add(types.InlineKeyboardButton("➕ Rasxod kiritish", callback_data="rasxod_add"))
    bot.send_message(message.chat.id, text, parse_mode="HTML", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data == "rasxod_add")
def rasxod_add_start(call):
    uid = call.from_user.id
    if not menejer_tasdiqlangan(uid):
        bot.answer_callback_query(call.id, "Faqat tasdiqlangan menejerlar uchun.", show_alert=True)
        return
    rasxod_holati[uid] = {"bosqich": "kategoriya"}
    bot.answer_callback_query(call.id)
    kb = types.InlineKeyboardMarkup(row_width=2)
    for i, nom in enumerate(RASXOD_KATEGORIYALARI):
        kb.add(types.InlineKeyboardButton(nom, callback_data=f"rasxod_kat:{i}"))
    bot.send_message(call.message.chat.id, "💸 Rasxod kategoriyasini tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("rasxod_kat:"))
def rasxod_category(call):
    uid = call.from_user.id
    if not menejer_tasdiqlangan(uid) or uid not in rasxod_holati:
        bot.answer_callback_query(call.id, "Jarayon topilmadi.", show_alert=True)
        return
    idx = int(call.data.split(":",1)[1])
    if not 0 <= idx < len(RASXOD_KATEGORIYALARI):
        return
    rasxod_holati[uid] = {"bosqich": "summa", "kategoriya": RASXOD_KATEGORIYALARI[idx]}
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id, "💰 Rasxod summasini so'mda kiriting.\nMasalan: 150000", reply_markup=orqaga_menyu_yaratish())

@bot.message_handler(func=lambda m: m.from_user.id in rasxod_holati)
def rasxod_input(message):
    uid = message.from_user.id
    st = rasxod_holati.get(uid)
    if not st or not menejer_tasdiqlangan(uid):
        return
    raw = (message.text or "").strip().replace(" ", "").replace(",", "").replace(".", "")
    if st["bosqich"] == "summa":
        try:
            summa = float(raw)
        except ValueError:
            bot.send_message(message.chat.id, "❌ Summani faqat raqam bilan kiriting. Masalan: 150000")
            return
        if summa <= 0:
            bot.send_message(message.chat.id, "❌ Summa 0 dan katta bo'lishi kerak.")
            return
        st["summa"] = summa
        st["bosqich"] = "izoh"
        bot.send_message(message.chat.id, "📝 Izoh yozing yoki '-' yuboring:", reply_markup=orqaga_menyu_yaratish())
        return
    izoh = (message.text or "").strip()
    if izoh == "-":
        izoh = ""
    m = menejer_ol(uid)
    rasxodlar.append({
        "id": len(rasxodlar) + 1,
        "manager_id": uid,
        "manager_name": m.get("ism", "") if m else "",
        "kategoriya": st["kategoriya"],
        "summa": st["summa"],
        "izoh": izoh,
        "status": "kutilmoqda",
        "sana": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })
    json_saqlash(RASXODLAR_FAYLI, rasxodlar)
    rasxod_holati.pop(uid, None)
    # Rahbarlarga tasdiqlash uchun yuboramiz.
    yangi = rasxodlar[-1]
    kb = types.InlineKeyboardMarkup()
    kb.add(types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"rasxod_ok:{yangi['id']}"),
           types.InlineKeyboardButton("❌ Rad etish", callback_data=f"rasxod_rad:{yangi['id']}"))
    xabar = (f"🆕 <b>Yangi rasxod</b>\n\n👤 {yangi['manager_name']}\n"
             f"💸 {yangi['kategoriya']}\n💰 {yangi['summa']:,.0f} so'm\n"
             f"📝 {yangi['izoh'] or '-'}\n🕐 {yangi['sana']}")
    for rid in ADMIN_IDLAR:
        try:
            bot.send_message(rid, xabar, parse_mode="HTML", reply_markup=kb)
        except Exception:
            log.exception("Rasxod rahbarga yuborilmadi: %s", rid)
    bot.send_message(message.chat.id, "✅ Rasxod saqlandi va rahbar tasdig'iga yuborildi.", reply_markup=menejer_menu())

@bot.callback_query_handler(func=lambda c: c.data.startswith("rasxod_ok:"))
def rasxod_approve(call):
    if not rahbar_mi(call.from_user.id):
        bot.answer_callback_query(call.id, "Ruxsat yo'q", show_alert=True); return
    rid = int(call.data.split(":",1)[1])
    x = next((z for z in rasxodlar if z.get("id") == rid), None)
    if not x or x.get("status") != "kutilmoqda":
        bot.answer_callback_query(call.id, "Bu rasxod allaqachon ko'rib chiqilgan.", show_alert=True); return
    x["status"] = "tasdiqlangan"
    x["tasdiqlagan"] = call.from_user.id
    json_saqlash(RASXODLAR_FAYLI, rasxodlar)
    bot.answer_callback_query(call.id, "Tasdiqlandi")
    try: bot.edit_message_text(call.message.text + "\n\n✅ TASDIQLANDI", call.message.chat.id, call.message.message_id)
    except Exception: pass
    try: bot.send_message(int(x["manager_id"]), f"✅ Rasxodingiz tasdiqlandi: {float(x['summa']):,.0f} so'm")
    except Exception: pass

@bot.callback_query_handler(func=lambda c: c.data.startswith("rasxod_rad:"))
def rasxod_reject(call):
    if not rahbar_mi(call.from_user.id):
        bot.answer_callback_query(call.id, "Ruxsat yo'q", show_alert=True); return
    rid = int(call.data.split(":",1)[1])
    x = next((z for z in rasxodlar if z.get("id") == rid), None)
    if not x or x.get("status") != "kutilmoqda":
        bot.answer_callback_query(call.id, "Bu rasxod allaqachon ko'rib chiqilgan.", show_alert=True); return
    x["status"] = "rad_etilgan"
    x["rad_etgan"] = call.from_user.id
    json_saqlash(RASXODLAR_FAYLI, rasxodlar)
    bot.answer_callback_query(call.id, "Rad etildi")
    try: bot.edit_message_text(call.message.text + "\n\n❌ RAD ETILDI", call.message.chat.id, call.message.message_id)
    except Exception: pass
    try: bot.send_message(int(x["manager_id"]), "❌ Rasxodingiz rahbar tomonidan rad etildi.")
    except Exception: pass

@bot.message_handler(func=lambda m: m.text == "📦 Buyurtmalar")
def my_orders(message):
    bot.send_message(message.chat.id,"📦 Buyurtmalar moduli keyingi bosqichda ulanadi.")

# ---------------- RAHBAR STATISTIKASI ----------------
@bot.message_handler(func=lambda m: m.text == "👨‍💼 Menejerlar" and rahbar_mi(m.from_user.id))
def managers_list(message):
    text="👨‍💼 <b>Menejerlar</b>\n\n"
    if not menedjerlar: text+="Hozircha menejer yo'q."
    else:
        for uid,m in menedjerlar.items():
            text += f"• {m.get('ism')} — {m.get('status')}\n"
    bot.send_message(message.chat.id,text,parse_mode="HTML")

@bot.message_handler(func=lambda m: m.text == "📊 Umumiy savdo" and rahbar_mi(m.from_user.id))
def total_sales(message):
    total=sum(x.get("summa",0) for x in savdolar)
    qty=sum(x.get("son",0) for x in savdolar)
    bot.send_message(message.chat.id,f"📊 <b>Umumiy savdo</b>\n\n📦 {qty} dona\n💰 ${total:,.0f}",parse_mode="HTML")

@bot.message_handler(func=lambda m: m.text == "🏆 Menejerlar reytingi" and rahbar_mi(m.from_user.id))
def manager_rating(message):
    oy=joriy_oy()
    stats=[]
    for uid,m in menedjerlar.items():
        if m.get("status")!="tasdiqlangan": continue
        plan=plan_olish(uid,oy)
        savdo=menejer_oylik_savdosi(uid,oy)
        foiz=(savdo/plan*100) if plan>0 else 0
        stats.append((foiz,savdo,plan,m.get("ism","")))
    stats.sort(key=lambda z:(z[0],z[1]),reverse=True)
    text=f"🏆 <b>Menejerlar reytingi — {oy}</b>\n\n"
    for i,(foiz,savdo,plan,name) in enumerate(stats,1):
        text+=f"{i}. <b>{name}</b> — {foiz:.1f}%\n   📊 ${savdo:,.0f} / ${plan:,.0f}\n"
    bot.send_message(message.chat.id,text if stats else "Hozircha tasdiqlangan menejer yo'q.",parse_mode="HTML")

@bot.message_handler(func=lambda m: m.text == "🎯 Planlar" and rahbar_mi(m.from_user.id))
def plans_admin(message):
    kb=types.InlineKeyboardMarkup()
    for uid,m in menedjerlar.items():
        if m.get("status")=="tasdiqlangan":
            kb.add(types.InlineKeyboardButton(m.get("ism","Noma'lum"),callback_data=f"plan_mgr:{uid}"))
    bot.send_message(message.chat.id,"🎯 Qaysi menejerga oylik plan berasiz?",reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("plan_mgr:"))
def plan_manager_selected(call):
    if not rahbar_mi(call.from_user.id): return
    uid=call.data.split(":",1)[1]
    if not menejer_tasdiqlangan(uid):
        bot.answer_callback_query(call.id,"Menejer topilmadi",show_alert=True); return
    call.message.chat.id
    dokon_holati[call.from_user.id]={"bosqich":"plan","manager_id":int(uid)}
    bot.answer_callback_query(call.id)
    m=menejer_ol(uid)
    bot.send_message(call.message.chat.id,
        f"🎯 <b>{m.get('ism')}</b> uchun {joriy_oy()} oy planini kiriting.\n\n"
        f"Masalan: 500000000",
        parse_mode="HTML",reply_markup=orqaga_menyu_yaratish())



# ============================================================
# 7-BOSQICH: RAHBAR DASHBOARD VA KENGAYTIRILGAN ANALITIKA
# ============================================================

def _float(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0

def _dokon_nomi(dokon_id):
    d = dokonlar.get(str(dokon_id)) or dokonlar.get(dokon_id) or {}
    return d.get("nomi", f"Do'kon #{dokon_id}")

def _menejer_ismi(uid):
    m = menejer_ol(uid) or menejer_ol(str(uid))
    return (m or {}).get("ism", f"Menejer #{uid}")

def rahbar_oylik_savdo_statistikasi(oy=None):
    oy = oy or joriy_oy()
    return [x for x in savdolar if str(x.get("sana", ""))[:7] == oy]

def rahbar_dashboard_matni(oy=None):
    oy = oy or joriy_oy()
    rows = rahbar_oylik_savdo_statistikasi(oy)
    jami = sum(_float(x.get("summa")) for x in rows)
    dona = sum(int(_float(x.get("son"))) for x in rows)

    mgr, kat, brand, model, store = {}, {}, {}, {}, {}
    excel_data = None

    for x in rows:
        uid = str(x.get("manager_id"))
        mgr[uid] = mgr.get(uid, 0) + _float(x.get("summa"))

        k = str(x.get("kategoriya") or "Noma'lum")
        kat[k] = kat.get(k, 0) + _float(x.get("summa"))

        b = str(x.get("brend") or "").strip()
        if not b:
            b = "Noma'lum"
            try:
                if excel_data is None:
                    excel_data = ombor_malumotlarini_oqish()
                for item in excel_data.get(k, []):
                    if item[0] == x.get("model"):
                        b = item[4] or "Noma'lum"
                        break
            except Exception:
                pass
        brand[b] = brand.get(b, 0) + _float(x.get("summa"))

        md = str(x.get("model") or "Noma'lum")
        model[md] = model.get(md, 0) + _float(x.get("summa"))

        did = str(x.get("dokon_id"))
        store[did] = store.get(did, 0) + _float(x.get("summa"))

    matn = (
        f"📊 <b>RAHBAR DASHBOARD — {oy}</b>\n\n"
        f"💰 Jami savdo: <b>${jami:,.0f}</b>\n"
        f"📦 Jami mahsulot: <b>{dona:,} dona</b>\n"
        f"🧾 Savdo soni: <b>{len(rows)}</b>\n\n"
        f"👨‍💼 <b>Menejerlar:</b>\n"
    )

    for uid, summa in sorted(mgr.items(), key=lambda z: z[1], reverse=True):
        matn += f"• {_menejer_ismi(uid)} — ${summa:,.0f}\n"

    matn += "\n🏷 <b>Kategoriyalar:</b>\n"
    for k, summa in sorted(kat.items(), key=lambda z: z[1], reverse=True)[:10]:
        matn += f"• {k} — ${summa:,.0f}\n"

    matn += "\n🔵 <b>Brendlar:</b>\n"
    for b, summa in sorted(brand.items(), key=lambda z: z[1], reverse=True)[:10]:
        matn += f"• {b} — ${summa:,.0f}\n"

    matn += "\n📦 <b>Top modellar:</b>\n"
    for md, summa in sorted(model.items(), key=lambda z: z[1], reverse=True)[:10]:
        matn += f"• {md} — ${summa:,.0f}\n"

    matn += "\n🏪 <b>Top do'konlar:</b>\n"
    for did, summa in sorted(store.items(), key=lambda z: z[1], reverse=True)[:10]:
        matn += f"• {_dokon_nomi(did)} — ${summa:,.0f}\n"

    return matn

@bot.message_handler(func=lambda m: m.text == "📈 Savdo analitikasi" and rahbar_mi(m.from_user.id))
def rahbar_dashboard(message):
    bot.send_message(
        message.chat.id,
        rahbar_dashboard_matni(),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

@bot.message_handler(func=lambda m: m.text == "🏆 Menejerlar reytingi" and rahbar_mi(m.from_user.id))
def rahbar_menejer_reytingi_yangi(message):
    oy = joriy_oy()
    natija = []

    for uid, m in menedjerlar.items():
        if m.get("status") != "tasdiqlangan":
            continue
        plan = plan_olish(uid, oy)
        savdo = menejer_oylik_savdosi(uid, oy)
        foiz = (savdo / plan * 100) if plan > 0 else 0
        natija.append((foiz, savdo, plan, m.get("ism", "Noma'lum")))

    natija.sort(key=lambda x: (x[0], x[1]), reverse=True)

    if not natija:
        bot.send_message(message.chat.id, "Hozircha tasdiqlangan menejer yo'q.")
        return

    matn = f"🏆 <b>MENEJERLAR REYTINGI — {oy}</b>\n\n"
    for i, (foiz, savdo, plan, ism) in enumerate(natija, 1):
        belgi = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "▪️"
        matn += (
            f"{belgi} <b>{i}. {ism}</b>\n"
            f"   🎯 Plan: ${plan:,.0f}\n"
            f"   📊 Savdo: ${savdo:,.0f}\n"
            f"   📈 Bajarilish: <b>{foiz:.1f}%</b>\n\n"
        )

    bot.send_message(message.chat.id, matn, parse_mode="HTML", reply_markup=rahbar_menu())

@bot.message_handler(commands=["dashboard"])
def dashboard_command(message):
    if not rahbar_mi(message.from_user.id):
        bot.send_message(message.chat.id, "❌ Bu bo'lim faqat rahbarlar uchun.")
        return
    bot.send_message(
        message.chat.id,
        rahbar_dashboard_matni(),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

# ============================================================
# 8-BOSQICH: MENEJER BUYURTMALARI
# Faqat tasdiqlangan menejerlar buyurtma yaratadi.
# Buyurtma do'kon + menejer + mahsulot + miqdor bilan saqlanadi.
# ============================================================

MENEJER_BUYURTMALAR_FAYLI = "menedjer_buyurtmalar.json"
menejer_buyurtma_holati = {}
menejer_buyurtma_savat = {}

def menejer_buyurtmalarini_yuklash():
    if not os.path.exists(MENEJER_BUYURTMALAR_FAYLI):
        return []
    try:
        with open(MENEJER_BUYURTMALAR_FAYLI, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        log.exception("Menejer buyurtmalarini yuklashda xatolik")
        return []

def menejer_buyurtmalarini_saqlash(data):
    tmp = MENEJER_BUYURTMALAR_FAYLI + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, MENEJER_BUYURTMALAR_FAYLI)

def menejer_buyurtma_raqami():
    data = menejer_buyurtmalarini_yuklash()
    eng_katta = 0
    for x in data:
        try:
            eng_katta = max(eng_katta, int(x.get("id", 0)))
        except Exception:
            pass
    return eng_katta + 1

def menejer_buyurtma_savatcha_matni(user_id):
    items = menejer_buyurtma_savat.get(user_id, [])
    if not items:
        return "🛒 Buyurtma savati bo'sh.", 0

    jami = 0
    matn = "🛒 <b>Buyurtma savati</b>\n\n"
    for i, item in enumerate(items, 1):
        summa = _float(item["narx"]) * int(item["son"])
        jami += summa
        matn += (
            f"{i}. {item['model']}\n"
            f"   {item['son']} dona × ${_float(item['narx']):,.0f} = ${summa:,.0f}\n"
        )
    matn += f"\n💰 <b>Jami: ${jami:,.0f}</b>"
    return matn, jami

@bot.message_handler(func=lambda m: m.text == "📦 Buyurtmalar" and menejer_mi(m.from_user.id))
def menejer_buyurtmalar_menu(message):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("➕ Buyurtma yaratish", "📋 Mening buyurtmalarim")
    kb.add("🏠 Bosh menyu")
    bot.send_message(message.chat.id, "📦 Buyurtmalar bo'limi:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "➕ Buyurtma yaratish" and menejer_mi(m.from_user.id))
def menejer_buyurtma_boshlash(message):
    stores = menejer_dokonlari(message.from_user.id)
    if not stores:
        bot.send_message(message.chat.id, "❌ Sizga hali do'kon biriktirilmagan.")
        return

    kb = types.InlineKeyboardMarkup()
    for did, d in stores:
        kb.add(types.InlineKeyboardButton(
            f"🏪 {d.get('nomi', 'Nomsiz do‘kon')}",
            callback_data=f"mbstore:{did}"
        ))
    bot.send_message(message.chat.id, "Buyurtma qaysi do'kon uchun?", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbstore:"))
def menejer_buyurtma_dokon_tanlash(call):
    bot.answer_callback_query(call.id)
    if not menejer_mi(call.from_user.id):
        return
    did = call.data.split(":", 1)[1]
    allowed = {str(x[0]) for x in menejer_dokonlari(call.from_user.id)}
    if str(did) not in allowed:
        bot.send_message(call.message.chat.id, "❌ Bu do'kon sizga biriktirilmagan.")
        return

    menejer_buyurtma_savat[call.from_user.id] = []
    menejer_buyurtma_holati[call.from_user.id] = {"dokon_id": did}

    data = ombor_malumotlarini_oqish()
    kb = types.InlineKeyboardMarkup()
    for kat in data.keys():
        kb.add(types.InlineKeyboardButton(kat, callback_data=f"mbkat:{kat}"))
    bot.send_message(call.message.chat.id, "📦 Kategoriyani tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbkat:"))
def menejer_buyurtma_kategoriya(call):
    bot.answer_callback_query(call.id)
    if not menejer_mi(call.from_user.id):
        return
    kat = call.data.split(":", 1)[1]
    data = ombor_malumotlarini_oqish()
    products = [x for x in data.get(kat, []) if _float(x[1]) > 0]

    if not products:
        bot.send_message(call.message.chat.id, "Bu kategoriyada mavjud mahsulot yo'q.")
        return

    kb = types.InlineKeyboardMarkup()
    for item in products[:50]:
        model, son, narx = item[0], item[1], item[2]
        kb.add(types.InlineKeyboardButton(
            f"{model} — ${_float(narx):,.0f} ({int(_float(son))} dona)",
            callback_data=f"mbprod:{kat}|{model}"
        ))
    bot.send_message(call.message.chat.id, "Modelni tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbprod:"))
def menejer_buyurtma_model(call):
    bot.answer_callback_query(call.id)
    if not menejer_mi(call.from_user.id):
        return
    kat, model = call.data.split(":", 1)[1].split("|", 1)
    data = ombor_malumotlarini_oqish()
    found = next((x for x in data.get(kat, []) if x[0] == model), None)
    if not found:
        bot.send_message(call.message.chat.id, "❌ Mahsulot topilmadi.")
        return

    menejer_buyurtma_holati[call.from_user.id]["bosqich"] = "son"
    menejer_buyurtma_holati[call.from_user.id]["kategoriya"] = kat
    menejer_buyurtma_holati[call.from_user.id]["model"] = model
    menejer_buyurtma_holati[call.from_user.id]["mavjud"] = int(_float(found[1]))
    menejer_buyurtma_holati[call.from_user.id]["narx"] = _float(found[2])

    bot.send_message(
        call.message.chat.id,
        f"📦 <b>{model}</b>\n"
        f"💰 Narx: ${_float(found[2]):,.0f}\n"
        f"📊 Omborda: {int(_float(found[1]))} dona\n\n"
        f"Nechta dona kerak?",
        parse_mode="HTML"
    )

@bot.message_handler(func=lambda m: m.from_user.id in menejer_buyurtma_holati)
def menejer_buyurtma_son_qabul(message):
    uid = message.from_user.id
    state = menejer_buyurtma_holati.get(uid)
    if not state or state.get("bosqich") != "son":
        return

    txt = (message.text or "").strip()
    if not txt.isdigit() or int(txt) <= 0:
        bot.send_message(message.chat.id, "❌ Musbat butun son kiriting. Masalan: 5")
        return

    son = int(txt)
    if son > state["mavjud"]:
        bot.send_message(
            message.chat.id,
            f"❌ Omborda faqat {state['mavjud']} dona bor."
        )
        return

    menejer_buyurtma_savat.setdefault(uid, []).append({
        "kategoriya": state["kategoriya"],
        "model": state["model"],
        "son": son,
        "narx": state["narx"],
    })
    state["bosqich"] = "yana"

    matn, jami = menejer_buyurtma_savatcha_matni(uid)
    kb = types.InlineKeyboardMarkup()
    kb.add(types.InlineKeyboardButton("➕ Yana mahsulot", callback_data="mb_yana"))
    kb.add(types.InlineKeyboardButton("✅ Buyurtmani yuborish", callback_data="mb_yuborish"))
    kb.add(types.InlineKeyboardButton("🗑 Bekor qilish", callback_data="mb_bekor"))
    bot.send_message(message.chat.id, matn, parse_mode="HTML", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data == "mb_yana")
def menejer_buyurtma_yana(call):
    bot.answer_callback_query(call.id)
    if not menejer_mi(call.from_user.id):
        return
    data = ombor_malumotlarini_oqish()
    kb = types.InlineKeyboardMarkup()
    for kat in data.keys():
        kb.add(types.InlineKeyboardButton(kat, callback_data=f"mbkat:{kat}"))
    bot.send_message(call.message.chat.id, "📦 Kategoriyani tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data == "mb_bekor")
def menejer_buyurtma_bekor(call):
    bot.answer_callback_query(call.id)
    menejer_buyurtma_savat.pop(call.from_user.id, None)
    menejer_buyurtma_holati.pop(call.from_user.id, None)
    bot.send_message(call.message.chat.id, "❌ Buyurtma bekor qilindi.", reply_markup=menejer_menu())

@bot.callback_query_handler(func=lambda c: c.data == "mb_yuborish")
def menejer_buyurtma_yuborish(call):
    bot.answer_callback_query(call.id)
    uid = call.from_user.id
    if not menejer_mi(uid):
        return

    state = menejer_buyurtma_holati.get(uid)
    items = menejer_buyurtma_savat.get(uid, [])
    if not state or not items:
        bot.send_message(call.message.chat.id, "❌ Buyurtma topilmadi.")
        return

    did = str(state["dokon_id"])
    allowed = {str(x[0]) for x in menejer_dokonlari(uid)}
    if did not in allowed:
        bot.send_message(call.message.chat.id, "❌ Do'kon sizga biriktirilmagan.")
        return

    jami = sum(_float(x["narx"]) * int(x["son"]) for x in items)
    bid = menejer_buyurtma_raqami()
    data = menejer_buyurtmalarini_yuklash()

    data.append({
        "id": bid,
        "manager_id": uid,
        "manager_ism": (menejer_ol(uid) or {}).get("ism", ""),
        "dokon_id": did,
        "dokon_nomi": _dokon_nomi(did),
        "itemlar": items,
        "jami_summa": jami,
        "sana": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": "yangi",
    })
    menejer_buyurtmalarini_saqlash(data)

    matn = (
        f"🆕 <b>Yangi menejer buyurtmasi #{bid}</b>\n\n"
        f"👨‍💼 Menejer: {(menejer_ol(uid) or {}).get('ism', '')}\n"
        f"🏪 Do'kon: {_dokon_nomi(did)}\n"
        f"💰 Jami: ${jami:,.0f}\n\n"
    )
    for x in items:
        matn += f"• {x['model']} — {x['son']} dona\n"

    bot.send_message(call.message.chat.id, "✅ Buyurtma rahbarga yuborildi.", reply_markup=menejer_menu())

    for rid in RAHBAR_IDS:
        try:
            kb = types.InlineKeyboardMarkup()
            kb.add(
                types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"mbok:{bid}"),
                types.InlineKeyboardButton("❌ Bekor qilish", callback_data=f"mbno:{bid}")
            )
            bot.send_message(rid, matn, parse_mode="HTML", reply_markup=kb)
        except Exception:
            log.exception("Rahbarga menejer buyurtmasini yuborishda xatolik")

    menejer_buyurtma_savat.pop(uid, None)
    menejer_buyurtma_holati.pop(uid, None)

def _menejer_buyurtmani_status(bid, status):
    data = menejer_buyurtmalarini_yuklash()
    for x in data:
        if int(x.get("id", 0)) == int(bid):
            x["status"] = status
            x["status_vaqti"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            menejer_buyurtmalarini_saqlash(data)
            return x
    return None

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbok:"))
def menejer_buyurtma_tasdiq(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return
    bid = int(call.data.split(":", 1)[1])
    x = _menejer_buyurtmani_status(bid, "tasdiqlandi")
    if not x:
        bot.send_message(call.message.chat.id, "❌ Buyurtma topilmadi.")
        return
    try:
        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    bot.send_message(
        int(x["manager_id"]),
        f"✅ Buyurtma #{bid} tasdiqlandi.\n🏪 {x['dokon_nomi']}\n💰 ${_float(x['jami_summa']):,.0f}"
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbno:"))
def menejer_buyurtma_rad(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return
    bid = int(call.data.split(":", 1)[1])
    x = _menejer_buyurtmani_status(bid, "bekor_qilindi")
    if not x:
        bot.send_message(call.message.chat.id, "❌ Buyurtma topilmadi.")
        return
    try:
        bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
    except Exception:
        pass
    bot.send_message(int(x["manager_id"]), f"❌ Buyurtma #{bid} bekor qilindi.")

@bot.message_handler(func=lambda m: m.text == "📋 Mening buyurtmalarim" and menejer_mi(m.from_user.id))
def mening_menejer_buyurtmalarim(message):
    data = menejer_buyurtmalarini_yuklash()
    mine = [x for x in data if str(x.get("manager_id")) == str(message.from_user.id)]
    if not mine:
        bot.send_message(message.chat.id, "Hozircha buyurtmalaringiz yo'q.")
        return

    statuslar = {
        "yangi": "🕐 Yangi",
        "tasdiqlandi": "✅ Tasdiqlandi",
        "bekor_qilindi": "❌ Bekor qilindi",
    }
    matn = "📋 <b>Mening buyurtmalarim</b>\n\n"
    for x in reversed(mine[-20:]):
        matn += (
            f"#{x['id']} — {statuslar.get(x.get('status'), x.get('status'))}\n"
            f"🏪 {x.get('dokon_nomi')}\n"
            f"💰 ${_float(x.get('jami_summa')):,.0f}\n"
            f"🕐 {x.get('sana')}\n\n"
        )
    bot.send_message(message.chat.id, matn, parse_mode="HTML", reply_markup=menejer_menu())

# ============================================================
# 9-BOSQICH: BUYURTMA + OMBOR NAZORATI
# Rahbar buyurtmani tasdiqlaganda ombor avtomatik kamaymaydi.
# Ombor faqat "yetkazildi" bosqichiga o'tganda kamaytiriladi.
# Shu bilan birga bir xil buyurtma ikki marta hisoblanishidan himoya bor.
# ============================================================

def menejer_buyurtma_top(bid):
    data = menejer_buyurtmalarini_yuklash()
    for x in data:
        try:
            if int(x.get("id", 0)) == int(bid):
                return x
        except Exception:
            pass
    return None

def ombordagi_model(kategoriya, model):
    data = ombor_malumotlarini_oqish()
    for item in data.get(kategoriya, []):
        if str(item[0]) == str(model):
            return item
    return None

def ombor_buyurtma_uchun_tekshir(order):
    xatolar = []
    for item in order.get("itemlar", []):
        row = ombor_malum_model = ombor_malumotlarini_oqish().get(item["kategoriya"], [])
        found = next((r for r in row if str(r[0]) == str(item["model"])), None)
        if not found:
            xatolar.append(f"❌ {item['model']} — omborda topilmadi")
            continue
        mavjud = int(_float(found[1]))
        kerak = int(item["son"])
        if kerak > mavjud:
            xatolar.append(
                f"❌ {item['model']} — kerak: {kerak}, mavjud: {mavjud}"
            )
    return xatolar

def ombor_soni_ayirish(order):
    """
    Buyurtma 'yetkazildi' bo'lganda Excel omboridan ayiradi.
    Natijada buyurtma ikki marta yetkazildi holatiga o'tkazilsa,
    qayta ayirish amalga oshirilmaydi.
    """
    if order.get("ombor_ayirildi"):
        return True, []

    data = ombor_malumotlarini_oqish()
    xatolar = []

    # Avval barcha pozitsiyalarni tekshiramiz.
    for item in order.get("itemlar", []):
        rows = data.get(item["kategoriya"], [])
        found = next((r for r in rows if str(r[0]) == str(item["model"])), None)
        if not found:
            xatolar.append(f"{item['model']} — topilmadi")
            continue
        mavjud = int(_float(found[1]))
        kerak = int(item["son"])
        if kerak > mavjud:
            xatolar.append(f"{item['model']} — {mavjud} dona mavjud, {kerak} dona kerak")

    if xatolar:
        return False, xatolar

    # Keyin ayiramiz.
    for item in order.get("itemlar", []):
        rows = data.get(item["kategoriya"], [])
        found = next((r for r in rows if str(r[0]) == str(item["model"])), None)
        found[1] = int(_float(found[1])) - int(item["son"])

    # Mavjud botning Excel saqlash funksiyasidan foydalanishga harakat.
    saqlangan = False
    try:
        if "ombor_malumotlarini_saqlash" in globals():
            ombor_malumotlarini_saqlash(data)
            saqlangan = True
    except Exception:
        log.exception("Omborni saqlashda xatolik")

    if not saqlangan:
        # Baza.xlsx to'g'ridan-to'g'ri yangilanadi.
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
            for row in range(2, ws.max_row + 1):
                kat = str(ws.cell(row, 2).value or "").strip()
                model = str(ws.cell(row, 3).value or "").strip()
                for item in order.get("itemlar", []):
                    if kat == str(item["kategoriya"]) and model == str(item["model"]):
                        eski = int(_float(ws.cell(row, 4).value))
                        ws.cell(row, 4).value = eski - int(item["son"])
            wb.save(EXCEL_FILE)
            saqlangan = True
        except Exception:
            log.exception("Baza.xlsx ni yangilashda xatolik")

    if not saqlangan:
        return False, ["Excel omborini saqlab bo'lmadi"]

    return True, []

def menejer_buyurtma_status_yangila(bid, yangi_status, rahbar_id=None):
    data = menejer_buyurtmalarini_yuklash()
    for order in data:
        if int(order.get("id", 0)) != int(bid):
            continue

        eski = order.get("status", "yangi")

        # Yetkazilgan buyurtma qayta o'zgartirilmaydi.
        if eski == "yetkazildi":
            return order, False, "Bu buyurtma allaqachon yetkazilgan."

        if yangi_status == "yetkazildi":
            ok, errors = ombor_soni_ayirish(order)
            if not ok:
                return order, False, "Ombor yetarli emas:\n" + "\n".join(errors)
            order["ombor_ayirildi"] = True

        order["status"] = yangi_status
        order["status_vaqti"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if rahbar_id:
            order["status_rahbar_id"] = rahbar_id

        menejer_buyurtmalarini_saqlash(data)
        return order, True, None

    return None, False, "Buyurtma topilmadi."

def menejer_buyurtma_status_kb(bid, status):
    kb = types.InlineKeyboardMarkup()
    if status == "yangi":
        kb.add(
            types.InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"mboq:{bid}"),
            types.InlineKeyboardButton("❌ Bekor qilish", callback_data=f"mbcan:{bid}")
        )
    elif status == "tasdiqlandi":
        kb.add(types.InlineKeyboardButton("📦 Tayyorlanmoqda", callback_data=f"mbprep:{bid}"))
    elif status == "tayyorlanmoqda":
        kb.add(types.InlineKeyboardButton("🚚 Yo'lda", callback_data=f"mbroad:{bid}"))
    elif status == "yolda":
        kb.add(types.InlineKeyboardButton("✅ Yetkazildi", callback_data=f"mbdone:{bid}"))
    return kb

def rahbar_buyurtma_matni(order):
    statuslar = {
        "yangi": "🕐 Yangi",
        "tasdiqlandi": "✅ Tasdiqlandi",
        "tayyorlanmoqda": "📦 Tayyorlanmoqda",
        "yolda": "🚚 Yo'lda",
        "yetkazildi": "🏁 Yetkazildi",
        "bekor_qilindi": "❌ Bekor qilindi",
    }
    matn = (
        f"📦 <b>Buyurtma #{order['id']}</b>\n"
        f"👨‍💼 Menejer: {order.get('manager_ism', '')}\n"
        f"🏪 Do'kon: {order.get('dokon_nomi', '')}\n"
        f"📅 Sana: {order.get('sana', '')}\n"
        f"📌 Status: {statuslar.get(order.get('status'), order.get('status'))}\n\n"
    )
    for item in order.get("itemlar", []):
        summa = _float(item["narx"]) * int(item["son"])
        matn += f"• {item['model']} — {item['son']} dona — ${summa:,.0f}\n"
    matn += f"\n💰 <b>Jami: ${_float(order.get('jami_summa')):,.0f}</b>"
    return matn

@bot.message_handler(func=lambda m: m.text == "📦 Buyurtmalar" and rahbar_mi(m.from_user.id))
def rahbar_buyurtmalar_menu(message):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("🆕 Yangi buyurtmalar", "📋 Barcha buyurtmalar")
    kb.add("🏠 Bosh menyu")
    bot.send_message(message.chat.id, "📦 Rahbar buyurtmalar bo'limi:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "🆕 Yangi buyurtmalar" and rahbar_mi(m.from_user.id))
def rahbar_yangi_buyurtmalar(message):
    data = menejer_buyurtmalarini_yuklash()
    yangi = [x for x in data if x.get("status") == "yangi"]
    if not yangi:
        bot.send_message(message.chat.id, "🕐 Yangi buyurtmalar yo'q.")
        return

    for order in reversed(yangi[-20:]):
        bot.send_message(
            message.chat.id,
            rahbar_buyurtma_matni(order),
            parse_mode="HTML",
            reply_markup=menejer_buyurtma_status_kb(order["id"], order["status"])
        )

@bot.message_handler(func=lambda m: m.text == "📋 Barcha buyurtmalar" and rahbar_mi(m.from_user.id))
def rahbar_barcha_buyurtmalar(message):
    data = menejer_buyurtmalarini_yuklash()
    if not data:
        bot.send_message(message.chat.id, "Hozircha buyurtmalar yo'q.")
        return

    for order in reversed(data[-30:]):
        bot.send_message(
            message.chat.id,
            rahbar_buyurtma_matni(order),
            parse_mode="HTML",
            reply_markup=menejer_buyurtma_status_kb(order["id"], order.get("status"))
        )

def _rahbar_order_status_callback(call, target_status):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    bid = int(call.data.split(":", 1)[1])
    order, changed, error = menejer_buyurtma_status_yangila(
        bid, target_status, call.from_user.id
    )
    if not order:
        bot.send_message(call.message.chat.id, "❌ Buyurtma topilmadi.")
        return

    if not changed:
        bot.send_message(call.message.chat.id, f"❌ {error}")
        return

    try:
        bot.edit_message_text(
            rahbar_buyurtma_matni(order),
            call.message.chat.id,
            call.message.message_id,
            parse_mode="HTML",
            reply_markup=menejer_buyurtma_status_kb(order["id"], order["status"])
        )
    except Exception:
        pass

    statuslar = {
        "tasdiqlandi": "✅ Buyurtmangiz tasdiqlandi.",
        "tayyorlanmoqda": "📦 Buyurtmangiz tayyorlanmoqda.",
        "yolda": "🚚 Buyurtmangiz yo'lga chiqdi.",
        "yetkazildi": "🏁 Buyurtmangiz yetkazildi.",
        "bekor_qilindi": "❌ Buyurtmangiz bekor qilindi.",
    }
    try:
        bot.send_message(
            int(order["manager_id"]),
            f"{statuslar.get(order['status'], 'Buyurtma statusi o‘zgardi.')}\n"
            f"📦 Buyurtma #{order['id']}\n"
            f"🏪 {order['dokon_nomi']}"
        )
    except Exception:
        log.exception("Menejerga buyurtma statusini yuborishda xatolik")

@bot.callback_query_handler(func=lambda c: c.data.startswith("mboq:"))
def order_tasdiqlash_callback(call):
    _rahbar_order_status_callback(call, "tasdiqlandi")

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbcan:"))
def order_bekor_callback(call):
    _rahbar_order_status_callback(call, "bekor_qilindi")

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbprep:"))
def order_tayyorlash_callback(call):
    _rahbar_order_status_callback(call, "tayyorlanmoqda")

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbroad:"))
def order_yolda_callback(call):
    _rahbar_order_status_callback(call, "yolda")

@bot.callback_query_handler(func=lambda c: c.data.startswith("mbdone:"))
def order_yetkazildi_callback(call):
    _rahbar_order_status_callback(call, "yetkazildi")

# ============================================================
# 10-BOSQICH: MENEJERLAR REYTINGI + KPI
# KPI: plan bajarilishi, savdo, buyurtmalar, qarzdorlik.
# Reyting asosiy mezoni — plan bajarilish foizi.
# ============================================================

def joriy_oy():
    return datetime.now().strftime("%Y-%m")

def menejer_savdosi_oy(uid, oy=None):
    oy = oy or joriy_oy()
    jami = 0
    soni = 0
    for x in savdolar:
        if str(x.get("manager_id")) != str(uid):
            continue
        sana = str(x.get("sana", ""))
        if sana.startswith(oy):
            jami += _float(x.get("total", x.get("summa", 0)))
            soni += int(_float(x.get("qty", x.get("son", 0))))
    return jami, soni

def menejer_buyurtma_statistikasi(uid, oy=None):
    oy = oy or joriy_oy()
    data = menejer_buyurtmalarini_yuklash()
    jami = 0
    soni = 0
    yetkazilgan = 0
    tasdiqlangan = 0

    for x in data:
        if str(x.get("manager_id")) != str(uid):
            continue
        if not str(x.get("sana", "")).startswith(oy):
            continue

        jami += _float(x.get("jami_summa", 0))
        soni += 1
        status = x.get("status")
        if status == "yetkazildi":
            yetkazilgan += 1
        if status in {"tasdiqlandi", "tayyorlanmoqda", "yolda", "yetkazildi"}:
            tasdiqlangan += 1

    return jami, soni, yetkazilgan, tasdiqlangan

def menejer_qarzi_jami(uid):
    jami = 0
    for did, d in dokonlar.items():
        if str(d.get("manager_id")) != str(uid):
            continue
        for x in qarzdorlik:
            if str(x.get("dokon_id")) != str(did):
                continue
            if x.get("tur") == "qarz":
                jami += _float(x.get("summa"))
            elif x.get("tur") == "tolov":
                jami -= _float(x.get("summa"))
    return max(jami, 0)

def menejer_oylik_plan(uid, oy=None):
    oy = oy or joriy_oy()
    key = f"{uid}_{oy}"
    try:
        return _float(planlar.get(key, 0))
    except Exception:
        return 0

def menejer_kpi(uid, oy=None):
    oy = oy or joriy_oy()
    plan = menejer_oylik_plan(uid, oy)
    savdo, qty = menejer_savdosi_oy(uid, oy)
    _, buyurtma_soni, yetkazilgan, tasdiqlangan = menejer_buyurtma_statistikasi(uid, oy)

    bajarilish = (savdo / plan * 100) if plan > 0 else 0

    return {
        "plan": plan,
        "savdo": savdo,
        "qty": qty,
        "bajarilish": bajarilish,
        "buyurtma": buyurtma_soni,
        "yetkazilgan": yetkazilgan,
        "tasdiqlangan": tasdiqlangan,
        "qarz": menejer_qarzi_jami(uid),
    }

def tasdiqlangan_menejerlar():
    natija = []
    for uid, info in menedjerlar.items():
        try:
            if info.get("status") == "tasdiqlangan":
                natija.append((str(uid), info))
        except Exception:
            pass
    return natija

def kpi_reyting_matni(oy=None):
    oy = oy or joriy_oy()
    rows = []

    for uid, info in tasdiqlangan_menejerlar():
        k = menejer_kpi(uid, oy)
        rows.append((uid, info, k))

    rows.sort(
        key=lambda z: (
            z[2]["bajarilish"],
            z[2]["savdo"],
            z[2]["yetkazilgan"]
        ),
        reverse=True
    )

    if not rows:
        return "🏆 Tasdiqlangan menejerlar yo'q."

    medal = ["🥇", "🥈", "🥉"]
    matn = f"🏆 <b>MENEJERLAR REYTINGI — {oy}</b>\n\n"

    for i, (_, info, k) in enumerate(rows, 1):
        belgi = medal[i-1] if i <= 3 else f"{i}."
        matn += (
            f"{belgi} <b>{info.get('ism', 'Nomsiz')}</b>\n"
            f"   🎯 Plan: ${k['plan']:,.0f}\n"
            f"   💰 Savdo: ${k['savdo']:,.0f}\n"
            f"   📈 Bajarilish: <b>{k['bajarilish']:.1f}%</b>\n"
            f"   📦 Buyurtmalar: {k['buyurtma']} ta\n"
            f"   🏁 Yetkazilgan: {k['yetkazilgan']} ta\n"
            f"   💳 Qarzdorlik: ${k['qarz']:,.0f}\n\n"
        )
    return matn

@bot.message_handler(func=lambda m: m.text == "🏆 Menejerlar reytingi" and rahbar_mi(m.from_user.id))
def rahbar_kpi_reyting(message):
    bot.send_message(
        message.chat.id,
        kpi_reyting_matni(),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

@bot.message_handler(func=lambda m: m.text == "👨‍💼 Menejerlar" and rahbar_mi(m.from_user.id))
def rahbar_menejerlar_kpi(message):
    rows = tasdiqlangan_menejerlar()
    if not rows:
        bot.send_message(message.chat.id, "Tasdiqlangan menejerlar yo'q.")
        return

    kb = types.InlineKeyboardMarkup()
    for uid, info in rows:
        kb.add(types.InlineKeyboardButton(
            f"👨‍💼 {info.get('ism', 'Nomsiz')}",
            callback_data=f"mkpi:{uid}"
        ))
    bot.send_message(message.chat.id, "Menejerning KPI hisobotini tanlang:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("mkpi:"))
def rahbar_menejer_kpi_callback(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    uid = call.data.split(":", 1)[1]
    info = menejer_ol(uid)
    if not info or info.get("status") != "tasdiqlangan":
        bot.send_message(call.message.chat.id, "❌ Menejer topilmadi.")
        return

    k = menejer_kpi(uid)
    matn = (
        f"👨‍💼 <b>{info.get('ism', 'Nomsiz')}</b>\n"
        f"📅 Oy: {joriy_oy()}\n\n"
        f"🎯 Plan: ${k['plan']:,.0f}\n"
        f"💰 Savdo: ${k['savdo']:,.0f}\n"
        f"📈 Plan bajarilishi: <b>{k['bajarilish']:.1f}%</b>\n"
        f"📦 Sotilgan dona: {k['qty']}\n"
        f"🛒 Buyurtmalar: {k['buyurtma']} ta\n"
        f"🏁 Yetkazilgan: {k['yetkazilgan']} ta\n"
        f"💳 Qarzdorlik: ${k['qarz']:,.0f}"
    )
    bot.send_message(call.message.chat.id, matn, parse_mode="HTML")

@bot.message_handler(commands=["kpi"])
def kpi_command(message):
    if not rahbar_mi(message.from_user.id):
        return
    bot.send_message(
        message.chat.id,
        kpi_reyting_matni(),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

# ============================================================
# 11-BOSQICH: RAHBAR ANALITIK DASHBOARD 2.0
# Bitta markazdan: savdo, menejer, do'kon, kategoriya, brend,
# model, qarzdorlik va tasdiqlangan rasxodlar.
# ============================================================

def dashboard_savdolar(oy=None):
    oy = oy or joriy_oy()
    rows = [x for x in savdolar if str(x.get("sana", "")).startswith(oy)]
    jami = sum(_float(x.get("total", x.get("summa", 0))) for x in rows)
    qty = sum(int(_float(x.get("qty", x.get("son", 0)))) for x in rows)
    return rows, jami, qty

def dashboard_qarz_jami():
    jami = 0
    for x in qarzdorlik:
        if x.get("tur") == "qarz":
            jami += _float(x.get("summa"))
        elif x.get("tur") == "tolov":
            jami -= _float(x.get("summa"))
    return max(jami, 0)

def dashboard_rasxod_jami(oy=None):
    oy = oy or joriy_oy()
    jami = 0
    for x in rasxodlar:
        if x.get("status") != "tasdiqlangan":
            continue
        if not str(x.get("sana", "")).startswith(oy):
            continue
        jami += _float(x.get("summa"))
    return jami

def dashboard_guruh(rows, key_func):
    result = {}
    for row in rows:
        key = key_func(row) or "Noma'lum"
        result[key] = result.get(key, 0) + _float(row.get("total", row.get("summa", 0)))
    return sorted(result.items(), key=lambda x: x[1], reverse=True)

def dashboard_menejer_nomi(uid):
    info = menejer_ol(uid)
    return info.get("ism", str(uid)) if info else str(uid)

def dashboard_dokon_nomi(did):
    return _dokon_nomi(did)

def dashboard_brand(row):
    brand = row.get("brand") or row.get("brend")
    if brand:
        return str(brand)
    model = str(row.get("model", ""))
    kat = str(row.get("category", row.get("kategoriya", "")))
    try:
        data = ombor_malumotlarini_oqish()
        for x in data.get(kat, []):
            if str(x[0]) == model:
                # Excel tuzilmasi A=Brend, B=Kategoriya, C=Model
                return str(x[3] if len(x) > 3 and x[3] else x[0])
    except Exception:
        pass
    return "Noma'lum"

def dashboard_matn(oy=None):
    oy = oy or joriy_oy()
    rows, savdo_jami, qty = dashboard_savdolar(oy)
    qarz = dashboard_qarz_jami()
    rasxod = dashboard_rasxod_jami(oy)
    sof = savdo_jami - rasxod

    matn = (
        f"📊 <b>RAHBAR DASHBOARD</b>\n"
        f"📅 {oy}\n\n"
        f"💰 <b>Jami savdo:</b> ${savdo_jami:,.0f}\n"
        f"📦 <b>Sotilgan dona:</b> {qty:,}\n"
        f"🧾 <b>Savdo operatsiyalari:</b> {len(rows):,} ta\n"
        f"💳 <b>Jami qarzdorlik:</b> ${qarz:,.0f}\n"
        f"💸 <b>Tasdiqlangan rasxod:</b> ${rasxod:,.0f}\n"
        f"📈 <b>Savdo − rasxod:</b> ${sof:,.0f}\n"
    )

    if rows:
        matn += "\n🏆 <b>TOP 5 MENEJER</b>\n"
        for i, (name, total) in enumerate(
            dashboard_guruh(rows, lambda r: dashboard_menejer_nomi(r.get("manager_id")))[:5], 1
        ):
            matn += f"{i}. {name} — ${total:,.0f}\n"

        matn += "\n🏪 <b>TOP 5 DO'KON</b>\n"
        for i, (name, total) in enumerate(
            dashboard_guruh(rows, lambda r: dashboard_dokon_nomi(r.get("store_id")))[:5], 1
        ):
            matn += f"{i}. {name} — ${total:,.0f}\n"

    return matn

@bot.message_handler(func=lambda m: m.text == "📊 Umumiy savdo" and rahbar_mi(m.from_user.id))
def rahbar_umumiy_savdo_dashboard(message):
    bot.send_message(
        message.chat.id,
        dashboard_matn(),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

@bot.message_handler(func=lambda m: m.text == "📈 Savdo analitikasi" and rahbar_mi(m.from_user.id))
def rahbar_analitika_menu(message):
    kb = types.InlineKeyboardMarkup()
    kb.add(
        types.InlineKeyboardButton("👨‍💼 Menejer", callback_data="dash:manager"),
        types.InlineKeyboardButton("🏪 Do'kon", callback_data="dash:store")
    )
    kb.add(
        types.InlineKeyboardButton("🏷 Brend", callback_data="dash:brand"),
        types.InlineKeyboardButton("📂 Kategoriya", callback_data="dash:category")
    )
    kb.add(types.InlineKeyboardButton("📦 Model", callback_data="dash:model"))
    bot.send_message(
        message.chat.id,
        "📈 <b>Savdo analitikasi</b>\nQaysi kesim kerak?",
        parse_mode="HTML",
        reply_markup=kb
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("dash:"))
def rahbar_dashboard_taqsimot(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    tur = call.data.split(":", 1)[1]
    rows, jami, qty = dashboard_savdolar()
    if not rows:
        bot.send_message(call.message.chat.id, "Joriy oyda savdo ma'lumotlari yo'q.")
        return

    if tur == "manager":
        title = "👨‍💼 MENEJERLAR"
        groups = dashboard_guruh(rows, lambda r: dashboard_menejer_nomi(r.get("manager_id")))
    elif tur == "store":
        title = "🏪 DO'KONLAR"
        groups = dashboard_guruh(rows, lambda r: dashboard_dokon_nomi(r.get("store_id")))
    elif tur == "brand":
        title = "🏷 BRENDLAR"
        groups = dashboard_guruh(rows, dashboard_brand)
    elif tur == "category":
        title = "📂 KATEGORIYALAR"
        groups = dashboard_guruh(rows, lambda r: r.get("category", r.get("kategoriya", "Noma'lum")))
    else:
        title = "📦 MODELLAR"
        groups = dashboard_guruh(rows, lambda r: r.get("model", "Noma'lum"))

    matn = f"📊 <b>{title}</b> — {joriy_oy()}\n\n"
    for i, (name, total) in enumerate(groups[:30], 1):
        ulush = (total / jami * 100) if jami else 0
        matn += f"{i}. {name} — ${total:,.0f} ({ulush:.1f}%)\n"

    bot.send_message(call.message.chat.id, matn, parse_mode="HTML")

@bot.message_handler(commands=["dashboard"])
def dashboard_command_2(message):
    if not rahbar_mi(message.from_user.id):
        return
    bot.send_message(message.chat.id, dashboard_matn(), parse_mode="HTML", reply_markup=rahbar_menu())

# ============================================================
# 12-BOSQICH: KUNLIK / OYLIK HISOBOTLAR
# Rahbar uchun tayyor hisobot: savdo, buyurtma, qarz, rasxod,
# menejerlar va top mahsulotlar.
# ============================================================

def sana_bugun():
    return datetime.now().strftime("%Y-%m-%d")

def hisobot_savdo_oraliq(start_date, end_date):
    rows = []
    for x in savdolar:
        sana = str(x.get("sana", ""))[:10]
        if start_date <= sana <= end_date:
            rows.append(x)
    return rows

def hisobot_buyurtma_oraliq(start_date, end_date):
    data = menejer_buyurtmalarini_yuklash()
    return [
        x for x in data
        if start_date <= str(x.get("sana", ""))[:10] <= end_date
    ]

def hisobot_rasxod_oraliq(start_date, end_date):
    return [
        x for x in rasxodlar
        if x.get("status") == "tasdiqlangan"
        and start_date <= str(x.get("sana", ""))[:10] <= end_date
    ]

def hisobot_qarz_balansi():
    return dashboard_qarz_jami()

def umumiy_hisobot(start_date, end_date, sarlavha):
    sales = hisobot_savdo_oraliq(start_date, end_date)
    orders = hisobot_buyurtma_oraliq(start_date, end_date)
    expenses = hisobot_rasxod_oraliq(start_date, end_date)

    savdo = sum(_float(x.get("total", x.get("summa", 0))) for x in sales)
    qty = sum(int(_float(x.get("qty", x.get("son", 0)))) for x in sales)
    rasxod = sum(_float(x.get("summa", 0)) for x in expenses)

    yangi = sum(1 for x in orders if x.get("status") == "yangi")
    tasdiq = sum(1 for x in orders if x.get("status") == "tasdiqlandi")
    tayyor = sum(1 for x in orders if x.get("status") == "tayyorlanmoqda")
    yolda = sum(1 for x in orders if x.get("status") == "yolda")
    yetkazildi = sum(1 for x in orders if x.get("status") == "yetkazildi")
    bekor = sum(1 for x in orders if x.get("status") == "bekor_qilindi")

    matn = (
        f"📑 <b>{sarlavha}</b>\n"
        f"📅 {start_date} → {end_date}\n\n"
        f"💰 <b>Savdo:</b> ${savdo:,.0f}\n"
        f"📦 <b>Sotilgan:</b> {qty:,} dona\n"
        f"🧾 <b>Savdo operatsiyasi:</b> {len(sales)} ta\n"
        f"💸 <b>Tasdiqlangan rasxod:</b> ${rasxod:,.0f}\n"
        f"📈 <b>Savdo − rasxod:</b> ${savdo - rasxod:,.0f}\n\n"
        f"🛒 <b>Buyurtmalar:</b> {len(orders)} ta\n"
        f"   🕐 Yangi: {yangi}\n"
        f"   ✅ Tasdiqlangan: {tasdiq}\n"
        f"   📦 Tayyorlanmoqda: {tayyor}\n"
        f"   🚚 Yo'lda: {yolda}\n"
        f"   🏁 Yetkazilgan: {yetkazildi}\n"
        f"   ❌ Bekor qilingan: {bekor}\n\n"
        f"💳 <b>Jami qarzdorlik:</b> ${hisobot_qarz_balansi():,.0f}\n"
    )

    # Menejerlar kesimi
    if sales:
        matn += "\n🏆 <b>MENEJERLAR</b>\n"
        groups = dashboard_guruh(
            sales,
            lambda r: dashboard_menejer_nomi(r.get("manager_id"))
        )
        for i, (name, total) in enumerate(groups[:10], 1):
            matn += f"{i}. {name} — ${total:,.0f}\n"

        # Top mahsulotlar
        products = dashboard_guruh(
            sales,
            lambda r: r.get("model", "Noma'lum")
        )
        matn += "\n📦 <b>TOP MAHSULOTLAR</b>\n"
        for i, (name, total) in enumerate(products[:10], 1):
            matn += f"{i}. {name} — ${total:,.0f}\n"

    return matn

def rahbar_hisobot_kb():
    kb = types.InlineKeyboardMarkup()
    kb.add(
        types.InlineKeyboardButton("📅 Bugun", callback_data="report:today"),
        types.InlineKeyboardButton("📆 Kecha", callback_data="report:yesterday")
    )
    kb.add(types.InlineKeyboardButton("🗓 Shu oy", callback_data="report:month"))
    return kb

@bot.message_handler(func=lambda m: m.text == "📑 Hisobotlar" and rahbar_mi(m.from_user.id))
def rahbar_hisobotlar_menu(message):
    bot.send_message(
        message.chat.id,
        "📑 <b>Hisobotlar</b>\n\nKerakli davrni tanlang:",
        parse_mode="HTML",
        reply_markup=rahbar_hisobot_kb()
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("report:"))
def rahbar_hisobot_callback(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    today = datetime.now().date()
    tur = call.data.split(":", 1)[1]

    if tur == "today":
        start = end = today
        title = "BUGUNGI HISOBOT"
    elif tur == "yesterday":
        start = end = today - timedelta(days=1)
        title = "KECHAGI HISOBOT"
    else:
        start = today.replace(day=1)
        end = today
        title = "SHU OY HISOBOTI"

    matn = umumiy_hisobot(
        start.strftime("%Y-%m-%d"),
        end.strftime("%Y-%m-%d"),
        title
    )
    bot.send_message(call.message.chat.id, matn, parse_mode="HTML")

@bot.message_handler(commands=["hisobot"])
def hisobot_command(message):
    if not rahbar_mi(message.from_user.id):
        return
    today = datetime.now().date()
    start = today.strftime("%Y-%m-%d")
    bot.send_message(
        message.chat.id,
        umumiy_hisobot(start, start, "BUGUNGI HISOBOT"),
        parse_mode="HTML",
        reply_markup=rahbar_menu()
    )

# ============================================================
# 13-BOSQICH: VAZIFALAR + DEADLINE + ESLATMA
# Rahbar menejerga vazifa beradi. Menejer faqat o'z vazifalarini
# ko'radi. Deadline yaqinlashganda bot eslatma yuboradi.
# ============================================================

VAZIFALAR_FAYLI = "vazifalar.json"
vazifa_holati = {}

def vazifalarni_yuklash():
    if not os.path.exists(VAZIFALAR_FAYLI):
        return []
    try:
        with open(VAZIFALAR_FAYLI, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        log.exception("Vazifalarni yuklashda xatolik")
        return []

def vazifalarni_saqlash(data):
    tmp = VAZIFALAR_FAYLI + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, VAZIFALAR_FAYLI)

def yangi_vazifa_id():
    data = vazifalarni_yuklash()
    mx = 0
    for x in data:
        try:
            mx = max(mx, int(x.get("id", 0)))
        except Exception:
            pass
    return mx + 1

def vazifa_status(status):
    return {
        "yangi": "🕐 Yangi",
        "jarayonda": "🔄 Jarayonda",
        "bajarildi": "✅ Bajarildi",
        "muddati_otdi": "🔴 Muddati o'tdi",
        "bekor_qilingan": "❌ Bekor qilingan",
    }.get(status, status)

def vazifa_matni(x):
    deadline = x.get("deadline", "Belgilanmagan")
    return (
        f"🎯 <b>Vazifa #{x['id']}</b>\n\n"
        f"📝 {x.get('matn', '')}\n"
        f"👨‍💼 Menejer: {x.get('menejer_ism', '')}\n"
        f"📅 Deadline: {deadline}\n"
        f"📌 Status: {vazifa_status(x.get('status'))}"
    )

@bot.message_handler(func=lambda m: m.text == "🎯 Vazifalar" and rahbar_mi(m.from_user.id))
def rahbar_vazifalar_menu(message):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("➕ Vazifa berish", "📋 Barcha vazifalar")
    kb.add("🏠 Bosh menyu")
    bot.send_message(
        message.chat.id,
        "🎯 <b>Vazifalar boshqaruvi</b>",
        parse_mode="HTML",
        reply_markup=kb
    )

@bot.message_handler(func=lambda m: m.text == "➕ Vazifa berish" and rahbar_mi(m.from_user.id))
def rahbar_vazifa_menejer_tanlash(message):
    rows = tasdiqlangan_menejerlar()
    if not rows:
        bot.send_message(message.chat.id, "Tasdiqlangan menejerlar yo'q.")
        return

    kb = types.InlineKeyboardMarkup()
    for uid, info in rows:
        kb.add(types.InlineKeyboardButton(
            f"👨‍💼 {info.get('ism', 'Nomsiz')}",
            callback_data=f"taskmgr:{uid}"
        ))
    bot.send_message(message.chat.id, "Vazifa qaysi menejerga?", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("taskmgr:"))
def rahbar_vazifa_menejer_callback(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    uid = call.data.split(":", 1)[1]
    if not menejer_mi(uid):
        bot.send_message(call.message.chat.id, "❌ Menejer topilmadi.")
        return

    vazifa_holati[call.from_user.id] = {
        "bosqich": "matn",
        "menejer_id": uid
    }
    bot.send_message(
        call.message.chat.id,
        "📝 Vazifani yozing.\nMasalan: 20 ta do'kon bilan bog'lanish"
    )

@bot.message_handler(func=lambda m: m.from_user.id in vazifa_holati)
def rahbar_vazifa_kiritish(message):
    uid = message.from_user.id
    state = vazifa_holati.get(uid)
    if not state:
        return

    if state.get("bosqich") == "matn":
        state["matn"] = (message.text or "").strip()
        if not state["matn"]:
            bot.send_message(message.chat.id, "❌ Vazifa matni bo'sh bo'lmasin.")
            return
        state["bosqich"] = "deadline"
        bot.send_message(
            message.chat.id,
            "📅 Deadline kiriting:\n"
            "Format: YYYY-MM-DD HH:MM\n"
            "Masalan: 2026-09-15 18:00"
        )
        return

    if state.get("bosqich") == "deadline":
        deadline = (message.text or "").strip()
        try:
            dt = datetime.strptime(deadline, "%Y-%m-%d %H:%M")
        except ValueError:
            bot.send_message(
                message.chat.id,
                "❌ Format noto'g'ri.\n"
                "Masalan: 2026-09-15 18:00"
            )
            return

        if dt <= datetime.now():
            bot.send_message(message.chat.id, "❌ Deadline kelajakdagi vaqt bo'lishi kerak.")
            return

        menejer_id = state["menejer_id"]
        info = menejer_ol(menejer_id)
        task = {
            "id": yangi_vazifa_id(),
            "menejer_id": int(menejer_id),
            "menejer_ism": (info or {}).get("ism", ""),
            "matn": state["matn"],
            "deadline": dt.strftime("%Y-%m-%d %H:%M"),
            "bergan_rahbar_id": uid,
            "berilgan_vaqt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "yangi",
            "eslatma_1soat": False,
            "eslatma_24soat": False,
        }

        data = vazifalarni_yuklash()
        data.append(task)
        vazifalarni_saqlash(data)
        vazifa_holati.pop(uid, None)

        bot.send_message(
            message.chat.id,
            "✅ Vazifa yaratildi va menejerga yuborildi.",
            reply_markup=rahbar_menu()
        )

        try:
            kb = types.InlineKeyboardMarkup()
            kb.add(
                types.InlineKeyboardButton("🔄 Jarayonda", callback_data=f"taskstart:{task['id']}"),
                types.InlineKeyboardButton("✅ Bajarildi", callback_data=f"taskdone:{task['id']}")
            )
            bot.send_message(
                int(menejer_id),
                vazifa_matni(task),
                parse_mode="HTML",
                reply_markup=kb
            )
        except Exception:
            log.exception("Menejerga vazifa yuborishda xatolik")

@bot.message_handler(func=lambda m: m.text == "🎯 Mening vazifalarim" and menejer_mi(m.from_user.id))
def menejer_vazifalarim(message):
    data = vazifalarni_yuklash()
    mine = [x for x in data if str(x.get("menejer_id")) == str(message.from_user.id)]

    if not mine:
        bot.send_message(message.chat.id, "🎯 Sizga berilgan vazifalar yo'q.")
        return

    for x in reversed(mine[-30:]):
        kb = types.InlineKeyboardMarkup()
        if x.get("status") == "yangi":
            kb.add(types.InlineKeyboardButton(
                "🔄 Jarayonda", callback_data=f"taskstart:{x['id']}"
            ))
        if x.get("status") in {"yangi", "jarayonda"}:
            kb.add(types.InlineKeyboardButton(
                "✅ Bajarildi", callback_data=f"taskdone:{x['id']}"
            ))
        bot.send_message(
            message.chat.id,
            vazifa_matni(x),
            parse_mode="HTML",
            reply_markup=kb
        )

def vazifa_status_ozgartir(uid, task_id, status):
    data = vazifalarni_yuklash()
    for x in data:
        if int(x.get("id", 0)) != int(task_id):
            continue

        if str(x.get("menejer_id")) != str(uid) and not rahbar_mi(uid):
            return None, False, "Bu vazifa sizga tegishli emas."

        if x.get("status") == "bajarildi":
            return x, False, "Vazifa allaqachon bajarilgan."

        x["status"] = status
        x["status_vaqti"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if status == "bajarildi":
            x["bajarilgan_vaqt"] = x["status_vaqti"]

        vazifalarni_saqlash(data)
        return x, True, None

    return None, False, "Vazifa topilmadi."

@bot.callback_query_handler(func=lambda c: c.data.startswith("taskstart:"))
def vazifa_jarayonda_callback(call):
    bot.answer_callback_query(call.id)
    uid = call.from_user.id
    task_id = int(call.data.split(":", 1)[1])
    x, ok, error = vazifa_status_ozgartir(uid, task_id, "jarayonda")
    if not x:
        bot.send_message(call.message.chat.id, f"❌ {error}")
        return
    if not ok:
        bot.send_message(call.message.chat.id, f"ℹ️ {error}")
        return

    try:
        bot.edit_message_text(
            vazifa_matni(x),
            call.message.chat.id,
            call.message.message_id,
            parse_mode="HTML"
        )
    except Exception:
        pass

@bot.callback_query_handler(func=lambda c: c.data.startswith("taskdone:"))
def vazifa_bajarildi_callback(call):
    bot.answer_callback_query(call.id)
    uid = call.from_user.id
    task_id = int(call.data.split(":", 1)[1])
    x, ok, error = vazifa_status_ozgartir(uid, task_id, "bajarildi")
    if not x:
        bot.send_message(call.message.chat.id, f"❌ {error}")
        return
    if not ok:
        bot.send_message(call.message.chat.id, f"ℹ️ {error}")
        return

    try:
        bot.edit_message_text(
            vazifa_matni(x),
            call.message.chat.id,
            call.message.message_id,
            parse_mode="HTML"
        )
    except Exception:
        pass

    # Vazifani bergan rahbarga bajarilganligi haqida xabar.
    try:
        bot.send_message(
            int(x["bergan_rahbar_id"]),
            f"✅ Menejer vazifani bajardi!\n\n{vazifa_matni(x)}",
            parse_mode="HTML"
        )
    except Exception:
        log.exception("Rahbarga vazifa natijasini yuborishda xatolik")

@bot.message_handler(func=lambda m: m.text == "📋 Barcha vazifalar" and rahbar_mi(m.from_user.id))
def rahbar_barcha_vazifalar(message):
    data = vazifalarni_yuklash()
    if not data:
        bot.send_message(message.chat.id, "Vazifalar yo'q.")
        return

    for x in reversed(data[-50:]):
        bot.send_message(
            message.chat.id,
            vazifa_matni(x),
            parse_mode="HTML"
        )

@bot.message_handler(commands=["vazifalar"])
def vazifalar_command(message):
    if rahbar_mi(message.from_user.id):
        bot.send_message(
            message.chat.id,
            "🎯 Vazifalar bo'limini menyudan oching.",
            reply_markup=rahbar_menu()
        )
    elif menejer_mi(message.from_user.id):
        menejer_vazifalarim(message)

def vazifalarni_avtomatik_tekshirish():
    """
    Bot ishga tushganda va keyinchalik davriy chaqirilishi mumkin.
    24 soat va 1 soat qolganida bir martadan eslatma yuboradi.
    Muddati o'tgan vazifa avtomatik 'muddati_otdi' bo'ladi.
    """
    data = vazifalarni_yuklash()
    now = datetime.now()
    ozgardi = False

    for x in data:
        if x.get("status") == "bajarildi" or x.get("status") == "bekor_qilingan":
            continue

        try:
            deadline = datetime.strptime(x["deadline"], "%Y-%m-%d %H:%M")
        except Exception:
            continue

        farq = deadline - now

        if farq.total_seconds() <= 0:
            if x.get("status") != "muddati_otdi":
                x["status"] = "muddati_otdi"
                x["muddat_otgan_vaqt"] = now.strftime("%Y-%m-%d %H:%M:%S")
                ozgardi = True
                try:
                    bot.send_message(
                        int(x["menejer_id"]),
                        f"🔴 <b>Vazifa muddati o'tdi!</b>\n\n{vazifa_matni(x)}",
                        parse_mode="HTML"
                    )
                except Exception:
                    pass
            continue

        if farq.total_seconds() <= 3600 and not x.get("eslatma_1soat"):
            x["eslatma_1soat"] = True
            ozgardi = True
            try:
                bot.send_message(
                    int(x["menejer_id"]),
                    f"⏰ <b>1 soat qoldi!</b>\n\n{vazifa_matni(x)}",
                    parse_mode="HTML"
                )
            except Exception:
                pass

        elif farq.total_seconds() <= 86400 and not x.get("eslatma_24soat"):
            x["eslatma_24soat"] = True
            ozgardi = True
            try:
                bot.send_message(
                    int(x["menejer_id"]),
                    f"🔔 <b>24 soat qoldi!</b>\n\n{vazifa_matni(x)}",
                    parse_mode="HTML"
                )
            except Exception:
                pass

    if ozgardi:
        vazifalarni_saqlash(data)

# ============================================================
# 14-BOSQICH: FOTO / VIDEO HISOBOT
# Menejer vazifaga foto yoki video yuboradi.
# Rahbar tasdiqlaydi yoki qayta topshirishni so'raydi.
# ============================================================

HISOBOTLAR_FAYLI = "vazifa_hisobotlar.json"

def hisobotlarni_yuklash():
    if not os.path.exists(HISOBOTLAR_FAYLI):
        return []
    try:
        with open(HISOBOTLAR_FAYLI, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        log.exception("Hisobotlarni yuklashda xatolik")
        return []

def hisobotlarni_saqlash(data):
    tmp = HISOBOTLAR_FAYLI + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, HISOBOTLAR_FAYLI)

def yangi_hisobot_id():
    data = hisobotlarni_yuklash()
    mx = 0
    for x in data:
        try:
            mx = max(mx, int(x.get("id", 0)))
        except Exception:
            pass
    return mx + 1

def vazifa_hisobot_tanlash(task_id, user_id):
    data = vazifalarni_yuklash()
    for x in data:
        if int(x.get("id", 0)) == int(task_id):
            if str(x.get("menejer_id")) == str(user_id):
                return x
            return None
    return None

@bot.callback_query_handler(func=lambda c: c.data.startswith("taskreport:"))
def vazifa_hisobot_boshlash(call):
    bot.answer_callback_query(call.id)
    uid = call.from_user.id
    task_id = int(call.data.split(":", 1)[1])

    task = vazifa_hisobot_tanlash(task_id, uid)
    if not task:
        bot.send_message(call.message.chat.id, "❌ Bu vazifa sizga tegishli emas.")
        return

    if task.get("status") not in {"yangi", "jarayonda", "muddati_otdi"}:
        bot.send_message(
            call.message.chat.id,
            "❌ Bu vazifa uchun hisobot yuborish holati mavjud emas."
        )
        return

    vazifa_holati[uid] = {
        "bosqich": "hisobot_media",
        "task_id": task_id
    }
    bot.send_message(
        call.message.chat.id,
        "📸 Foto yoki 🎥 video yuboring.\n\n"
        "Bu vazifa bo'yicha rahbarga hisobot sifatida yuboriladi."
    )

@bot.message_handler(
    content_types=["photo", "video"],
    func=lambda m: m.from_user.id in vazifa_holati
)
def vazifa_media_qabul(message):
    uid = message.from_user.id
    state = vazifa_holati.get(uid)

    if not state or state.get("bosqich") != "hisobot_media":
        return

    task_id = int(state["task_id"])
    task = vazifa_hisobot_tanlash(task_id, uid)
    if not task:
        vazifa_holati.pop(uid, None)
        bot.send_message(message.chat.id, "❌ Vazifa topilmadi.")
        return

    media_type = "photo" if message.content_type == "photo" else "video"

    if media_type == "photo":
        file_id = message.photo[-1].file_id
    else:
        file_id = message.video.file_id

    hid = yangi_hisobot_id()
    report = {
        "id": hid,
        "task_id": task_id,
        "menejer_id": uid,
        "menejer_ism": task.get("menejer_ism", ""),
        "rahbar_id": task.get("bergan_rahbar_id"),
        "media_type": media_type,
        "file_id": file_id,
        "sana": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": "kutilmoqda",
        "izoh": ""
    }

    data = hisobotlarni_yuklash()
    data.append(report)
    hisobotlarni_saqlash(data)

    # Vazifani avtomatik "hisobot_kutilmoqda" holatiga o'tkazmaymiz:
    # menejer bajarildi tugmasini alohida bosishi mumkin.
    vazifa_holati.pop(uid, None)

    bot.send_message(
        message.chat.id,
        f"📤 Hisobot #{hid} rahbarga yuborildi.\n"
        f"🎯 Vazifa #{task_id}"
    )

    rahbar_id = task.get("bergan_rahbar_id")
    try:
        kb = types.InlineKeyboardMarkup()
        kb.add(
            types.InlineKeyboardButton(
                "✅ Hisobotni tasdiqlash",
                callback_data=f"hrok:{hid}"
            ),
            types.InlineKeyboardButton(
                "🔄 Qayta yuborish",
                callback_data=f"hrretry:{hid}"
            )
        )

        caption = (
            f"📸 <b>Vazifa hisoboti #{hid}</b>\n"
            f"🎯 Vazifa #{task_id}\n"
            f"👨‍💼 Menejer: {task.get('menejer_ism', '')}\n"
            f"📝 {task.get('matn', '')}\n"
            f"📅 {task.get('deadline', '')}"
        )

        if media_type == "photo":
            bot.send_photo(
                int(rahbar_id),
                file_id,
                caption=caption,
                parse_mode="HTML",
                reply_markup=kb
            )
        else:
            bot.send_video(
                int(rahbar_id),
                file_id,
                caption=caption,
                parse_mode="HTML",
                reply_markup=kb
            )
    except Exception:
        log.exception("Rahbarga media hisobot yuborishda xatolik")

def hisobot_status_yangila(hid, status):
    data = hisobotlarni_yuklash()
    for x in data:
        if int(x.get("id", 0)) == int(hid):
            x["status"] = status
            x["status_vaqti"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            hisobotlarni_saqlash(data)
            return x
    return None

@bot.callback_query_handler(func=lambda c: c.data.startswith("hrok:"))
def hisobot_tasdiqlash_callback(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    hid = int(call.data.split(":", 1)[1])
    report = hisobot_status_yangila(hid, "tasdiqlandi")
    if not report:
        bot.send_message(call.message.chat.id, "❌ Hisobot topilmadi.")
        return

    try:
        bot.edit_message_reply_markup(
            call.message.chat.id,
            call.message.message_id,
            reply_markup=None
        )
    except Exception:
        pass

    # Hisobot tasdiqlanganda tegishli vazifa bajarildi qilinadi.
    task_id = report.get("task_id")
    tasks = vazifalarni_yuklash()
    for task in tasks:
        if int(task.get("id", 0)) == int(task_id):
            task["status"] = "bajarildi"
            task["bajarilgan_vaqt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            task["hisobot_id"] = hid
            break
    vazifalarni_saqlash(tasks)

    try:
        bot.send_message(
            int(report["menejer_id"]),
            f"✅ Hisobotingiz tasdiqlandi.\n"
            f"📸 Hisobot #{hid}\n"
            f"🎯 Vazifa #{task_id} bajarilgan deb qabul qilindi."
        )
    except Exception:
        pass

@bot.callback_query_handler(func=lambda c: c.data.startswith("hrretry:"))
def hisobot_qayta_callback(call):
    bot.answer_callback_query(call.id)
    if not rahbar_mi(call.from_user.id):
        return

    hid = int(call.data.split(":", 1)[1])
    report = hisobot_status_yangila(hid, "qayta_yuborish")
    if not report:
        bot.send_message(call.message.chat.id, "❌ Hisobot topilmadi.")
        return

    try:
        bot.edit_message_reply_markup(
            call.message.chat.id,
            call.message.message_id,
            reply_markup=None
        )
    except Exception:
        pass

    try:
        bot.send_message(
            int(report["menejer_id"]),
            f"🔄 Hisobot #{hid} qayta yuborilishi kerak.\n"
            f"🎯 Vazifa #{report['task_id']}\n"
            f"📸 Yangi foto yoki video yuboring."
        )
    except Exception:
        pass

@bot.message_handler(func=lambda m: m.text == "📸 Hisobotlarim" and menejer_mi(m.from_user.id))
def menejer_hisobotlarim(message):
    data = hisobotlarni_yuklash()
    mine = [x for x in data if str(x.get("menejer_id")) == str(message.from_user.id)]

    if not mine:
        bot.send_message(message.chat.id, "📸 Hozircha hisobotlaringiz yo'q.")
        return

    statuslar = {
        "kutilmoqda": "🕐 Kutilmoqda",
        "tasdiqlandi": "✅ Tasdiqlandi",
        "qayta_yuborish": "🔄 Qayta yuborish",
    }

    matn = "📸 <b>Hisobotlarim</b>\n\n"
    for x in reversed(mine[-30:]):
        matn += (
            f"#{x['id']} — 🎯 Vazifa #{x['task_id']}\n"
            f"📅 {x['sana']}\n"
            f"📌 {statuslar.get(x.get('status'), x.get('status'))}\n\n"
        )
    bot.send_message(message.chat.id, matn, parse_mode="HTML", reply_markup=menejer_menu())
