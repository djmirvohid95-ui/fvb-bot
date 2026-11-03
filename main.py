import os
import re
import datetime
import telebot
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI

# Telegram Bot Tokeningiz
TOKEN = "9140584759:AAFI2r78alD7d4mBaulU_8xMd-Ob8A_nvLw"
bot = telebot.TeleBot(TOKEN)

# OpenAI API kalitingiz
client = OpenAI(api_key="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")

user_data = {}

def analyze_voice_with_chatgpt(audio_path):
    """Ovozli xabarni rasmiy va xatosiz o'zbek tilida tahlil qilish"""
    try:
        with open(audio_path, "rb") as audio_file:
            transcript = client.audio.transcriptions.create(
                model="whisper-1",
                file=audio_file
            )
        
        raw_text = transcript.text
        
        # ChatGPT ga mahalla nomi va imlolarni qat'iy to'g'rilash bo'yicha buyruq
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Siz Jizzax shahar FVB yong'in xavfsizligi inspektorining malakali yordamchisiz. "
                        "Inspektor (leytenant Akramov Mirvohid) aytgan ovozli xabarni matnga o'girib, quyidagi qoidalarga qat'iy amal qilgan holda rasmiy dalolatnoma ma'lumotlarini tayyorlaysiz:\n"
                        "QOIDALAR:\n"
                        "1. Mahalla nomi eshitilishidan qat'iy nazar FAKAT 'Kuyovboshi' deb yozilsin (Qiyobosh yoki boshqa narsa yozilmasin).\n"
                        "2. Aholi soni '6 kishi' formatida yozilsin ('odam' emas).\n"
                        "3. Kamchiliklar qismida: oshxona va gaz plitasi atrofida yong'in xavfsizligi qoidalariga rioya etilmagani, dudbo'ron (vytyajka/dudbo'ron) nosozligi yoki yo'qligi hamda is gazi xavfining oldini olish bo'yicha tushunarli va professional tilda bayon qilinsin.\n"
                        "4. Javobni FAQAT quyidagi aniq tartibda bering:\n"
                        "1. Mahalla nomi:\n"
                        "2. Ko'cha va uy raqami:\n"
                        "3. Xonadon egasining F.I.Sh.:\n"
                        "4. Aholi soni:\n"
                        "5. Aniqlangan kamchiliklar:"
                    )
                },
                {
                    "role": "user",
                    "content": f"Ovozli xabardan olingan matn: {raw_text}"
                }
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"ChatGPT orqali tahlil qilishda xatolik: {str(e)}"

def create_word_document(ai_extracted_text, photo_paths):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Sarlavha
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("DALOLATNOMA\n")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.name = "Times New Roman"

    run_sub = p_title.add_run("Xonadonlarni yong'in xavfsizligi qoidalari bo'yicha o'rganish haqida")
    run_sub.bold = True
    run_sub.font.size = Pt(12)
    run_sub.font.name = "Times New Roman"

    # Sana va joy
    today_str = datetime.datetime.now().strftime("%d.%m.%Y yil")
    p_date = doc.add_paragraph()
    p_date.add_run(f"Jizzax shahri                                                                 {today_str}").font.size = Pt(11)

    # Kirish qismi
    p_body = doc.add_paragraph()
    p_body.paragraph_format.line_spacing = 1.15
    p_body.paragraph_format.space_after = Pt(6)
    
    run_body = p_body.add_run(
        "   Jizzax shahar FVB FVPB inspektori, leytenant M.Sh.Akramov tomonidan fuqarolarning "
        "uy-joylarida yong'in xavfsizligini ta'minlash, is gazi va ovqatdan zaharlanishning "
        "oldini olish borasidagi profilaktik tadbirlar doirasida quyidagi xonadon ko'zdan kechirildi:\n"
    )
    run_body.font.name = "Times New Roman"
    run_body.font.size = Pt(11)

    # Natijalar
    p_item = doc.add_paragraph()
    p_item.paragraph_format.space_after = Pt(6)
    r_text = p_item.add_run(ai_extracted_text)
    r_text.font.size = Pt(11)

    # Yoriqnoma berilganligi
    p_ins = doc.add_paragraph()
    p_ins.paragraph_format.space_after = Pt(8)
    r_ins = p_ins.add_run(
        "Profilaktik tushuntirish ishlari:\n"
        "O'rganish davomida xonadon egasiga hamda oila a'zolariga yong'in xavfsizligi qoidalari, "
        "is gazi va ovqatdan zaharlanishning oldini olish bo'yicha zaruriy yo'riqnomalar berildi."
    )
    r_ins.font.size = Pt(11)

    # Rasmlar sarlavhasi
    p_img_title = doc.add_paragraph()
    p_img_title.paragraph_format.space_after = Pt(4)
    r_img = p_img_title.add_run("Xonadon holati va aniqlangan kamchilik fotolavhalari:")
    r_img.bold = True
    r_img.font.size = Pt(11)

    # Rasmlarni jadvalga yonma-yon joylash
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False

    for i, path in enumerate(photo_paths[:2]):
        if os.path.exists(path):
            cell = table.cell(0, i)
            cell_p = cell.paragraphs[0]
            cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_p.add_run().add_picture(path, width=Inches(2.8))

    # Imzo qismi
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(16)
    r_sign1 = p_sign.add_run("Jizzax shahar FVB FVPB inspektori\nleytenant                                                                           M.Sh.Akramov")
    r_sign1.bold = True
    r_sign1.font.size = Pt(11)

    # Fayl nomini Mahalla_KuchaUy_Familiya shaklida yasash
    mahalla = "Mahalla"
    kucha_uy = "Manzil"
    familiya = "Fuqaro"
    
    for line in ai_extracted_text.split('\n'):
        if "1. Mahalla nomi:" in line:
            parts = line.split(":")
            if len(parts) > 1 and parts[1].strip():
                mahalla = re.sub(r'[^a-zA-Z0-9og\'g\'shch\s]', '', parts[1]).strip().replace(' ', '_')
        elif "2. Ko'cha va uy raqami:" in line:
            parts = line.split(":")
            if len(parts) > 1 and parts[1].strip():
                kucha_uy = re.sub(r'[^a-zA-Z0-9og\'g\'shch\s]', '', parts[1]).strip().replace(' ', '_')
        elif "3. Xonadon egasining F.I.Sh.:" in line:
            parts = line.split(":")
            if len(parts) > 1 and parts[1].strip():
                fio_parts = parts[1].strip().split()
                if fio_parts:
                    familiya = re.sub(r'[^a-zA-Z0-9]', '', fio_parts[0])

    filename = f"{mahalla}_{kucha_uy}_{familiya}.docx"
    doc.save(filename)
    return filename

@bot.message_handler(commands=['start'])
def send_welcome(message):
    user_data[message.chat.id] = {'photos': [], 'ai_text': '', 'step': 'waiting_voice'}
    bot.reply_to(
        message, 
        "🤖 Salom! Bot tayyor.\n\n"
        "Marhamat, xonadon bo'yicha **ovozli xabar** yuboring.",
        parse_mode="HTML"
    )

@bot.message_handler(content_types=['voice'])
def handle_voice(message):
    chat_id = message.chat.id
    if chat_id not in user_data:
        user_data[chat_id] = {'photos': [], 'ai_text': '', 'step': 'waiting_voice'}

    bot.reply_to(message, "🎙 Ovozli xabar qabul qilindi. Tahlil qilinmoqda, biroz kuting...")

    try:
        file_info = bot.get_file(message.voice.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        if not os.path.exists("temp"):
            os.makedirs("temp")

        voice_path = f"temp/{message.voice.file_id}.ogg"
        with open(voice_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        extracted_info = analyze_voice_with_chatgpt(voice_path)

        user_data[chat_id]['voice_path'] = voice_path
        user_data[chat_id]['ai_text'] = extracted_info
        user_data[chat_id]['step'] = 'waiting_photos'
        user_data[chat_id]['photos'] = []

        bot.reply_to(message, f"✅ **Aniqlangan ma'lumotlar:**\n\n{extracted_info}\n\n📸 Endi ushbu xonadon bo'yicha **2 ta rasm** yuboring.")
    except Exception as e:
        bot.reply_to(message, f"⚠️ Xatolik yuz berdi: {str(e)}")

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    chat_id = message.chat.id
    if chat_id not in user_data or user_data[chat_id].get('step') != 'waiting_photos':
        bot.reply_to(message, "⚠️ Avval ovozli xabar yuboring!")
        return

    try:
        file_info = bot.get_file(message.photo[-1].file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        if not os.path.exists("temp"):
            os.makedirs("temp")

        photo_path = f"temp/{message.photo[-1].file_id}.jpg"
        with open(photo_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        user_data[chat_id]['photos'].append(photo_path)
        count = len(user_data[chat_id]['photos'])

        if count < 2:
            bot.reply_to(message, f"📸 {count}-rasm qabul qilindi. Yana **1 ta rasm** yuboring.")
        else:
            bot.reply_to(message, "⏳ Dalolatnoma tayyorlanmoqda, kuting...")
            
            ai_text = user_data[chat_id]['ai_text']
            photos = user_data[chat_id]['photos']
            
            doc_file = create_word_document(ai_text, photos)
            
            with open(doc_file, 'rb') as doc:
                bot.send_document(chat_id, doc, caption=f"📄 Tayyor Dalolatnoma: {doc_file}")

            for p in photos:
                if os.path.exists(p): os.remove(p)
            if 'voice_path' in user_data[chat_id] and os.path.exists(user_data[chat_id]['voice_path']):
                os.remove(user_data[chat_id]['voice_path'])
            
            user_data[chat_id] = {'photos': [], 'ai_text': '', 'step': 'waiting_voice'}
            bot.send_message(chat_id, "✅ Tayyor! Keyingi xonadon uchun yana ovozli xabar yuborishingiz mumkin.")
    except Exception as e:
        bot.reply_to(message, f"⚠️ Xatolik yuz berdi: {str(e)}")

print("Bot mukammal imlo rejimida ishga tushdi...")
bot.infinity_polling(none_stop=True, skip_pending=True)
