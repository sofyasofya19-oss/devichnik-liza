#!/usr/bin/env python3
"""Convert the bachelorette party presentation to PowerPoint."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

PHOTOS = "photos"
STOCK = "stock"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

BG_PURPLE = RGBColor(0x1a, 0x0a, 0x2e)
BG_DARK = RGBColor(0x0f, 0x05, 0x1a)
PINK = RGBColor(0xf8, 0xb4, 0xd9)
VIOLET = RGBColor(0xa7, 0x8b, 0xfa)
GOLD = RGBColor(0xfb, 0xbf, 0x24)
WHITE = RGBColor(0xff, 0xff, 0xff)
LIGHT_PURPLE = RGBColor(0xe9, 0xd5, 0xff)


def add_bg(slide, color=BG_PURPLE):
    """Fill slide background with solid color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_title_slide(title, subtitle="", block_num=""):
    """Create a title/section slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    if block_num:
        tb = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.5), Inches(12), Inches(1)
        )
        tf = tb.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = block_num
        run.font.size = Pt(24)
        run.font.color.rgb = VIOLET
        run.font.bold = True
    tb = slide.shapes.add_textbox(
        Inches(0.5), Inches(2.5), Inches(12), Inches(2)
    )
    tf = tb.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = title
    run.font.size = Pt(54)
    run.font.color.rgb = WHITE
    run.font.bold = True
    if subtitle:
        tb2 = slide.shapes.add_textbox(
            Inches(1), Inches(4.5), Inches(11), Inches(2)
        )
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
        run2 = tf2.paragraphs[0].add_run()
        run2.text = subtitle
        run2.font.size = Pt(24)
        run2.font.color.rgb = LIGHT_PURPLE
    return slide


def add_question_slide(num, text):
    """Create a question slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    tb = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(11), Inches(1)
    )
    tf = tb.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = f"Вопрос {num} из 13"
    run.font.size = Pt(20)
    run.font.color.rgb = VIOLET
    tb2 = slide.shapes.add_textbox(
        Inches(1), Inches(2.5), Inches(11), Inches(3)
    )
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
    run2 = tf2.paragraphs[0].add_run()
    run2.text = text
    run2.font.size = Pt(36)
    run2.font.color.rgb = WHITE
    run2.font.bold = True
    return slide


def add_video_slide(num):
    """Create a placeholder slide for video."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    tb = slide.shapes.add_textbox(
        Inches(2), Inches(3), Inches(9), Inches(2)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = f"🎬 Видео-ответ Саши #{num}"
    run.font.size = Pt(32)
    run.font.color.rgb = PINK
    run.font.bold = True
    p = tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    r2 = p.add_run()
    r2.text = "(воспроизвести на устройстве)"
    r2.font.size = Pt(18)
    r2.font.color.rgb = LIGHT_PURPLE
    return slide


def add_tinder_slide(name, fields, photo_path):
    """Create a tinder profile slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, RGBColor(0x1a, 0x0a, 0x2e))
    if os.path.exists(photo_path):
        slide.shapes.add_picture(
            photo_path, Inches(0.8), Inches(1.2),
            width=Inches(4), height=Inches(5)
        )
    tb = slide.shapes.add_textbox(
        Inches(5.3), Inches(0.8), Inches(7.5), Inches(6.5)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = name
    run.font.size = Pt(32)
    run.font.color.rgb = PINK
    run.font.bold = True
    for label, value in fields:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"{label}: "
        r.font.size = Pt(18)
        r.font.color.rgb = VIOLET
        r.font.bold = True
        r2 = p.add_run()
        r2.text = value
        r2.font.size = Pt(18)
        r2.font.color.rgb = LIGHT_PURPLE
    return slide


def add_quote_slide(pair_num, quote_a, quote_b):
    """Create a quotes comparison slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    tb = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.5), Inches(12), Inches(1)
    )
    tf = tb.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = f"Пара {pair_num} — Какая цитата Сашина?"
    run.font.size = Pt(24)
    run.font.color.rgb = VIOLET
    run.font.bold = True
    tb_a = slide.shapes.add_textbox(
        Inches(0.5), Inches(1.8), Inches(6), Inches(4.5)
    )
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True
    p = tf_a.paragraphs[0]
    r = p.add_run()
    r.text = "Цитата А"
    r.font.size = Pt(18)
    r.font.color.rgb = GOLD
    r.font.bold = True
    p2 = tf_a.add_paragraph()
    p2.space_before = Pt(12)
    r2 = p2.add_run()
    r2.text = quote_a
    r2.font.size = Pt(20)
    r2.font.color.rgb = WHITE
    tb_b = slide.shapes.add_textbox(
        Inches(6.8), Inches(1.8), Inches(6), Inches(4.5)
    )
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p = tf_b.paragraphs[0]
    r = p.add_run()
    r.text = "Цитата Б"
    r.font.size = Pt(18)
    r.font.color.rgb = GOLD
    r.font.bold = True
    p2 = tf_b.add_paragraph()
    p2.space_before = Pt(12)
    r2 = p2.add_run()
    r2.text = quote_b
    r2.font.size = Pt(20)
    r2.font.color.rgb = WHITE
    return slide


def add_result_slide(answer_text):
    """Create a result/answer slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, RGBColor(0x15, 0x08, 0x25))
    tb = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(11), Inches(1.5)
    )
    tf = tb.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = "🎉 Ответ:"
    run.font.size = Pt(36)
    run.font.color.rgb = GOLD
    run.font.bold = True
    tb2 = slide.shapes.add_textbox(
        Inches(1), Inches(3), Inches(11), Inches(3.5)
    )
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
    run2 = tf2.paragraphs[0].add_run()
    run2.text = answer_text
    run2.font.size = Pt(28)
    run2.font.color.rgb = WHITE
    return slide


def add_nhie_slide(text, bg_img=None):
    """Create a Never Have I Ever slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if bg_img and os.path.exists(bg_img):
        slide.shapes.add_picture(
            bg_img, Inches(0), Inches(0),
            width=prs.slide_width, height=prs.slide_height
        )
    else:
        add_bg(slide)
    tb = slide.shapes.add_textbox(
        Inches(1), Inches(1.5), Inches(11), Inches(1.5)
    )
    tf = tb.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = "Never have I ever..."
    run.font.size = Pt(28)
    run.font.color.rgb = PINK
    run.font.bold = True
    tb2 = slide.shapes.add_textbox(
        Inches(1), Inches(3.5), Inches(11), Inches(3)
    )
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
    run2 = tf2.paragraphs[0].add_run()
    run2.text = text
    run2.font.size = Pt(32)
    run2.font.color.rgb = WHITE
    run2.font.bold = True
    return slide


def add_hot_slide(photo_path):
    """Create a Hot or Not screenshot slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    if os.path.exists(photo_path):
        slide.shapes.add_picture(
            photo_path, Inches(3.5), Inches(0.5),
            width=Inches(6), height=Inches(6.5)
        )
    return slide


# ===== BUILD PRESENTATION =====

# Slide 0: Title
add_title_slide(
    "Девичник Лизы",
    "Она выходит замуж — и мы счастливы, "
    "что собрались по этому поводу.\nПоехали!"
)

# Round 1: Лиза vs Саша
add_title_slide(
    "Лиза vs Саша",
    "Лиза отвечает на вопрос — "
    "а потом смотрим, что думает Саша.",
    "Раунд 1"
)

questions = [
    "Кто из вас первый написал после первого свидания?",
    "Какое самое странное блюдо Лиза готовила?",
    "Из-за чего была ваша самая тупая ссора?",
    "Какую привычку Лизы ты сначала терпел, а потом полюбил?",
    "Что Лиза делает, когда думает, что ты не видишь?",
    "Кто из вас храпит?",
    "Какой фильм Лиза пересматривала больше всего?",
    "Что Черри любит больше — тебя или Лизу?",
    "Какой Лизин «контрол-фрик» момент тебя "
    "больше всего веселит?",
    "Какой подарок от Лизы был самым неожиданным?",
    "Если бы Лиза не стала продюсером — кем бы она была?",
    "Самый неловкий момент на свидании?",
    "Что бы ты сказал Лизе-из-прошлого, "
    "которая ещё тебя не знает?",
]

for i, q in enumerate(questions, 1):
    add_question_slide(i, q)
    add_video_slide(i)

# Round 2: Лиза выбирает мужа
add_title_slide(
    "Лиза выбирает мужа",
    "Саша, конечно, очень хорош. Но мы нашли самых "
    "красивых мужчин в стране — ты просто обязана "
    "взглянуть.\nНадеемся, они заслужили свайп вправо "
    "напоследок от незамужней женщины",
    "Раунд 2"
)

candidates = [
    ("Кандидат №1 — Дмитрий, 30", "photos/dmitriy.png", [
        ("Город", "Москва (удалёнка, так что "
         "технически — диван)"),
        ("Работа", 'Senior-разработчик / "работаю 4 часа '
         'в день, остальное — думаю"'),
        ("О себе", '"Код пишу в халате. Дурь иногда '
         "помогает сконцентрироваться — это не я "
         "сказал, это Стив Джобс. Однажды починил прод "
         'в 3 ночи, не вставая с дивана. Герой."'),
        ("Плюс", "Зарплата 400к, философски относится "
         "к проблемам"),
        ("Минус", '"Щас доделаю" может означать от 5 '
         "минут до 3 дней"),
        ("Ищу", "Ту, которая не трогает мой второй "
         "монитор и не спрашивает зачем мне "
         "три клавиатуры"),
    ]),
    ("Кандидат №2 — Родриго, 28", "photos/rodrigo.png", [
        ("Город", "Мадрид"),
        ("Работа", "Таинственный предприниматель"),
        ("О себе", '"Люблю текилу, маленьких собак и '
         'женщин, которые всё контролируют"'),
        ("Плюс", "Сам готовит и убирает"),
        ("Минус", "Засыпает в 22:00"),
        ("Ищу", "Ту, которая составит мне "
         "расписание на жизнь"),
    ]),
    ("Кандидат №3 — Ахмед, 31", "photos/ahmed.png", [
        ("Город", "Нальчик → Москва"),
        ("Работа", '"Бизнес, сестра, не спрашивай"'),
        ("О себе", '"Увезу в горы. Там тихо, красиво, '
         'и никто не найдёт. Шучу. Или нет."'),
        ("Плюс", "Решает любой вопрос за 10 минут"),
        ("Минус", "Мама всегда на первом месте"),
        ("Ищу", "Ту, за которую не стыдно перед родом"),
    ]),
    ("Кандидат №4 — Алехандро, 27",
     "photos/alejandro.png", [
        ("Город", "Барселона"),
        ("Работа", '"Художник... пишу, рисую, живу"'),
        ("О себе", '"Я покажу тебе такой закат, что ты '
         "забудешь как тебя зовут. Потом я тоже забуду "
         'как тебя зовут, но это не важно."'),
        ("Плюс", "Говорит так, что хочется верить"),
        ("Минус", "Исчезает на 3 недели после "
         "идеального свидания"),
        ("Ищу", "Музу. На одну ночь. Максимум три."),
    ]),
    ("Кандидат №5 — Алижон, 38", "photos/alijon.png", [
        ("Город", "Ташкент → Москва (с 2009)"),
        ("Работа", "Акушер-гинеколог / "
         '"делаю кесарево за 500 долларов, гарантия"'),
        ("О себе", '"Руки золотые. В прямом смысле. '
         "3000 детей принял. Чипсы с беконом не ем — "
         'харам. Но тебя на руках носить буду."'),
        ("Плюс", "Никогда не упадёт в обморок "
         "при виде крови"),
        ("Минус", "Все разговоры сводит к родам"),
        ("Ищу", "Ту, которая родит минимум четверых"),
    ]),
    ("Кандидат №6 — Геннадий, 45",
     "photos/gennadiy.png", [
        ("Город", "Воронеж (переезжать не планирует)"),
        ("Работа", 'Начальник отдела в "Водоканале"'),
        ("О себе", '"Могу достать любые трубы. Дача — '
         "15 соток, баня с бассейном. По пятницам — "
         "шашлык. Люблю русское кино, особенно "
         '«Русский роман» и сериал «Гадалка»."'),
        ("Плюс", "Всё починит, всё построит, "
         "кран не течёт"),
        ("Минус", "В отпуск только в Анапу. "
         "Принципиально."),
        ("Ищу", "Хозяйственную. Чтоб борщ "
         "и чтоб не пилила."),
    ]),
]

for name, photo, fields in candidates:
    add_tinder_slide(name, fields, photo)

# Round 3: Цитаты
add_title_slide(
    "Саша или Писатель?",
    "Две цитаты. Одна от Саши, другая от "
    "латиноамериканского писателя.\n"
    "Лиза угадывает. Ошиблась — шот 🥃",
    "Раунд 3"
)

quotes = [
    (
        "«Любовь — это не то, что находишь. "
        "Это то, что строишь каждый день заново.»",
        "«Нууууу, любовь — это, понимаете... типо "
        "мехового треугольника. Вот вечером лежим "
        "вместе и лежим спокойненько.»",
        "Цитата Б — это Саша!\n"
        "А цитата А — Габриэль Гарсиа Маркес"
    ),
    (
        "«Кто-то вот любит окрошку с квасом, а кто-то "
        "с кефиром, ну или с айраном. Чисто "
        "гипотетически, абстрактно выражаясь. Но вы "
        "вместе всё равно, в одну сторону смотрите!»",
        "«Любовь так коротка, а забвение так длинно.»",
        "Цитата А — это Саша! Да, с окрошкой.\n"
        "А цитата Б — Пабло Неруда"
    ),
    (
        "«Люблю тебя не за то, кто ты, а за то, "
        "кем я становлюсь рядом с тобой.»",
        "«Люблю тебя потому что самая лучшая "
        "женщина во вселенной!»",
        "Цитата Б — это Саша! Коротко и по делу.\n"
        "А цитата А — Габриэль Гарсиа Маркес"
    ),
    (
        "«Я пью потому что ты пьёшь. "
        "Это любовь же и есть.»",
        "«Я пью, чтобы окружающие "
        "становились интереснее.»",
        "Цитата А — это Саша! Романтик и алкоголик "
        "в одном флаконе.\n"
        "А цитата Б — Эрнест Хемингуэй"
    ),
]

for i, (qa, qb, ans) in enumerate(quotes, 1):
    add_quote_slide(i, qa, qb)
    add_result_slide(ans)

# Round 4: Never Have I Ever
add_title_slide(
    "Never Have I Ever",
    "Кто делал — пьёт! 🍹",
    "Раунд 4"
)

nhie_items = [
    ("...напивалась и писала бывшему",
     "stock/couple1.png"),
    ("...просыпалась в чужой квартире и не помнила "
     "как туда попала", "stock/s1.png"),
    ("...пьяная признавалась кому-то в любви",
     "stock/s2.png"),
    ("...флиртовала ради бесплатного коктейля "
     "(и получала его)", "stock/s3.png"),
    ("...целовалась с кем-то просто потому что "
     "было скучно", "stock/s4.png"),
    ("...занималась сексом в общественном месте",
     "stock/s5.png"),
    ("...отправляла ню не тому человеку",
     "stock/s6.png"),
    ("...сталкерила человека настолько глубоко, что "
     "случайно лайкнула фото трёхлетней давности",
     "stock/s7.png"),
    ("...придумывала дебильную отмазку, чтобы уйти "
     "с плохого свидания", "stock/s8.png"),
    ("...блевала в такси и делала вид что "
     "всё нормально", "stock/s9.png"),
    ("...уходила из бара с незнакомцем, а подругам "
     'писала «я дома»', "stock/s10.png"),
    ("...плакала пьяная в туалете клуба, а через "
     "10 минут танцевала как ни в чём не бывало",
     "stock/s11.png"),
    ("...секстилась с человеком, с которым даже "
     "не встречалась вживую", "stock/couple3.png"),
]

for text, bg in nhie_items:
    add_nhie_slide(text, bg)

# Final slide
add_title_slide(
    "За Лизу и Сашу!",
    "За любовь, за окрошку и за "
    "меховые треугольники! 🤍"
)

# P.S. Hot or Not
add_title_slide(
    "Hot or Not",
    "Бонус-раунд! Лиза, последний шанс оценить, "
    "от чего ты отказываешься. Реальные анкеты, "
    "реальный ужас. Свайпаем! 🔥",
    "P.S."
)

for i in range(1, 10):
    add_hot_slide(f"photos/hot_{i}.png")

# Save
output = "devichnik_liza.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
