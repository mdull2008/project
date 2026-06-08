from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
LOGO_PATH = ROOT / "assets" / "spbstu_logo_02.png"
OUTPUT_PATHS = (
    Path.home() / "Desktop" / "my_project.pptx",
    ROOT / "my_project.pptx",
)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

GREEN = RGBColor(58, 170, 53)      # SPbPU primary green, HEX #3AA935
DARK_GREEN = RGBColor(36, 65, 40)  # SPbPU green gradient
MID_GREEN = RGBColor(86, 151, 91)
ORANGE = RGBColor(219, 73, 40)     # SPbPU orange gradient, humanities pair
LIGHT_ORANGE = RGBColor(243, 152, 105)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
LIGHT_BG = RGBColor(246, 248, 244)
MUTED_TEXT = RGBColor(92, 103, 95)

FONT = "Arial"
TOTAL_SLIDES = 10


def apply_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = LIGHT_BG


def add_rect(slide, left, top, width, height, color, transparency=0):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.fill.transparency = transparency
    shape.line.fill.background()
    return shape


def style_run(run, size, color=BLACK, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold


def add_text(slide, left, top, width, height, text, size=24, color=BLACK, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = align
    style_run(paragraph.runs[0], size=size, color=color, bold=bold)
    return box


def add_multiline(slide, left, top, width, height, lines, size=20, color=BLACK, bold_first=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.space_after = Pt(8)
        style_run(paragraph.runs[0], size=size, color=color, bold=bold_first and index == 0)
    return box


def add_bullets(slide, left, top, width, height, bullets, size=19, color=BLACK, marker_color=GREEN):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True

    for index, bullet in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.space_after = Pt(9)
        marker = paragraph.add_run()
        marker.text = "■ "
        style_run(marker, size=size, color=marker_color, bold=True)
        run = paragraph.add_run()
        run.text = bullet
        style_run(run, size=size, color=color)

    return box


def add_logo(slide):
    if LOGO_PATH.exists():
        slide.shapes.add_picture(str(LOGO_PATH), Inches(0.55), Inches(0.34), width=Inches(1.75))
    else:
        add_text(slide, Inches(0.55), Inches(0.36), Inches(2.2), Inches(0.35), "СПбПУ", 16, GREEN, True)


def add_header(slide, title, section, number, timing):
    add_logo(slide)
    add_text(slide, Inches(3.0), Inches(0.35), Inches(6.8), Inches(0.34), section, 10, MUTED_TEXT)
    add_text(slide, Inches(3.0), Inches(0.69), Inches(7.8), Inches(0.52), title, 24, DARK_GREEN, True)
    add_rect(slide, Inches(0.55), Inches(1.28), Inches(12.2), Inches(0.035), GREEN)
    add_text(slide, Inches(10.35), Inches(0.45), Inches(2.4), Inches(0.34), timing, 10, MUTED_TEXT, align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(12.05), Inches(6.95), Inches(0.7), Inches(0.25), f"{number}/{TOTAL_SLIDES}", 10, MUTED_TEXT, align=PP_ALIGN.RIGHT)


def add_card(slide, left, top, width, height, title, body, accent=GREEN):
    add_rect(slide, left, top, width, height, WHITE)
    add_rect(slide, left, top, Inches(0.12), height, accent)
    add_text(slide, left + Inches(0.28), top + Inches(0.23), width - Inches(0.45), Inches(0.35), title, 15, DARK_GREEN, True)
    add_text(slide, left + Inches(0.28), top + Inches(0.72), width - Inches(0.45), height - Inches(0.86), body, 17, BLACK)


def add_metric(slide, left, top, width, label, value, color=GREEN, note=""):
    add_text(slide, left, top, width, Inches(0.25), label, 14, BLACK, True)
    add_rect(slide, left, top + Inches(0.35), width, Inches(0.18), RGBColor(218, 226, 216))
    add_rect(slide, left, top + Inches(0.35), int(width * value / 100), Inches(0.18), color)
    add_text(slide, left + width - Inches(0.75), top + Inches(0.2), Inches(0.75), Inches(0.25), f"{value:g}%", 13, color, True, PP_ALIGN.RIGHT)
    if note:
        add_text(slide, left, top + Inches(0.62), width, Inches(0.28), note, 10, MUTED_TEXT)


def add_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)

    add_rect(slide, Inches(8.2), Inches(0), Inches(5.13), SLIDE_H, DARK_GREEN)
    add_rect(slide, Inches(8.2), Inches(0), Inches(5.13), Inches(2.1), GREEN)
    add_rect(slide, Inches(8.2), Inches(2.1), Inches(3.35), Inches(1.2), MID_GREEN)
    add_rect(slide, Inches(11.55), Inches(2.1), Inches(1.78), Inches(1.2), LIGHT_ORANGE)
    add_rect(slide, Inches(8.2), Inches(3.3), Inches(2.2), Inches(4.2), ORANGE)
    add_rect(slide, Inches(10.4), Inches(3.3), Inches(2.93), Inches(4.2), DARK_GREEN)

    add_logo(slide)
    add_text(slide, Inches(0.62), Inches(1.34), Inches(6.75), Inches(0.35), "Курсовая работа · выступление 5–7 минут", 14, MUTED_TEXT)
    add_text(
        slide,
        Inches(0.58),
        Inches(1.9),
        Inches(7.25),
        Inches(1.65),
        "Топонимы Санкт-Петербурга\nв лексиконе носителя русского языка",
        34,
        DARK_GREEN,
        True,
    )
    add_text(
        slide,
        Inches(0.62),
        Inches(3.9),
        Inches(6.8),
        Inches(0.65),
        "Как официальные, разговорные и исторические названия распределяются между активным и пассивным запасом речи",
        18,
        BLACK,
    )
    add_multiline(
        slide,
        Inches(0.62),
        Inches(5.45),
        Inches(6.4),
        Inches(1.1),
        [
            "Гуманитарный институт · Высшая школа лингвистики и педагогики",
            "Выполнил: К.А. Шин · Руководитель: В.А. Белов",
            "Санкт-Петербург, 2026",
        ],
        size=12,
        color=MUTED_TEXT,
    )
    add_text(slide, Inches(11.75), Inches(6.95), Inches(0.9), Inches(0.25), f"1/{TOTAL_SLIDES}", 10, WHITE, align=PP_ALIGN.RIGHT)


def add_intro_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Введение", "Актуальность и объект исследования", 2, "0:30–1:15")

    add_rect(slide, Inches(0.75), Inches(1.65), Inches(7.15), Inches(4.75), WHITE)
    add_text(slide, Inches(1.08), Inches(1.95), Inches(6.3), Inches(0.4), "Почему тема важна", 20, DARK_GREEN, True)
    add_bullets(
        slide,
        Inches(1.08),
        Inches(2.55),
        Inches(6.25),
        Inches(2.6),
        [
            "Санкт-Петербург воспринимается через сеть городских названий: улиц, площадей, районов, станций метро.",
            "Один объект может иметь официальное, разговорное и историческое имя.",
            "Для лингвистики важно понять, какие названия реально живут в речи, а какие остаются культурной отсылкой.",
        ],
        size=17,
    )
    add_text(slide, Inches(1.1), Inches(5.65), Inches(6.35), Inches(0.35), "Проблемный вопрос", 16, ORANGE, True)
    add_text(slide, Inches(1.1), Inches(5.98), Inches(6.35), Inches(0.35), "Почему одни говорят «Питер» и «Лиговка», а другие — «Санкт-Петербург» и «Лиговский проспект»?", 14, BLACK)

    add_rect(slide, Inches(8.25), Inches(1.65), Inches(4.25), Inches(4.75), DARK_GREEN)
    add_text(slide, Inches(8.65), Inches(2.0), Inches(3.45), Inches(0.35), "Объект", 18, WHITE, True)
    add_text(slide, Inches(8.65), Inches(2.45), Inches(3.45), Inches(0.85), "топонимическая лексика в речи носителей русского языка", 16, WHITE)
    add_rect(slide, Inches(8.65), Inches(3.55), Inches(3.35), Inches(0.05), LIGHT_ORANGE)
    add_text(slide, Inches(8.65), Inches(3.9), Inches(3.45), Inches(0.35), "Предмет", 18, WHITE, True)
    add_text(slide, Inches(8.65), Inches(4.35), Inches(3.45), Inches(1.0), "включение петербургских топонимов в активный и пассивный компоненты лексикона", 16, WHITE)


def add_goal_tasks_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Цель, задачи и гипотеза", "Логика исследования", 3, "1:15–2:00")

    add_card(
        slide,
        Inches(0.75),
        Inches(1.65),
        Inches(5.35),
        Inches(1.15),
        "Цель",
        "Описать, как петербургские топонимы представлены в лексиконе современных носителей русского языка.",
    )
    add_card(
        slide,
        Inches(6.45),
        Inches(1.65),
        Inches(5.35),
        Inches(1.15),
        "Гипотеза",
        "Топонимы Санкт-Петербурга образуют многослойную систему: официальные, разговорные и исторические названия распределены по разным сферам общения.",
        accent=ORANGE,
    )

    add_rect(slide, Inches(0.75), Inches(3.25), Inches(11.05), Inches(2.6), WHITE)
    add_text(slide, Inches(1.08), Inches(3.55), Inches(10.3), Inches(0.35), "Основные задачи", 19, DARK_GREEN, True)
    add_bullets(
        slide,
        Inches(1.08),
        Inches(4.05),
        Inches(10.1),
        Inches(1.3),
        [
            "рассмотреть подходы к изучению топонимов в отечественной лингвистике;",
            "обосновать методику эмпирического исследования;",
            "провести анкетирование и интерпретировать полученные данные;",
            "сформулировать выводы о связи топонимов с возрастом, опытом города и сферой общения.",
        ],
        size=15,
    )


def add_theory_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Теоретическая база", "Ключевые понятия", 4, "2:00–2:45")

    add_card(
        slide,
        Inches(0.75),
        Inches(1.65),
        Inches(3.65),
        Inches(2.0),
        "Урбаноним",
        "Название внутригородского объекта: улицы, площади, моста, станции метро или района.",
        accent=GREEN,
    )
    add_card(
        slide,
        Inches(4.85),
        Inches(1.65),
        Inches(3.65),
        Inches(2.0),
        "Активный топонимикон",
        "Названия, которые человек спонтанно употребляет в речи и переписке.",
        accent=MID_GREEN,
    )
    add_card(
        slide,
        Inches(8.95),
        Inches(1.65),
        Inches(3.65),
        Inches(2.0),
        "Пассивный топонимикон",
        "Названия, которые человек узнаёт, но почти не использует самостоятельно.",
        accent=ORANGE,
    )

    add_rect(slide, Inches(0.75), Inches(4.15), Inches(11.85), Inches(1.65), WHITE)
    add_text(slide, Inches(1.05), Inches(4.43), Inches(10.9), Inches(0.35), "Научная опора", 18, DARK_GREEN, True)
    add_bullets(
        slide,
        Inches(1.05),
        Inches(4.88),
        Inches(10.8),
        Inches(0.75),
        [
            "А.В. Суперанская: имена собственные включены в общую систему языка и получают коннотации.",
            "Е.В. Иванцова: количество узнаваемых топонимов часто больше числа активно используемых.",
            "О.Н. Минюшова: топонимы могут быть логоэпистемами — носителями культурно-исторического знания.",
        ],
        size=13,
    )


def add_method_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Методика и материал", "Как собирались данные", 5, "2:45–3:30")

    add_card(
        slide,
        Inches(0.75),
        Inches(1.65),
        Inches(3.7),
        Inches(1.55),
        "Выборка",
        "54 полные анкеты; доминируют респонденты 18–25 лет, в основном молодые жители и приезжие студенты.",
    )
    add_card(
        slide,
        Inches(4.8),
        Inches(1.65),
        Inches(3.7),
        Inches(1.55),
        "Материал анкеты",
        "12 топонимов: Эрмитаж, Невский, Купчино, Лиговка, Питер, Ленинград и др.",
        accent=ORANGE,
    )
    add_card(
        slide,
        Inches(8.85),
        Inches(1.65),
        Inches(3.7),
        Inches(1.55),
        "Формат",
        "Множественный выбор сфер употребления для каждого названия.",
    )

    add_rect(slide, Inches(0.75), Inches(3.75), Inches(11.8), Inches(2.25), WHITE)
    add_text(slide, Inches(1.05), Inches(4.0), Inches(4.7), Inches(0.35), "Контексты употребления", 18, DARK_GREEN, True)
    add_bullets(
        slide,
        Inches(1.05),
        Inches(4.48),
        Inches(5.25),
        Inches(1.15),
        [
            "семья и близкие",
            "друзья / приятели",
            "работа или учёба",
            "соцсети и мессенджеры",
            "пассивное знание / незнание",
        ],
        size=14,
    )
    add_text(slide, Inches(7.0), Inches(4.0), Inches(4.8), Inches(0.35), "Ограничение", 18, DARK_GREEN, True)
    add_text(
        slide,
        Inches(7.0),
        Inches(4.48),
        Inches(4.85),
        Inches(1.15),
        "Выводы не описывают всех петербуржцев, а прежде всего показывают речевую практику молодой городской аудитории.",
        16,
        BLACK,
    )
    add_rect(slide, Inches(0.75), Inches(6.35), Inches(7.2), Inches(0.18), GREEN)
    add_rect(slide, Inches(7.95), Inches(6.35), Inches(4.6), Inches(0.18), ORANGE)


def add_results_core_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Результаты: активное ядро", "N = 54 · проценты от числа респондентов", 6, "3:30–4:30")

    add_rect(slide, Inches(0.72), Inches(1.65), Inches(7.0), Inches(4.8), WHITE)
    add_text(slide, Inches(1.0), Inches(1.95), Inches(6.3), Inches(0.35), "Самые активные названия", 18, DARK_GREEN, True)
    add_metric(slide, Inches(1.0), Inches(2.55), Inches(5.8), "Питер · общение с друзьями", 92.6, GREEN, "ядро неформального городского лексикона")
    add_metric(slide, Inches(1.0), Inches(3.45), Inches(5.8), "Невский · общение с друзьями", 83.3, GREEN, "универсальный центральный навигатор")
    add_metric(slide, Inches(1.0), Inches(4.35), Inches(5.8), "Питер · соцсети и мессенджеры", 79.6, MID_GREEN, "цифровая коммуникация закрепляет разговорную форму")
    add_metric(slide, Inches(1.0), Inches(5.25), Inches(5.8), "Невский · работа / учёба", 63.0, MID_GREEN, "название проходит и через более формальные сферы")

    add_rect(slide, Inches(8.05), Inches(1.65), Inches(4.55), Inches(4.8), DARK_GREEN)
    add_text(slide, Inches(8.45), Inches(1.98), Inches(3.75), Inches(0.38), "Что важно проговорить", 18, WHITE, True)
    add_bullets(
        slide,
        Inches(8.45),
        Inches(2.55),
        Inches(3.75),
        Inches(3.2),
        [
            "«Питер» — не просто сокращение, а эмоционально нейтральное имя «своего» города.",
            "«Невский» совмещает разговорность и официальную узнаваемость.",
            "Активное ядро связано с ежедневными маршрутами и частой коммуникацией.",
        ],
        size=15,
        color=WHITE,
        marker_color=LIGHT_ORANGE,
    )
    add_rect(slide, Inches(8.05), Inches(6.1), Inches(3.1), Inches(0.18), GREEN)
    add_rect(slide, Inches(11.15), Inches(6.1), Inches(1.45), Inches(0.18), ORANGE)


def add_results_passive_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Результаты: пассив и различия", "Что знают, но не всегда используют", 7, "4:30–5:30")

    add_rect(slide, Inches(0.72), Inches(1.65), Inches(6.15), Inches(4.7), WHITE)
    add_text(slide, Inches(1.0), Inches(1.95), Inches(5.55), Inches(0.35), "Пассивный запас шире активного", 18, DARK_GREEN, True)
    add_metric(slide, Inches(1.0), Inches(2.55), Inches(5.2), "Ленинград · «знаю, но не использую»", 51.9, ORANGE, "для молодёжи — исторический и поколенческий маркер")
    add_metric(slide, Inches(1.0), Inches(3.45), Inches(5.2), "Рыбацкое · «знаю, но не использую»", 44.4, ORANGE, "название известно, но часто не включено в личный маршрут")
    add_metric(slide, Inches(1.0), Inches(4.35), Inches(5.2), "Пять углов · пассивное знание", 22.2, ORANGE, "локальное название с низкой активностью")
    add_metric(slide, Inches(1.0), Inches(5.25), Inches(5.2), "Лиговка · незнание", 20.4, ORANGE, "показывает зависимость от опыта района")

    add_rect(slide, Inches(7.25), Inches(1.65), Inches(5.15), Inches(4.7), DARK_GREEN)
    add_text(slide, Inches(7.65), Inches(1.98), Inches(4.35), Inches(0.35), "Интерпретация", 18, WHITE, True)
    add_bullets(
        slide,
        Inches(7.65),
        Inches(2.55),
        Inches(4.2),
        Inches(2.55),
        [
            "Пассивное знание не равно незнанию: название может быть культурно знакомым, но не речевым.",
            "Стаж проживания расширяет активный запас локальных названий.",
            "У старших горожан «Ленинград» и районные названия чаще остаются живыми словами.",
        ],
        size=15,
        color=WHITE,
        marker_color=LIGHT_ORANGE,
    )
    add_text(slide, Inches(7.65), Inches(5.62), Inches(4.2), Inches(0.42), "Главный вывод: топонимикон зависит от опыта города.", 15, WHITE, True)


def add_conclusion_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Заключение", "Итоги курсовой работы", 8, "5:30–6:20")

    add_rect(slide, Inches(0.75), Inches(1.65), Inches(3.55), Inches(3.25), GREEN)
    add_rect(slide, Inches(4.55), Inches(1.65), Inches(3.55), Inches(3.25), MID_GREEN)
    add_rect(slide, Inches(8.35), Inches(1.65), Inches(3.55), Inches(3.25), ORANGE)

    add_text(slide, Inches(1.05), Inches(2.0), Inches(2.95), Inches(0.45), "Официальное", 20, WHITE, True)
    add_text(slide, Inches(1.05), Inches(2.62), Inches(2.95), Inches(1.55), "Дворцовая площадь, Эрмитаж, Исаакиевский собор: высокая узнаваемость и культурный слой.", 16, WHITE)
    add_text(slide, Inches(4.85), Inches(2.0), Inches(2.95), Inches(0.45), "Разговорное", 20, WHITE, True)
    add_text(slide, Inches(4.85), Inches(2.62), Inches(2.95), Inches(1.55), "Питер, Невский, Лиговка: повседневная навигация, друзья, соцсети, маршруты.", 16, WHITE)
    add_text(slide, Inches(8.65), Inches(2.0), Inches(2.95), Inches(0.45), "Историческое", 20, WHITE, True)
    add_text(slide, Inches(8.65), Inches(2.62), Inches(2.95), Inches(1.55), "Ленинград: знаком почти всем, но у молодых респондентов чаще находится в пассиве.", 16, WHITE)

    add_rect(slide, Inches(0.75), Inches(5.35), Inches(11.15), Inches(1.05), WHITE)
    add_text(slide, Inches(1.05), Inches(5.58), Inches(2.05), Inches(0.4), "Итог", 18, DARK_GREEN, True)
    add_text(
        slide,
        Inches(2.1),
        Inches(5.52),
        Inches(9.45),
        Inches(0.55),
        "Границы между активным и пассивным топонимиконом задаются не грамматикой, а возрастом, опытом города, кругом общения и цифровыми практиками.",
        15,
        BLACK,
    )
    add_text(slide, Inches(0.78), Inches(6.72), Inches(10.0), Inches(0.28), "Перспектива: расширить выборку, добавить старшее поколение и сравнить Петербург с другими миллионниками.", 11, MUTED_TEXT)


def add_sources_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Источники литературы", "Основные работы, использованные в курсовой", 9, "6:20–6:45")

    add_rect(slide, Inches(0.75), Inches(1.65), Inches(11.5), Inches(4.9), WHITE)
    sources = [
        "Агеева Р.А. Ономастический код культуры. М.: Наука, 1990.",
        "Иванцова Е.В. Топонимия в идиолексиконе диалектоносителя // Вестник ЧелГУ. 2007. № 17.",
        "Кузнецова А.В., Петрулевич И.А. Топонимика как инструмент конструирования городской идентичности // Социолингвистика. 2021. № 2.",
        "Минюшова О.Н. Топонимы-логоэпистемы в русской ономастике. Казань, 2006.",
        "Никонов В.А. Введение в топонимику. М.: Наука, 1965.",
        "Суперанская А.В. Общая теория имён собственных. М.: Наука, 1973.",
        "Зыкова И.В. Полевой принцип описания значения топонима // Вопросы лингвистики. 2022. № 4.",
        "Анкета «Топонимы Санкт-Петербурга в лексиконе носителя русского языка», 2026.",
    ]
    add_bullets(slide, Inches(1.08), Inches(1.98), Inches(10.85), Inches(4.1), sources, size=12, marker_color=ORANGE)


def add_thanks_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, DARK_GREEN)
    add_rect(slide, Inches(0), Inches(0), Inches(4.15), SLIDE_H, GREEN)
    add_rect(slide, Inches(4.15), Inches(0), Inches(2.05), Inches(3.75), MID_GREEN)
    add_rect(slide, Inches(4.15), Inches(3.75), Inches(2.05), Inches(3.75), ORANGE)
    add_logo(slide)
    add_text(slide, Inches(6.8), Inches(2.35), Inches(5.3), Inches(0.85), "Спасибо за внимание!", 38, WHITE, True)
    add_text(slide, Inches(6.85), Inches(3.35), Inches(4.95), Inches(0.55), "Готов ответить на вопросы", 22, WHITE)
    add_text(slide, Inches(6.85), Inches(5.55), Inches(5.2), Inches(0.35), "Курсовая работа · Топонимы Санкт-Петербурга", 12, RGBColor(223, 237, 222))
    add_text(slide, Inches(11.75), Inches(6.95), Inches(0.9), Inches(0.25), f"10/{TOTAL_SLIDES}", 10, WHITE, align=PP_ALIGN.RIGHT)


def build_presentation(output_paths=OUTPUT_PATHS):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    add_title_slide(prs)
    add_intro_slide(prs)
    add_goal_tasks_slide(prs)
    add_theory_slide(prs)
    add_method_slide(prs)
    add_results_core_slide(prs)
    add_results_passive_slide(prs)
    add_conclusion_slide(prs)
    add_sources_slide(prs)
    add_thanks_slide(prs)

    saved_paths = []
    for output_path in output_paths:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(output_path)
        saved_paths.append(output_path)

    return saved_paths


if __name__ == "__main__":
    for saved_path in build_presentation():
        print(f"Presentation saved to: {saved_path}")
