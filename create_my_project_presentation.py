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

FONT = "Onest"
FALLBACK_FONT = "Arial"


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
    add_text(slide, Inches(12.05), Inches(6.95), Inches(0.7), Inches(0.25), f"{number}/4", 10, MUTED_TEXT, align=PP_ALIGN.RIGHT)


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
    add_text(slide, Inches(11.75), Inches(6.95), Inches(0.9), Inches(0.25), "1/4", 10, WHITE, align=PP_ALIGN.RIGHT)


def add_research_frame_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Исследовательская рамка", "Тема, материал и методика", 2, "1:00–2:30")

    add_card(
        slide,
        Inches(0.75),
        Inches(1.75),
        Inches(3.7),
        Inches(1.55),
        "Цель",
        "Описать, как петербургские топонимы представлены в лексиконе современных носителей русского языка.",
    )
    add_card(
        slide,
        Inches(4.8),
        Inches(1.75),
        Inches(3.7),
        Inches(1.55),
        "Гипотеза",
        "Официальные, разговорные и исторические названия работают в разных функциональных зонах.",
        accent=ORANGE,
    )
    add_card(
        slide,
        Inches(8.85),
        Inches(1.75),
        Inches(3.7),
        Inches(1.55),
        "Материал",
        "Онлайн-анкета Google Forms: 54 полных ответа, 12 топонимов, множественный выбор контекстов.",
    )

    add_rect(slide, Inches(0.75), Inches(3.85), Inches(11.8), Inches(2.4), WHITE)
    add_text(slide, Inches(1.05), Inches(4.1), Inches(4.7), Inches(0.35), "Как фиксировалось употребление", 18, DARK_GREEN, True)
    add_bullets(
        slide,
        Inches(1.05),
        Inches(4.58),
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
    add_text(slide, Inches(7.0), Inches(4.08), Inches(4.8), Inches(0.35), "Фокус интерпретации", 18, DARK_GREEN, True)
    add_text(
        slide,
        Inches(7.0),
        Inches(4.58),
        Inches(4.85),
        Inches(1.15),
        "Выборка в основном отражает речевую практику молодых городских носителей и приезжих студентов, поэтому результаты читаются как модель молодежного топонимикона.",
        16,
        BLACK,
    )
    add_rect(slide, Inches(0.75), Inches(6.55), Inches(7.2), Inches(0.18), GREEN)
    add_rect(slide, Inches(7.95), Inches(6.55), Inches(4.6), Inches(0.18), ORANGE)


def add_results_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Ключевые результаты анкетирования", "N = 54 · проценты от числа респондентов", 3, "2:30–5:00")

    add_rect(slide, Inches(0.72), Inches(1.65), Inches(7.0), Inches(4.8), WHITE)
    add_text(slide, Inches(1.0), Inches(1.95), Inches(6.3), Inches(0.35), "Активность и пассивность топонимов", 18, DARK_GREEN, True)
    add_metric(slide, Inches(1.0), Inches(2.55), Inches(5.8), "Питер · общение с друзьями", 92.6, GREEN, "ядро неформального городского лексикона")
    add_metric(slide, Inches(1.0), Inches(3.45), Inches(5.8), "Невский · общение с друзьями", 83.3, GREEN, "универсальный центральный навигатор")
    add_metric(slide, Inches(1.0), Inches(4.35), Inches(5.8), "Ленинград · «знаю, но не использую»", 51.9, ORANGE, "историческое название у молодёжи уходит в пассив")
    add_metric(slide, Inches(1.0), Inches(5.25), Inches(5.8), "Рыбацкое · «знаю, но не использую»", 44.4, ORANGE, "локальная периферия без личного маршрута")

    add_rect(slide, Inches(8.05), Inches(1.65), Inches(4.55), Inches(4.8), DARK_GREEN)
    add_text(slide, Inches(8.45), Inches(1.98), Inches(3.75), Inches(0.38), "Что важно проговорить", 18, WHITE, True)
    add_bullets(
        slide,
        Inches(8.45),
        Inches(2.55),
        Inches(3.75),
        Inches(3.2),
        [
            "«Питер» и «Невский» проходят через почти все сферы общения.",
            "Культурные символы известны, но не всегда нужны в ежедневной навигации.",
            "Разговорные и районные названия зависят от маршрутов, возраста и круга общения.",
        ],
        size=15,
        color=WHITE,
        marker_color=LIGHT_ORANGE,
    )
    add_rect(slide, Inches(8.05), Inches(6.1), Inches(3.1), Inches(0.18), GREEN)
    add_rect(slide, Inches(11.15), Inches(6.1), Inches(1.45), Inches(0.18), ORANGE)


def add_conclusion_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(slide, "Вывод: лексикон горожанина многослоен", "Итоговая модель и перспективы", 4, "5:00–7:00")

    add_rect(slide, Inches(0.75), Inches(1.68), Inches(3.6), Inches(3.55), GREEN)
    add_rect(slide, Inches(4.55), Inches(1.68), Inches(3.6), Inches(3.55), MID_GREEN)
    add_rect(slide, Inches(8.35), Inches(1.68), Inches(3.6), Inches(3.55), ORANGE)

    add_text(slide, Inches(1.05), Inches(2.0), Inches(3.0), Inches(0.45), "1. Официальное", 20, WHITE, True)
    add_text(slide, Inches(1.05), Inches(2.62), Inches(3.0), Inches(1.7), "Дворцовая площадь, Эрмитаж, Исаакиевский собор: высокая узнаваемость и культурный слой.", 17, WHITE)
    add_text(slide, Inches(4.85), Inches(2.0), Inches(3.0), Inches(0.45), "2. Разговорное", 20, WHITE, True)
    add_text(slide, Inches(4.85), Inches(2.62), Inches(3.0), Inches(1.7), "Питер, Невский, Лиговка: активная навигация, друзья, соцсети, повседневные маршруты.", 17, WHITE)
    add_text(slide, Inches(8.65), Inches(2.0), Inches(3.0), Inches(0.45), "3. Историческое", 20, WHITE, True)
    add_text(slide, Inches(8.65), Inches(2.62), Inches(3.0), Inches(1.7), "Ленинград: почти всем знаком, но у молодых респондентов чаще находится в пассиве.", 17, WHITE)

    add_rect(slide, Inches(0.75), Inches(5.65), Inches(11.2), Inches(0.95), WHITE)
    add_text(slide, Inches(1.05), Inches(5.89), Inches(10.55), Inches(0.4), "Защитный тезис", 17, DARK_GREEN, True)
    add_text(
        slide,
        Inches(3.05),
        Inches(5.88),
        Inches(8.55),
        Inches(0.46),
        "Границы между активным и пассивным топонимиконом задаются не грамматикой, а возрастом, опытом города, кругом общения и цифровыми практиками.",
        15,
        BLACK,
    )
    add_text(slide, Inches(0.78), Inches(6.78), Inches(10.0), Inches(0.28), "Дальше: расширить выборку, добавить старшее поколение и сравнить Петербург с другими миллионниками.", 11, MUTED_TEXT)


def build_presentation(output_paths=OUTPUT_PATHS):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    add_title_slide(prs)
    add_research_frame_slide(prs)
    add_results_slide(prs)
    add_conclusion_slide(prs)

    saved_paths = []
    for output_path in output_paths:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(output_path)
        saved_paths.append(output_path)

    return saved_paths


if __name__ == "__main__":
    for saved_path in build_presentation():
        print(f"Presentation saved to: {saved_path}")
