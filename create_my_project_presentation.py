from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


OUTPUT_PATH = Path.home() / "Desktop" / "my_project.pptx"


def set_text(frame, text, font_size=24, bold=False, color=RGBColor(255, 255, 255)):
    frame.clear()
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = PP_ALIGN.LEFT
    run = paragraph.runs[0]
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Arial"


def add_text_box(slide, left, top, width, height, text, font_size=24, bold=False, color=RGBColor(255, 255, 255)):
    box = slide.shapes.add_textbox(left, top, width, height)
    set_text(box.text_frame, text, font_size=font_size, bold=bold, color=color)
    return box


def add_bullet_list(slide, left, top, width, height, title, bullets):
    add_text_box(slide, left, top, width, Inches(0.5), title, font_size=28, bold=True)
    box = slide.shapes.add_textbox(left, top + Inches(0.7), width, height - Inches(0.7))
    frame = box.text_frame
    frame.clear()

    for index, bullet in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = bullet
        paragraph.level = 0
        paragraph.font.size = Pt(21)
        paragraph.font.color.rgb = RGBColor(235, 241, 255)
        paragraph.font.name = "Arial"


def add_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_accent_bar(slide, color):
    shape = slide.shapes.add_shape(
        1,
        Inches(0),
        Inches(0),
        Inches(0.18),
        Inches(7.5),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_footer(slide, number):
    footer = add_text_box(
        slide,
        Inches(11.8),
        Inches(7.0),
        Inches(1.1),
        Inches(0.3),
        f"{number}/4",
        font_size=12,
        color=RGBColor(195, 210, 235),
    )
    footer.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT


def build_presentation(output_path=OUTPUT_PATH):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    slides = [
        {
            "title": "Мой проект",
            "subtitle": "Краткая презентация идеи, ценности и плана развития",
            "background": RGBColor(20, 35, 70),
            "accent": RGBColor(91, 192, 190),
        },
        {
            "title": "Идея и цель",
            "bullets": [
                "Создать понятное и полезное решение для выбранной аудитории",
                "Сократить ручные действия и упростить ежедневные процессы",
                "Собрать обратную связь и быстро улучшать продукт",
            ],
            "background": RGBColor(28, 48, 92),
            "accent": RGBColor(255, 185, 95),
        },
        {
            "title": "Ключевые компоненты",
            "bullets": [
                "Интерфейс: простой пользовательский сценарий",
                "Логика: обработка данных и автоматизация задач",
                "Результат: измеримые показатели пользы и качества",
            ],
            "background": RGBColor(32, 58, 67),
            "accent": RGBColor(115, 219, 180),
        },
        {
            "title": "Следующие шаги",
            "bullets": [
                "Подготовить прототип и показать его первым пользователям",
                "Определить метрики успеха и критерии готовности",
                "Запланировать улучшения на основе реальных отзывов",
            ],
            "background": RGBColor(44, 38, 78),
            "accent": RGBColor(244, 117, 117),
        },
    ]

    for index, spec in enumerate(slides, start=1):
        slide = prs.slides.add_slide(blank_layout)
        add_background(slide, spec["background"])
        add_accent_bar(slide, spec["accent"])

        if index == 1:
            add_text_box(
                slide,
                Inches(0.9),
                Inches(2.25),
                Inches(9.5),
                Inches(0.9),
                spec["title"],
                font_size=52,
                bold=True,
            )
            add_text_box(
                slide,
                Inches(0.95),
                Inches(3.3),
                Inches(8.6),
                Inches(0.7),
                spec["subtitle"],
                font_size=24,
                color=RGBColor(218, 232, 252),
            )
        else:
            add_bullet_list(
                slide,
                Inches(0.9),
                Inches(1.15),
                Inches(11.2),
                Inches(4.8),
                spec["title"],
                spec["bullets"],
            )

        add_footer(slide, index)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path


if __name__ == "__main__":
    saved_path = build_presentation()
    print(f"Presentation saved to: {saved_path}")
