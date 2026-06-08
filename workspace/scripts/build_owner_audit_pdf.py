#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path

import pandas as pd
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


BASE = Path(__file__).resolve().parents[1]
OUT_DIR = BASE / "outputs" / "wb_audit"
PDF_PATH = OUT_DIR / "K-Beauty_Hub_WB_owner_audit_FINAL_v4_irp_formula_caution_2026-05-28.pdf"

FONT_DIR = Path("/System/Library/Fonts/Supplemental")
FONT = "AuditArial"
FONT_BOLD = "AuditArialBold"
FONT_ITALIC = "AuditArialItalic"

NAVY = colors.HexColor("#17324D")
TEAL = colors.HexColor("#0F766E")
MINT = colors.HexColor("#E7F5F0")
BLUE_TINT = colors.HexColor("#EAF2FF")
CORAL = colors.HexColor("#C85250")
AMBER = colors.HexColor("#F4B740")
INK = colors.HexColor("#1F2937")
MUTED = colors.HexColor("#6B7280")
LINE = colors.HexColor("#D8DEE9")
SOFT = colors.HexColor("#F6F8FB")
WHITE = colors.white


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont(FONT, str(FONT_DIR / "Arial.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, str(FONT_DIR / "Arial Bold.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_ITALIC, str(FONT_DIR / "Arial Italic.ttf")))


def styles():
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=base["Title"],
            fontName=FONT_BOLD,
            fontSize=31,
            leading=34,
            textColor=WHITE,
            alignment=TA_LEFT,
            spaceAfter=12,
        ),
        "cover_subtitle": ParagraphStyle(
            "cover_subtitle",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=12,
            leading=16,
            textColor=colors.HexColor("#DDEBEE"),
            spaceAfter=8,
        ),
        "h1": ParagraphStyle(
            "h1",
            parent=base["Heading1"],
            fontName=FONT_BOLD,
            fontSize=17,
            leading=21,
            textColor=NAVY,
            spaceBefore=8,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=base["Heading2"],
            fontName=FONT_BOLD,
            fontSize=12.5,
            leading=15,
            textColor=TEAL,
            spaceBefore=8,
            spaceAfter=5,
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=9.4,
            leading=12.3,
            textColor=INK,
            spaceAfter=5,
        ),
        "small": ParagraphStyle(
            "small",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=7.6,
            leading=9.4,
            textColor=INK,
            spaceAfter=2,
        ),
        "small_muted": ParagraphStyle(
            "small_muted",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=7.3,
            leading=9,
            textColor=MUTED,
            spaceAfter=2,
        ),
        "table": ParagraphStyle(
            "table",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=7.4,
            leading=8.8,
            textColor=INK,
        ),
        "table_bold": ParagraphStyle(
            "table_bold",
            parent=base["BodyText"],
            fontName=FONT_BOLD,
            fontSize=7.4,
            leading=8.8,
            textColor=INK,
        ),
        "metric": ParagraphStyle(
            "metric",
            parent=base["BodyText"],
            fontName=FONT_BOLD,
            fontSize=13.5,
            leading=15.5,
            textColor=NAVY,
            alignment=TA_CENTER,
        ),
        "metric_label": ParagraphStyle(
            "metric_label",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=7.3,
            leading=9,
            textColor=MUTED,
            alignment=TA_CENTER,
        ),
        "callout": ParagraphStyle(
            "callout",
            parent=base["BodyText"],
            fontName=FONT_BOLD,
            fontSize=10,
            leading=13,
            textColor=NAVY,
        ),
        "center": ParagraphStyle(
            "center",
            parent=base["BodyText"],
            fontName=FONT,
            fontSize=8,
            leading=10,
            textColor=INK,
            alignment=TA_CENTER,
        ),
    }


def clean_text(text: str) -> str:
    return str(text).replace("₽", "руб.")


def p(text: str, style: ParagraphStyle) -> Paragraph:
    text = clean_text(text).replace("&", "&amp;")
    return Paragraph(text, style)


def rub(value: float | int | None, digits: int = 0) -> str:
    if value is None or pd.isna(value):
        return ""
    if abs(float(value)) >= 1_000_000 and digits == 0:
        number = f"{float(value) / 1_000_000:.2f}".replace(".", ",")
        return f"{number} млн руб."
    number = f"{float(value):,.{digits}f}".replace(",", " ").replace(".", ",")
    return f"{number} руб."


def qty(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{float(value):,.0f}".replace(",", " ")


def pct(value: float | int | None, digits: int = 1) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{float(value):.{digits}f}%".replace(".", ",")


def short(text: str, limit: int = 58) -> str:
    text = str(text)
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def paragraph_table(rows, col_widths, header=True, font_size=7.4, leading=8.8):
    st = styles()
    normal = ParagraphStyle("tbl_dynamic", parent=st["table"], fontSize=font_size, leading=leading)
    bold = ParagraphStyle("tbl_dynamic_bold", parent=st["table_bold"], fontSize=font_size, leading=leading)
    header_style = ParagraphStyle(
        "tbl_dynamic_header",
        parent=st["table_bold"],
        fontSize=font_size,
        leading=leading,
        textColor=WHITE,
    )
    data = []
    for i, row in enumerate(rows):
        data.append([p(cell, header_style if header and i == 0 else normal) for cell in row])
    table = Table(data, colWidths=col_widths, repeatRows=1 if header else 0, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY if header else WHITE),
                ("TEXTCOLOR", (0, 0), (-1, 0), WHITE if header else INK),
                ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.35, LINE),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, SOFT]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def metric_cards(items, cols=4, total_width=None):
    st = styles()
    total_width = total_width or (A4[0] - 36 * mm)
    rows = []
    for i in range(0, len(items), cols):
        chunk = items[i : i + cols]
        while len(chunk) < cols:
            chunk.append(("", ""))
        rows.append(
            [
                [
                    p(value, st["metric"]),
                    Spacer(1, 1.5 * mm),
                    p(label, st["metric_label"]),
                ]
                for value, label in chunk
            ]
        )
    table = Table(rows, colWidths=[total_width / cols] * cols, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), WHITE),
                ("BOX", (0, 0), (-1, -1), 0.4, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table


def callout(text: str, bg=MINT):
    st = styles()
    table = Table([[p(text, st["callout"])]], colWidths=[A4[0] - 36 * mm], hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), bg),
                ("BOX", (0, 0), (-1, -1), 0.4, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return table


def bar_chart(title: str, rows, width=500, bar_color=TEAL, value_fmt=qty):
    rows = [(str(label), float(value or 0)) for label, value in rows if value is not None and not pd.isna(value)]
    height = 42 + max(1, len(rows)) * 22
    drawing = Drawing(width, height)
    drawing.add(String(0, height - 14, clean_text(title), fontName=FONT_BOLD, fontSize=10, fillColor=NAVY))
    max_value = max([value for _, value in rows] + [1])
    y = height - 36
    label_w = width * 0.40
    bar_w = width * 0.43
    for idx, (label, value) in enumerate(rows):
        drawing.add(String(0, y + 3, clean_text(short(label, 32)), fontName=FONT, fontSize=7.3, fillColor=INK))
        drawing.add(Rect(label_w, y, bar_w, 9, fillColor=colors.HexColor("#E5E7EB"), strokeColor=None))
        drawing.add(Rect(label_w, y, max(1, bar_w * value / max_value), 9, fillColor=bar_color, strokeColor=None))
        drawing.add(String(label_w + bar_w + 8, y + 1, value_fmt(value), fontName=FONT_BOLD, fontSize=7.2, fillColor=INK))
        if idx < len(rows) - 1:
            drawing.add(Rect(0, y - 7, width, 0.25, fillColor=LINE, strokeColor=None))
        y -= 22
    return drawing


def page_chrome(canvas, doc) -> None:
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(NAVY)
    canvas.rect(0, height - 9 * mm, width, 9 * mm, stroke=0, fill=1)
    canvas.setFillColor(WHITE)
    canvas.setFont(FONT_BOLD, 7.5)
    canvas.drawString(doc.leftMargin, height - 6 * mm, "K-Beauty Hub / аудит WB / 28.05.2026")
    canvas.setFont(FONT, 7.2)
    canvas.drawRightString(width - doc.rightMargin, height - 6 * mm, f"стр. {doc.page}")
    canvas.setFillColor(MUTED)
    canvas.setFont(FONT, 7)
    canvas.drawString(doc.leftMargin, 8 * mm, "Основа: выгрузки WB, себестоимость из шаблона, данные кабинета WB и проверенные справки WB/MPStats.")
    canvas.restoreState()


def read_csv(name: str) -> pd.DataFrame:
    return pd.read_csv(OUT_DIR / name)


def build_story():
    st = styles()
    story = []

    supply = read_csv("supply_recommendations_by_sku_summary_2026-05-28.csv")
    irp = read_csv("irp_overpay_by_sku_actual_0_60pct.csv")
    storage = read_csv("paid_storage_beauty_by_sku_2026-04-01_2026-05-28.csv")
    ab = read_csv("ab_test_problem_cards_2026-05-28.csv")

    width = A4[0] - 36 * mm

    # Cover
    cover_box = Table(
        [
            [
                [
                    p("K-Beauty Hub", st["cover_title"]),
                    p("Финальный аудит кабинета Wildberries", st["cover_subtitle"]),
                    p("Итоги работы за 28 мая 2026: экономика, ИРП, хранение, поставки, реклама, карточки и план роста.", st["cover_subtitle"]),
                    Spacer(1, 8 * mm),
                    metric_cards(
                        [
                            ("≈5,09 млн ₽", "потенциал прибыли от правильной дозагрузки на 30 дней"),
                            ("≈172 тыс ₽/мес", "оценка надбавки логистики по косметике при ИРП 0,6%"),
                            ("≈118,6 тыс ₽/мес", "текущее хранение по категории «Красота»"),
                            ("55,7%", "фактическая локальность"),
                        ],
                        cols=2,
                        total_width=width - 28 * mm,
                    ),
                ]
            ]
        ],
        colWidths=[width],
        rowHeights=[210 * mm],
        hAlign="LEFT",
    )
    cover_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                ("LEFTPADDING", (0, 0), (-1, -1), 14 * mm),
                ("RIGHTPADDING", (0, 0), (-1, -1), 14 * mm),
                ("TOPPADDING", (0, 0), (-1, -1), 19 * mm),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 14 * mm),
            ]
        )
    )
    story.append(cover_box)
    story.append(PageBreak())

    # Executive summary
    story.append(p("1. Главный вывод", st["h1"]))
    story.append(
        callout(
            "Кабинет не выглядит слабым: есть товары с сильным спросом и положительной юнит-экономикой. "
            "Главная проблема сейчас управленческая: часть товара лежит не там, где спрос, фактическая локальность 55,7%, "
            "реклама местами ведется без опоры на фактическую прибыль, а карточки надо довести до 10/10."
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(
        metric_cards(
            [
                ("34,26 млн ₽", "заказы по воронке за 28.04-28.05"),
                ("-17,7%", "динамика суммы заказов к прошлому периоду"),
                ("55,7%", "фактическая локальность"),
                ("0,6%", "ИРП на 28.05.2026"),
                ("20 015 шт", "остаток на складе WB"),
                ("94%", "выкуп по воронке"),
                ("5%", "CTR магазина"),
                ("30%", "конверсия в заказ"),
            ],
            cols=4,
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(p("Деньги и эффект", st["h2"]))
    story.append(
        paragraph_table(
            [
                ["Блок", "Что показывает аудит", "Оценка эффекта", "Что делать"],
                ["Распределение остатков", "Спрос есть, но часть заказов нелокальна или упирается в дефицит по регионам.", "≈5,09 млн ₽ прибыли до рекламы при продаже рекомендованной 30-дневной дозагрузки.", "Сделать первую волну 7 452 шт и дальше считать поставки через MPStats."],
                ["ИРП/локальность", "Факт локальности 55,7%. ИРП на 28.05.2026 — 0,6%.", "Ориентир надбавки к логистике: ≈172 тыс ₽/мес по косметике; ≈204 тыс ₽/мес по кабинету.", "Поднять локальность через региональные поставки."],
                ["Хранение", "Хранение по красоте уже ≈118,6 тыс ₽/мес.", "Это не критично для оборота, но расход надо контролировать на SKU.", "Отделить ходовые остатки от залежалых, следить за объемом >1 л."],
                ["Карточки", "Средняя заполненность 8-8,5/10, есть карточки с большим трафиком и CTR 1-2%.", "Рост CTR/конверсии даст продажи без пропорционального роста бюджета.", "Контент до 10/10 и регулярные A/B-тесты."],
            ],
            [27 * mm, 58 * mm, 48 * mm, 44 * mm],
            font_size=7.15,
            leading=8.7,
        )
    )

    story.append(PageBreak())

    # Funnel and finance
    story.append(p("2. Динамика продаж и воронки", st["h1"]))
    story.append(
        paragraph_table(
            [
                ["Метрика", "Текущий период", "Предыдущий период", "Динамика"],
                ["Показы", "10,01 млн", "13,50 млн", "-25,9%"],
                ["Переходы в карточку", "531 тыс", "760 тыс", "-30,1%"],
                ["Заказы, шт", "24 673", "27 844", "-11,4%"],
                ["Сумма заказов", "34,26 млн ₽", "41,64 млн ₽", "-17,7%"],
                ["CTR", "5%", "—", "держится на среднем уровне"],
                ["Конверсия в корзину", "15%", "—", "нормальная, но есть проблемные SKU"],
                ["Конверсия в заказ", "30%", "—", "сильная сторона кабинета"],
                ["Выкуп", "94%", "—", "сильная сторона кабинета"],
            ],
            [48 * mm, 42 * mm, 42 * mm, 45 * mm],
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(
        callout(
            "Просадка начинается сверху: стало меньше показов и переходов. Значит главный рычаг роста — видимость, наличие в нужных регионах, кликабельность карточек и управляемая реклама."
        )
    )

    story.append(PageBreak())

    # IPR and locality
    story.append(p("3. ИЛ, ИРП и логистика", st["h1"]))
    story.append(
        metric_cards(
            [
                ("1,03", "индекс локализации"),
                ("0,6%", "ИРП на 28.05.2026"),
                ("55,7%", "фактическая локальность"),
                ("10,9 тыс", "нелокальные заказы"),
            ],
            cols=4,
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(
        callout(
            "Фактическая локальность сейчас 55,7%. "
            "ИРП не рассчитывается как простая разница от локальности: WB считает его через распределение заказов и коэффициенты по артикулам. "
            "Оценка влияния текущего ИРП на логистику: около 204 тыс ₽/мес по кабинету и 172 тыс ₽/мес по косметике."
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        bar_chart(
            "Оценка надбавки логистики по косметике при ИРП 0,6%, ₽/мес",
            [
                (row["Название"], row["Переплата_ИРП_месяц"])
                for _, row in irp[irp["Косметика"].astype(str).str.lower().eq("true")].head(7).iterrows()
            ],
            width=width,
            bar_color=CORAL,
            value_fmt=rub,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        paragraph_table(
            [
                ["Регион", "Доля заказов", "Локальность", "Вывод"],
                ["Дальневосточный и Сибирский", "15,5%", "1,4%", "Ключевой провал: спрос есть, локального товара почти нет."],
                ["Северо-Западный", "8,7%", "29,3%", "Нужно довезти ходовые SKU, особенно в Петербург/Шушары."],
                ["Уральский", "6,5%", "49,7%", "Есть потенциал быстро улучшить локальность точечной поставкой."],
                ["Приволжский", "—", "59,4%", "Точечные поставки могут дать быстрый эффект."],
            ],
            [52 * mm, 28 * mm, 28 * mm, 69 * mm],
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(p("Управленческий вывод", st["h2"]))
    story.append(p("Задача не в том, чтобы просто нарастить общий остаток. Нужно переложить запас туда, где уже есть спрос: Сибирь/Дальний Восток, Северо-Запад, Урал, Приволжье и Юг. Пока это не сделано, реклама будет разгонять спрос в тех регионах, где WB везет товар дороже и дольше.", st["body"]))

    story.append(PageBreak())

    # Stock distribution
    story.append(p("4. Куда поставлять и сколько", st["h1"]))
    story.append(
        metric_cards(
            [
                ("15 757 шт", "30-дневная региональная дозагрузка с буфером 20%"),
                ("7 452 шт", "первая аккуратная волна на 14 дней"),
                ("≈5,09 млн ₽", "прибыль до рекламы при продаже 30-дневного объема"),
                ("≈130 тыс ₽/мес", "ориентир потенциальной экономии логистики"),
            ],
            cols=4,
        )
    )
    story.append(Spacer(1, 5 * mm))
    supply_rows = [["SKU", "Товар", "Куда в первую очередь", "30 дней", "14 дней", "Прибыль"]]
    for _, row in supply.head(8).iterrows():
        profit = rub(row["profit_if_recommended_qty_sells"]) if row["profit_if_recommended_qty_sells"] else "нужна себ-ть"
        supply_rows.append(
            [
                str(int(row["nm"])),
                short(row["Название_sku"], 36),
                short(row["regions"], 55),
                qty(row["recommended_supply_qty"]),
                qty(row["first_wave_14d_qty"]),
                profit,
            ]
        )
    story.append(paragraph_table(supply_rows, [22 * mm, 39 * mm, 59 * mm, 20 * mm, 20 * mm, 17 * mm], font_size=6.8, leading=8.1))
    story.append(Spacer(1, 5 * mm))
    story.append(
        bar_chart(
            "Потенциал прибыли от рекомендованного объема, топ SKU",
            [
                (row["Название_sku"], row["profit_if_recommended_qty_sells"])
                for _, row in supply.head(6).iterrows()
                if float(row["profit_if_recommended_qty_sells"] or 0) > 0
            ],
            width=width,
            bar_color=TEAL,
            value_fmt=rub,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        callout(
            "Особый приоритет — Celimax Dual Barrier Creamy Toner (nm 868516027): товар дает оборот и прибыль, но остаток низкий и локальность слабая. "
            "Рекомендация: 5 608 шт на 30 дней, первая волна 2 622 шт, фокус на Юг, Центр, Сибирь/Дальний Восток, Приволжье и Северо-Запад."
        )
    )

    story.append(PageBreak())

    # Storage
    story.append(p("5. Платное хранение", st["h1"]))
    story.append(
        metric_cards(
            [
                ("105 148 ₽", "факт 01.05-28.05"),
                ("≈112 659 ₽/мес", "30-дневный пересчет майского периода"),
                ("262 601 ₽", "точный SKU-отчет за 58 дней"),
                ("≈118 644 ₽/мес", "категория «Красота»"),
            ],
            cols=4,
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(
        bar_chart(
            "Топ SKU по хранению в категории «Красота», ₽/мес",
            [
                (row["Артикул продавца"], row["Хранение_месяц30"])
                for _, row in storage.head(8).iterrows()
            ],
            width=width,
            bar_color=AMBER,
            value_fmt=rub,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        paragraph_table(
            [
                ["Причина", "Сумма за 58 дней", "Что делать"],
                ["До 1 литра / ниже базы", "182 182 ₽", "Контроль оборачиваемости, не завозить медленные SKU сверх потребности."],
                ["Свыше 1 литра / выше базы", "47 196 ₽ по красоте", "Особенно следить за шампунями/пенками и крупными упаковками."],
                ["Основные склады", "Казань, Коледино, Краснодар, Тула, Екатеринбург", "Сопоставить хранение со спросом региона, убрать перекосы."],
            ],
            [49 * mm, 45 * mm, 83 * mm],
        )
    )

    story.append(PageBreak())

    # Content and AB
    story.append(p("6. Карточки, A/B-тесты и «Оригинал»", st["h1"]))
    story.append(
        callout(
            "Средняя заполненность карточек сейчас примерно 8-8,5/10. Для ключевых SKU цель — 10/10: фото, инфографика, видео/обложка, характеристики, описание, SEO-ключи и ответы на возражения. Это влияет на позиции в поиске и эффективность рекламы."
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(p("С чего начать A/B-тесты", st["h2"]))
    ab_rows = [["SKU", "Товар", "Проблема", "Показы", "CTR", "Заказ"]]
    cosmetic_ab = ab[~ab["Предмет"].astype(str).str.contains("Роботы", case=False, na=False)].head(8)
    for _, row in cosmetic_ab.iterrows():
        ab_rows.append(
            [
                str(int(row["Артикул WB"])),
                short(row["Название"], 36),
                short(row["Проблема"], 38),
                qty(row["Показы"]),
                pct(row["CTR"], 0),
                pct(row["Конверсия в заказ, %"], 0),
            ]
        )
    story.append(paragraph_table(ab_rows, [22 * mm, 52 * mm, 47 * mm, 23 * mm, 16 * mm, 17 * mm], font_size=6.9, leading=8.2))
    story.append(Spacer(1, 4 * mm))
    story.append(p("Что тестировать", st["h2"]))
    for item in [
        "Главное фото: товар крупнее, понятный эффект, бренд и объем читаются в выдаче.",
        "Первые 3 слайда: эффект, кому подходит, как применять, чем отличается.",
        "Обложка видео и порядок слайдов.",
        "Название и первые SEO-ключи без переспама.",
    ]:
        story.append(p("• " + item, st["body"]))
    story.append(Spacer(1, 3 * mm))
    story.append(p("Значок «Оригинал»", st["h2"]))
    story.append(p("На части брендов и карточек значок «Оригинал» уже может быть подключен. Задача — расширить его по возможности на все брендовые карточки, где есть подтверждающие документы: Celimax, Farm Stay, Round Lab, ENOUGH и другие бренды. Это усиливает доверие к карточке и помогает отстроиться от продавцов без подтвержденной оригинальности.", st["body"]))
    story.append(
        paragraph_table(
            [
                ["Что подготовить", "Комментарий"],
                ["Договор с официальным дистрибьютором / дилером / поставщиком", "Лучший вариант для подтверждения цепочки поставки."],
                ["УПД, ТТН, инвойсы, импортные документы", "Нужны актуальные документы на партии товара."],
                ["Цепочка документов от правообладателя до продавца", "Особенно если нет прямого договора с правообладателем."],
                ["Проверка карточек", "Бренд на фото/упаковке и бренд в карточке должны совпадать с документами."],
            ],
            [67 * mm, 110 * mm],
        )
    )

    return story


def main() -> None:
    register_fonts()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(PDF_PATH),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title="K-Beauty Hub WB Audit 2026-05-28",
        author="Codex",
    )
    story = build_story()
    doc.build(story, onFirstPage=page_chrome, onLaterPages=page_chrome)
    print(PDF_PATH)


if __name__ == "__main__":
    main()
