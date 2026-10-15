#!/usr/bin/env python3
"""Приложения к договору купли-продажи от 18.09.2026: сборка данных.

Приложения переделываются под факт. Раньше приложение 1 было
«первоначальная партия» без цены, приложение 2 — весь перечень с
ценами, приложение 3 — перечень без цен. Теперь так:

  № 1 — задвижки, отгружено в Керчь (УПД № 55 от 23.09.2026)
  № 2 — фильтры, отгружено в Петербург (УПД № 56 от 23.09.2026)
  № 3 — запчасти, отгружено в Петербург (УПД № 57 от 29.09.2026)
  № 4 — перечень к отгрузке контейнером, еще не отгружено

То есть первые три приложения описывают уже переданный товар и цены,
по которым он передан, а четвертое — то, что поедет следующим.

Факт по отгрузкам лежит в upd_shipments.py и взят с подписанных УПД.
Перечень к отгрузке — тот же свод, что и в бланке для склада: список
менеджера плюс склад с вычетом задвоений, уже отгруженное снято.

Цена по договору для приложения 4 берется из действующих приложений
2 и 3 — по номеру чертежа, иначе по очищенному наименованию. Где цены
нет, ставится «по согласованию»: по пункту 3.5 договора такая позиция
не считается проданной до согласования цены. Своих цен не выдумываем
и учетную цену склада в договор не тащим — она внутренняя и сделками
не проверена.

Результат: outputs/marine_equipment/Приложения_к_договору_18.09.2026.docx
"""

import pathlib
import re
import sys
from collections import defaultdict

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import price_fixes  # noqa: E402
import upd_shipments  # noqa: E402
from build_container_spb import plural  # noqa: E402
from build_logistics_compare import rows_all  # noqa: E402
from build_picking_list import TAIL, group_order, ru_of  # noqa: E402
from build_spb_list_xlsx import group_of, read as read_spb  # noqa: E402
from build_weight_xlsx import du as du_of  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "outputs" / "marine_equipment"

CONTRACT = "Договору купли-продажи товара с отсрочкой платежа и " \
           "последующей реализацией Покупателем от «18» сентября 2026 г."
AGREE = "по согласованию"


def norm_draw(v):
    return re.sub(r"[^0-9a-zа-я.\-]", "", str(v or "").lower())


def clean(name):
    s = re.sub(r"\(.*?\)", " ", str(name or "").lower())
    s = re.sub(r"[^0-9a-zа-я]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def price_index():
    """Цены из действующих приложений 2 и 3: по чертежу и по названию."""
    by_draw, by_name = {}, {}
    for r in read_spb():
        p = (r["Цена по договору"] or "").strip()
        if not p:
            continue
        d = norm_draw(r["Обозначение"])
        if d:
            by_draw.setdefault(d, p)
        by_name.setdefault(clean(r["Наименование"]), p)
    return by_draw, by_name


def normalize_price(text):
    """Цена из договора к одному виду: рубли, а не тысячи.

    В действующих приложениях цена записана как попало: «250-300»,
    «85-90000», «15-20», «3000 - 4000». В колонке, озаглавленной
    «руб.», «250-300» читается как двести пятьдесят рублей, а имелось
    в виду двести пятьдесят тысяч. Правило простое и проверяемое:
    граница меньше 1 000 означает тысячи, потому что самая дешевая
    позиция во всем договоре стоит 1 500 руб. Судовой арматуры по
    250 рублей не бывает.

    Возвращает готовый текст и признак, что запись пришлось привести.
    """
    raw = (text or "").strip()
    nums = re.findall(r"\d+(?:[  ]\d{3})*", raw)
    if not nums:
        return raw, False
    vals = [int(n.replace(" ", "").replace("\u00a0", "")) for n in nums]
    fixed = [v * 1000 if v < 1000 else v for v in vals]
    changed = fixed != vals
    out = " - ".join(f"{v:,.0f}".replace(",", " ") for v in fixed)
    return out, changed


def price_of(r, by_draw, by_name):
    """Цена по договору и откуда она взята."""
    fixed = price_fixes.fix(r["Наименование"])
    p = by_draw.get(norm_draw(r["Чертеж"]))
    how = "по чертежу"
    if not p:
        p = by_name.get(clean(r["Наименование"]))
        how = "по наименованию"
    if not p:
        return AGREE, "цены в договоре нет"
    if fixed and re.search(r"\d+\s*-\s*\d{3,}", p):
        # подтвержденная опечатка: «55-650» это 55-65 тысяч
        return f"55 000 - {fixed:,.0f}".replace(",", " "), \
            "опечатка в договоре исправлена"
    if AGREE in p.lower():
        return AGREE, "цены в договоре нет"
    out, changed = normalize_price(p)
    return out, (how + ", тысячи переведены в рубли" if changed else how)


def annex4():
    by_draw, by_name = price_index()
    rows = []
    for r in rows_all():
        price, how = price_of(r, by_draw, by_name)
        rows.append({
            "name": r["Наименование"],
            "draw": r.get("Чертеж") or "",
            "du": int(r.get("Ду") or du_of(r["Наименование"]) or 0),
            "ru": ru_of(r),
            "qty": int(r["Штук"]),
            "kg": round(r["вес"]),
            "price": price,
            "how": how,
            "group": group_of(r["Наименование"]),
        })
    by_group = defaultdict(list)
    for r in rows:
        by_group[r["group"]].append(r)
    # group_order ждет строки с ключами «Ду» и «вес» — отдаем их в
    # его виде, чтобы порядок групп совпадал с бланком для склада
    shaped = {g: [{"Ду": i["du"], "вес": i["kg"]} for i in items]
              for g, items in by_group.items()}
    blocks = []
    for g in group_order(shaped):
        items = sorted(by_group[g], key=lambda x: (-x["du"], x["name"]))
        blocks.append({
            "group": g,
            "qty": sum(i["qty"] for i in items),
            "kg": sum(i["kg"] for i in items),
            "items": items,
        })
    return rows, blocks




FONT = "Times New Roman"


def setup(doc):
    """А4, поля как в договоре, шрифт Times."""
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(2)
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(1.5)
        sec.top_margin = sec.bottom_margin = Cm(1.5)


def para(doc, text="", size=10, bold=False, italic=False, align=None,
         after=2, before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def shade(cell, color):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(el)


def fill(cell, text, size=9, bold=False, align=None, wrap=True):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align is not None:
        p.alignment = align
    run = p.add_run(str(text))
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if not wrap:
        cell.paragraphs[0].paragraph_format.keep_together = True


def table(doc, headers, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.autofit = False
    for i, (h, w) in enumerate(zip(headers, widths)):
        t.columns[i].width = Cm(w)
        cell = t.rows[0].cells[i]
        cell.width = Cm(w)
        fill(cell, h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        shade(cell, "D9E2F3")
    return t


def row(t, values, widths, bold=False, color=None, center=()):
    cells = t.add_row().cells
    for i, (v, w) in enumerate(zip(values, widths)):
        cells[i].width = Cm(w)
        fill(cells[i], v, bold=bold,
             align=WD_ALIGN_PARAGRAPH.CENTER if i in center else None)
        if color:
            shade(cells[i], color)
    return cells


def span(t, text, widths, color="D9E2F3", bold=True):
    """Строка-заголовок на всю ширину таблицы."""
    cells = t.add_row().cells
    merged = cells[0]
    for c in cells[1:]:
        merged = merged.merge(c)
    fill(merged, text, bold=bold)
    shade(merged, color)
    return merged


def signatures(doc):
    para(doc, before=10)
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    for i, text in enumerate([
        "ПРОДАВЕЦ:\n\n__________________ / Склярова Т.Е. /",
        "ПОКУПАТЕЛЬ:\n\n________________ / Белова Е.В. /",
    ]):
        c = t.rows[0].cells[i]
        c.width = Cm(9)
        c.text = ""
        for j, line in enumerate(text.split("\n")):
            p = c.paragraphs[0] if j == 0 else c.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(line)
            run.font.name = FONT
            run.font.size = Pt(10)
            run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def rub(n):
    return f"{n:,.0f}".replace(",", " ")


def header(doc, number, title, first=False):
    if not first:
        doc.add_page_break()
    para(doc, f"Приложение № {number}", size=11, bold=True,
         align=WD_ALIGN_PARAGRAPH.RIGHT)
    para(doc, title, size=12, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=2)
    para(doc, f"к {CONTRACT}", size=9,
         italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)


def sheet_shipped(doc, s, first=False):
    """Приложение по уже отгруженной партии: факт по УПД."""
    header(doc, s["annex"], s["title"], first)
    para(doc, f"Продавец: {upd_shipments.SELLER}", size=9)
    para(doc, f"Покупатель: {upd_shipments.BUYER}", size=9, after=6)
    para(doc, f"Товар передан по универсальному передаточному документу "
              f"№ {s['upd']} от {s['upd_date']}. Дата отгрузки "
              f"{s['shipped']}. Место получения: {s['city']}. "
              f"Получатель: ООО «ПКФ «АРМА».", size=9)
    para(doc, f"Перевозка: {s['carrier']}. {s['places']}. Объявленная "
              f"ценность {rub(s['declared'])} руб. Отправка за счет "
              f"Продавца (пункт 2.1 Договора).", size=9, after=8)

    # сумма ширин равна ширине полосы набора: 21 см минус поля 2 x 1,5
    w = [0.9, 6.0, 3.0, 1.6, 1.5, 2.5, 2.5]
    t = table(doc, ["№", "Наименование товара", "Обозначение",
                    "Состояние", "Кол-во, шт.", "Цена по Договору, руб.",
                    "Стоимость, руб."], w)
    for i, it in enumerate(s["items"], start=1):
        row(t, [i, it["name"], it["draw"] or "—", it.get("cond", ""),
                it["qty"], rub(it["price"]),
                rub(it["qty"] * it["price"])],
            w, center=(0, 3, 4, 5, 6))
    cells = t.add_row().cells
    merged = cells[0]
    for c in cells[1:4]:
        merged = merged.merge(c)
    fill(merged, "ИТОГО", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    for i, v in [(4, sum(it["qty"] for it in s["items"])), (5, ""),
                 (6, rub(s["total"]))]:
        cells[i].width = Cm(w[i])
        fill(cells[i], v, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for c in cells:
        shade(c, "D9E2F3")

    para(doc, before=6)
    notes = [it["note"] for it in s["items"] if it.get("note")]
    if notes:
        para(doc, "Состояние и особенности: " + "; ".join(notes) + ".",
             size=9)
    para(doc, f"Общая стоимость партии {rub(s['total'])} руб. Оплата "
              f"производится на условиях отсрочки платежа по разделу 4 "
              f"Договора: не позднее 5 рабочих дней с даты получения "
              f"оплаты от конечного покупателя.", size=9)
    para(doc, "Цена реализации конечному покупателю подлежит "
              "предварительному письменному согласованию с Продавцом "
              "(пункты 1.2 и 3.2 Договора).", size=9)
    signatures(doc)


def sheet_annex4(doc, data):
    header(doc, 4, "Перечень товара к отгрузке контейнером "
                   "в г. Санкт-Петербург")
    a = data["annex4"]
    para(doc, f"Продавец: {upd_shipments.SELLER}", size=9)
    para(doc, f"Покупатель: {upd_shipments.BUYER}", size=9, after=6)
    para(doc, f"Товар по настоящему приложению НЕ ОТГРУЖЕН. Отгрузка "
              f"производится контейнером из г. Владивостока в "
              f"г. Санкт-Петербург за счет Продавца (пункт 2.1 "
              f"Договора). Всего {a['positions']} позиций, "
              f"{rub(a['qty'])} шт, {a['kg'] / 1000:.2f} т.", size=9)
    para(doc, f"Цена по Договору определена по {a['agreed']} позициям. "
              f"По {a['open']} позициям цена не согласована и указана "
              f"«{AGREE}»: по пункту 3.5 Договора до согласования цены "
              f"такая позиция не считается проданной и не подлежит "
              f"передаче в составе партии.", size=9)
    para(doc, f"По {a['ranges']} позициям цена указана диапазоном. По "
              f"пункту 3.4 Договора конкретная цена за единицу "
              f"определяется отдельным письменным согласованием Сторон "
              f"до передачи соответствующей позиции Покупателю.", size=9,
         after=8)

    # сумма ширин равна ширине полосы набора: 21 см минус поля 2 x 1,5
    w = [0.9, 7.6, 3.3, 1.6, 4.6]
    t = table(doc, ["№", "Наименование", "Обозначение", "Кол-во, шт.",
                    "Цена по Договору, руб."], w)
    n = 0
    for b in data["annex4"]["blocks"]:
        span(t, f"{b['group']} — {len(b['items'])} "
                f"{plural(len(b['items']), 'позиция', 'позиции', 'позиций')}"
                f", {rub(b['qty'])} шт, {rub(b['kg'])} кг", w)
        for it in b["items"]:
            n += 1
            row(t, [n, it["name"], it["draw"] or "—", it["qty"],
                    it["price"]], w, center=(0, 3))
    span(t, f"ИТОГО {a['positions']} позиций, {rub(a['qty'])} шт, "
            f"{rub(a['kg'])} кг", w, color="BFBFBF")

    para(doc, before=6)
    para(doc, "Цены, записанные в приложениях к Договору в тысячах "
              "(например «250-300»), приведены здесь к рублям. Граница "
              "меньше 1 000 означает тысячи рублей: самая дешевая "
              "позиция Договора стоит 1 500 руб. Смысл цены при этом не "
              "меняется.", size=9)
    para(doc, "Порядок позиций соответствует порядку погрузки: сначала "
              "крупная арматура по убыванию диаметра, в конце "
              "вентиляционные головки.", size=9)
    para(doc, "Количество указано по учету на дату составления "
              "приложения и подлежит уточнению при погрузке. "
              "Фактически переданное количество фиксируется в "
              "универсальном передаточном документе на партию.", size=9)
    para(doc, "Настоящее приложение является неотъемлемой частью "
              "Договора (пункт 9.1).", size=9)
    signatures(doc)


def main():
    rows, blocks = annex4()
    agreed = [r for r in rows if r["price"] != AGREE]
    open_price = [r for r in rows if r["price"] == AGREE]
    ranges = [r for r in agreed if re.search(r"\d\s*-\s*\d", r["price"])]
    firm = [r for r in agreed if r not in ranges]
    data = {"annex4": {
        "blocks": blocks,
        "positions": len(rows),
        "qty": sum(r["qty"] for r in rows),
        "kg": sum(r["kg"] for r in rows),
        "agreed": len(agreed),
        "open": len(open_price),
        "open_kg": sum(r["kg"] for r in open_price),
        "ranges": len(ranges),
        "ranges_kg": sum(r["kg"] for r in ranges),
        "firm": len(firm),
        "firm_kg": sum(r["kg"] for r in firm),
    }}

    doc = Document()
    setup(doc)
    for i, s in enumerate(upd_shipments.SHIPMENTS):
        sheet_shipped(doc, dict(s, total=upd_shipments.total(s)),
                      first=(i == 0))
    sheet_annex4(doc, data)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Приложения_к_договору_18.09.2026.docx"
    doc.save(path)
    print(f"готово: {path}\n")
    for s in upd_shipments.SHIPMENTS:
        print(f"Приложение № {s['annex']}: {s['title']} — УПД № {s['upd']} "
              f"от {s['upd_date']}, {len(s['items'])} "
              f"{plural(len(s['items']), 'позиция', 'позиции', 'позиций')}, "
              f"{sum(i['qty'] for i in s['items'])} шт, "
              f"{rub(upd_shipments.total(s))} руб")
    print(f"Отгружено всего на "
          f"{rub(sum(upd_shipments.total(s) for s in upd_shipments.SHIPMENTS))}"
          f" руб")
    a = data["annex4"]
    print(f"\nПриложение № 4: {a['positions']} позиций, {rub(a['qty'])} шт, "
          f"{a['kg'] / 1000:.2f} т")
    print(f"  твердая цена    {a['firm']:>4} поз, "
          f"{a['firm_kg'] / 1000:>5.2f} т")
    print(f"  цена диапазоном {a['ranges']:>4} поз, "
          f"{a['ranges_kg'] / 1000:>5.2f} т — пункт 3.4")
    print(f"  «{AGREE}» {a['open']:>4} поз, "
          f"{a['open_kg'] / 1000:>5.2f} т — пункт 3.5, продано не считается")


if __name__ == "__main__":
    main()
