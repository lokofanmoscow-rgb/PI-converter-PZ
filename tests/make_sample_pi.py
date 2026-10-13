"""Собрать тестовый ПИ (.docx) из JSON-фикстуры.

Структура повторяет ПИ нового формата: абзацы «Общая информация», Таблица №1
(условия, 2 колонки), Таблица №2 (оборудование), Таблица №4 (компоненты,
первая ячейка с w:vMerge), Приложение №1 с фото и подписями.

    python tests/make_sample_pi.py tests/fixtures/pi_2610101.json out.docx
"""
import json
import struct
import sys
import zipfile
import zlib
from xml.sax.saxutils import escape

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def png(width: int, height: int, rgb=(40, 120, 80)) -> bytes:
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def run(text, bold=False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    parts = text.split("\n")
    inner = "<w:br/>".join(f'<w:t xml:space="preserve">{escape(x)}</w:t>' for x in parts)
    return f"<w:r>{rpr}{inner}</w:r>"


def para(text="", bold=False, jc=None):
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ""
    return f"<w:p>{ppr}{run(text, bold) if text else ''}</w:p>"


def cell(text, width=2000, vmerge=None, span=None):
    tcpr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span:
        tcpr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge == "restart":
        tcpr += '<w:vMerge w:val="restart"/>'
    elif vmerge == "continue":
        tcpr += "<w:vMerge/>"
    return f"<w:tc><w:tcPr>{tcpr}</w:tcPr>{para(text, jc='center')}</w:tc>"


def table(rows, grid):
    g = "".join(f'<w:gridCol w:w="{x}"/>' for x in grid)
    return ('<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'
            + "".join(f'<w:{b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                      for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
            + f"</w:tblBorders></w:tblPr><w:tblGrid>{g}</w:tblGrid>{''.join(rows)}</w:tbl>")


def tr(cells, header=False):
    trpr = "<w:trPr><w:tblHeader/></w:trPr>" if header else ""
    return f"<w:tr>{trpr}{''.join(cells)}</w:tr>"


def drawing(rid, cx, cy, n):
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Рисунок {n}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="0" name="img{n}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


def build(fx: dict, out: str) -> None:
    body = []
    for t in fx["paragraphs"]:
        body.append(para(t, bold=t.startswith(("ПРОТОКОЛ", "№"))))
    body.append(para("Таблица №1"))
    rows = [tr([cell("Наименование условий при испытании", 6000), cell("Значение показателей", 3000)], header=True)]
    rows += [tr([cell(a, 6000), cell(b, 3000)]) for a, b in fx["conditions"]]
    body.append(table(rows, [6000, 3000]))
    for t in fx["sample_paragraphs"]:
        body.append(para(t))
    body.append(para("4. Перечень оборудования, использованного при испытаниях"))
    body.append(para("Таблица №2"))
    rows = [tr([cell("Наименование средств измерений/ испытательного оборудования", 4000),
                cell("Заводской и (или) инвентарный номер", 2500),
                cell("Дата очередной поверки/аттестации", 2500)], header=True)]
    rows += [tr([cell(a, 4000), cell(b, 2500), cell(c, 2500)]) for a, b, c in fx["equipment"]]
    body.append(table(rows, [4000, 2500, 2500]))
    body.append(para("5. Процедура испытаний"))
    body.append(para(f"Подготовка: проведена полная разборка на комплектующие компоненты. "
                     f"Количество компонентов составило {len(fx['components'])} шт."))
    body.append(para("Таблица №3"))
    body.append(table([
        tr([cell("Нормативный документ на метод/методику испытаний", 3000),
            cell("Определяемая характеристика (показатель)", 3000), cell("Результаты испытаний", 3000)]),
        tr([cell("ГОСТ IEC 62321-2-2016, пп. 5.1, 5.4, 5.5, р.6, п.7.3.1, Приложение А, В", 3000),
            cell("Пробоподготовка", 3000), cell("готов", 3000)]),
    ], [3000, 3000, 3000]))
    body.append(para("Таблица №4"))
    grid = [1300, 2300, 1800, 1800, 1600]
    rows = [
        tr([cell("№ компонента/ материала", grid[0], "restart"),
            cell("Определяемая характеристика (Показатель)", grid[1], "restart"),
            cell("Результат испытаний по ГОСТ IEC 62321-3-1-2016", grid[2] + grid[3], span=2),
            cell("Интерпретация результата", grid[4], "restart")], header=True),
        tr([cell("", grid[0], "continue"), cell("", grid[1], "continue"),
            cell("мг/кг", grid[2]), cell("%", grid[3]), cell("", grid[4], "continue")], header=True),
    ]
    for comp in fx["components"]:
        for k, (el, mg, pct, interp) in enumerate(comp["elements"]):
            first = cell(f"{comp['number']}." if k == 0 else "", grid[0], "restart" if k == 0 else "continue")
            rows.append(tr([first, cell(el, grid[1]), cell(mg, grid[2]), cell(pct, grid[3]), cell(interp, grid[4])]))
    body.append(table(rows, grid))
    body.append(para("¹ Разъяснения по интерпретации результатов представлены в разделе «10. Мнения и интерпретации»"))
    body.append(para("Таблица №5"))
    body.append(table([tr([cell("Элемент"), cell("Полимеры"), cell("Металлы"), cell("Композиты")]),
                       tr([cell("Cd"), cell("BL ≤(70−3σ)<X"), cell("BL ≤(70−3σ)<X"), cell("LOD < X")])],
                      [2000, 2000, 2000, 2000]))
    body.append(para("Испытания провел(а):"))
    body.append(para(f"Дата выдачи протокола испытаний: {fx['date']}"))
    body.append(para("Приложение №1", bold=True, jc="right"))
    media = {}
    rels = []
    for k, (num, lines) in enumerate(fx["photos"]):
        rid = f"rIdImg{k + 1}"
        w_px, h_px = (30, 40) if k == 0 else (40, 30)
        media[f"word/media/image{k + 1}.png"] = png(w_px, h_px, (40 + k * 10, 120, 80))
        rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/image" Target="media/image{k + 1}.png"/>')
        cx, cy = (2700000, 3600000) if k == 0 else (4000000, 3000000)
        body.append(drawing(rid, cx, cy, k + 1))
        body.append(para(f"Фото №{num}", jc="center"))
        body.append(para("\n".join(lines), jc="center"))
    body.append(para("--------------------- конец протокола испытаний ---------------------"))

    ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}><w:body>'
                + "".join(body)
                + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>'
                  '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')
    header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {ns}>'
              + para("ИЛ ООО «Астра»") + para(f"Протокол испытаний № {fx['number']} от {fx['date']}") + "</w:hdr>")
    rels.append('<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/header" Target="header1.xml"/>')
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.header+xml"/></Types>'),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="word/document.xml"/></Relationships>'),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + "</Relationships>"),
        "word/document.xml": document,
        "word/header1.xml": header,
    }
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
        for name, data in media.items():
            z.writestr(name, data)


if __name__ == "__main__":
    with open(sys.argv[1], encoding="utf-8") as f:
        build(json.load(f), sys.argv[2])
    print("OK:", sys.argv[2])
