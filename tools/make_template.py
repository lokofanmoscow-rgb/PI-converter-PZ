"""Сделать ПЗ_template.docx из готового ПЗ.

    python tools/make_template.py ПЗ_2610101_от_05_10_2026.docx ПЗ_template.docx

В таблице компонентов остаются шапка, строка пробоподготовки и одна строка
компонента — по ней программа строит все блоки. Остальное (шапка, оборудование,
подписи, Приложение №1) остаётся как в исходном ПЗ и перезаписывается при генерации.
"""
import os
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from converter import para_text, table_rows, w  # noqa: E402


def main(src: str, dst: str) -> None:
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}
    doc = etree.fromstring(files["word/document.xml"])
    body = doc.find(w("body"))
    for tbl in body.iter(w("tbl")):
        rows = table_rows(tbl)
        comp_rows = [r for r in rows if "Содержание в компоненте" in "".join(para_text(p) for p in r.iter(w("p")))]
        if comp_rows:
            for r in comp_rows[1:]:
                r.getparent().remove(r)
            print(f"Строк компонентов было {len(comp_rows)}, оставлена 1")
            break
    files["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, files[n])
    print(f"Шаблон сохранён: {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
