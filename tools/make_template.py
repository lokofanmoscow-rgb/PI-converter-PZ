"""Сделать ПЗ_template.docx из готового ПЗ.

    python tools/make_template.py ПЗ_эталон.docx ПЗ_template.docx

В таблице компонентов остаются шапка, строка пробоподготовки и одна строка
компонента — по ней программа строит все блоки. Фото из Приложения №1 удаляются
(программа вставляет фото из ПИ). Остальное (шапка, оборудование, подписи,
штампы ЭП) остаётся как в исходном ПЗ и перезаписывается при генерации.
"""
import os
import re
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from converter import _Package, norm, para_text, table_rows, w  # noqa: E402


def main(src: str, dst: str) -> None:
    pkg = _Package(src)
    doc = etree.fromstring(pkg.files["word/document.xml"])
    body = doc.find(w("body"))
    for tbl in body.iter(w("tbl")):
        rows = table_rows(tbl)
        comp_rows = [r for r in rows if "Содержание в компоненте" in "".join(para_text(p) for p in r.iter(w("p")))]
        if comp_rows:
            for r in comp_rows[1:]:
                r.getparent().remove(r)
            print(f"Строк компонентов было {len(comp_rows)}, оставлена 1")
            break
    kids = list(body)
    head = max(i for i, el in enumerate(kids)
               if el.tag == w("p") and re.match(r"^Приложение\s*№\s*1$", norm(para_text(el))))
    end = next(i for i, el in enumerate(kids) if i > head and "Конец технических" in para_text(el))
    for el in kids[head + 1:end]:
        body.remove(el)
    print(f"Из Приложения №1 удалено элементов: {end - head - 1}")
    pkg.drop_unused_images(doc)
    pkg.save(dst, doc)
    print(f"Шаблон сохранён: {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
