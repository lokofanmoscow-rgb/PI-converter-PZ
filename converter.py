"""Конвертер ПИ (.docx) -> ПЗ (.docx) для ИЛ ООО «Астра».

Работает напрямую с XML документа Word через lxml (без python-docx).

Публичный API:
    parse_pi(path) -> PIData                  извлечь данные из ПИ
    generate_pz(pi, template, out, log=None)  собрать ПЗ по шаблону
    convert(pi_path, template, out_dir=None, log=None) -> путь к ПЗ
"""
from __future__ import annotations

import copy
import math
import os
import posixpath
import random
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Callable, List, Optional, Tuple

from lxml import etree

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

TEMPLATE_NAME = "ПЗ_template.docx"

# Оборудование набора Б (начиная с ПИ 2609093).
# (наименование, заводской/инвентарный номер, дата очередной поверки)
EQUIPMENT_SET_B = [
    (
        "Анализатор портативный спектрометрический Vanta, модификация: Анализатор Vanta C Core (V2CА), "
        "модель: Спектрометр портативный рентгенофлуоресцентный Vanta Core (Анализатор XRF). "
        "ПО: Vanta Software; номер версии: 4.12.107",
        "Зав. № 940289, инв. № СИ1",
        "23.09.2027",
    ),
    (
        "Термогигрометр автономный ИВА-6 модификации ИВА-6Н-Д. ПО: iva6n.txt, номер версии: 9.05",
        "Зав. № 43610, инв. № СИ2",
        "30.10.2026",
    ),
    (
        'Линейка измерительная металлическая торговой марки "Калиброн", иcполнения II',
        "Зав. № 230201677, инв. № СИ5",
        "05.11.2026",
    ),
    (
        "Штангенциркуль торговой марки «SHAN», с цифровым отсчетным устройством двусторонний с глубиномером",
        "Зав. № САА09097, инв. № СИ8",
        "05.11.2026",
    ),
    (
        "Мультиметр VERDO МН6100 модификация VERDO MH6104",
        "Зав. № H12G-C24127, инв. № СИ12",
        "27.06.2027",
    ),
]

# Если испытания идут несколько дней, а в ПИ указаны только начало и конец
# («ДД.ММ.ГГГГ с ЧЧ:ММ по ДД.ММ.ГГГГ ЧЧ:ММ»), рабочий день считается таким:
WORKDAY_START = "09:00"
WORKDAY_END = "18:00"

# Что писать в колонтитуле ПЗ после «№»: "napravlenie" (номер и дата направления,
# как в исходном шаблоне) или "pi" (номер и дата ПИ).
HEADER_NUMBER_SOURCE = "napravlenie"

# Диапазон случайной 3σ для элементов с результатом BL.
SIGMA_MIN, SIGMA_MAX = 3, 29

# Максимальный размер фото в Приложении №1 (EMU, 1 см = 360000 EMU).
PHOTO_MAX_W = int(10.2 * 360000)
PHOTO_MAX_H = int(10.0 * 360000)

# ---------------------------------------------------------------------------
# XML-утилиты
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V_NS = "urn:schemas-microsoft-com:vml"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

IMAGE_CONTENT_TYPES = {
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "png": "image/png",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
}

DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")
TIME_RE = re.compile(r"\b(\d{1,2}):(\d{2})\b")


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def norm(text: str) -> str:
    """Нормализовать пробелы (в т.ч. неразрывные) и обрезать края."""
    return re.sub(r"[ \t   ]+", " ", text or "").strip()


def _walk_text(el, out: list) -> None:
    """Собрать текст элемента, пропуская надписи (text box) и удалённый текст."""
    for ch in el:
        if not isinstance(ch.tag, str):
            continue
        ns, tag = etree.QName(ch).namespace, etree.QName(ch).localname
        if tag in ("txbxContent", "Fallback", "delText", "instrText") or tag.endswith("Pr"):
            continue
        if ns == W_NS and tag == "t":
            out.append(ch.text or "")
        elif ns == W_NS and tag == "tab":
            out.append("\t")
        elif ns == W_NS and tag in ("br", "cr"):
            out.append("\n")
        else:
            _walk_text(ch, out)


def para_text(p) -> str:
    out: list = []
    _walk_text(p, out)
    return "".join(out)


def text_nodes(p) -> List:
    """Все w:t абзаца (без надписей), в порядке документа."""
    res = []

    def rec(el):
        for ch in el:
            if not isinstance(ch.tag, str):
                continue
            tag = etree.QName(ch).localname
            if tag in ("txbxContent", "Fallback"):
                continue
            if ch.tag == w("t"):
                res.append(ch)
            else:
                rec(ch)

    rec(p)
    return res


def row_cells(tr) -> List:
    return tr.xpath("./w:tc | ./w:sdt/w:sdtContent/w:tc", namespaces={"w": W_NS})


def table_rows(tbl) -> List:
    return tbl.xpath("./w:tr | ./w:sdt/w:sdtContent/w:tr", namespaces={"w": W_NS})


def cell_paragraphs(tc) -> List:
    return tc.findall(w("p"))


def cell_text(tc) -> str:
    return "\n".join(para_text(p) for p in tc.iter(w("p")))


def table_text(tbl) -> str:
    return "\n".join(para_text(p) for p in tbl.iter(w("p")))


def make_run(text: str, rpr=None):
    r = etree.Element(w("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i:
            etree.SubElement(r, w("br"))
        t = etree.SubElement(r, w("t"))
        t.text = line
        t.set(XML_SPACE, "preserve")
    return r


def first_rpr(p):
    """rPr первого текстового прогона абзаца (или None)."""
    for r in p.iter(w("r")):
        if r.find(w("t")) is not None:
            return r.find(w("rPr"))
    for r in p.iter(w("r")):
        rpr = r.find(w("rPr"))
        if rpr is not None:
            return rpr
    return None


def clear_paragraph(p) -> None:
    for ch in list(p):
        if ch.tag != w("pPr"):
            p.remove(ch)


def set_par_text(p, text: str, rpr=None) -> None:
    """Заменить содержимое абзаца одним прогоном с текстом, сохранив pPr и rPr."""
    if rpr is None:
        rpr = first_rpr(p)
    rpr = copy.deepcopy(rpr) if rpr is not None else None
    clear_paragraph(p)
    if text:
        p.append(make_run(text, rpr))
    elif rpr is not None:
        r = etree.SubElement(p, w("r"))
        r.append(rpr)


def set_cell_text(tc, text: str) -> None:
    """Оставить в ячейке один абзац с заданным текстом."""
    ps = cell_paragraphs(tc)
    if not ps:
        ps = [etree.SubElement(tc, w("p"))]
    set_par_text(ps[0], text)
    for p in ps[1:]:
        tc.remove(p)


def regex_replace_in_paragraph(p, pattern, repl: str) -> int:
    """Замена по регулярному выражению с учётом разбиения текста на прогоны."""
    nodes = text_nodes(p)
    full = "".join(t.text or "" for t in nodes)
    matches = list(re.finditer(pattern, full))
    for m in reversed(matches):
        start, end = m.span()
        new = m.expand(repl)
        pos = 0
        first = True
        for t in nodes:
            s = t.text or ""
            n_start, n_end = pos, pos + len(s)
            pos = n_end
            if n_end <= start or n_start >= end and not (first and n_start == start == end):
                continue
            a = max(start, n_start) - n_start
            b = min(end, n_end) - n_start
            if first:
                t.text = s[:a] + new + s[b:]
                first = False
            else:
                t.text = s[:a] + s[b:]
            t.set(XML_SPACE, "preserve")
    return len(matches)


def strip_para_ids(el) -> None:
    """Убрать w14:paraId/textId у клонированных элементов (они должны быть уникальны)."""
    for node in el.iter():
        if not isinstance(node.tag, str):
            continue
        for attr in (f"{{{W14_NS}}}paraId", f"{{{W14_NS}}}textId"):
            if attr in node.attrib:
                del node.attrib[attr]


# ---------------------------------------------------------------------------
# Числа
# ---------------------------------------------------------------------------


def to_float(s: str) -> float:
    return float(s.strip().replace(",", ".").replace(" ", ""))


def decimals_of(s: str) -> int:
    s = s.strip()
    m = re.search(r"[.,](\d+)$", s)
    return len(m.group(1)) if m else 0


def fmt(v: float, dec: int) -> str:
    return f"{v:.{dec}f}".replace(".", ",")


def interp(mn_str: str, mx_str: str, i: int, total: int, dec: int) -> float:
    """Интерполяция от min к max по номеру компонента (последний = max)."""
    if i == total:
        return float(mx_str.replace(",", "."))
    mn = float(mn_str.replace(",", "."))
    mx = float(mx_str.replace(",", "."))
    v = mn + (mx - mn) * (i - 1) / (total - 1)
    return round(v * 10**dec) / 10**dec


# ---------------------------------------------------------------------------
# Модель данных
# ---------------------------------------------------------------------------

# ключ: (именительный падеж, родительный падеж)
ELEMENTS = {
    "Pb": ("Свинец (Pb)", "свинца (Pb)"),
    "Hg": ("Ртуть (Hg)", "ртути (Hg)"),
    "Cd": ("Кадмий (Cd)", "кадмия (Cd)"),
    "Cr": ("Общий хром", "общего хрома"),
    "Br": ("Общий бром", "общего брома"),
}

INTERPRETATION = {"BL": "Pass(BL)", "OL": "Fail(OL)", "X": "Inconclusive(X)"}


def element_key(text: str) -> Optional[str]:
    t = text.lower()
    if "свин" in t:
        return "Pb"
    if "ртут" in t:
        return "Hg"
    if "кадм" in t:
        return "Cd"
    if "хром" in t:
        return "Cr"
    if "бром" in t:
        return "Br"
    return None


@dataclass
class ElementResult:
    key: str
    is_bl: bool
    mg: str  # для BL — предел («10»), иначе значение («163»)
    pct: str  # для BL — предел («0,001»), иначе значение («0,0163»)
    mg_unc: str = ""  # «21» из «163 ± 21»
    pct_unc: str = ""
    interp: str = "BL"

    @property
    def name(self) -> str:
        return ELEMENTS[self.key][0]

    @property
    def genitive(self) -> str:
        return ELEMENTS[self.key][1]


@dataclass
class Component:
    number: str
    elements: List[ElementResult] = field(default_factory=list)


@dataclass
class Condition:
    mn: str
    mx: str

    @property
    def dec(self) -> int:
        return max(decimals_of(self.mn), decimals_of(self.mx))

    @property
    def text(self) -> str:
        return self.mn if self.mn == self.mx else f"{self.mn} – {self.mx}"


@dataclass
class DayWindow:
    day: date
    start: str  # ЧЧ:ММ
    end: str


@dataclass
class Photo:
    number: str
    lines: List[str]
    image: Optional[dict] = None  # {'data', 'ext', 'cx', 'cy', 'rot', 'src_rect'}


@dataclass
class PIData:
    number: str = ""
    date: str = ""
    napravlenie: str = ""
    napr_number: str = ""
    napr_date: str = ""
    sample_full: str = ""
    sample_name: str = ""
    sample_code: str = ""
    windows: List[DayWindow] = field(default_factory=list)
    room: str = ""
    conditions: dict = field(default_factory=dict)  # T, RH, P, U, f -> Condition
    equipment: List[Tuple[str, str, str]] = field(default_factory=list)
    components: List[Component] = field(default_factory=list)
    photos: List[Photo] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def start(self) -> Tuple[str, str]:
        return self.windows[0].day.strftime("%d.%m.%Y"), self.windows[0].start

    @property
    def end(self) -> Tuple[str, str]:
        return self.windows[-1].day.strftime("%d.%m.%Y"), self.windows[-1].end


class ConversionError(Exception):
    pass


# ---------------------------------------------------------------------------
# Разбор ПИ
# ---------------------------------------------------------------------------


class DocxPackage:
    """Минимальная обёртка над zip-пакетом .docx."""

    def __init__(self, path: str):
        self.path = path
        try:
            with zipfile.ZipFile(path) as z:
                self.files = {n: z.read(n) for n in z.namelist()}
        except zipfile.BadZipFile as e:
            raise ConversionError(f"Файл не является документом .docx: {os.path.basename(path)}") from e
        except FileNotFoundError as e:
            raise ConversionError(f"Файл не найден: {path}") from e
        if "word/document.xml" not in self.files:
            raise ConversionError("В файле нет word/document.xml — это не документ Word.")
        self.document = etree.fromstring(self.files["word/document.xml"])

    def rels(self, part: str = "word/document.xml") -> dict:
        d, f = posixpath.split(part)
        rel_name = posixpath.join(d, "_rels", f + ".rels")
        if rel_name not in self.files:
            return {}
        root = etree.fromstring(self.files[rel_name])
        res = {}
        for r in root:
            target = r.get("Target", "")
            if r.get("TargetMode") == "External":
                full = None
            else:
                full = posixpath.normpath(posixpath.join(d, target)) if not target.startswith("/") else target[1:]
            res[r.get("Id")] = (r.get("Type"), full)
        return res

    def header_parts(self) -> List[str]:
        return sorted(n for n in self.files if re.match(r"word/header\d*\.xml$", n))


def _all_paragraph_texts(root) -> List[str]:
    body = root.find(w("body"))
    return [norm(para_text(p)) for p in body.iter(w("p"))]


def _value_after_label(texts: List[str], pattern: str, accept: Callable[[str], bool] = lambda s: True) -> str:
    """Текст после абзаца-заголовка (или остаток самого заголовка, если это значение)."""
    for i, t in enumerate(texts):
        m = re.search(pattern, t, re.I)
        if not m:
            continue
        rest = t[m.end():].strip(" *:.\t")
        if len(rest) > 3 and accept(rest):
            return rest
        for nxt in texts[i + 1: i + 6]:
            if nxt and accept(nxt):
                return nxt
    return ""


def _parse_number_and_date(pkg: DocxPackage, texts: List[str]) -> Tuple[str, str]:
    pat = re.compile(r"№\s*(\d{5,})\s*от\s*(\d{2}\.\d{2}\.\d{4})")
    for t in texts + [" ".join(texts[:60])]:
        m = pat.search(t)
        if m:
            return m.group(1), m.group(2)
    for part in pkg.header_parts():
        root = etree.fromstring(pkg.files[part])
        for p in root.iter(w("p")):
            m = pat.search(norm(para_text(p)))
            if m:
                return m.group(1), m.group(2)
    m = re.search(r"(\d{7})\D+(\d{2})[._](\d{2})[._](\d{4})", os.path.basename(pkg.path))
    if m:
        return m.group(1), f"{m.group(2)}.{m.group(3)}.{m.group(4)}"
    raise ConversionError("Не найден номер и дата ПИ («№ XXXXXXX от ДД.ММ.ГГГГ»).")


def _hm(s: str) -> int:
    h, m = s.split(":")
    return int(h) * 60 + int(m)


def parse_test_period(text: str) -> List[DayWindow]:
    """Разобрать «02.10.2026 с 09:10 до 17:35» и многодневные варианты."""
    text = norm(text)
    cut = re.search(r"по\s+адресу|адрес", text, re.I)
    if cut:
        text = text[: cut.start()]
    dates = [datetime.strptime(d, "%d.%m.%Y").date() for d in DATE_RE.findall(text)]
    times = [f"{int(h):02d}:{m}" for h, m in TIME_RE.findall(text)]
    if not dates or len(times) < 2:
        raise ConversionError(f"Не удалось разобрать дату/время испытаний: «{text}»")
    if len(dates) == 1:
        return [DayWindow(dates[0], times[0], times[-1])]
    if len(times) == 2 * len(dates):
        # «ДД1 с ЧЧ:ММ до ЧЧ:ММ, ДД2 с ЧЧ:ММ до ЧЧ:ММ»
        return [DayWindow(d, times[2 * k], times[2 * k + 1]) for k, d in enumerate(dates)]
    # «ДД1 с ЧЧ:ММ по ДД2 ЧЧ:ММ» — начало и окончание, между ними рабочие дни
    d1, d2 = dates[0], dates[-1]
    t1, t2 = times[0], times[-1]
    if d2 < d1:
        raise ConversionError(f"Дата окончания раньше даты начала: «{text}»")
    if d1 == d2:
        return [DayWindow(d1, t1, t2)]
    windows = []
    day = d1
    while day <= d2:
        start = t1 if day == d1 else WORKDAY_START
        end = t2 if day == d2 else WORKDAY_END
        if _hm(end) <= _hm(start):  # начали после конца рабочего дня / закончили до его начала
            if day == d1:
                end = "23:59"
            else:
                start = "00:00"
        windows.append(DayWindow(day, start, end))
        day += timedelta(days=1)
    return windows


def _parse_conditions(tables) -> dict:
    keys = [("T", "температур"), ("RH", "влажност"), ("P", "давлени"), ("U", "напряжени"), ("f", "частот")]
    for tbl in tables:
        txt = table_text(tbl).lower()
        if "температур" not in txt or "влажност" not in txt or "свин" in txt:
            continue
        res = {}
        for tr in table_rows(tbl):
            cells = row_cells(tr)
            if len(cells) < 2:
                continue
            label = norm(cell_text(cells[-2])).lower()
            value = norm(cell_text(cells[-1]))
            for key, kw in keys:
                if kw in label and key not in res:
                    nums = re.findall(r"\d+(?:[.,]\d+)?", value)
                    if nums:
                        res[key] = Condition(nums[0].replace(".", ","), nums[-1].replace(".", ","))
        if len(res) == 5:
            return res
    raise ConversionError("Не найдена таблица условий проведения испытаний (Таблица №1).")


def _parse_equipment(tables) -> List[Tuple[str, str, str]]:
    for tbl in tables:
        rows = table_rows(tbl)
        if not rows:
            continue
        head = norm(cell_text(row_cells(rows[0])[0]) if row_cells(rows[0]) else "").lower()
        if "средств измерений" not in head:
            continue
        res = []
        for tr in rows[1:]:
            cells = row_cells(tr)
            if len(cells) >= 3:
                res.append(tuple(norm(cell_text(c)) for c in cells[:3]))
        return res
    return []


_NUM_TEXT = r"\d+(?:[.,]\d+)?"


def _parse_value(text: str) -> Tuple[bool, str, str]:
    """→ (is_bl, value, uncertainty)."""
    t = norm(text)
    m = re.match(rf"(?i)^(?:менее|<)\s*({_NUM_TEXT})", t)
    if m:
        return True, m.group(1).replace(".", ","), ""
    m = re.match(rf"^({_NUM_TEXT})\s*(?:±\s*({_NUM_TEXT}))?", t)
    if m:
        return False, m.group(1).replace(".", ","), (m.group(2) or "").replace(".", ",")
    raise ValueError(t)


def _pct_from_mg(mg: str) -> str:
    v = to_float(mg) / 10000
    return fmt(v, 4)


def _parse_components(tables, warnings: List[str]) -> List[Component]:
    comps: List[Component] = []
    for tbl in tables:
        txt = table_text(tbl).lower()
        if "компонент" not in txt and "менее" not in txt:
            continue
        if not any(kw in txt for kw in ("свин", "ртут", "кадм")):
            continue
        for tr in table_rows(tbl):
            cells = row_cells(tr)
            texts = [norm(cell_text(c)) for c in cells]
            el_idx = None
            for k, t in enumerate(texts):
                if element_key(t) and len(t) < 40:
                    el_idx = k
                    break
            if el_idx is None or len(texts) < el_idx + 3:
                continue
            key = element_key(texts[el_idx])
            first = cells[0]
            vm = first.find(f"{w('tcPr')}/{w('vMerge')}")
            restart = vm is not None and vm.get(w("val")) == "restart"
            num_m = re.match(r"^(\d+)\s*\.?$", texts[0]) if el_idx > 0 else None
            current = comps[-1] if comps else None
            is_new = (
                current is None
                or bool(num_m)
                or (restart and bool(texts[0]))
                or (restart and key == "Pb")
                or any(e.key == key for e in current.elements)
            )
            if is_new:
                number = num_m.group(1) if num_m else str(len(comps) + 1)
                current = Component(number)
                comps.append(current)
            mg_text = texts[el_idx + 1]
            pct_text = texts[el_idx + 2] if len(texts) > el_idx + 2 else ""
            interp_text = texts[el_idx + 3] if len(texts) > el_idx + 3 else "BL"
            try:
                is_bl, mg, mg_unc = _parse_value(mg_text)
            except ValueError:
                warnings.append(
                    f"Компонент {current.number}, {ELEMENTS[key][0]}: не удалось разобрать «{mg_text}» — пропущено."
                )
                continue
            try:
                _, pct, pct_unc = _parse_value(pct_text)
            except ValueError:
                pct, pct_unc = (_pct_from_mg(mg), _pct_from_mg(mg_unc) if mg_unc else "")
            interp_code = norm(interp_text).upper() or "BL"
            current.elements.append(ElementResult(key, is_bl, mg, pct, mg_unc, pct_unc, interp_code))
    comps = [c for c in comps if c.elements]
    for i, c in enumerate(comps, 1):
        if c.number != str(i):
            warnings.append(f"Нумерация компонентов в ПИ нарушена: на позиции {i} стоит № {c.number}.")
            break
    return comps


def _drawing_image(pkg: DocxPackage, rels: dict, drawing) -> Optional[dict]:
    blip = drawing.find(f".//{{{A_NS}}}blip")
    if blip is None:
        return None
    rid = blip.get(f"{{{R_NS}}}embed")
    if not rid or rid not in rels or not rels[rid][1] or rels[rid][1] not in pkg.files:
        return None
    target = rels[rid][1]
    container = drawing.find(f"{{{WP_NS}}}inline")
    if container is None:
        container = drawing.find(f"{{{WP_NS}}}anchor")
    extent = container.find(f"{{{WP_NS}}}extent") if container is not None else None
    cx = int(extent.get("cx")) if extent is not None else 3600000
    cy = int(extent.get("cy")) if extent is not None else 2700000
    eff_el = container.find(f"{{{WP_NS}}}effectExtent") if container is not None else None
    eff = tuple(int(eff_el.get(k, "0")) for k in ("l", "t", "r", "b")) if eff_el is not None else (0, 0, 0, 0)
    xfrm = drawing.find(f".//{{{PIC_NS}}}spPr/{{{A_NS}}}xfrm")
    rot = int(xfrm.get("rot", "0")) if xfrm is not None else 0
    src_rect = drawing.find(f".//{{{PIC_NS}}}blipFill/{{{A_NS}}}srcRect")
    return {
        "data": pkg.files[target],
        "ext": target.rsplit(".", 1)[-1].lower(),
        "cx": cx,
        "cy": cy,
        "rot": rot,
        "eff": eff,
        "src_rect": dict(src_rect.attrib) if src_rect is not None else {},
    }


def _vml_image(pkg: DocxPackage, rels: dict, imagedata, shape) -> Optional[dict]:
    rid = imagedata.get(f"{{{R_NS}}}id")
    if not rid or rid not in rels or not rels[rid][1] or rels[rid][1] not in pkg.files:
        return None
    target = rels[rid][1]
    style = shape.get("style", "") if shape is not None else ""

    def dim(name, default):
        m = re.search(rf"{name}:([\d.]+)pt", style)
        return int(float(m.group(1)) * 12700) if m else default

    return {
        "data": pkg.files[target],
        "ext": target.rsplit(".", 1)[-1].lower(),
        "cx": dim("width", 3600000),
        "cy": dim("height", 2700000),
        "rot": 0,
        "src_rect": {},
    }


def _parse_photos(pkg: DocxPackage, warnings: List[str]) -> List[Photo]:
    body = pkg.document.find(w("body"))
    paras = list(body.iter(w("p")))
    start = None
    for i, p in enumerate(paras):
        if re.match(r"^Приложение\s*№\s*1\.?$", norm(para_text(p)), re.I):
            start = i
    if start is None:
        return []
    rels = pkg.rels()
    events = []  # ('img', dict) | ('text', str)
    for p in paras[start + 1:]:
        for el in p.iter():
            if el.tag == w("drawing"):
                img = _drawing_image(pkg, rels, el)
                if img:
                    events.append(("img", img))
            elif el.tag == f"{{{V_NS}}}imagedata":
                img = _vml_image(pkg, rels, el, el.getparent())
                if img:
                    events.append(("img", img))
        t = para_text(p).strip()
        if t:
            events.append(("text", t))
    # Порядок: картинка над подписью (как в ПИ) или под ней
    first_kind = next((k for k, v in events if k == "img" or (k == "text" and v.startswith("Фото"))), None)
    image_first = first_kind == "img"
    photos: List[Photo] = []
    pending = None
    for kind, val in events:
        if kind == "img":
            if image_first:
                pending = val
            elif photos and photos[-1].image is None:
                photos[-1].image = val
            continue
        m = re.match(r"^Фото\s*№\s*(\d+)\.?\s*(.*)$", val, re.S)
        if m:
            photo = Photo(m.group(1), [])
            rest = m.group(2).strip()
            if rest:
                photo.lines.extend(norm(x) for x in rest.split("\n") if norm(x))
            if image_first:
                photo.image, pending = pending, None
            photos.append(photo)
        elif photos and not re.search(r"конец\s+протокола|^-{5,}", val, re.I):
            photos[-1].lines.extend(norm(x) for x in val.split("\n") if norm(x))
    for ph in photos:
        if ph.image is None:
            warnings.append(f"Фото №{ph.number}: изображение в ПИ не найдено — в ПЗ будет пустое место под фото.")
    return photos


def parse_pi(path: str) -> PIData:
    """Извлечь данные из файла ПИ (.docx)."""
    if path.lower().endswith(".pdf"):
        raise ConversionError(
            "Выбран PDF. Нужен исходный файл ПИ в формате Word (.docx) — "
            "в скане PDF первые страницы являются картинками и данные из них не извлекаются."
        )
    pkg = DocxPackage(path)
    root = pkg.document
    body = root.find(w("body"))
    texts = _all_paragraph_texts(root)
    tables = list(body.iter(w("tbl")))
    pi = PIData()
    pi.number, pi.date = _parse_number_and_date(pkg, texts)

    pi.napravlenie = _value_after_label(
        texts, r"Основание для проведения испытаний", lambda s: "№" in s or "направлен" in s.lower()
    ).rstrip(" .")
    if not pi.napravlenie:
        raise ConversionError("Не найдено «Основание для проведения испытаний» (направление).")
    m = re.search(r"№\s*(\S+)\s+от\s+(\d{2}\.\d{2}\.\d{4})", pi.napravlenie)
    if m:
        pi.napr_number, pi.napr_date = m.group(1), m.group(2)

    period = _value_after_label(
        texts, r"Место,?\s*дата\(?ы?\)?\s*и\s*время", lambda s: bool(DATE_RE.search(s) and TIME_RE.search(s))
    )
    if not period:
        raise ConversionError("Не найден раздел «Место, дата(ы) и время осуществления…».")
    pi.windows = parse_test_period(period)
    m = re.search(r"(комнат[аы])\s+([\d,\sи]+\d)", period, re.I)
    if m:
        pi.room = f"{m.group(1).capitalize()} {norm(m.group(2))}"

    sample = _value_after_label(texts, r"Наименование образца испытаний")
    if not sample:
        raise ConversionError("Не найдено «Наименование образца испытаний».")
    pi.sample_full = sample
    m = re.search(r"Шифр образца[^:]*:\s*(\S+)", sample, re.I)
    if not m:
        m = next((x for x in (re.search(r"Шифр образца[^:]*:\s*(\S+)", t, re.I) for t in texts) if x), None)
    if m:
        pi.sample_code = m.group(1).rstrip(".,;")
    else:
        pi.warnings.append("Шифр образца не найден в ПИ.")
    name = re.split(r"\s+(?=Образец\s|Шифр образца)", sample, maxsplit=1)[0].strip()
    pi.sample_name = name if name.endswith(".") else name + "."

    pi.conditions = _parse_conditions(tables)
    pi.equipment = _parse_equipment(tables)
    pi.components = _parse_components(tables, pi.warnings)
    if not pi.components:
        raise ConversionError("Не найдена таблица результатов по компонентам (Таблица №4).")
    pi.photos = _parse_photos(pkg, pi.warnings)
    return pi


# ---------------------------------------------------------------------------
# Расчёты для ПЗ
# ---------------------------------------------------------------------------


def component_datetimes(windows: List[DayWindow], n: int) -> List[str]:
    """Строка «Дата и время» для каждого компонента 1..n."""
    if n == 1:
        w0 = windows[0]
        return [f"{w0.day:%d.%m.%Y} {w0.start}"]
    minutes = [max(_hm(x.end) - _hm(x.start), 1) for x in windows]
    total = sum(minutes)
    day_of = []
    for i in range(1, n + 1):
        t = total * (i - 1) / (n - 1)
        acc = 0
        k = 0
        for k, m in enumerate(minutes):
            acc += m
            if t <= acc + 1e-9:
                break
        day_of.append(k)
    res = []
    for i, k in enumerate(day_of):
        d = windows[k].day.strftime("%d.%m.%Y")
        first_in_day = i == 0 or day_of[i - 1] != k
        last_in_day = i == n - 1 or day_of[i + 1] != k
        if i == 0 or (first_in_day and not (last_in_day and i == n - 1)):
            res.append(f"{d} {windows[k].start}")
        elif last_in_day:
            res.append(f"{d} {windows[k].end}")
        else:
            res.append(d)
    return res


def _format_conditions(pi: PIData, values: dict) -> dict:
    res = {}
    for key, v in values.items():
        dec = pi.conditions[key].dec
        if key == "P":
            res[key] = f"{fmt(v * 10, max(dec - 1, 0))} ({fmt(v, dec)} кПа)"
        else:
            res[key] = fmt(v, dec)
    return res


def component_conditions(pi: PIData, i: int, total: int) -> dict:
    """Значения условий для компонента i (строки в формате ПЗ)."""
    return _format_conditions(pi, {k: interp(c.mn, c.mx, i, total, c.dec) for k, c in pi.conditions.items()})


def start_conditions(pi: PIData) -> dict:
    """Условия на начало испытаний (строка пробоподготовки)."""
    return _format_conditions(pi, {k: to_float(c.mn) for k, c in pi.conditions.items()})


def generate_sigmas(pi: PIData) -> List[List[Optional[int]]]:
    """3σ для BL-элементов: seed = номер ПИ, по порядку компонентов и элементов."""
    rnd = random.Random(int(pi.number))
    return [[rnd.randint(SIGMA_MIN, SIGMA_MAX) if e.is_bl else None for e in c.elements] for c in pi.components]


def result_text(e: ElementResult) -> str:
    if e.is_bl:
        return f"{e.name} менее {e.mg} мг/кг ({e.pct} %)"
    if e.mg_unc:
        return f"{e.name} {e.mg} ± {e.mg_unc} мг/кг ({e.pct} ± {e.pct_unc} %)"
    return f"{e.name} {e.mg} мг/кг ({e.pct} %)"


def detail_lines(e: ElementResult, sigma: Optional[int]) -> List[str]:
    interp_txt = INTERPRETATION.get(e.interp, e.interp)
    if e.is_bl:
        return [
            f" - {e.name}: ND (Не зафиксировано его значение); 3σ < {sigma} ppm (мг/кг).",
            f"   Интерпретация: {interp_txt}.",
        ]
    if e.mg_unc:
        return [
            f" - {e.name}: {e.mg} ppm (мг/кг); 3σ = {e.mg_unc} ppm (мг/кг).",
            f"   Интерпретация: {interp_txt}.",
            f"   Перевод мг/кг в проценты (мг/кг/10000): {e.pct} % ± {e.pct_unc} %.",
        ]
    return [
        f" - {e.name}: {e.mg} ppm (мг/кг).",
        f"   Интерпретация: {interp_txt}.",
        f"   Перевод мг/кг в проценты (мг/кг/10000): {e.pct} %.",
    ]


def pages_word(n: int) -> str:
    """«на 1 странице», «на 4 страницах»."""
    return "странице" if n % 10 == 1 and n % 100 != 11 else "страницах"


# ---------------------------------------------------------------------------
# Генерация ПЗ
# ---------------------------------------------------------------------------

COND_PREFIXES = {
    "T": "Температура окружающей среды",
    "RH": "Относительная влажность",
    "P": "Атмосферное давление",
    "U": "Напряжение сети",
    "f": "Частота переменного тока",
}


def _update_conditions_in_cell(tc, date_text: str, cond: dict) -> None:
    for p in cell_paragraphs(tc):
        t = para_text(p)
        if t.startswith("Дата и время проведения испытаний"):
            set_par_text(p, f"Дата и время проведения испытаний {date_text}")
            continue
        for key, prefix in COND_PREFIXES.items():
            if t.startswith(prefix) and ":" in t:
                label = t.split(":", 1)[0]
                set_par_text(p, f"{label}:{cond[key]}")
                break


def _find_table(body, needle: str):
    for tbl in body.iter(w("tbl")):
        if needle in table_text(tbl):
            return tbl
    return None


def _fill_header_table(body, pi: PIData) -> None:
    tbl = _find_table(body, "Основание проведения испытаний")
    if tbl is None:
        raise ConversionError("В шаблоне не найдена таблица «Основание проведения испытаний».")
    (d1, t1), (d2, t2) = pi.start, pi.end
    sample = pi.sample_name + (f" Шифр образца {pi.sample_code}" if pi.sample_code else "")
    for tr in table_rows(tbl):
        cells = row_cells(tr)
        if len(cells) < 2:
            continue
        label = norm(cell_text(cells[0]))
        sub = norm(cell_text(cells[-2])).lower() if len(cells) >= 3 else ""
        value = cells[-1]
        if label.startswith("Основание проведения"):
            set_cell_text(value, pi.napravlenie)
        elif label.startswith("Описание объекта"):
            set_cell_text(value, sample)
        elif label.startswith("Дата и время проведения"):
            set_cell_text(value, f"с {d1} {t1} по {d2} {t2}")
        elif label.startswith("Место проведения") and pi.room:
            set_cell_text(value, pi.room)
        else:
            for key, kw in (("T", "температур"), ("RH", "влажност"), ("P", "давлени"), ("U", "напряжени"), ("f", "частот")):
                if kw in sub:
                    set_cell_text(value, pi.conditions[key].text)
                    break


def _rebuild_c0(tc, comp: Component) -> None:
    ps = cell_paragraphs(tc)
    k = next((i for i, p in enumerate(ps) if "Содержание в компоненте" in para_text(p)), None)
    if k is None:
        raise ConversionError("В шаблоне не найден абзац «Содержание в компоненте N:».")
    proto = ps[k]
    rpr = first_rpr(proto)
    for p in ps[k:]:
        tc.remove(p)
    lines = [f"Содержание в компоненте {comp.number}:"]
    lines += [f"{j}. {e.name}" for j, e in enumerate(comp.elements, 1)]
    for line in lines:
        p = copy.deepcopy(proto)
        set_par_text(p, line, rpr)
        tc.append(p)


def _rebuild_c1(tc, comp: Component) -> None:
    ps = cell_paragraphs(tc)
    proto_text = next((p for p in ps if para_text(p).strip()), ps[0])
    proto_empty = next((p for p in ps if not para_text(p).strip()), None)
    rpr = first_rpr(proto_text)
    for p in ps:
        tc.remove(p)
    for j, e in enumerate(comp.elements):
        if j:
            if proto_empty is not None:
                tc.append(copy.deepcopy(proto_empty))
            else:
                p = copy.deepcopy(proto_text)
                set_par_text(p, "", rpr)
                tc.append(p)
        p = copy.deepcopy(proto_text)
        set_par_text(p, result_text(e), rpr)
        tc.append(p)


def _rebuild_c2(tc, comp: Component, date_text: str, cond: dict, sigmas: List[Optional[int]]) -> None:
    _update_conditions_in_cell(tc, date_text, cond)
    ps = cell_paragraphs(tc)
    texts = [para_text(p) for p in ps]
    i_meas = next((i for i, t in enumerate(texts) if t.startswith("Произведены измерения")), None)
    i_res = next((i for i, t in enumerate(texts) if t.startswith("Результаты измерений")), None)
    if i_meas is None or i_res is None:
        raise ConversionError("В шаблоне не найдены абзацы «Произведены измерения…» / «Результаты измерений:».")
    names = ", ".join(e.genitive for e in comp.elements)
    set_par_text(ps[i_meas], f"Произведены измерения содержания {names} в компоненте {comp.number}.")
    proto = ps[i_res + 1] if i_res + 1 < len(ps) else ps[i_res]
    rpr = first_rpr(proto)
    proto = copy.deepcopy(proto)
    for p in ps[i_res + 1:]:
        tc.remove(p)
    for e, s in zip(comp.elements, sigmas):
        for line in detail_lines(e, s):
            p = copy.deepcopy(proto)
            set_par_text(p, line, rpr)
            tc.append(p)


def _fill_components_table(body, pi: PIData, log: Callable[[str], None]) -> None:
    tbl = _find_table(body, "Содержание в компоненте")
    if tbl is None:
        raise ConversionError("В шаблоне не найдена таблица компонентов.")
    rows = table_rows(tbl)
    comp_rows = [r for r in rows if "Содержание в компоненте" in "".join(para_text(p) for p in r.iter(w("p")))]
    prep_row = next(
        (r for r in rows if r not in comp_rows and "Пробоподготовка" in "".join(para_text(p) for p in r.iter(w("p")))),
        None,
    )
    proto = copy.deepcopy(comp_rows[0])
    strip_para_ids(proto)
    anchor = comp_rows[0].getprevious()
    parent = comp_rows[0].getparent()
    for r in comp_rows:
        r.getparent().remove(r)

    n = len(pi.components)
    dates = component_datetimes(pi.windows, n)
    sigmas = generate_sigmas(pi)

    if prep_row is not None:
        cells = row_cells(prep_row)
        d1, t1 = pi.start
        _update_conditions_in_cell(cells[-1], f"{d1} {t1}", start_conditions(pi))
        name_no_dot = pi.sample_name.rstrip(".")
        for p in cell_paragraphs(cells[-1]):
            regex_replace_in_paragraph(p, r"После анализа .*? \(далее по тексту «образец»\)",
                                       f"После анализа {_escape_repl(name_no_dot)} (далее по тексту «образец»)")
            regex_replace_in_paragraph(p, r"Количество компонентов составило \d+ шт", f"Количество компонентов составило {n} шт")

    insert_at = parent.index(anchor) + 1 if anchor is not None else len(parent)
    for i, comp in enumerate(pi.components, 1):
        row = copy.deepcopy(proto)
        cells = row_cells(row)
        if len(cells) < 3:
            raise ConversionError("Строка компонента в шаблоне должна содержать 3 ячейки.")
        _rebuild_c0(cells[0], comp)
        _rebuild_c1(cells[1], comp)
        _rebuild_c2(cells[2], comp, dates[i - 1], component_conditions(pi, i, n), sigmas[i - 1])
        parent.insert(insert_at, row)
        insert_at += 1
    log(f"Блоков компонентов в ПЗ: {n}")


def _escape_repl(s: str) -> str:
    return s.replace("\\", "\\\\")


def _equipment_rows(pi: PIData, log: Callable[[str], None]) -> List[Tuple[str, str, str]]:
    """Набор Б; дата поверки берётся из ПИ, если там она другая."""
    translit = str.maketrans("АВЕКМНОРСТХаеосрх", "ABEKMHOPCTXaeocpx")

    def key(s: str) -> str:
        return re.sub(r"[\s№.]", "", s.translate(translit)).upper()

    rows = []
    used = set()
    for name, number, due in EQUIPMENT_SET_B:
        serial = re.search(r"Зав\. № (\S+?),", number).group(1)
        match = None
        for k, (_, pi_number, pi_due) in enumerate(pi.equipment):
            if key(serial) in key(pi_number):
                match = (k, pi_due)
                break
        if match:
            used.add(match[0])
            pi_due = DATE_RE.search(match[1])
            if pi_due and pi_due.group(0) != due:
                log(f"Внимание: дата поверки «{name[:40]}…» в ПИ = {pi_due.group(0)} (в наборе Б {due}) — взята из ПИ.")
                due = pi_due.group(0)
        rows.append((name, number, due))
    for k, (name, number, _) in enumerate(pi.equipment):
        if k not in used:
            log(f"Внимание: в ПИ есть оборудование вне набора Б: «{name[:60]}» ({number}) — в ПЗ не добавлено.")
    return rows


def _fill_equipment_table(body, pi: PIData, log: Callable[[str], None]) -> None:
    tbl = _find_table(body, "Наименование средств измерений")
    if tbl is None:
        log("Внимание: в шаблоне не найдена таблица оборудования.")
        return
    data = _equipment_rows(pi, log)
    rows = table_rows(tbl)[1:]
    if not rows:
        return
    while len(rows) < len(data):
        new = copy.deepcopy(rows[-1])
        strip_para_ids(new)
        rows[-1].addnext(new)
        rows.append(new)
    for extra in rows[len(data):]:
        extra.getparent().remove(extra)
    for tr, values in zip(rows, data):
        for tc, value in zip(row_cells(tr), values):
            set_cell_text(tc, value)


def _update_signature_dates(body, pi: PIData) -> int:
    count = 0
    for p in body.findall(w("p")):
        t = para_text(p)
        if "_" in t and DATE_RE.search(t):
            count += regex_replace_in_paragraph(p, DATE_RE.pattern, pi.date)
    return count


def _update_appendix_pages(body, pages: int) -> None:
    for p in body.iter(w("p")):
        if re.search(r"Приложение\s*№\s*1\s+на\s+\d+", para_text(p)):
            regex_replace_in_paragraph(p, r"на\s+\d+\s+страниц\w*", f"на {pages} {pages_word(pages)}")


# --- фото -------------------------------------------------------------------

_INLINE_XML = (
    '<w:drawing xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}">'
    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="0" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="1" name="Рисунок"/>'
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Photo"/>'
    '<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed=""/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></pic:spPr>'
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
).format(W=W_NS, WP=WP_NS, A=A_NS, PIC=PIC_NS, R=R_NS)


def _make_drawing(rid: str, img: dict, doc_pr_id: int, number: str):
    cx, cy, rot = img["cx"], img["cy"], img.get("rot", 0)
    l, t, r, b = img.get("eff", (0, 0, 0, 0))
    vis_w, vis_h = max(cx + l + r, 1), max(cy + t + b, 1)
    scale = min(PHOTO_MAX_W / vis_w, PHOTO_MAX_H / vis_h)

    def sc(v: int) -> str:
        return str(int(round(v * scale)))

    def set_attrs(el, **attrs):
        for k, v in attrs.items():
            el.set(k, v)

    drawing = etree.fromstring(_INLINE_XML)
    inline = drawing[0]
    set_attrs(inline.find(f"{{{WP_NS}}}extent"), cx=sc(cx), cy=sc(cy))
    set_attrs(inline.find(f"{{{WP_NS}}}effectExtent"), l=sc(l), t=sc(t), r=sc(r), b=sc(b))
    doc_pr = inline.find(f"{{{WP_NS}}}docPr")
    doc_pr.set("id", str(doc_pr_id))
    doc_pr.set("name", f"Фото {number}")
    drawing.find(f".//{{{A_NS}}}blip").set(f"{{{R_NS}}}embed", rid)
    xfrm = drawing.find(f".//{{{A_NS}}}xfrm")
    set_attrs(xfrm.find(f"{{{A_NS}}}ext"), cx=sc(cx), cy=sc(cy))
    if rot:
        xfrm.set("rot", str(rot))
    if img.get("src_rect"):
        src = drawing.find(f".//{{{A_NS}}}srcRect")
        for k, v in img["src_rect"].items():
            src.set(k, v)
    return drawing


class _Package:
    """Пакет шаблона ПЗ, который мы изменяем и сохраняем."""

    def __init__(self, path: str):
        with zipfile.ZipFile(path) as z:
            self.order = z.namelist()
            self.files = {n: z.read(n) for n in self.order}
        self.rels_root = etree.fromstring(self.files["word/_rels/document.xml.rels"])
        self.ct_root = etree.fromstring(self.files["[Content_Types].xml"])
        self._media_counter = 0

    def add_image(self, data: bytes, ext: str) -> str:
        ext = "jpeg" if ext == "jpg" else ext
        self._media_counter += 1
        name = f"media/pz_photo{self._media_counter}.{ext}"
        while f"word/{name}" in self.files:
            self._media_counter += 1
            name = f"media/pz_photo{self._media_counter}.{ext}"
        self.files[f"word/{name}"] = data
        self.order.append(f"word/{name}")
        ids = {r.get("Id") for r in self.rels_root}
        k = 1
        while f"rIdPz{k}" in ids:
            k += 1
        rid = f"rIdPz{k}"
        etree.SubElement(self.rels_root, f"{{{PKG_REL_NS}}}Relationship", Id=rid, Type=IMAGE_REL_TYPE, Target=name)
        if not any(d.get("Extension", "").lower() == ext for d in self.ct_root.findall(f"{{{CT_NS}}}Default")):
            etree.SubElement(
                self.ct_root, f"{{{CT_NS}}}Default", Extension=ext,
                ContentType=IMAGE_CONTENT_TYPES.get(ext, f"image/{ext}"),
            )
        return rid

    def drop_unused_images(self, document) -> None:
        used = set()
        for el in document.iter():
            if not isinstance(el.tag, str):
                continue
            for attr, val in el.attrib.items():
                if attr.startswith(f"{{{R_NS}}}"):
                    used.add(val)
        other_targets = set()
        for name, data in self.files.items():
            if name.endswith(".rels") and name != "word/_rels/document.xml.rels":
                for r in etree.fromstring(data):
                    other_targets.add(posixpath.normpath(posixpath.join(posixpath.dirname(posixpath.dirname(name)), r.get("Target", ""))))
        for r in list(self.rels_root):
            if r.get("Type") == IMAGE_REL_TYPE and r.get("Id") not in used:
                target = posixpath.normpath(posixpath.join("word", r.get("Target")))
                self.rels_root.remove(r)
                still = any(posixpath.normpath(posixpath.join("word", x.get("Target"))) == target for x in self.rels_root)
                if not still and target not in other_targets and target in self.files:
                    del self.files[target]
                    self.order.remove(target)

    def save(self, path: str, document) -> None:
        self.files["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.files["word/_rels/document.xml.rels"] = etree.tostring(
            self.rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.files["[Content_Types].xml"] = etree.tostring(
            self.ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            names = ["[Content_Types].xml"] + [n for n in self.order if n != "[Content_Types].xml"]
            for n in names:
                if n in self.files:
                    z.writestr(n, self.files[n])
        os.replace(tmp, path)


def _caption_paragraph(proto_p, lines: List[str], small_rpr) -> None:
    rpr = first_rpr(proto_p)
    rpr = copy.deepcopy(rpr) if rpr is not None else None
    clear_paragraph(proto_p)
    for i, line in enumerate(lines):
        is_range = bool(re.match(r"^\d+\s*(÷|-|–)\s*\d+|^\d+\s+компонент", line))
        r = make_run(line, small_rpr if (is_range and small_rpr is not None) else rpr)
        if i:
            br_run = etree.Element(w("r"))
            if rpr is not None:
                br_run.append(copy.deepcopy(rpr))
            etree.SubElement(br_run, w("br"))
            proto_p.append(br_run)
        proto_p.append(r)


def _default_photos(pi: PIData) -> List[Photo]:
    n = len(pi.components)
    return [
        Photo("1", [pi.sample_name]),
        Photo("2", [f"Компоненты. {pi.sample_name}", f"1÷{n} компоненты"]),
    ]


def _fill_photos(pkg: _Package, body, pi: PIData, log: Callable[[str], None]) -> int:
    blocks = list(body)
    head_idx = None
    for i, el in enumerate(blocks):
        if el.tag == w("p") and re.match(r"^Приложение\s*№\s*1\.?$", norm(para_text(el))):
            head_idx = i
    if head_idx is None:
        log("Внимание: в шаблоне не найден заголовок «Приложение №1» — фото не обновлены.")
        return 0
    photo_tables = [el for el in blocks[head_idx + 1:] if el.tag == w("tbl") and "Фото" in table_text(el)]
    if not photo_tables:
        log("Внимание: в шаблоне нет таблиц с фото — фото не обновлены.")
        return 0
    proto_tbl = copy.deepcopy(photo_tables[0])
    strip_para_ids(proto_tbl)
    spacer = photo_tables[0].getnext()
    spacer = copy.deepcopy(spacer) if spacer is not None and spacer.tag == w("p") and not para_text(spacer).strip() else None
    small_rpr = None
    if len(photo_tables) > 1:
        last_p = cell_paragraphs(row_cells(table_rows(photo_tables[1])[-1])[0])[-1]
        runs = [r for r in last_p.findall(w("r")) if r.find(w("t")) is not None]
        if runs and runs[-1].find(w("rPr")) is not None:
            small_rpr = copy.deepcopy(runs[-1].find(w("rPr")))
    # удаляем старые таблицы фото и пустые абзацы между ними
    insert_at = body.index(photo_tables[0])
    last = photo_tables[-1]
    el = photo_tables[0]
    while el is not None:
        nxt = el.getnext()
        body.remove(el)
        if el is last:
            break
        el = nxt

    photos = pi.photos or _default_photos(pi)
    if not pi.photos:
        log("Внимание: в ПИ не найдено Приложение №1 с фото — подписи сформированы по умолчанию, места под фото пустые.")
    doc_pr_id = 900000
    for k, ph in enumerate(photos):
        tbl = copy.deepcopy(proto_tbl)
        rows = table_rows(tbl)
        img_cell = row_cells(rows[0])[0]
        img_p = cell_paragraphs(img_cell)[0]
        for r in list(img_p.findall(w("r"))):
            img_p.remove(r)
        if ph.image:
            try:
                rid = pkg.add_image(ph.image["data"], ph.image["ext"])
                run = etree.SubElement(img_p, w("r"))
                rpr = etree.SubElement(run, w("rPr"))
                etree.SubElement(rpr, w("noProof"))
                doc_pr_id += 1
                run.append(_make_drawing(rid, ph.image, doc_pr_id, ph.number))
            except Exception as e:  # noqa: BLE001 — фото не должно ломать весь документ
                log(f"Внимание: не удалось вставить фото №{ph.number}: {e}")
        cap_cell = row_cells(rows[-1])[0]
        cps = cell_paragraphs(cap_cell)
        num_p = next((p for p in cps if para_text(p).strip().startswith("Фото")), None)
        if num_p is None:
            raise ConversionError("В шаблоне в таблице фото нет абзаца «Фото №…».")
        set_par_text(num_p, f"Фото №{ph.number}")
        cap_ps = cps[cps.index(num_p) + 1:]
        if cap_ps:
            _caption_paragraph(cap_ps[0], ph.lines or [""], small_rpr)
            for extra in cap_ps[1:]:
                cap_cell.remove(extra)
        else:
            p = copy.deepcopy(num_p)
            _caption_paragraph(p, ph.lines or [""], small_rpr)
            cap_cell.append(p)
        body.insert(insert_at, tbl)
        insert_at += 1
        if spacer is not None and k < len(photos) - 1:
            body.insert(insert_at, copy.deepcopy(spacer))
            insert_at += 1
    return len(photos)


def _appendix_pages(photos_count: int) -> int:
    # Альбомный лист: одно фото (до 10 см) с подписью на страницу.
    return max(photos_count, 1)


def _update_headers(pkg: _Package, pi: PIData, est_pages: int) -> None:
    if HEADER_NUMBER_SOURCE == "pi" or not pi.napr_number:
        ref = f"№ {pi.number} от {pi.date}"
    else:
        ref = f"№ {pi.napr_number} от {pi.napr_date}"
    for name in [n for n in pkg.files if re.match(r"word/header\d*\.xml$", n)]:
        root = etree.fromstring(pkg.files[name])
        changed = False
        for p in root.iter(w("p")):
            t = norm(para_text(p))
            if re.match(r"^№\s*\S+\s+от\s+\d{2}\.\d{2}\.\d{4}$", t):
                set_par_text(p, ref)
                changed = True
            elif t.startswith("Общее количество страниц") and "NUMPAGES" not in etree.tostring(p, encoding="unicode"):
                _make_numpages(p, est_pages)
                changed = True
        if changed:
            pkg.files[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _make_numpages(p, est_pages: int) -> None:
    """«Общее количество страниц: NN» → поле NUMPAGES (Word посчитает сам)."""
    runs = p.findall(w("r"))
    label_run = next((r for r in runs if "Общее количество страниц" in "".join(t.text or "" for t in r.iter(w("t")))), None)
    if label_run is None:
        return
    for r in runs[runs.index(label_run) + 1:]:
        p.remove(r)
    for t in label_run.iter(w("t")):
        t.text = re.sub(r"\s*\d+\s*$", " ", t.text or "")
        t.set(XML_SPACE, "preserve")
    rpr = label_run.find(w("rPr"))

    def run(child_tag, text=None, **attrs):
        r = etree.Element(w("r"))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        c = etree.SubElement(r, w(child_tag))
        for k, v in attrs.items():
            c.set(w(k), v)
        if text is not None:
            c.text = text
            c.set(XML_SPACE, "preserve")
        return r

    p.append(run("fldChar", fldCharType="begin"))
    p.append(run("instrText", " NUMPAGES  \\* Arabic  \\* MERGEFORMAT "))
    p.append(run("fldChar", fldCharType="separate"))
    p.append(run("t", str(est_pages)))
    p.append(run("fldChar", fldCharType="end"))


def _estimate_pages(n_components: int, n_photos: int) -> int:
    # грубая оценка для кэшированного значения поля (Word пересчитает)
    return 3 + math.ceil(n_components / 1.45) + _appendix_pages(n_photos)


def _dedupe_para_ids(document) -> None:
    seen = set()
    pid, tid = f"{{{W14_NS}}}paraId", f"{{{W14_NS}}}textId"
    for el in document.iter(w("p"), w("tr")):
        v = el.get(pid)
        if v is None:
            continue
        if v in seen:
            del el.attrib[pid]
            el.attrib.pop(tid, None)
        else:
            seen.add(v)


def generate_pz(pi: PIData, template_path: str, out_path: str, log: Optional[Callable[[str], None]] = None) -> str:
    log = log or (lambda s: None)
    if not os.path.exists(template_path):
        raise ConversionError(f"Не найден шаблон ПЗ: {template_path}")
    pkg = _Package(template_path)
    document = etree.fromstring(pkg.files["word/document.xml"])
    body = document.find(w("body"))

    _fill_header_table(body, pi)
    _fill_components_table(body, pi, log)
    _fill_equipment_table(body, pi, log)
    n_sig = _update_signature_dates(body, pi)
    if n_sig < 2:
        log(f"Внимание: обновлено дат подписей: {n_sig} (ожидалось 2).")
    n_photos = _fill_photos(pkg, body, pi, log)
    pages = _appendix_pages(n_photos)
    _update_appendix_pages(body, pages)
    _update_headers(pkg, pi, _estimate_pages(len(pi.components), n_photos))
    _dedupe_para_ids(document)
    pkg.drop_unused_images(document)
    pkg.save(out_path, document)
    return out_path


def output_name(pi: PIData) -> str:
    return f"ПЗ_{pi.number}_от_{pi.date.replace('.', '_')}.docx"


def summary(pi: PIData) -> List[str]:
    (d1, t1), (d2, t2) = pi.start, pi.end
    n_num = sum(1 for c in pi.components for e in c.elements if not e.is_bl)
    lines = [
        f"ПИ № {pi.number} от {pi.date}",
        f"Направление: {pi.napravlenie}",
        f"Образец: {pi.sample_name} Шифр: {pi.sample_code or '—'}",
        f"Испытания: с {d1} {t1} по {d2} {t2}",
        "Условия: " + "; ".join(f"{k} {c.text}" for k, c in pi.conditions.items()),
        f"Компонентов: {len(pi.components)} (числовых результатов: {n_num})",
        f"Фото в Приложении №1: {len(pi.photos)} (с изображением: {sum(1 for p in pi.photos if p.image)})",
    ]
    return lines


def convert(pi_path: str, template_path: str, out_dir: Optional[str] = None,
            log: Optional[Callable[[str], None]] = None) -> str:
    log = log or (lambda s: None)
    pi = parse_pi(pi_path)
    for line in summary(pi):
        log(line)
    for wmsg in pi.warnings:
        log("Внимание: " + wmsg)
    out_dir = out_dir or os.path.dirname(os.path.abspath(pi_path))
    out_path = os.path.join(out_dir, output_name(pi))
    generate_pz(pi, template_path, out_path, log)
    log(f"Готово: {out_path}")
    return out_path


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Использование: python converter.py ПИ.docx [ПЗ_template.docx] [папка_результата]")
        sys.exit(1)
    here = os.path.dirname(os.path.abspath(__file__))
    tpl = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, TEMPLATE_NAME)
    out = sys.argv[3] if len(sys.argv) > 3 else None
    convert(sys.argv[1], tpl, out, print)
