"""Тесты конвертера: python -m unittest discover -s tests -v"""
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIXTURES = os.path.join(HERE, "fixtures")
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

import converter  # noqa: E402
from converter import para_text, row_cells, table_rows, w  # noqa: E402
from make_sample_pi import build  # noqa: E402

TEMPLATE = os.path.join(ROOT, converter.TEMPLATE_NAME)
with open(os.path.join(FIXTURES, "pi_2610101.json"), encoding="utf-8") as f:
    FIXTURE = json.load(f)
REAL_PI = os.path.join(FIXTURES, "ПИ_2609082_от_15.09.2026.docx")
with open(os.path.join(FIXTURES, "pz_2609082_expected.json"), encoding="utf-8") as f:
    REAL_EXPECTED = json.load(f)

TXBX = w("txbxContent")


def run_conversion(fx, tmp):
    pi_path = os.path.join(tmp, f"ПИ_{fx['number']}.docx")
    build(fx, pi_path)
    logs = []
    out = converter.convert(pi_path, TEMPLATE, tmp, logs.append)
    return out, logs


def read_doc(path, part="word/document.xml"):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(part)), z.namelist()


def plain_texts(root, stop_at_appendix=False):
    """Тексты абзацев (без надписей/штампов), как в эталонном JSON."""
    out = []
    container = root.find(w("body")) if root.find(w("body")) is not None else root
    for p in container.iter(w("p")):
        if any(a.tag == TXBX for a in p.iterancestors()):
            continue
        t = "".join(("\n" if x.tag == w("br") else x.text or "") for x in p.iter(w("t"), w("br"))
                    if not any(a.tag == TXBX for a in x.iterancestors())).rstrip("\n")
        if stop_at_appendix and re.match(r"^Приложение\s*№\s*1$", t.strip()):
            break
        if t.strip():
            out.append(t)
    return out


def stamps(doc):
    return [["".join(x.text or "" for x in p.iter(w("t"))) for p in tx.findall(w("p"))] for tx in doc.iter(TXBX)]


def component_rows(doc):
    body = doc.find(w("body"))
    tbl = next(t for t in body.iter(w("tbl")) if "Содержание в компоненте" in "".join(
        para_text(p) for p in t.iter(w("p"))))
    res = []
    for tr in table_rows(tbl):
        cells = row_cells(tr)
        c = [[para_text(p) for p in cells[k].findall(w("p"))] for k in range(len(cells))]
        if any("Содержание в компоненте" in x for x in c[0]):
            res.append(c)
    return res


def wrap_pct(result: str) -> str:
    """Числовой результат: проценты в скобках — с новой строки (BL-строки без изменений)."""
    return result if "менее" in result else result.replace(" мг/кг (", " мг/кг\n(")


def field(c2, prefix):
    return next(x for x in c2 if x.startswith(prefix)).split(":", 1)[1]


def appendix_texts(doc):
    texts = plain_texts(doc)
    k = max(i for i, t in enumerate(texts) if re.match(r"^Приложение\s*№\s*1$", t.strip()))
    return texts[k:]


class RealProtocolTest(unittest.TestCase):
    """Настоящий ПИ 2609082 (Word) → ПЗ совпадает с эталонным ПЗ 46ТР-НИ до Приложения №1."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        pi_path = os.path.join(cls.tmp.name, os.path.basename(REAL_PI))
        shutil.copy(REAL_PI, pi_path)
        cls.logs = []
        cls.out = converter.convert(pi_path, TEMPLATE, cls.tmp.name, cls.logs.append)
        cls.doc, cls.names = read_doc(cls.out)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_file_name(self):
        self.assertEqual(os.path.basename(self.out), "ПЗ_2609082_от_15_09_2026.docx")

    def test_body_matches_reference(self):
        got = plain_texts(self.doc, stop_at_appendix=True)
        exp = REAL_EXPECTED["body"]
        for i, (g, e) in enumerate(zip(got, exp)):
            self.assertEqual(g, e, f"абзац {i}")
        self.assertEqual(len(got), len(exp))

    def test_stamps_match_reference(self):
        self.assertEqual(stamps(self.doc), REAL_EXPECTED["stamps"])

    def test_header_part(self):
        hdr, _ = read_doc(self.out, "word/header1.xml")
        texts = plain_texts(hdr)
        self.assertIn("№ 46ТР-НИ от 05.08.2026", texts)
        self.assertIn("NUMPAGES", etree.tostring(hdr, encoding="unicode"))

    def test_numeric_results_wrap_percent(self):
        rows = component_rows(self.doc)
        c1 = [x for x in rows[-1][1] if x]
        self.assertIn("Общий бром 74 ± 7 мг/кг\n(0,0074 ± 0,0007 %)", c1)
        self.assertIn("Свинец (Pb) менее 10 мг/кг (0,001 %)", c1)

    def test_component_times(self):
        rows = component_rows(self.doc)
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 14.09.2026 13:30")
        self.assertEqual(rows[-1][2][0], "Дата и время проведения испытаний 14.09.2026 16:40")

    def test_appendix(self):
        app = appendix_texts(self.doc)
        self.assertEqual(app[0].strip(), "Приложение № 1")
        self.assertEqual([t for t in app if t.startswith("Фото")], ["Фото №1", "Фото №2", "Фото №3", "Фото №4"])
        self.assertIn("Компоненты. Электрическая турка Торговой марки: Aceline Модель: ETCM-600", app)
        self.assertIn("51÷65 компоненты", app)
        self.assertIn("Приложение № 1 на 4 страницах.", plain_texts(self.doc))
        self.assertEqual(len(self.doc.findall(f".//{{{converter.A_NS}}}blip")), 4)
        # каждое фото после первого — с новой страницы, приложение — с новой страницы
        breaks = self.doc.findall(f".//{w('pageBreakBefore')}")
        self.assertEqual(len(breaks), 4)

    def test_photos_fit_page(self):
        body = self.doc.find(w("body"))
        page_h, _ = converter._page_geometry(body)
        for ext in self.doc.iter(f"{{{converter.WP_NS}}}extent"):
            self.assertLessEqual(int(ext.get("cy")), converter.PHOTO_MAX_H)
            self.assertLessEqual(int(ext.get("cx")), converter.PHOTO_MAX_W)
            self.assertLess(int(ext.get("cy")) / 12700, page_h)

    def test_no_warnings(self):
        self.assertFalse([x for x in self.logs if "Внимание" in x or "ВНИМАНИЕ" in x], self.logs)


class SampleProtocolTest(unittest.TestCase):
    """ПИ 2610101 (собранный из PDF) → все 173 блока, другой инженер."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out, cls.logs = run_conversion(FIXTURE, cls.tmp.name)
        cls.doc, cls.names = read_doc(cls.out)
        cls.texts = plain_texts(cls.doc)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_components_match_reference(self):
        rows = component_rows(self.doc)
        expected = FIXTURE["expected"]["components"]
        self.assertEqual(len(rows), len(expected))
        for i, (row, exp) in enumerate(zip(rows, expected), 1):
            c0, c1, c2 = row
            with self.subTest(component=i):
                self.assertIn(f"Содержание в компоненте {i}:", c0)
                if 1 < i < len(expected):
                    self.assertEqual(c2[0], f"Дата и время проведения испытаний {exp['date']}")
                for key in ("T", "RH", "P", "U", "f"):
                    self.assertEqual(field(c2, converter.COND_PREFIXES[key]), exp[key], key)
                self.assertEqual([x for x in c1 if x], [wrap_pct(r) for r in exp["results"]])
                self.assertEqual(c2[c2.index("Результаты измерений:") + 1:], exp["details"])
                self.assertIn(exp["measured"], c2)
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 02.10.2026 09:40")
        self.assertEqual(rows[-1][2][0], "Дата и время проведения испытаний 02.10.2026 17:30")

    def test_header_table(self):
        for value in FIXTURE["expected"]["header"].values():
            self.assertIn(value, self.texts)

    def test_prep_row(self):
        self.assertIn("Количество компонентов составило 173 шт.", self.texts)
        self.assertIn("Дата и время проведения испытаний 02.10.2026 09:10", self.texts)

    def test_equipment_from_pi_then_fixed(self):
        tbl = next(t for t in self.doc.iter(w("tbl")) if "Наименование средств измерений" in "".join(
            para_text(p) for p in t.iter(w("p"))))
        rows = [[" ".join(para_text(p) for p in c.findall(w("p"))) for c in row_cells(tr)] for tr in table_rows(tbl)[1:]]
        self.assertEqual(len(rows), 11)
        self.assertEqual(rows[0][1:], ["Зав. № 940289, инв. № СИ1", "23.09.2027"])
        self.assertTrue(rows[0][0].endswith("4.12.107"))
        self.assertEqual(rows[3][1], "Зав. № CAA09097, инв. № СИ8")
        self.assertEqual([r[1] for r in rows[5:]], ["Инв. № ВО1", "Инв. № ВО18", "Инв. № ВО12",
                                                    "Инв. № ВО22", "Инв. № ВО26", "Инв. № ВО30"])

    def test_engineer_and_stamp(self):
        sig = [t for t in self.texts if "_" in t and ("Инженер" in t or "Руководитель" in t)]
        self.assertEqual(len(sig), 2)
        self.assertIn("М.С. Свистунов", sig[0])
        self.assertIn("Д.Р. Хасянов", sig[1])
        for t in sig:
            self.assertIn("02.10.2026", t)
        st = stamps(self.doc)
        self.assertIn("Владелец: Свистунов Максим Сергеевич", st[0])
        self.assertIn("Сертификат: 022779a70091b3e59d4a3c5d9bb6704b48", st[0])
        self.assertEqual(st[0], st[1])  # Choice + Fallback
        self.assertIn("Владелец: Хасянов Дамир Рафикович", st[2])

    def test_engineer_line_width_kept(self):
        line = next(t for t in self.texts if "Инженер-испытатель" in t and "_" in t)
        ref = next(t for t in REAL_EXPECTED["body"] if "Инженер-испытатель" in t and "_" in t)
        self.assertEqual(len(line), len(ref))

    def test_photos(self):
        captions = [t for t in self.texts if t.startswith("Фото №")]
        self.assertEqual(captions, [f"Фото №{k}" for k in range(1, 13)])
        self.assertIn("126÷173 компоненты", self.texts)
        media = [n for n in self.names if n.startswith("word/media/") and not n.endswith("/")]
        self.assertEqual(len(media), 12)
        self.assertIn("Приложение № 1 на 12 страницах.", self.texts)

    def test_no_duplicate_para_ids(self):
        ids = [el.get(f"{{{converter.W14_NS}}}paraId") for el in self.doc.iter(w("p"))]
        ids = [x for x in ids if x]
        self.assertEqual(len(ids), len(set(ids)))


class EdgeCasesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def fixture(self, n_components=3):
        fx = copy.deepcopy(FIXTURE)
        fx["components"] = fx["components"][:n_components]
        fx["photos"] = fx["photos"][:2]
        return fx

    def test_small_n_and_no_photos(self):
        fx = self.fixture(3)
        fx["photos"] = []
        out, logs = run_conversion(fx, self.tmp.name)
        doc, names = read_doc(out)
        rows = component_rows(doc)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 02.10.2026 09:40")
        self.assertEqual(rows[1][2][0], "Дата и время проведения испытаний 02.10.2026")
        self.assertEqual(rows[2][2][0], "Дата и время проведения испытаний 02.10.2026 17:30")
        self.assertEqual(field(rows[2][2], "Температура"), "23,5")
        texts = plain_texts(doc)
        self.assertIn("Фото №1", texts)
        self.assertIn("1÷3 компоненты", texts)
        self.assertIn("Приложение № 1 на 2 страницах.", texts)
        self.assertEqual(doc.findall(f".//{{{converter.A_NS}}}blip"), [])
        self.assertEqual([n for n in names if n.startswith("word/media/") and not n.endswith("/")], [])
        self.assertTrue(any("Приложение №1 с фото" in x for x in logs))

    def test_single_component(self):
        out, _ = run_conversion(self.fixture(1), self.tmp.name)
        rows = component_rows(read_doc(out)[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 02.10.2026 09:40")

    def test_unknown_engineer_warns(self):
        fx = self.fixture(2)
        fx["engineer"] = "А.Б. Новиков"
        out, logs = run_conversion(fx, self.tmp.name)
        self.assertTrue(any("А.Б. Новиков" in x and "штамп" in x for x in logs))
        texts = plain_texts(read_doc(out)[0])
        self.assertTrue(any("А.Б. Новиков" in t and "_" in t for t in texts))

    def test_engineers_file(self):
        with open(os.path.join(self.tmp.name, converter.ENGINEERS_FILE), "w", encoding="utf-8") as f:
            json.dump({"А.Б.Новиков": {"owner": "Новиков Андрей Борисович", "certificate": "abc",
                                        "valid": "с 01.01.2026 по 01.01.2027"}}, f, ensure_ascii=False)
        eng = converter.load_engineers((self.tmp.name,))
        self.assertEqual(eng["А.Б. Новиков"]["owner"], "Новиков Андрей Борисович")
        self.assertIn("С.Д. Рождественский", eng)

    def test_pressure_two_decimals(self):
        fx = self.fixture(3)
        fx["conditions"][2][1] = "98,67 – 99,96"
        out, _ = run_conversion(fx, self.tmp.name)
        rows = component_rows(read_doc(out)[0])
        self.assertEqual(field(rows[0][2], "Атмосферное"), "986,7 (98,67 кПа)")
        self.assertEqual(field(rows[2][2], "Атмосферное"), "999,6 (99,96 кПа)")

    def test_two_day_period(self):
        fx = self.fixture(10)
        fx["paragraphs"] = [p.replace("02.10.2026 с 09:10 до 17:35", "01.10.2026 с 14:00 по 02.10.2026 12:00")
                            for p in fx["paragraphs"]]
        out, _ = run_conversion(fx, self.tmp.name)
        doc = read_doc(out)[0]
        dates = [r[2][0].replace("Дата и время проведения испытаний ", "") for r in component_rows(doc)]
        self.assertEqual(dates[0], "01.10.2026 14:30")
        self.assertEqual(dates[-1], "02.10.2026 11:55")
        boundary = max(i for i, d in enumerate(dates) if d.startswith("01.10.2026"))
        self.assertEqual(dates[boundary], f"01.10.2026 {converter.WORKDAY_END}")
        self.assertEqual(dates[boundary + 1], f"02.10.2026 {converter.WORKDAY_START}")
        self.assertTrue(all(len(d) == 10 for d in dates[1:boundary] + dates[boundary + 2:-1]))
        texts = plain_texts(doc)
        self.assertIn("с 01.10.2026 14:00 по 02.10.2026 12:00", texts)
        sig = [t for t in texts if "_" in t and "Инженер" in t]
        self.assertIn("02.10.2026", sig[0])  # дата подписи = дата окончания испытаний

    def test_explicit_day_windows(self):
        windows = converter.parse_test_period(
            "Испытания проводились 01.10.2026 с 10:00 до 17:00, 02.10.2026 с 09:30 до 11:15 по адресу: …")
        self.assertEqual([(x.day.day, x.start, x.end) for x in windows], [(1, "10:00", "17:00"), (2, "09:30", "11:15")])

    def test_component_table_split_across_tables(self):
        fx = self.fixture(4)
        pi_path = os.path.join(self.tmp.name, "pi.docx")
        build(fx, pi_path)
        with zipfile.ZipFile(pi_path) as z:
            files = {n: z.read(n) for n in z.namelist()}
        doc = etree.fromstring(files["word/document.xml"])
        tbl = next(t for t in doc.iter(w("tbl")) if "Ртуть" in "".join(para_text(p) for p in t.iter(w("p"))))
        rows = table_rows(tbl)
        cut = 2 + 4 + 2  # 2 строки шапки + компонент 1 (4 строки) + 2 строки компонента 2
        new_tbl = etree.Element(w("tbl"))
        for el in tbl:
            if el.tag != w("tr"):
                new_tbl.append(copy.deepcopy(el))
        for r in rows[:2]:
            new_tbl.append(copy.deepcopy(r))
        for r in rows[cut:]:
            new_tbl.append(r)
        first_cell = row_cells(rows[cut])[0]
        first_cell.find(w("tcPr")).find(w("vMerge")).set(w("val"), "restart")
        tbl.addnext(new_tbl)
        files["word/document.xml"] = etree.tostring(doc)
        with zipfile.ZipFile(pi_path, "w") as z:
            for n, d in files.items():
                z.writestr(n, d)
        pi = converter.parse_pi(pi_path)
        self.assertEqual([len(c.elements) for c in pi.components],
                         [len(c["elements"]) for c in fx["components"]])

    def test_auto_numbered_components(self):
        """В настоящих ПИ номер компонента — автонумерация, ячейка пустая."""
        pi = converter.parse_pi(REAL_PI)
        self.assertEqual(len(pi.components), 65)
        self.assertEqual([c.number for c in pi.components[:3]], ["1", "2", "3"])
        self.assertEqual(pi.engineer, "С.Д. Рождественский")

    def test_pdf_rejected(self):
        with self.assertRaises(converter.ConversionError):
            converter.parse_pi("ПИ_2610101.pdf")

    def test_interp_rule(self):
        self.assertEqual(converter.interp("23,0", "23,5", 1, 173, 1), 23.0)
        self.assertEqual(converter.interp("23,0", "23,5", 173, 173, 1), 23.5)
        self.assertEqual(converter.interp("23,0", "23,5", 5, 5, 1), 23.5)

    def test_sigma_reproducible(self):
        path = os.path.join(self.tmp.name, "pi.docx")
        build(self.fixture(3), path)
        pi = converter.parse_pi(path)
        self.assertEqual(converter.generate_sigmas(pi), converter.generate_sigmas(pi))
        self.assertEqual(converter.generate_sigmas(pi)[0], [9, 27, 21, 15])


if __name__ == "__main__":
    unittest.main()
