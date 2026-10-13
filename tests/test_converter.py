"""Тесты конвертера: python -m unittest discover -s tests -v"""
import copy
import json
import os
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

import converter  # noqa: E402
from converter import para_text, row_cells, table_rows, w  # noqa: E402
from make_sample_pi import build  # noqa: E402

TEMPLATE = os.path.join(ROOT, converter.TEMPLATE_NAME)
with open(os.path.join(HERE, "fixtures", "pi_2610101.json"), encoding="utf-8") as f:
    FIXTURE = json.load(f)


def run_conversion(fx, tmp):
    pi_path = os.path.join(tmp, f"ПИ_{fx['number']}.docx")
    build(fx, pi_path)
    logs = []
    out = converter.convert(pi_path, TEMPLATE, tmp, logs.append)
    return out, logs


def read_doc(path, part="word/document.xml"):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(part)), z.namelist()


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


def field(c2, prefix):
    return next(x for x in c2 if x.startswith(prefix)).split(":", 1)[1]


class SampleProtocolTest(unittest.TestCase):
    """ПИ 2610101 → ПЗ совпадает с эталонным ПЗ по всем 173 блокам."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out, cls.logs = run_conversion(FIXTURE, cls.tmp.name)
        cls.doc, cls.names = read_doc(cls.out)
        cls.texts = [para_text(p) for p in cls.doc.iter(w("p"))]

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_file_name(self):
        self.assertEqual(os.path.basename(self.out), "ПЗ_2610101_от_05_10_2026.docx")

    def test_components_match_reference(self):
        rows = component_rows(self.doc)
        expected = FIXTURE["expected"]["components"]
        self.assertEqual(len(rows), len(expected))
        for i, (row, exp) in enumerate(zip(rows, expected), 1):
            c0, c1, c2 = row
            with self.subTest(component=i):
                self.assertIn(f"Содержание в компоненте {i}:", c0)
                self.assertEqual(c2[0], f"Дата и время проведения испытаний {exp['date']}")
                for key in ("T", "RH", "P", "U", "f"):
                    prefix = converter.COND_PREFIXES[key]
                    self.assertEqual(field(c2, prefix), exp[key], key)
                self.assertEqual([x for x in c1 if x], exp["results"])
                self.assertEqual(c2[c2.index("Результаты измерений:") + 1:], exp["details"])
                self.assertIn(exp["measured"], c2)

    def test_header_table(self):
        exp = FIXTURE["expected"]["header"]
        for value in exp.values():
            self.assertIn(value, self.texts)
        for value in ("23,0 – 23,5", "43,8 – 44,2", "101,4 – 101,5", "222,6 – 226,2", "49,99"):
            self.assertIn(value, self.texts)

    def test_prep_row(self):
        self.assertIn("Количество компонентов составило 173 шт.", self.texts)
        self.assertTrue(any(t.startswith("Пробоподготовка. После анализа Блок питания, Торговой марки: GAMDIAS "
                                         "Модель: THOR P2-1350P (далее по тексту «образец»)") for t in self.texts))

    def test_equipment_set_b(self):
        for _, number, due in converter.EQUIPMENT_SET_B:
            self.assertIn(number, self.texts)
            self.assertIn(due, self.texts)
        self.assertNotIn("24.09.2026", self.texts)

    def test_signature_dates(self):
        sig = [t for t in self.texts if "Свистунов" in t and "_" in t or "Хасянов" in t and "_" in t]
        self.assertEqual(len(sig), 2)
        for t in sig:
            self.assertIn("05.10.2026", t)
            self.assertNotIn("17.07.2026", t)

    def test_photos(self):
        captions = [t for t in self.texts if t.startswith("Фото №")]
        self.assertEqual(captions, [f"Фото №{k}" for k in range(1, 13)])
        self.assertIn("Компоненты. Блок питания, Торговой марки: GAMDIAS Модель: THOR P2-1350P.\n126÷173 компоненты",
                      self.texts)
        self.assertNotIn("DEXP", " ".join(self.texts))
        media = [n for n in self.names if n.startswith("word/media/") and not n.endswith("/")]
        self.assertEqual(len(media), 12)
        blips = self.doc.findall(f".//{{{converter.A_NS}}}blip")
        self.assertEqual(len(blips), 12)
        self.assertIn("Приложение № 1 на 12 страницах.", self.texts)

    def test_header_part(self):
        hdr, _ = read_doc(self.out, "word/header1.xml")
        texts = [para_text(p) for p in hdr.iter(w("p"))]
        self.assertIn("№ 63ТР-НИ от 28.08.2026", texts)
        xml = etree.tostring(hdr, encoding="unicode")
        self.assertIn("NUMPAGES", xml)

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
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 02.10.2026 09:10")
        self.assertEqual(rows[1][2][0], "Дата и время проведения испытаний 02.10.2026")
        self.assertEqual(rows[2][2][0], "Дата и время проведения испытаний 02.10.2026 17:35")
        self.assertEqual(field(rows[2][2], "Температура"), "23,5")
        texts = [para_text(p) for p in doc.iter(w("p"))]
        self.assertIn("Фото №1", texts)
        self.assertIn("Компоненты. Блок питания, Торговой марки: GAMDIAS Модель: THOR P2-1350P.\n1÷3 компоненты", texts)
        self.assertEqual(doc.findall(f".//{{{converter.A_NS}}}blip"), [])
        self.assertEqual([n for n in names if n.startswith("word/media/") and not n.endswith("/")], [])
        self.assertTrue(any("Приложение №1 с фото" in x for x in logs))

    def test_single_component(self):
        out, _ = run_conversion(self.fixture(1), self.tmp.name)
        rows = component_rows(read_doc(out)[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][2][0], "Дата и время проведения испытаний 02.10.2026 09:10")

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
        # день 1: 14:00–18:00 (240 мин), день 2: 09:00–12:00 (180 мин)
        self.assertEqual(dates[0], "01.10.2026 14:00")
        self.assertEqual(dates[-1], "02.10.2026 12:00")
        boundary = max(i for i, d in enumerate(dates) if d.startswith("01.10.2026"))
        self.assertEqual(dates[boundary], f"01.10.2026 {converter.WORKDAY_END}")
        self.assertEqual(dates[boundary + 1], f"02.10.2026 {converter.WORKDAY_START}")
        self.assertTrue(all(len(d) == 10 for d in dates[1:boundary] + dates[boundary + 2:-1]))
        texts = [para_text(p) for p in doc.iter(w("p"))]
        self.assertIn("с 01.10.2026 14:00 по 02.10.2026 12:00", texts)

    def test_explicit_day_windows(self):
        windows = converter.parse_test_period(
            "Испытания проводились 01.10.2026 с 10:00 до 17:00, 02.10.2026 с 09:30 до 11:15 по адресу: …")
        self.assertEqual([(x.day.day, x.start, x.end) for x in windows], [(1, "10:00", "17:00"), (2, "09:30", "11:15")])

    def test_component_table_split_across_tables(self):
        fx = self.fixture(4)
        pi_path = os.path.join(self.tmp.name, "pi.docx")
        build(fx, pi_path)
        # разрезаем таблицу компонентов посередине компонента 2
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

    def test_pdf_rejected(self):
        with self.assertRaises(converter.ConversionError):
            converter.parse_pi("ПИ_2610101.pdf")

    def test_interp_rule(self):
        self.assertEqual(converter.interp("23,0", "23,5", 1, 173, 1), 23.0)
        self.assertEqual(converter.interp("23,0", "23,5", 173, 173, 1), 23.5)
        self.assertEqual(converter.interp("23,0", "23,5", 5, 5, 1), 23.5)

    def test_sigma_reproducible(self):
        pi = converter.parse_pi(self._pi(3))
        self.assertEqual(converter.generate_sigmas(pi), converter.generate_sigmas(pi))
        self.assertEqual(converter.generate_sigmas(pi)[0], [9, 27, 21, 15])

    def _pi(self, n):
        path = os.path.join(self.tmp.name, "pi.docx")
        build(self.fixture(n), path)
        return path


if __name__ == "__main__":
    unittest.main()
