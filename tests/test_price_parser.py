"""
Юнит-тесты для price_parser_sheets.py.

Покрывают чистую логику (извлечение цены/толщины/цвета из текста,
сопоставление толщины, парсинг табличных цен из HTML, стабильность
ключа сохраненного выбора пользователя) без обращения к реальным
сайтам и без запуска настоящего браузера.

Selenium и webdriver_manager не обязательны для запуска этих тестов:
если они не установлены, вместо них подставляются заглушки-модули,
поскольку price_parser_sheets.py импортирует их на уровне модуля, но
код, покрытый тестами ниже, их не использует.

Запуск:
    python -m unittest discover -s tests
"""
import importlib.util
import os
import re
import sys
import tempfile
import types
import unittest
from unittest.mock import patch, MagicMock

from openpyxl import Workbook, load_workbook


def _install_selenium_stubs():
    """Подставляет минимальные заглушки selenium/webdriver_manager,
    если они не установлены, чтобы модуль можно было импортировать."""
    try:
        import selenium  # noqa: F401
        return
    except ImportError:
        pass

    for mod_name in [
        'selenium', 'selenium.webdriver', 'selenium.webdriver.common.by',
        'selenium.webdriver.support.ui', 'selenium.webdriver.support',
        'selenium.webdriver.chrome.service', 'selenium.webdriver.chrome.options',
        'selenium.common.exceptions', 'webdriver_manager', 'webdriver_manager.chrome',
    ]:
        sys.modules[mod_name] = types.ModuleType(mod_name)

    sys.modules['selenium'].webdriver = types.SimpleNamespace(Chrome=object)
    sys.modules['selenium.webdriver.common.by'].By = types.SimpleNamespace(
        CSS_SELECTOR='css selector', XPATH='xpath', TAG_NAME='tag name'
    )
    sys.modules['selenium.webdriver.support.ui'].WebDriverWait = object
    sys.modules['selenium.webdriver.support'].expected_conditions = types.SimpleNamespace()
    sys.modules['selenium.webdriver.chrome.service'].Service = object
    sys.modules['selenium.webdriver.chrome.options'].Options = object
    sys.modules['selenium.common.exceptions'].TimeoutException = Exception
    sys.modules['webdriver_manager.chrome'].ChromeDriverManager = object


_install_selenium_stubs()

_MODULE_PATH = os.path.join(os.path.dirname(__file__), '..', 'price_parser_sheets.py')
_spec = importlib.util.spec_from_file_location('price_parser_sheets', _MODULE_PATH)
pps = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pps)


def make_parser():
    p = pps.PriceParserWithSheets.__new__(pps.PriceParserWithSheets)
    p.rounding_mode = 'ceil'
    return p


def make_workbook(tmpdir, name, url, selector=""):
    """Creates a one-row test Прайс-лист workbook and returns its path."""
    xlsx_path = os.path.join(tmpdir, "test.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Прайс-лист"
    ws.append(["Название", "URL", "Цена", "Селектор", "Характеристика", "Дата обновления"])
    ws.append([name, url, "", selector, "", ""])
    wb.save(xlsx_path)
    return xlsx_path


def make_fake_response(url, text, status_code=200):
    resp = MagicMock()
    resp.status_code = status_code
    resp.url = url
    resp.encoding = 'utf-8'
    resp.text = text
    return resp


def make_parser_for_product(tmpdir, name, url, selector=""):
    """Convenience: build a workbook + loaded parser in one call."""
    xlsx_path = make_workbook(tmpdir, name, url, selector=selector)
    p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
    p.load_excel_data()
    return p


def no_selenium():
    """
    bestly.ru always tries Selenium first in parse_single_product(). Without
    this patch, tests using a bestly.ru URL would make a real
    webdriver_manager network call (to look up a chromedriver version)
    before falling back to the mocked requests.get — slow, and a violation
    of this suite's "no network needed" guarantee. Forces that fallback
    immediately instead.
    """
    return patch.object(pps.PriceParserWithSheets, 'get_with_selenium', return_value=None)


class SafeStrTests(unittest.TestCase):
    def test_none_returns_default(self):
        self.assertEqual(pps.safe_str(None), '')

    def test_nan_returns_default(self):
        self.assertEqual(pps.safe_str(float('nan')), '')

    def test_regular_value(self):
        self.assertEqual(pps.safe_str('hello'), 'hello')
        self.assertEqual(pps.safe_str(123), '123')


class ExtractPriceFromTextTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_price_with_currency_and_thousands_separator(self):
        self.assertEqual(self.p.extract_price_from_text("6 388.80₽"), 6388.80)

    def test_price_with_comma_decimal(self):
        self.assertEqual(self.p.extract_price_from_text("от 1 234,56 руб."), 1234.56)

    def test_plain_integer(self):
        self.assertEqual(self.p.extract_price_from_text("7686"), 7686.0)

    def test_price_with_abbreviation(self):
        self.assertEqual(self.p.extract_price_from_text("Цена: 999 р."), 999.0)

    def test_none_and_empty(self):
        self.assertIsNone(self.p.extract_price_from_text(None))
        self.assertIsNone(self.p.extract_price_from_text(""))


class ExtractThicknessTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_thickness_with_mm_suffix(self):
        self.assertEqual(
            self.p.extract_thickness_from_product_name("Орг.стекло PLAZCRYL прозрачный 4мм"),
            "4.0",
        )

    def test_thickness_before_color(self):
        self.assertEqual(
            self.p.extract_thickness_from_product_name("Орг.стекло 8мм прозрачный"),
            "8.0",
        )


class ExtractColorMaterialBrandTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_color_direct_match(self):
        self.assertEqual(
            self.p.extract_color_from_product_name("Орг.стекло PLAZCRYL прозрачный 4мм"),
            "прозрачный",
        )

    def test_color_variation_match(self):
        self.assertEqual(
            self.p.extract_color_from_product_name("Орг.стекло матовое 4мм"),
            "матовый",
        )

    def test_material_type(self):
        self.assertEqual(
            self.p.extract_material_type_from_product_name("Поликарбонат POLYGAL сотовый 4мм"),
            "поликарбонат",
        )

    def test_brand_uppercase(self):
        self.assertEqual(
            self.p.extract_brand_from_product_name("Орг.стекло PLAZCRYL прозрачный 4мм"),
            "PLAZCRYL",
        )

    def test_product_type(self):
        self.assertEqual(
            self.p.extract_product_type_from_product_name("Поликарбонат сотовый 4мм"),
            "сотовый",
        )


class RoundPriceTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_ceil_mode(self):
        self.assertEqual(self.p.round_price(10.001, mode='ceil'), 10.01)

    def test_floor_mode(self):
        self.assertEqual(self.p.round_price(10.999, mode='floor'), 10.99)

    def test_no_decimal_mode(self):
        self.assertEqual(self.p.round_price(10.01, mode='no_decimal'), 11)


class IsReasonablePriceTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_within_configured_range(self):
        self.assertTrue(self.p.is_reasonable_price(7686, "Поликарбонат POLYGAL сотовый 4мм"))

    def test_outside_configured_range(self):
        self.assertFalse(self.p.is_reasonable_price(500, "Поликарбонат POLYGAL сотовый 4мм"))

    def test_unknown_product_uses_wide_range(self):
        self.assertTrue(self.p.is_reasonable_price(500, "какой-то неизвестный товар"))

    def test_non_positive_price_rejected(self):
        self.assertFalse(self.p.is_reasonable_price(-1, "что угодно"))
        self.assertFalse(self.p.is_reasonable_price(0, "что угодно"))


class ThicknessMatchingTests(unittest.TestCase):
    """Regression tests for the thickness-matching helpers (dedup + bugfix stages)."""

    def setUp(self):
        self.p = make_parser()

    def test_matches_exact_mm_mention(self):
        self.assertTrue(self.p.check_thickness_in_text("Толщина: 4мм, цена 500р", "4"))

    def test_does_not_match_different_thickness(self):
        self.assertFalse(self.p.check_thickness_in_text("6мм лист", "4"))

    def test_no_false_positive_on_substring_of_longer_number(self):
        # Regression test: "4" must not match inside "14мм"/"24мм"/"144мм".
        self.assertFalse(self.p.check_thickness_in_text("14мм что-то не то", "4"))
        self.assertFalse(self.p.check_thickness_in_text("24мм странная толщина", "4"))
        self.assertFalse(self.p.check_thickness_in_text("144мм странное совпадение", "4"))

    def test_exact_mode_requires_whole_cell(self):
        self.assertTrue(self.p._text_matches_thickness("4", "4.0", mode='exact'))
        self.assertFalse(self.p._text_matches_thickness("14", "4.0", mode='exact'))

    def test_find_best_match_by_thickness_picks_correct_candidate(self):
        elements = [
            {'text': '4мм прозрачный 1234 руб.'},
            {'text': '6мм прозрачный 2345 руб.'},
        ]
        best = self.p.find_best_match_by_thickness(elements, "4.0")
        self.assertIsNotNone(best)
        self.assertEqual(best['price'], 1234.0)


class BestlyTableParsingTests(unittest.TestCase):
    """Parse synthetic HTML tables shaped like the real bestly.ru pages."""

    def setUp(self):
        self.p = make_parser()

    def test_universal_table_parsing_bestly(self):
        html = """
        <html><body>
        <table>
        <tr><th>Толщина</th><th>Размер</th><th>Цена</th></tr>
        <tr><td>4 мм</td><td>2050x3050</td><td>7 686 ₽</td></tr>
        <tr><td>6 мм</td><td>2050x3050</td><td>10 500 ₽</td></tr>
        </table>
        </body></html>
        """
        price = self.p.universal_table_parsing_bestly(html, "Поликарбонат POLYGAL сотовый 4мм")
        self.assertEqual(price, 7686.0)

    def test_parse_orgsteklo_table_improved(self):
        html = """
        <html><body>
        <table>
        <tr><td>4мм</td><td>2050x3050</td><td>12 345.00</td></tr>
        <tr><td>6мм</td><td>2050x3050</td><td>15 000.00</td></tr>
        </table>
        </body></html>
        """
        price = self.p.parse_orgsteklo_table_improved(html, "Орг.стекло PLAZCRYL прозрачный 4мм")
        self.assertEqual(price, 12345.0)

    def test_parse_bestly_orgsteklo_table(self):
        html = """
        <html><body>
        <table>
        <tr><td></td><td>цвет</td><td>толщина</td><td>размер</td><td>цена</td></tr>
        <tr><td><input></td><td>прозрачный</td><td>4</td><td>2050x3050</td><td>9 999 ₽</td></tr>
        <tr><td><input></td><td>прозрачный</td><td>6</td><td>2050x3050</td><td>13 000 ₽</td></tr>
        </table>
        </body></html>
        """
        price = self.p.parse_bestly_orgsteklo_table(html, "Орг.стекло PLAZCRYL прозрачный 4мм")
        self.assertEqual(price, 9999.0)


class UserSelectionPersistenceTests(unittest.TestCase):
    """Regression test for the hash()-instability bugfix (stage 1)."""

    def setUp(self):
        self._old_cwd = os.getcwd()
        self._tmpdir = tempfile.mkdtemp()
        os.chdir(self._tmpdir)

    def tearDown(self):
        os.chdir(self._old_cwd)

    def test_same_selection_produces_stable_key_across_instances(self):
        p1 = make_parser()
        p1.user_selections = {}
        p1.save_user_selection("https://bestly.ru/x.html", ".price", "1234 руб.", 1234.0)
        key1 = list(p1.user_selections.keys())[0]

        p2 = make_parser()
        p2.user_selections = {}
        p2.load_user_selections()
        p2.save_user_selection("https://bestly.ru/x.html", ".price", "1234 руб.", 1234.0)
        key2 = list(p2.user_selections.keys())[0]

        self.assertEqual(key1, key2)
        self.assertEqual(len(p2.user_selections), 1)


class ExtractDomainTests(unittest.TestCase):
    def setUp(self):
        self.p = make_parser()

    def test_strips_www_and_scheme(self):
        self.assertEqual(self.p.extract_domain("https://www.bestly.ru/catalog/x.html"), "bestly.ru")

    def test_empty_for_invalid_input(self):
        self.assertEqual(self.p.extract_domain(None), '')


class StripTrackingParamsTests(unittest.TestCase):
    def test_removes_ysclid(self):
        self.assertEqual(
            pps.strip_tracking_params("https://expo-torg.ru/product/x/?ysclid=mgz4hztunr807680454"),
            "https://expo-torg.ru/product/x/",
        )

    def test_removes_utm_params_keeps_real_ones(self):
        self.assertEqual(
            pps.strip_tracking_params("https://site.ru/p/?utm_source=yandex&utm_medium=cpc&real=1"),
            "https://site.ru/p/?real=1",
        )

    def test_keeps_legitimate_query_param(self):
        self.assertEqual(
            pps.strip_tracking_params("https://site.ru/p/?id=123&ysclid=abc"),
            "https://site.ru/p/?id=123",
        )

    def test_no_query_string_unchanged(self):
        self.assertEqual(pps.strip_tracking_params("https://site.ru/p/"), "https://site.ru/p/")

    def test_empty_and_none(self):
        self.assertEqual(pps.strip_tracking_params(""), "")
        self.assertIsNone(pps.strip_tracking_params(None))


class GetWithRequestsUrlHealingTests(unittest.TestCase):
    """Self-healing of a product URL when the site redirects to a new address."""

    def setUp(self):
        self.p = make_parser()

    def test_captures_redirect_target(self):
        fake_response = MagicMock()
        fake_response.status_code = 200
        fake_response.url = "https://expo-torg.ru/catalog/ldsp/new-slug/"
        fake_response.encoding = 'utf-8'
        fake_response.text = "<html>ok</html>"

        with patch.object(pps.requests, 'get', return_value=fake_response):
            html = self.p.get_with_requests("https://expo-torg.ru/catalog/ldsp/old-slug/")

        self.assertEqual(html, "<html>ok</html>")
        self.assertEqual(self.p.last_status_code, 200)
        self.assertEqual(self.p.last_fetched_url, "https://expo-torg.ru/catalog/ldsp/new-slug/")

    def test_no_redirect_leaves_url_unchanged(self):
        fake_response = MagicMock()
        fake_response.status_code = 200
        fake_response.url = "https://expo-torg.ru/catalog/ldsp/same/"
        fake_response.encoding = 'utf-8'
        fake_response.text = "<html>ok</html>"

        with patch.object(pps.requests, 'get', return_value=fake_response):
            self.p.get_with_requests("https://expo-torg.ru/catalog/ldsp/same/")

        self.assertEqual(self.p.last_fetched_url, "https://expo-torg.ru/catalog/ldsp/same/")

    def test_404_is_recorded_and_returns_none(self):
        fake_response = MagicMock()
        fake_response.status_code = 404

        with patch.object(pps.requests, 'get', return_value=fake_response):
            html = self.p.get_with_requests("https://expo-torg.ru/catalog/ldsp/gone/")

        self.assertIsNone(html)
        self.assertEqual(self.p.last_status_code, 404)
        self.assertIsNone(self.p.last_fetched_url)


class ParseSingleProductRegressionTests(unittest.TestCase):
    """
    Regression tests for a real bug found while testing the URL self-heal
    feature: the "standard search" block in parse_single_product() was
    nested entirely inside `if selector and selector.strip():`, so:
      1. a product with NO selector set (the common case before a selector
         is manually chosen) crashed with UnboundLocalError on `result`
         instead of searching for the price at all;
      2. a product whose specified selector failed, but whose price WAS
         found via the bestly.ru-specific fallback selectors, had that
         price silently wiped back to None/"Цена не найдена" by a
         duplicated, mis-nested copy of the same search block.
    Both cases are exercised here with mocked HTTP responses (no network).
    """

    def test_no_selector_does_not_crash_and_finds_price(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Композитная панель белая 4мм",
                "https://bestly.ru/catalog/composite/x.html", selector="",
            )
            fake_response = make_fake_response(
                "https://bestly.ru/catalog/composite/x.html",
                "<html><body><div class='item_price'>5000 руб.</div></body></html>",
            )
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5000.0)
        self.assertNotIn('Ошибка', result['status'])

    def test_wrong_selector_falls_back_without_wiping_found_price(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Композитная панель белая 4мм",
                "https://bestly.ru/catalog/composite/x.html",
                selector=".wrong-selector-does-not-exist",
            )
            fake_response = make_fake_response(
                "https://bestly.ru/catalog/composite/x.html",
                "<html><body><div class='item_price'>5000 руб.</div></body></html>",
            )
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5000.0)

    def test_correct_selector_still_works(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Композитная панель белая 4мм",
                "https://bestly.ru/catalog/composite/x.html", selector=".item_price",
            )
            fake_response = make_fake_response(
                "https://bestly.ru/catalog/composite/x.html",
                "<html><body><div class='item_price'>5000 руб.</div></body></html>",
            )
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5000.0)
        self.assertIn('указанный селектор', result['status'])

    def test_nothing_found_is_a_clean_status_not_a_crash(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП белый 16мм", "https://expo-torg.ru/catalog/ldsp/y.html", selector="",
            )
            fake_response = make_fake_response(
                "https://expo-torg.ru/catalog/ldsp/y.html",
                "<html><body><p>ничего интересного</p></body></html>",
            )
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertIsNone(result['price'])
        self.assertNotIn('Ошибка', result['status'])

    def test_redirected_url_is_resolved_and_written_back_to_excel(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП белый 16мм",
                "https://expo-torg.ru/catalog/ldsp/old-slug/?ysclid=abc123", selector="",
            )
            fake_response = make_fake_response(
                "https://expo-torg.ru/catalog/ldsp/new-slug-16mm/",
                "<html><body><div class='price'>1234 руб.</div></body></html>",
            )
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

            self.assertEqual(result['url'], "https://expo-torg.ru/catalog/ldsp/new-slug-16mm/")

            p.results = [result]
            self.assertTrue(p.save_results_to_excel())

            wb2 = load_workbook(p.excel_file)
            saved_url = wb2["Прайс-лист"].cell(row=2, column=2).value
            self.assertEqual(saved_url, "https://expo-torg.ru/catalog/ldsp/new-slug-16mm/")

    def test_empty_result_url_does_not_blank_existing_cell(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Прайс-лист"
            ws.append(["Название", "URL", "Цена", "Селектор", "Характеристика", "Дата обновления"])
            ws.append(["Товар без URL", "", "", "", "", ""])
            wb.save(xlsx_path)

            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            p.results = [{
                'index': 0, 'name': 'Товар без URL', 'url': '', 'price': None,
                'characteristic': '', 'selector_used': '', 'best_found_selector': '',
                'status': 'URL не указан', 'rounding_mode': 'ceil',
                'timestamp': '2026-01-01 00:00:00',
            }]
            self.assertTrue(p.save_results_to_excel())

            wb2 = load_workbook(xlsx_path)
            cell = wb2["Прайс-лист"].cell(row=2, column=2).value
            self.assertIn(cell, (None, ''))


class WrongPriceRegressionTests(unittest.TestCase):
    """
    Regression tests for real "wrong price" cases reported from a live
    bestly.ru parsing run:
      1. The generic '[data-price]' fallback selector sometimes matches a
         quantity-stepper widget (data-price="1") instead of the actual
         price element, so "1" was accepted as the price outright.
      2. A "soft 404" — the site returns HTTP 200 but the page content is
         an actual "not found" page — had its title/h1 text ("Страница не
         найдена (404 Not Found)") auto-detected as a price of 404.
    Both are fixed by (a) sanity-checking every extracted price with
    is_reasonable_price() before accepting it in
    find_price_with_selector_and_name()/find_price_and_name_on_page(), and
    (b) looks_like_error_page() short-circuiting parse_single_product()
    before any price search runs.
    """

    def test_data_price_quantity_widget_is_rejected_in_favor_of_real_price(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Иглопробивной ковролин EXPORADU",
                "https://bestly.ru/catalog/exporadu.html", selector="",
            )
            html = (
                "<html><body>"
                "<div class='quantity-stepper' data-price='1'>шт.</div>"
                "<div class='item_price'>2450 руб.</div>"
                "</body></html>"
            )
            fake_response = make_fake_response("https://bestly.ru/catalog/exporadu.html", html)
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 2450.0)

    def test_data_price_one_alone_is_reported_as_not_found_not_as_price_one(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Товар без реальной цены на странице",
                "https://bestly.ru/catalog/nothing_real.html", selector="",
            )
            html = "<html><body><div class='quantity-stepper' data-price='1'>шт.</div></body></html>"
            fake_response = make_fake_response("https://bestly.ru/catalog/nothing_real.html", html)
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertIsNone(result['price'])

    def test_plausible_data_price_still_works(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Товар с ценой в data-price",
                "https://bestly.ru/catalog/real_dataprice.html", selector="",
            )
            html = "<html><body><div data-price='3500'>3500 руб.</div></body></html>"
            fake_response = make_fake_response("https://bestly.ru/catalog/real_dataprice.html", html)
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 3500.0)

    def test_soft_404_page_is_not_mistaken_for_a_price(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Баннерная ткань литая Blackback GLP",
                "https://bestly.ru/catalog/bannernaya_tkan_litaya_blackback_glp.html", selector="",
            )
            html = (
                "<html><head><title>Страница не найдена (404 Not Found)</title></head>"
                "<body><h1>Страница не найдена (404 Not Found)</h1></body></html>"
            )
            fake_response = make_fake_response(
                "https://bestly.ru/catalog/bannernaya_tkan_litaya_blackback_glp.html", html,
            )
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertIsNone(result['price'])
        self.assertIn('заглушка', result['status'])

    def test_sku_containing_404_does_not_false_positive(self):
        """A product whose SKU happens to contain '404' must still parse normally."""
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Артикул 404XYZ товар", "https://bestly.ru/catalog/sku404.html", selector="",
            )
            html = (
                "<html><head><title>Товар 404XYZ - каталог</title></head>"
                "<body><h1>Товар 404XYZ</h1><div class='item_price'>999 руб.</div></body></html>"
            )
            fake_response = make_fake_response("https://bestly.ru/catalog/sku404.html", html)
            with patch.object(pps.requests, 'get', return_value=fake_response), no_selenium():
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 999.0)

    def test_looks_like_error_page_direct(self):
        p = make_parser()
        self.assertTrue(p.looks_like_error_page(
            "<html><head><title>Страница не найдена</title></head><body></body></html>"
        ))
        self.assertFalse(p.looks_like_error_page(
            "<html><head><title>Товар XYZ</title></head><body><div class='price'>100</div></body></html>"
        ))
        self.assertFalse(p.looks_like_error_page(""))
        self.assertFalse(p.looks_like_error_page(None))


class DeadLinkRecoveryTests(unittest.TestCase):
    """
    Tests for recovering a dead product link by browsing its catalog
    category (expo-torg.ru): when a product URL 404s or shows a "not
    found" page, the category it lived in usually still exists, so we
    open that category's listing and look for a product with a matching
    name instead of just giving up.

    The category markup used here (.product-item > .product-item-title >
    a[href][title]) is the real markup the user copied from expo-torg.ru.
    """

    REAL_CATEGORY_HTML = """
    <div class="product-item">
    <a class="product-item-image-wrapper merl_sect_img" href="/catalog/plastik/bumazhno-sloistyy-plastik-hpl/kompakt-plita-3150-vl-st9-12-4200-1860-mm-svetlo-seryy-gentas/" title="Компакт-плита 3150 VL/ST9 12*4200*1860 мм Светло-серый GENTAS" data-entity="image-wrapper"></a>
    <div class="product-item-title">
        <a href="/catalog/plastik/bumazhno-sloistyy-plastik-hpl/kompakt-plita-3150-vl-st9-12-4200-1860-mm-svetlo-seryy-gentas/" title="Компакт-плита 3150 VL/ST9 12*4200*1860 мм Светло-серый GENTAS">Компакт-плита 3150 VL/ST9 12*4200*1860 мм Светло-серый GENTAS</a>
    </div>
    </div>
    <div class="product-item">
    <div class="product-item-title">
        <a href="/catalog/plastik/bumazhno-sloistyy-plastik-hpl/kompakt-plita-3155-vl-st9-12-4200-1860-mm-antratsit-seryy-gentas/" title="Компакт-плита 3155 VL/ST9 12*4200*1860 мм Антрацит серый GENTAS">Компакт-плита 3155 VL/ST9 12*4200*1860 мм Антрацит серый GENTAS</a>
    </div>
    </div>
    """
    CATEGORY_URL = "https://expo-torg.ru/catalog/plastik/bumazhno-sloistyy-plastik-hpl/"

    def setUp(self):
        self.p = make_parser()

    def test_derive_category_url(self):
        self.assertEqual(
            self.p.derive_category_url(
                "https://expo-torg.ru/catalog/plastik/bumazhno-sloistyy-plastik-hpl/"
                "bsp-3190-vl-st9-0-8-3050-1300-mm-chyernyy-gentas/"
            ),
            "https://expo-torg.ru/catalog/plastik/bumazhno-sloistyy-plastik-hpl/",
        )
        self.assertIsNone(self.p.derive_category_url("https://expo-torg.ru/catalog/"))

    def test_parse_category_product_cards_real_markup(self):
        cards = self.p.parse_category_product_cards(self.REAL_CATEGORY_HTML, self.CATEGORY_URL)
        self.assertEqual(len(cards), 2)
        self.assertTrue(cards[0]['url'].endswith('kompakt-plita-3150-vl-st9-12-4200-1860-mm-svetlo-seryy-gentas/'))
        self.assertEqual(
            cards[1]['name'],
            'Компакт-плита 3155 VL/ST9 12*4200*1860 мм Антрацит серый GENTAS',
        )

    def test_match_disambiguates_near_duplicate_products(self):
        """
        Regression test for the exact real-world case reported: two
        products in the same category differ only by a numeric decor code
        (3150 vs 3155) and color (Светло-серый vs Антрацит серый). Picking
        the wrong one would silently record the wrong product's price.
        """
        cards = self.p.parse_category_product_cards(self.REAL_CATEGORY_HTML, self.CATEGORY_URL)
        target = "Компакт-плита 3155 VL/ST9 12*4200*1860 мм Антрацит серый GENTAS"
        best, scored = self.p.find_best_category_match(cards, target)
        self.assertIsNotNone(best, scored)
        self.assertIn('3155', best['url'])
        self.assertNotIn('3150', best['url'])

    def test_no_confident_match_returns_none(self):
        cards = [{'url': 'https://expo-torg.ru/catalog/x/unrelated/', 'name': 'Совершенно другой товар'}]
        best, scored = self.p.find_best_category_match(cards, "Компакт-плита 3155 GENTAS")
        self.assertIsNone(best)

    def test_recover_dead_link_skips_bestly(self):
        """bestly.ru has no category hierarchy to browse this way — must not attempt it."""
        result = self.p.recover_dead_link("https://bestly.ru/catalog/exporadu.html", "Ковролин EXPORADU")
        self.assertIsNone(result)

    def test_end_to_end_dead_link_is_recovered_in_parse_single_product(self):
        class FakeDriver:
            def __init__(self, html):
                self.page_source = html

            def get(self, url):
                pass

            def find_elements(self, by, selector):
                return [1, 2] if 'product-item' in selector else []

        with tempfile.TemporaryDirectory() as tmpdir:
            dead_url = (
                "https://expo-torg.ru/catalog/plastik/bumazhno-sloistyy-plastik-hpl/"
                "kompakt-plita-3155-vl-st9-12-4200-1860-mm-antratsit-seryy-gentas-OLD/"
            )
            replacement_url = (
                "https://expo-torg.ru/catalog/plastik/bumazhno-sloistyy-plastik-hpl/"
                "kompakt-plita-3155-vl-st9-12-4200-1860-mm-antratsit-seryy-gentas/"
            )
            p = make_parser_for_product(
                tmpdir, "Компакт-плита 3155 VL/ST9 12*4200*1860 мм Антрацит серый GENTAS",
                dead_url, selector="",
            )

            def fake_init_driver():
                p.driver = FakeDriver(self.REAL_CATEGORY_HTML)
                return True

            def fake_requests_get(url, **kwargs):
                if url == dead_url:
                    return make_fake_response(url, "", status_code=404)
                if url == replacement_url:
                    return make_fake_response(url, "<html><body><div class='price-integer'>7 450</div></body></html>")
                raise AssertionError(f"Unexpected URL requested: {url}")

            with patch.object(pps.PriceParserWithSheets, 'get_with_selenium', return_value=None), \
                 patch.object(p, 'init_selenium_driver', side_effect=fake_init_driver), \
                 patch.object(pps.requests, 'get', side_effect=fake_requests_get):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['url'], replacement_url)
        self.assertEqual(result['price'], 7450.0)

    def test_recovery_failure_falls_back_to_dead_link_status(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            dead_url = "https://expo-torg.ru/catalog/some-category/dead-product/"
            p = make_parser_for_product(tmpdir, "Совсем другой товар XYZ", dead_url, selector="")

            with patch.object(pps.PriceParserWithSheets, 'get_with_selenium', return_value=None), \
                 patch.object(p, 'init_selenium_driver', return_value=False), \
                 patch.object(pps.requests, 'get', return_value=make_fake_response(dead_url, "", status_code=404)):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertIsNone(result['price'])
        self.assertIn('404', result['status'])


class EditOrAddProductTests(unittest.TestCase):
    """
    Tests for the manual find/edit/add-product menu flow (menu item 7):
    searching by (partial) product name, editing any field including the
    URL — which the previous version of this menu item could not do at
    all — and appending a brand new product row.
    """

    def _make_two_row_workbook(self, tmpdir):
        xlsx_path = os.path.join(tmpdir, "test.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Прайс-лист"
        ws.append(["Название", "URL", "Цена", "Селектор", "Характеристика", "Дата обновления"])
        ws.append(["ЛДСП W908 ST2 16мм Белый Базовый Egger", "https://expo-torg.ru/old/", "2940", ".old-sel", "char", ""])
        ws.append(["ЛДСП W960 SM 18 Белый классический Egger", "https://expo-torg.ru/other/", "4035", "", "", ""])
        wb.save(xlsx_path)
        return xlsx_path

    def test_get_column_indices(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(
                p.get_column_indices(ws),
                {'Название': 1, 'URL': 2, 'Цена': 3, 'Селектор': 4, 'Характеристика': 5, 'Дата обновления': 6},
            )

    def test_write_product_fields_edits_url_and_clears_selector(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            ok = p._write_product_fields(0, {'URL': 'https://expo-torg.ru/new-url/', 'Селектор': ''})
            self.assertTrue(ok)

            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(ws.cell(row=2, column=2).value, 'https://expo-torg.ru/new-url/')
            self.assertIn(ws.cell(row=2, column=4).value, (None, ''))
            # Untouched fields (name, characteristic) survive unchanged
            self.assertEqual(ws.cell(row=2, column=1).value, "ЛДСП W908 ST2 16мм Белый Базовый Egger")
            self.assertEqual(ws.cell(row=2, column=5).value, 'char')

    def test_add_new_product_appends_row_and_strips_tracking_params(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            with patch('builtins.input', side_effect=[
                "Новый товар МДФ 10мм",
                "https://expo-torg.ru/new-product/?ysclid=abc123",
                "",  # unit -> defaults to 'шт.'
                "",  # selector
            ]):
                p._add_new_product()

            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(ws.max_row, 4)  # header + 2 existing + 1 new
            self.assertEqual(ws.cell(row=4, column=1).value, "Новый товар МДФ 10мм")
            self.assertEqual(ws.cell(row=4, column=2).value, "https://expo-torg.ru/new-product/")
            col_indices = p.get_column_indices(ws)
            self.assertEqual(ws.cell(row=4, column=col_indices['Единица']).value, 'шт.')

    def test_add_new_product_rejects_empty_name_or_url(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            with patch('builtins.input', side_effect=["", ]):
                p._add_new_product()  # empty name -> bail out before asking for URL
            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(ws.max_row, 3)  # unchanged: header + 2 existing rows

    def test_add_new_product_fills_first_empty_row_not_after_blank_gap(self):
        """
        Regression test: a real report showed a new product appended way
        below several pre-formatted-but-empty rows instead of into the
        first of them, because ws.max_row in openpyxl counts any touched
        (even just styled) cell, not just rows with real data.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            wb = load_workbook(xlsx_path)
            ws = wb["Прайс-лист"]
            # Simulate leftover "touched but empty" rows (e.g. just a unit
            # pre-filled, like in the reported spreadsheet) after row 3
            for r in (4, 5, 6):
                ws.cell(row=r, column=2).value = None  # touch the row without real data
                ws.cell(row=r, column=1).value = None
            ws.cell(row=4, column=3).value = None
            wb.save(xlsx_path)

            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            with patch('builtins.input', side_effect=[
                "Новый товар", "https://expo-torg.ru/gap-fill/", "", "",
            ]):
                p._add_new_product()

            ws2 = load_workbook(xlsx_path)["Прайс-лист"]
            # Must land in row 4 (the first empty one), not after row 6
            self.assertEqual(ws2.cell(row=4, column=1).value, "Новый товар")
            self.assertEqual(ws2.cell(row=4, column=2).value, "https://expo-torg.ru/gap-fill/")

    def test_find_and_edit_product_by_search_updates_url(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            p.load_excel_data()
            with patch('builtins.input', side_effect=[
                "w908",   # search query matches only the first row
                "1",      # pick it
                "",       # keep Название
                "https://expo-torg.ru/found-and-fixed/",  # new URL
                "-",      # clear Селектор
                "",       # keep Характеристика
            ]):
                p._find_and_edit_product()

            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(ws.cell(row=2, column=2).value, "https://expo-torg.ru/found-and-fixed/")
            self.assertIn(ws.cell(row=2, column=4).value, (None, ''))
            self.assertEqual(ws.cell(row=2, column=5).value, 'char')

    def test_find_and_edit_product_no_match_does_not_crash(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = self._make_two_row_workbook(tmpdir)
            p = pps.PriceParserWithSheets(xlsx_path, sheet_name="Прайс-лист")
            p.load_excel_data()
            with patch('builtins.input', side_effect=["совершенно левый запрос без совпадений"]):
                p._find_and_edit_product()  # must not raise

            # Nothing changed
            ws = load_workbook(xlsx_path)["Прайс-лист"]
            self.assertEqual(ws.cell(row=2, column=2).value, "https://expo-torg.ru/old/")


class GetWithSeleniumScrollTests(unittest.TestCase):
    """
    Regression tests for the reduced-depth scroll in get_with_selenium()'s
    bestly.ru branch: a real report showed the parser sometimes picking up
    a DIFFERENT product's price after scrolling deep into the page (e.g. a
    "similar products" section further down). Scrolling is now capped at
    ~50% of the page height, and stops as soon as price-like elements
    appear instead of being forced through a fixed minimum number of
    scrolls first (previously i >= 3, i.e. at least 4 scrolls).
    """

    def _make_fake_driver(self, appear_after_calls, scroll_height=10000):
        from selenium.common.exceptions import NoSuchElementException

        class FakeDriver:
            def __init__(self):
                self.calls = 0
                self.scroll_positions = []

            def get(self, url):
                pass

            def set_page_load_timeout(self, t):
                pass

            def delete_all_cookies(self):
                pass

            def execute_script(self, script):
                if 'scrollHeight' in script:
                    return scroll_height
                if 'scrollTo' in script:
                    m = re.search(r'scrollTo\(0,\s*([\d.]+)\)', script)
                    if m:
                        self.scroll_positions.append(float(m.group(1)))
                    return None
                return None

            def find_element(self, by, selector):
                # Используется WebDriverWait/EC.presence_of_element_located
                # для начальной проверки — всегда "не найдено", чтобы тест
                # обязательно прошёл по ветке прокрутки.
                raise NoSuchElementException("not yet")

            def find_elements(self, by, selector):
                self.calls += 1
                return [object()] if self.calls >= appear_after_calls else []

            @property
            def page_source(self):
                return "<html>" + ("x" * 2000) + "</html>"

        return FakeDriver()

    def test_scroll_never_exceeds_half_of_page_height(self):
        p = make_parser()
        p.driver = self._make_fake_driver(appear_after_calls=999, scroll_height=10000)
        with patch.object(pps.time, 'sleep', return_value=None):
            p.get_with_selenium("https://bestly.ru/catalog/x.html", wait_time=1)
        self.assertTrue(p.driver.scroll_positions, "expected at least one scroll")
        self.assertTrue(all(pos <= 5000 for pos in p.driver.scroll_positions), p.driver.scroll_positions)

    def test_scroll_stops_as_soon_as_prices_found(self):
        p = make_parser()
        # Цены "находятся" уже на первой прокрутке — раньше код всё равно
        # требовал минимум 4 прокрутки (i >= 3), прежде чем остановиться.
        p.driver = self._make_fake_driver(appear_after_calls=1, scroll_height=10000)
        with patch.object(pps.time, 'sleep', return_value=None):
            p.get_with_selenium("https://bestly.ru/catalog/x.html", wait_time=1)
        self.assertEqual(len(p.driver.scroll_positions), 1, p.driver.scroll_positions)


class MskStandartProSiteTests(unittest.TestCase):
    """
    Support for msk.standart.pro: it shares the exact same price markup
    convention as expo-torg.ru (.product-item-detail-price-current), so
    registering it in site_configs (method 'requests', browser headers)
    is enough — no site-specific parsing code is needed. Verifies both
    the manually-set-selector path and the generic auto-detect fallback
    against the real markup the user provided.
    """

    REAL_PRICE_HTML = """
    <html><body>
    <div class="catalog-item__price">
        <div class="product-item-detail-price-current" id="bx_117848907_14005_price">
            <span>5 095</span> руб.
        </div>
    </div>
    </body></html>
    """
    PRODUCT_URL = (
        "https://msk.standart.pro/catalog/laminirovannaya_plita/ldsp_egger/"
        "ldsp_egger_2_80_2_07_16_mm_u968_st9_seryy_ugol/"
    )

    def test_domain_recognized_and_configured_for_requests(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП U968 ST9 16мм Серый угол Egger", self.PRODUCT_URL, selector="",
            )
        self.assertEqual(p.extract_domain(self.PRODUCT_URL), "msk.standart.pro")
        config = p.site_configs.get("msk.standart.pro")
        self.assertIsNotNone(config)
        self.assertEqual(config.get('method'), 'requests')

    def test_price_found_via_explicit_selector(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП U968 ST9 16мм Серый угол Egger", self.PRODUCT_URL,
                selector=".product-item-detail-price-current",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, self.REAL_PRICE_HTML)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5095.0)
        self.assertIn('указанный селектор', result['status'])

    def test_price_found_without_any_selector_via_known_site_selector(self):
        """
        With price_selectors configured for the domain, a product with no
        selector set at all should go through the known-selector path
        (site_specific_with_name), not fall through to the much less
        reliable generic auto-detect scan.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП U968 ST9 16мм Серый угол Egger", self.PRODUCT_URL, selector="",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, self.REAL_PRICE_HTML)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5095.0)
        self.assertIn('известный селектор сайта', result['status'])

    def test_known_selector_wins_over_decoy_number_elsewhere_on_page(self):
        """
        Regression test for the real bug reported: a page also showing an
        unrelated number (e.g. a per-sheet quantity, "135") alongside the
        real price used to have the generic auto-detect fallback pick up
        that decoy instead of the real price from
        .product-item-detail-price-current. Now that selector is tried
        first (as a configured site selector) and wins.
        """
        page_with_decoy_number = """
        <html><body>
        <div class="catalog-detail-add-to-cart d-flex">
            <div class="catalog-item__price">
                <div class="product-item-detail-price-current" id="bx_117848907_14005_price">
                    <span>5 095</span> руб.
                </div>
            </div>
            <div class="product-quantity d-flex">
                <span>135</span>
            </div>
            <small class="mt-1">Цена за 1 лист</small>
        </div>
        </body></html>
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "ЛДСП 8 мм Бетон Чикаго светло-серый", self.PRODUCT_URL, selector="",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, page_with_decoy_number)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 5095.0)
        self.assertNotEqual(result['price'], 135.0)
        self.assertEqual(result['best_found_selector'], '.product-item-detail-price-current')


class SiteSpecificSelectorsGeneralizedTests(unittest.TestCase):
    """
    price_selectors used to only be consulted for bestly.ru (hardcoded),
    even though other sites (expo-torg.ru) had known-good selectors
    configured too. Both call sites — the standard-search fallback in
    parse_single_product() and find_price_and_name_on_page() — now look
    up site_configs for whatever domain is actually being parsed.
    """

    def test_expo_torg_price_selectors_configured(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(tmpdir, "Товар", "https://expo-torg.ru/x/", selector="")
        selectors = p.site_configs.get('expo-torg.ru', {}).get('price_selectors', [])
        self.assertIn('.product-item-detail-price-current', selectors)

    def test_find_price_and_name_on_page_uses_domain_specific_selectors(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(tmpdir, "Товар", "https://expo-torg.ru/x/", selector="")
        html = """
        <html><body>
        <div class="product-item-detail-price-current">2 940 руб.</div>
        <div class="unrelated">100500</div>
        </body></html>
        """
        prices_found, _ = p.find_price_and_name_on_page(
            html, "https://expo-torg.ru/catalog/plity/ldsp/some-product/", selector=None, product_name="ЛДСП тест",
        )
        self.assertTrue(prices_found)
        self.assertEqual(prices_found[0]['price'], 2940.0)
        self.assertEqual(prices_found[0]['method'], 'site_specific_with_name')


class LedpremiumSiteTests(unittest.TestCase):
    """
    Support for ledpremium.ru: its price element carries Schema.org
    microdata (itemprop="price"), the real markup the user provided:
        <span class="item_price" itemprop="price" content="722.85">722.85</span>
    Registered with both [itemprop="price"] (preferred — semantic markup
    tends to survive redesigns better than a CSS class) and .item_price
    as a fallback, method 'requests' like the other simple/static sites.
    """

    PRODUCT_URL = (
        "https://ledpremium.ru/catalog/svetodiodnye_lenty_cob/"
        "svetodiodnaya_lenta_sirius_cob_lp480_lh_24v_cri_85_cob_10_w_m_100_lm_w_ip33_/"
    )
    REAL_PRICE_HTML = """
    <html><body>
    <div class="product-price">
        <span class="item_price" itemprop="price" content="722.85">722.85</span>
    </div>
    </body></html>
    """

    def test_domain_recognized_and_configured(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(tmpdir, "Товар", self.PRODUCT_URL, selector="")
        self.assertEqual(p.extract_domain(self.PRODUCT_URL), "ledpremium.ru")
        config = p.site_configs.get("ledpremium.ru")
        self.assertIsNotNone(config)
        self.assertEqual(config.get('method'), 'requests')
        self.assertIn('[itemprop="price"]', config.get('price_selectors', []))

    def test_price_found_without_any_selector(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Светодиодная лента SIRIUS COB LP480 LH 24V", self.PRODUCT_URL, selector="",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, self.REAL_PRICE_HTML)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 722.85)
        self.assertEqual(result['best_found_selector'], '[itemprop="price"]')

    def test_price_found_via_explicit_selector(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Светодиодная лента SIRIUS COB LP480 LH 24V", self.PRODUCT_URL,
                selector='[itemprop="price"]',
            )
            fake_response = make_fake_response(self.PRODUCT_URL, self.REAL_PRICE_HTML)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 722.85)
        self.assertIn('указанный селектор', result['status'])


class UnregisteredSiteHeadersTests(unittest.TestCase):
    """
    Regression test for a real bug found while adding ros-met.com: every
    call site building headers for get_with_requests() did
    site_configs.get(domain, {}).get('headers', {}) — for a domain with
    no site_configs entry at all (or no 'headers' key) this evaluated to
    {} (an empty dict), not None. get_with_requests() only substitutes
    its sensible default browser headers when headers is None, so an
    empty dict was passed straight to requests.get() as-is: no
    User-Agent, no Accept-* headers at all. Many sites' bot/anti-scraping
    protection rejects such bare requests outright — this matches a real
    report where every ros-met.com product failed to load. Fixed by
    dropping the {} default so an unconfigured site's headers resolve to
    None and get the same sensible defaults as any other request.
    """

    def test_unconfigured_domain_gets_default_browser_headers(self):
        captured = {}

        def fake_get(url, headers=None, **kwargs):
            captured['headers'] = headers
            resp = MagicMock()
            resp.status_code = 200
            resp.url = url
            resp.encoding = 'utf-8'
            resp.text = "<html><body><span class='price'>1234 руб.</span></body></html>"
            return resp

        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Товар с незарегистрированного сайта",
                "https://some-unregistered-site.example/product/", selector="",
            )
            with patch.object(pps.requests, 'get', side_effect=fake_get):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertIsNotNone(captured['headers'])
        self.assertIn('User-Agent', captured['headers'])
        self.assertEqual(result['price'], 1234.0)


class RosMetSiteTests(unittest.TestCase):
    """
    Support for ros-met.com (metal/aluminum pipes, WooCommerce-based).
    Real markup provided:
        <span class="woocommerce-Price-amount amount">
            <bdi>68&nbsp;<span class="woocommerce-Price-currencySymbol">₽</span></bdi>
        </span>

    Also a regression test for a second real bug found via this exact
    price: 68 руб. was being rejected by is_reasonable_price()'s
    DEFAULT_PRICE_RANGE (100-50000), calibrated for sheet materials
    priced in the thousands — pipes/metal profile are commonly priced
    per running metre at tens of rubles. Added a PRODUCT_PRICE_RANGES
    entry for "труба" with a much lower floor (10) so genuine low prices
    like this aren't discarded as implausible, while a real decoy value
    like a quantity stepper's "1" is still rejected.
    """

    PRODUCT_URL = "https://ros-met.com/truba-20x20h2/"
    REAL_PRICE_HTML = """
    <html><body>
    <p class="price">
    <span class="woocommerce-Price-amount amount"><bdi>68&nbsp;<span class="woocommerce-Price-currencySymbol">₽</span></bdi></span>
    </p>
    </body></html>
    """

    def test_domain_recognized_and_configured(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(tmpdir, "Товар", self.PRODUCT_URL, selector="")
        self.assertEqual(p.extract_domain(self.PRODUCT_URL), "ros-met.com")
        config = p.site_configs.get("ros-met.com")
        self.assertIsNotNone(config)
        self.assertEqual(config.get('method'), 'requests')
        self.assertIn('.woocommerce-Price-amount', config.get('price_selectors', []))

    def test_low_pipe_price_not_rejected_as_unreasonable(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Труба профильная 20х20х2", self.PRODUCT_URL, selector="",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, self.REAL_PRICE_HTML)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 68.0)
        self.assertEqual(result['best_found_selector'], '.woocommerce-Price-amount')

    def test_decoy_quantity_still_rejected_alongside_low_real_price(self):
        html_with_decoy = self.REAL_PRICE_HTML.replace(
            '</body>', '<div class="qty">1</div></body>',
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Труба профильная 20х20х2", self.PRODUCT_URL, selector="",
            )
            fake_response = make_fake_response(self.PRODUCT_URL, html_with_decoy)
            with patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 68.0)
        self.assertNotEqual(result['price'], 1.0)

    def test_is_reasonable_price_range_for_truba(self):
        p = make_parser()
        self.assertTrue(p.is_reasonable_price(68, "Труба профильная 20х20х2"))
        self.assertFalse(p.is_reasonable_price(1, "Труба профильная 20х20х2"))


class SelectorMissingDotAndScriptTagRegressionTests(unittest.TestCase):
    """
    Regression tests for two real bugs found from a full parsing log on a
    bestly.ru composite-panel product:

    1. A saved selector of "table__price-current" (missing its leading
       '.') matches nothing via soup.select() — without a CSS prefix,
       that string is parsed as an (nonexistent) HTML tag name, not a
       class. The code silently treated this as "no elements" and fell
       all the way through to the generic auto-detect fallback instead
       of the real price element.
    2. That auto-detect fallback's text-node scan
       (soup.find_all(text=re.compile(...))) matched digits INSIDE
       <script> tag content too — BeautifulSoup treats script contents as
       an ordinary text node. The real log showed exactly this: a price
       of "4223" was "found" with method=auto_detected, selector=script,
       picked from a JS data blob rather than any visible page content.
    """

    def test_selector_missing_leading_dot_falls_back_to_class_match(self):
        p = make_parser()
        html = '<html><body><div class="table__price-current">4 223 руб.</div></body></html>'
        price = p.find_price_with_selector_and_name(html, 'table__price-current', 'Композитная панель GROSSBOND')
        self.assertEqual(price, 4223.0)

    def test_selector_that_is_a_real_tag_name_is_not_broken(self):
        """A selector that's genuinely a tag name (e.g. 'span') must still work as-is."""
        p = make_parser()
        html = '<html><body><span>999 руб.</span></body></html>'
        price = p.find_price_with_selector_and_name(html, 'span', None)
        self.assertEqual(price, 999.0)

    def test_script_tag_content_excluded_from_auto_detect(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(tmpdir, "x", "https://bestly.ru/catalog/x.html", selector="")
        html = """
        <html><body>
        <script>var relatedProducts = [{"id":1,"price":4223},{"id":2,"price":10800}];</script>
        <div class="something-unrelated">нет цены тут</div>
        </body></html>
        """
        prices_found, _ = p.find_price_and_name_on_page(
            html, "https://bestly.ru/catalog/x.html", selector=None, product_name="Композитная панель",
        )
        self.assertFalse(any(c['price'] in (4223.0, 10800.0) for c in prices_found), prices_found)

    def test_full_report_scenario_end_to_end(self):
        """
        Mirrors the real report: a <table> present but not matched by the
        structured table parser, decoy numbers embedded in <script> (as if
        from a "related products" JS data blob), and the real price only
        reachable via the saved selector once the missing dot is tolerated.
        """
        url = "https://bestly.ru/catalog/kompozitnaya_panel_grossbond_matovaya_g1.html"
        real_page_html = """
        <html><body>
        <script>var relatedProducts = [{"id":1,"price":4223},{"id":2,"price":10800}];</script>
        <div class="product">
            <div class="table__price-current">4750 руб.</div>
        </div>
        </body></html>
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            p = make_parser_for_product(
                tmpdir, "Композитная панель GROSSBOND матовая 4мм 1220x4000мм", url,
                selector="table__price-current",
            )
            fake_response = make_fake_response(url, real_page_html)
            with patch.object(pps.PriceParserWithSheets, 'get_with_selenium', return_value=None), \
                 patch.object(pps.requests, 'get', return_value=fake_response):
                result = p.parse_single_product(0, p.df.iloc[0])

        self.assertEqual(result['price'], 4750.0)
        self.assertNotIn(result['price'], (4223.0, 10800.0))


if __name__ == '__main__':
    unittest.main()
