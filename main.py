from concurrent.futures import ThreadPoolExecutor, as_completed
import math
import time
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

import requests


class WildberriesParser:
    SEARCH_QUERY = "пальто из натуральной шерсти"
    DESTINATION = "-1257786"
    FULL_XLSX_PATH = "wb_catalog.xlsx"
    FILTERED_XLSX_PATH = "wb_catalog_filtered.xlsx"

    SEARCH_URL = "https://search.wb.ru/exactmatch/ru/common/v18/search"
    DETAIL_URL = "https://card.wb.ru/cards/v4/detail"
    CONTENT_HOSTS = (
        "rst-basket-cdn-11.geobasket.ru",
        "rst-basket-cdn-10.geobasket.ru",
        "rst-basket-cdn-12.geobasket.ru",
    )

    REQUEST_TIMEOUT = 30
    CONTENT_TIMEOUT = 8
    MAX_RETRIES = 4
    CONTENT_RETRIES = 2
    PROCESS_WORKERS = 6
    RETRY_DELAY = 1.5
    PAGE_DELAY = 0.35

    HEADERS = [
        "Ссылка на товар",
        "Артикул",
        "Название",
        "Цена",
        "Описание",
        "Ссылки на изображения",
        "Все характеристики",
        "Название селлера",
        "Ссылка на селлера",
        "Размеры товара",
        "Остатки по товару",
        "Рейтинг",
        "Количество отзывов",
    ]

    session = None

    @staticmethod
    def build_headers():
        return {
            "Accept": "*/*",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Origin": "https://www.wildberries.ru",
            "Referer": "https://www.wildberries.ru/",
            "Sec-CH-UA": '"Chromium";v="123", "Not:A-Brand";v="8"',
            "Sec-CH-UA-Mobile": "?0",
            "Sec-CH-UA-Platform": '"Linux"',
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-site",
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
            ),
        }

    @staticmethod
    def create_session():
        session = requests.Session()
        session.trust_env = False
        session.headers.update(WildberriesParser.build_headers())
        return session

    def prepare(self):
        if self.session is None:
            self.session = self.create_session()

    @staticmethod
    def request_json(
        session: requests.Session,
        url: str,
        *,
        params=None,
        timeout=0,
        retries=0,
        allow_not_found=False,
    ):
        last_error = None

        for attempt in range(1, retries + 1):
            try:
                response = session.get(url, params=params, timeout=timeout)
                if allow_not_found and response.status_code == 404:
                    return {}
                response.raise_for_status()
                return response.json()
            except Exception as error:
                last_error = error
                if attempt == retries:
                    break
                time.sleep(WildberriesParser.RETRY_DELAY * attempt)

        raise RuntimeError(f"Не удалось получить JSON: {url}") from last_error

    def fetch_all_search_products(self):
        products = []
        seen_nm_ids = set()
        page = 1

        while True:
            print(f"[info] Загружаю страницу поиска {page}")
            params = {
                "ab_testing": "false",
                "appType": "1",
                "curr": "rub",
                "dest": self.DESTINATION,
                "lang": "ru",
                "page": page,
                "query": self.SEARCH_QUERY,
                "resultset": "catalog",
                "sort": "popular",
                "spp": "30",
                "suppressSpellcheck": "false",
            }
            data = self.request_json(
                self.session,
                self.SEARCH_URL,
                params=params,
                timeout=self.REQUEST_TIMEOUT,
                retries=self.MAX_RETRIES,
            )
            page_products = data.get("products", [])
            if not page_products:
                break

            added = 0
            for product in page_products:
                nm_id = int(product["id"])
                if nm_id in seen_nm_ids:
                    continue
                seen_nm_ids.add(nm_id)
                products.append(product)
                added += 1

            if added == 0:
                break

            page += 1
            time.sleep(self.PAGE_DELAY)

        return products
    def fetch_content(self, nm_id):
        last_error = None
        volume = nm_id // 100000
        part = nm_id // 1000

        for host in self.CONTENT_HOSTS:
            url = f"https://{host}/vol{volume}/part{part}/{nm_id}/info/ru/card.json"
            try:
                return self.request_json(
                    self.session,
                    url,
                    timeout=self.CONTENT_TIMEOUT,
                    retries=self.CONTENT_RETRIES,
                    allow_not_found=True,
                )
            except Exception as error:
                last_error = error

        print(
            "[warn] Не удалось получать описание/характеристики через "
            f"geobasket card.json: {last_error}. Продолжаю без этих полей."
        )
        return {}

    @staticmethod
    def normalize_price(value):
        if not value:
            return 0.0
        return round(float(value) / 100, 2)

    @staticmethod
    def select_price(detail_product, search_product):
        for source in (detail_product, search_product):
            for size in source.get("sizes", []):
                price = WildberriesParser.normalize_price(
                    size.get("price", {}).get("product")
                )
                if price:
                    return price
        return 0.0

    @staticmethod
    def collect_sizes(product):
        names = []
        for size in product.get("sizes", []):
            value = str(size.get("origName") or size.get("name") or "").strip()
            if value and value not in names:
                names.append(value)
        return ", ".join(names)

    @staticmethod
    def collect_total_stock(product):
        total_quantity = product.get("totalQuantity")
        if isinstance(total_quantity, int):
            return total_quantity

        total = 0
        for size in product.get("sizes", []):
            for stock in size.get("stocks", []):
                qty = stock.get("qty")
                if isinstance(qty, int):
                    total += qty
        return total

    @staticmethod
    def extract_characteristics(content):
        grouped_options = content.get("grouped_options") or []
        if grouped_options:
            return grouped_options

        raw_options = content.get("options") or content.get("characteristics") or []
        if not raw_options:
            return []

        return [{"group_name": "Характеристики", "options": raw_options}]

    @staticmethod
    def format_characteristics(characteristics):
        if not characteristics:
            return ""

        lines = []
        for group in characteristics:
            group_name = str(group.get("group_name") or "Характеристики").strip()
            options = group.get("options", [])
            if group_name:
                lines.append(f"{group_name}:")

            if not isinstance(options, list):
                continue

            for item in options:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("name") or "").strip()
                value = str(item.get("value") or "").strip()
                if not name and not value:
                    continue
                if name and value:
                    lines.append(f"{name}: {value}")
                elif name:
                    lines.append(name)
                else:
                    lines.append(value)

            lines.append("")

        while lines and not lines[-1]:
            lines.pop()

        return "\n".join(lines)

    @staticmethod
    def basket_host(nm_id):
        volume = nm_id // 100000
        mapping = (
            (143, "01"),
            (287, "02"),
            (431, "03"),
            (719, "04"),
            (1007, "05"),
            (1061, "06"),
            (1115, "07"),
            (1169, "08"),
            (1313, "09"),
            (1601, "10"),
            (1655, "11"),
            (1919, "12"),
            (2045, "13"),
            (2189, "14"),
            (2405, "15"),
            (2621, "16"),
            (2837, "17"),
        )

        for upper_bound, bucket in mapping:
            if volume <= upper_bound:
                return f"basket-{bucket}.wbbasket.ru"
        return "basket-18.wbbasket.ru"

    @staticmethod
    def build_image_links(nm_id, pics_count):
        if pics_count <= 0:
            return ""

        host = WildberriesParser.basket_host(nm_id)
        volume = nm_id // 100000
        part = nm_id // 1000
        links = []

        for index in range(1, pics_count + 1):
            links.append(
                f"https://{host}/vol{volume}/part{part}/{nm_id}/images/big/{index}.webp"
            )

        return ", ".join(links)

    def build_row(self, search_product):
        nm_id = int(search_product["id"])
        params = {
            "appType": "1",
            "curr": "rub",
            "dest": self.DESTINATION,
            "nm": str(nm_id),
            "spp": "30",
        }
        try:
            data = self.request_json(
                self.session,
                self.DETAIL_URL,
                params=params,
                timeout=self.REQUEST_TIMEOUT,
                retries=self.MAX_RETRIES,
            )
            products = data.get("products", [])
            detail_product = products[0] if products else {}
        except Exception as error:
            print(f"[warn] Не удалось получить карточку {nm_id}: {error}")
            detail_product = {}

        content = self.fetch_content(nm_id)
        characteristics = self.extract_characteristics(content)

        rating = (
            detail_product.get("reviewRating")
            or search_product.get("reviewRating")
            or search_product.get("nmReviewRating")
            or 0
        )
        reviews = (
            detail_product.get("feedbacks")
            or search_product.get("feedbacks")
            or search_product.get("nmFeedbacks")
            or 0
        )
        supplier_id = detail_product.get("supplierId") or search_product.get("supplierId")
        pics_count = (
            content.get("media", {}).get("photo_count")
            or detail_product.get("pics")
            or search_product.get("pics")
            or search_product.get("picsCount")
            or 0
        )
        source_product = detail_product or search_product
        sizes = self.collect_sizes(source_product)
        if not sizes:
            names = []
            values = content.get("sizes_table", {}).get("values", [])
            if isinstance(values, list):
                for item in values:
                    if not isinstance(item, dict):
                        continue
                    value = str(item.get("tech_size") or "").strip()
                    if value and value not in names:
                        names.append(value)
            sizes = ", ".join(names)

        country = ""
        for group in characteristics:
            options = group.get("options", [])
            if not isinstance(options, list):
                continue
            for item in options:
                name = str(item.get("name", "")).strip().lower()
                value = str(item.get("value", "")).strip()
                if name == "страна производства" and value:
                    country = value
                    break
            if country:
                break

        return {
            "Ссылка на товар": f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx",
            "Артикул": nm_id,
            "Название": str(source_product.get("name") or "").strip(),
            "Цена": self.select_price(detail_product, search_product),
            "Описание": str(content.get("description") or "").strip(),
            "Ссылки на изображения": self.build_image_links(nm_id, int(pics_count)),
            "Все характеристики": self.format_characteristics(characteristics),
            "Название селлера": str(
                detail_product.get("supplier") or search_product.get("supplier") or ""
            ).strip(),
            "Ссылка на селлера": (
                f"https://www.wildberries.ru/seller/{int(supplier_id)}"
                if supplier_id
                else ""
            ),
            "Размеры товара": sizes,
            "Остатки по товару": self.collect_total_stock(source_product),
            "Рейтинг": float(rating or 0),
            "Количество отзывов": int(reviews or 0),
            "_country": country,
        }

    def filter_rows(self, rows):
        filtered = []

        for row in rows:
            country = str(row.get("_country", "")).strip().lower()
            rating = float(row.get("Рейтинг", 0) or 0)
            price = float(row.get("Цена", 0) or 0)

            if rating < 4.5:
                continue
            if price > 10000:
                continue
            if country != "россия":
                continue

            filtered.append(row)

        return filtered

    @staticmethod
    def column_letter(index):
        letters = []
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            letters.append(chr(65 + remainder))
        return "".join(reversed(letters))

    @staticmethod
    def excel_cell(value):
        if value is None:
            return "inlineStr", "<is><t></t></is>"

        if isinstance(value, bool):
            return "b", "1" if value else "0"

        if isinstance(value, int):
            return "n", str(value)

        if isinstance(value, float):
            if math.isnan(value) or math.isinf(value):
                return "inlineStr", "<is><t></t></is>"
            return "n", format(value, ".15g")

        text = str(value)
        preserve = (
            ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") else ""
        )
        return "inlineStr", f"<is><t{preserve}>{escape(text)}</t></is>"

    @staticmethod
    def build_sheet_xml(
        rows,
        headers,
        widths,
    ):
        xml_rows = []
        matrix = [headers]

        for row in rows:
            matrix.append([row.get(header, "") for header in headers])

        for row_index, row_values in enumerate(matrix, start=1):
            cells = []
            for column_index, value in enumerate(row_values, start=1):
                cell_ref = f"{WildberriesParser.column_letter(column_index)}{row_index}"
                cell_type, cell_value = WildberriesParser.excel_cell(value)
                if cell_type == "inlineStr":
                    cells.append(f'<c r="{cell_ref}" t="inlineStr">{cell_value}</c>')
                else:
                    cells.append(f'<c r="{cell_ref}" t="{cell_type}"><v>{cell_value}</v></c>')
            xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

        cols_xml = "".join(
            f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
            for index, width in enumerate(widths, start=1)
        )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<cols>{cols_xml}</cols>"
            f"<sheetData>{''.join(xml_rows)}</sheetData>"
            "</worksheet>"
        )

    def write_xlsx(self, path, rows):
        widths = []
        for header in self.HEADERS:
            max_length = len(header)
            for row in rows:
                max_length = max(max_length, len(str(row.get(header, ""))))
            widths.append(min(max(max_length + 2, 12), 90))

        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Каталог" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>"
        )
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            "</Relationships>"
        )
        workbook_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            "</Relationships>"
        )
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>"
        )
        content_types_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        )

        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types_xml)
            archive.writestr("_rels/.rels", rels_xml)
            archive.writestr("xl/workbook.xml", workbook_xml)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            archive.writestr("xl/styles.xml", styles_xml)
            archive.writestr(
                "xl/worksheets/sheet1.xml",
                WildberriesParser.build_sheet_xml(rows, self.HEADERS, widths),
            )

    def run(self):
        self.prepare()
        search_products = self.fetch_all_search_products()
        print(f"[info] Найдено товаров: {len(search_products)}")

        rows = [None] * len(search_products)
        with ThreadPoolExecutor(max_workers=self.PROCESS_WORKERS) as executor:
            futures = {
                executor.submit(self.build_row, search_product): (index, search_product)
                for index, search_product in enumerate(search_products)
            }

            completed = 0
            total = len(search_products)
            for future in as_completed(futures):
                index, search_product = futures[future]
                nm_id = int(search_product["id"])
                rows[index] = future.result()
                completed += 1
                print(f"[info] Обрабатываю {completed}/{total}: {nm_id}")

        filtered_rows = self.filter_rows(rows)
        self.write_xlsx(self.FULL_XLSX_PATH, rows)
        self.write_xlsx(self.FILTERED_XLSX_PATH, filtered_rows)

        print(f"[done] Полный каталог сохранен в {self.FULL_XLSX_PATH}")
        print(f"[done] Отфильтрованный каталог сохранен в {self.FILTERED_XLSX_PATH}")


if __name__ == "__main__":
    WildberriesParser().run()
