"""Парсер цен: собирает товары с сайта и сохраняет их в CSV и Excel.

Под каждый сайт создаётся свой JSON-конфиг в папке configs/: адрес каталога
и селекторы (где на странице цена, ссылка, название). Код при этом не меняется.

Запуск:
    .\\.venv\\Scripts\\python.exe parse_prices.py
    .\\.venv\\Scripts\\python.exe parse_prices.py --config configs/books.json
    .\\.venv\\Scripts\\python.exe parse_prices.py --config configs/webscraper-laptops.json --out ноутбуки
    .\\.venv\\Scripts\\python.exe parse_prices.py --pages 5 --delay 0.5
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup, Tag
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

LOG = logging.getLogger("price-parser")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0 Safari/537.36 PriceParser/1.0"
)

DEFAULT_START_URL = "https://books.toscrape.com/"

# ---------------------------------------------------------------- настройки

# Селекторы по умолчанию (сайт books.toscrape.com).
# Для нового сайта создаётся свой JSON в configs/ и передаётся через --config.
DEFAULT_SELECTORS = {
    "card": "article.product_pod",      # карточка товара
    "title": "h3 a",                    # название
    "title_attr": "title",              # откуда брать название: атрибут или текст
    "price": "p.price_color",           # цена
    "link": "h3 a",                     # ссылка на товар
    "link_attr": "href",
    "rating": "p.star-rating",          # рейтинг (класс star-rating Three)
    "availability": "p.instock.availability",  # наличие
    "next_page": "li.next a",           # кнопка «следующая страница»
}

RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


def load_config(path: str | None) -> dict:
    """Читает JSON-конфиг сайта. Пустой путь = настройки по умолчанию."""
    if not path:
        return {}
    config_path = Path(path)
    if not config_path.exists():
        raise SystemExit(f"Файл настроек не найден: {config_path}")
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Ошибка в JSON {config_path}: {exc}")
    if not isinstance(data, dict):
        raise SystemExit(f"В {config_path} должен быть объект JSON")
    return data


def pick(node: Tag | BeautifulSoup, selector: str) -> Tag | None:
    """Ищет элемент по селектору. Пустая строка значит «такого поля нет на сайте»."""
    return node.select_one(selector) if selector else None

# ---------------------------------------------------------------- модель


@dataclass
class Product:
    """Одна строка будущей таблицы."""

    title: str
    price: float | None
    price_text: str
    rating: int | None
    availability: str
    url: str

    def as_row(self) -> dict:
        data = asdict(self)
        data.pop("price_text", None)
        return data


# ---------------------------------------------------------------- парсинг


def fetch(session: requests.Session, url: str, retries: int = 3) -> str:
    """Скачивает страницу с повторами при сбоях сети."""
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, timeout=20)
            response.raise_for_status()
            response.encoding = response.apparent_encoding or "utf-8"
            return response.text
        except requests.RequestException as exc:
            last_error = exc
            LOG.warning("Попытка %s/%s не удалась: %s", attempt, retries, exc)
            time.sleep(1.5 * attempt)
    raise RuntimeError(f"Не удалось скачать {url}: {last_error}")


def to_number(text: str) -> float | None:
    """'£51.77' -> 51.77, '1 234,50 руб.' -> 1234.5."""
    cleaned = re.sub(r"[^\d,.\-]", "", text or "").replace(",", ".")
    if cleaned.count(".") > 1:  # 1.234.567 -> 1234567
        cleaned = cleaned.replace(".", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def text_of(node: Tag | None) -> str:
    return " ".join(node.get_text(" ", strip=True).split()) if node else ""


def parse_page(html: str, base_url: str, selectors: dict) -> list[Product]:
    """Достаёт товары с одной страницы, используя селекторы сайта."""
    soup = BeautifulSoup(html, "html.parser")
    products: list[Product] = []

    card_selector = selectors.get("card") or ""
    if not card_selector:
        return products

    for card in soup.select(card_selector):
        title_node = pick(card, selectors.get("title", ""))
        price_node = pick(card, selectors.get("price", ""))
        link_node = pick(card, selectors.get("link", ""))
        rating_node = pick(card, selectors.get("rating", ""))
        stock_node = pick(card, selectors.get("availability", ""))

        title = ""
        if title_node:
            attr = selectors.get("title_attr") or ""
            if attr:
                title = title_node.get(attr, "")
            title = title or text_of(title_node)

        price_text = text_of(price_node)
        link_attr = selectors.get("link_attr") or "href"
        href = link_node.get(link_attr, "") if link_node else ""

        rating = None
        if rating_node:
            classes = rating_node.get("class", [])
            rating = next((RATING_WORDS[c] for c in classes if c in RATING_WORDS), None)

        if not title:
            continue

        products.append(
            Product(
                title=title,
                price=to_number(price_text),
                price_text=price_text,
                rating=rating,
                availability=text_of(stock_node),
                url=urljoin(base_url, href) if href else base_url,
            )
        )
    return products


def next_page_url(html: str, base_url: str, selectors: dict) -> str | None:
    soup = BeautifulSoup(html, "html.parser")
    node = pick(soup, selectors.get("next_page", ""))
    if not node:
        return None
    href = node.get("href")
    return urljoin(base_url, href) if href else None


# ---------------------------------------------------------------- сохранение

COLUMNS = [
    ("title", "Название", 50),
    ("price", "Цена", 12),
    ("rating", "Рейтинг", 10),
    ("availability", "Наличие", 18),
    ("url", "Ссылка", 60),
]


def save_csv(products: list[Product], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=[c[0] for c in COLUMNS], delimiter=";")
        writer.writerow({c[0]: c[1] for c in COLUMNS})
        for product in products:
            row = product.as_row()
            writer.writerow({key: row.get(key, "") for key in row})


def save_xlsx(products: list[Product], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Цены"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF")

    for col, (_, label, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = width

    for row_idx, product in enumerate(products, start=2):
        row = product.as_row()
        for col, (key, _, _) in enumerate(COLUMNS, start=1):
            value = row.get(key)
            cell = ws.cell(row=row_idx, column=col, value=value)
            if key == "price":
                cell.number_format = '#,##0.00'
            if key == "url" and value:
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")
            if key in {"rating", "availability", "price"}:
                cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(products) + 1}"
    wb.save(path)


def summarize(products: list[Product]) -> str:
    prices = [p.price for p in products if p.price is not None]
    if not prices:
        return "Цены не найдены"
    return (
        f"Товаров: {len(products)} | "
        f"цена мин/сред/макс: {min(prices):.2f} / "
        f"{sum(prices) / len(prices):.2f} / {max(prices):.2f}"
    )


# ---------------------------------------------------------------- CLI


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Собирает цены с сайта в CSV и Excel")
    parser.add_argument("--config", default=None, help="JSON-конфиг сайта из папки configs/")
    parser.add_argument("--start-url", default=None, help="страница, с которой начать (перекрывает конфиг)")
    parser.add_argument("--pages", type=int, default=3, help="сколько страниц обойти (по умолчанию 3)")
    parser.add_argument("--delay", type=float, default=1.0, help="пауза между запросами, сек")
    parser.add_argument("--out", default=None, help="имя файлов без расширения")
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    args = build_parser().parse_args(argv)

    config = load_config(args.config)
    selectors = {**DEFAULT_SELECTORS, **(config.get("selectors") or {})}
    start_url = args.start_url or config.get("start_url") or DEFAULT_START_URL

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out_name = args.out or f"prices_{stamp}"
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)

    if config.get("name"):
        LOG.info("Сайт: %s", config["name"])

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    products: list[Product] = []
    url: str | None = start_url
    page = 0
    started = time.perf_counter()

    while url and page < args.pages:
        page += 1
        LOG.info("Страница %s: %s", page, url)
        html = fetch(session, url)
        found = parse_page(html, url, selectors)
        LOG.info("  найдено товаров: %s", len(found))
        products.extend(found)
        url = next_page_url(html, url, selectors)
        if url and page < args.pages:
            time.sleep(max(args.delay, 0))

    if not products:
        LOG.error("Ничего не найдено. Проверь селекторы в конфиге и стартовый URL.")
        return 1

    csv_path = out_dir / f"{out_name}.csv"
    xlsx_path = out_dir / f"{out_name}.xlsx"
    save_csv(products, csv_path)
    save_xlsx(products, xlsx_path)

    elapsed = time.perf_counter() - started
    LOG.info("Готово за %.1f сек. %s", elapsed, summarize(products))
    LOG.info("CSV:   %s", csv_path.resolve())
    LOG.info("Excel: %s", xlsx_path.resolve())
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        LOG.warning("Остановлено пользователем")
        sys.exit(130)
