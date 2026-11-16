"""Парсер цен: собирает товары с сайта и сохраняет их в CSV и Excel.

По умолчанию работает на демо-сайте https://books.toscrape.com —
он создан специально для обучения парсингу и не блокирует запросы.
Под реальный сайт меняются только SELECTORS ниже и стартовый URL.

Запуск:
    .\\.venv\\Scripts\\python.exe parse_prices.py
    .\\.venv\\Scripts\\python.exe parse_prices.py --pages 5 --delay 0.5
    .\\.venv\\Scripts\\python.exe parse_prices.py --start-url https://example.com/catalog --out конкуренты
"""

from __future__ import annotations

import argparse
import csv
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

# Селекторы под конкретный сайт. Для нового сайта правится только этот блок:
# открыть страницу в браузере, ПКМ -> «Посмотреть код» и подставить классы.
SELECTORS = {
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


def parse_page(html: str, base_url: str) -> list[Product]:
    """Достаёт товары с одной страницы."""
    soup = BeautifulSoup(html, "html.parser")
    products: list[Product] = []

    for card in soup.select(SELECTORS["card"]):
        title_node = card.select_one(SELECTORS["title"])
        price_node = card.select_one(SELECTORS["price"])
        link_node = card.select_one(SELECTORS["link"])
        rating_node = card.select_one(SELECTORS["rating"])
        stock_node = card.select_one(SELECTORS["availability"])

        title = ""
        if title_node:
            if SELECTORS["title_attr"]:
                title = title_node.get(SELECTORS["title_attr"], "")
            title = title or text_of(title_node)

        price_text = text_of(price_node)
        href = link_node.get(SELECTORS["link_attr"], "") if link_node else ""

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


def next_page_url(html: str, base_url: str) -> str | None:
    soup = BeautifulSoup(html, "html.parser")
    node = soup.select_one(SELECTORS["next_page"])
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
    parser.add_argument("--start-url", default=DEFAULT_START_URL, help="страница, с которой начать")
    parser.add_argument("--pages", type=int, default=3, help="сколько страниц обойти (по умолчанию 3)")
    parser.add_argument("--delay", type=float, default=1.0, help="пауза между запросами, сек")
    parser.add_argument("--out", default=None, help="имя файлов без расширения")
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    args = build_parser().parse_args(argv)

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out_name = args.out or f"prices_{stamp}"
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    products: list[Product] = []
    url: str | None = args.start_url
    page = 0
    started = time.perf_counter()

    while url and page < args.pages:
        page += 1
        LOG.info("Страница %s: %s", page, url)
        html = fetch(session, url)
        found = parse_page(html, url)
        LOG.info("  найдено товаров: %s", len(found))
        products.extend(found)
        url = next_page_url(html, url)
        if url and page < args.pages:
            time.sleep(max(args.delay, 0))

    if not products:
        LOG.error("Ничего не найдено. Проверь SELECTORS и стартовый URL.")
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
