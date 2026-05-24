#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Wildberries Public Product Parser
Author: AI Assistant (Antigravity)
Description: A hybrid CLI-based parser for Wildberries.ru using storefront APIs.
             Supports fetching by SKU list or search keyword, proxy rotation,
             warehouse ID to name translation, and exporting to Excel/CSV.
"""

import sys
import io
import os
import time
import random
import argparse
import requests
import pandas as pd
from typing import List, Dict, Any, Optional

# Set terminal encoding to UTF-8 for safe Cyrillic display on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Import Rich for modern, beautiful terminal aesthetics
try:
    from rich.console import Console
    from rich.table import Table
    from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeRemainingColumn
    from rich.panel import Panel
    from rich.prompt import Prompt, Confirm
    from rich import box
except ImportError:
    # Fallback class if rich is not installed (though we will instruct the user to install it)
    class Console:
        def print(self, *args, **kwargs):
            print(*args, **kwargs)
        def rule(self, *args, **kwargs):
            print("-" * 50)
    Console = Console()

# Global Console Object
console = Console()

# Constant parameters
DEFAULT_DEST = "-2888067"  # Grodno destination ID for pricing/stocks
MOSCOW_DEST = "-1257786"  # Moscow destination ID — needed separately for stock data
DEFAULT_CURR = "byn"  # Currency (byn for Belarusian rubles, rub for Russian)
DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Origin': 'https://www.wildberries.ru',
    'Referer': 'https://www.wildberries.ru/'
}

# Cache for x_wbaas_token to avoid fetching it for every product
_WBAAS_TOKEN = None

# Global session for keep-alive
_SESSION = requests.Session()
_SESSION.headers.update(DEFAULT_HEADERS)

_VOL_BASKET_CACHE = {}
_BASKET_EXECUTOR = None


def get_basket_dynamically(vol: int, part: int, sku: int) -> str:
    """
    Dynamically discovers the correct basket for a given volume by checking the image URL concurrently.
    Caches the result to avoid redundant network calls.
    """
    if vol in _VOL_BASKET_CACHE:
        return _VOL_BASKET_CACHE[vol]
        
    import requests
    from concurrent.futures import ThreadPoolExecutor, as_completed
    global _BASKET_EXECUTOR
    if _BASKET_EXECUTOR is None:
        _BASKET_EXECUTOR = ThreadPoolExecutor(max_workers=15)
    
    # Static fallback
    if vol <= 143: guess = "01"
    elif vol <= 287: guess = "02"
    elif vol <= 431: guess = "03"
    elif vol <= 719: guess = "04"
    elif vol <= 1007: guess = "05"
    elif vol <= 1061: guess = "06"
    elif vol <= 1115: guess = "07"
    elif vol <= 1169: guess = "08"
    elif vol <= 1313: guess = "09"
    elif vol <= 1601: guess = "10"
    elif vol <= 1655: guess = "11"
    elif vol <= 1919: guess = "12"
    elif vol <= 2045: guess = "13"
    elif vol <= 2189: guess = "14"
    elif vol <= 2405: guess = "15"
    else: guess = f"{(16 + (vol - 2406) // 216):02d}"
    
    def check_image(b_id_str):
        url = f"https://basket-{b_id_str}.wbbasket.ru/vol{vol}/part{part}/{sku}/images/big/1.webp"
        try:
            if requests.head(url, timeout=1.5).status_code == 200:
                return b_id_str
        except:
            pass
        return None

    # Fast track: try guess first
    if check_image(guess):
        _VOL_BASKET_CACHE[vol] = guess
        return guess

    # Also fast track 39, 40, 41 since WB recently migrated huge volume ranges there
    for common_basket in ['39', '40', '41']:
        if common_basket != guess and check_image(common_basket):
            _VOL_BASKET_CACHE[vol] = common_basket
            return common_basket

    # Search the rest concurrently without blocking on timeouts
    baskets_to_check = [f"{i:02d}" for i in range(1, 201) if f"{i:02d}" not in [guess, '39', '40', '41']]
    
    futures = [_BASKET_EXECUTOR.submit(check_image, b) for b in baskets_to_check]
    for future in as_completed(futures):
        result = future.result()
        if result:
            _VOL_BASKET_CACHE[vol] = result
            return result

    _VOL_BASKET_CACHE[vol] = guess
    return guess

def make_request(url: str, params: Optional[Dict[str, Any]] = None, headers: Optional[Dict[str, Any]] = None, cookies: Optional[Dict[str, str]] = None, proxies: Optional[List[str]] = None, max_retries: int = 3, timeout: int = 15) -> requests.Response:
    """
    Makes an HTTP request with automatic handling of 429 (Too Many Requests)
    using exponential backoff and the Retry-After header.
    Uses global session for keep-alive connections.
    """
    delay = 3.0
    for attempt in range(max_retries):
        try:
            if proxies:
                selected = random.choice(proxies)
                proxy = {'http': selected, 'https': selected}
            else:
                proxy = None
            response = _SESSION.get(url, params=params, headers=headers, cookies=cookies, proxies=proxy, timeout=timeout)
            
            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if (retry_after and retry_after.isdigit()) else int(delay)
                console.print(f"[yellow]\n[!] Предупреждение (429): Слишком много запросов. Ожидание {wait_time} сек. перед повторной попыткой...[/yellow]")
                time.sleep(wait_time)
                delay *= 2
                continue
                
            response.raise_for_status()
            return response
        except requests.exceptions.RequestException as e:
            # Do not retry on 404 Not Found
            if isinstance(e, requests.exceptions.HTTPError) and e.response is not None and e.response.status_code == 404:
                raise e
            if attempt == max_retries - 1:
                raise e
            time.sleep(delay)
            delay *= 2
    raise requests.exceptions.HTTPError("Превышено количество попыток запроса из-за ограничений 429")


def fetch_warehouse_map(proxies: Optional[List[str]] = None) -> Dict[int, str]:
    """
    Fetches the static warehouse name mapping from Wildberries CDN to translate
    raw store IDs (int) to human-readable Cyrillic names.
    """
    url = "https://static-basket-01.wbbasket.ru/vol0/data/stores-data.json"
    
    try:
        response = make_request(url, headers=DEFAULT_HEADERS, proxies=proxies, timeout=10)
        response.encoding = 'utf-8'
        stores_list = response.json()
        # Map id to name
        return {store['id']: store['name'] for store in stores_list if 'id' in store and 'name' in store}
    except Exception as e:
        console.print(f"[yellow]Предупреждение: Не удалось загрузить карту складов ({e}). ID складов будут отображаться без имён.[/yellow]")
    return {}

def parse_single_product(product: Dict[str, Any], warehouse_map: Dict[int, str]) -> Dict[str, Any]:
    """
    Parses a single raw product card dictionary from WB JSON response and returns
    a flattened dictionary with required basic and advanced fields.
    """
    product_id = product.get('id')
    name = product.get('name', 'Неизвестно')
    brand = product.get('brand', 'Нет бренда')
    supplier = product.get('supplier', 'Неизвестный продавец')
    supplier_id = product.get('supplierId')
    rating = product.get('rating', 0)
    feedbacks = product.get('feedbacks', 0)
    
    # Prices (represented in minor units like kopecks, divide by 100)
    price_u = product.get('priceU')
    sale_price_u = product.get('salePriceU')
    
    # Try to extract from sizes if root fields are missing (common on card detail API v4)
    sizes = product.get('sizes', [])
    if (price_u is None or sale_price_u is None) and sizes:
        first_size = sizes[0]
        price_obj = first_size.get('price', {})
        if price_obj:
            price_u = price_obj.get('basic')
            sale_price_u = price_obj.get('product')
            
    price_original = price_u / 100.0 if price_u else 0.0
    price_discounted = sale_price_u / 100.0 if sale_price_u else 0.0
    
    # Sizes and stocks
    sizes_list = []
    stocks_detail = []
    total_stock = 0
    
    sizes = product.get('sizes', [])
    for size_obj in sizes:
        size_name = size_obj.get('origName') or size_obj.get('name') or 'No Size'
        sizes_list.append(size_name)
        
        stocks = size_obj.get('stocks', [])
        for stock_obj in stocks:
            wh_id = stock_obj.get('wh')
            qty = stock_obj.get('qty', 0)
            total_stock += qty
            wh_name = warehouse_map.get(wh_id, f"Склад ID {wh_id}")
            stocks_detail.append({
                'size': size_name,
                'wh_id': wh_id,
                'wh_name': wh_name,
                'qty': qty
            })
            
    sizes_str = ", ".join(map(str, sizes_list))
    
    # Use totalQuantity from API when available (it's the real total across all warehouses)
    total_quantity = product.get('totalQuantity')
    if total_quantity is not None and total_quantity > total_stock:
        total_stock = total_quantity
    
    # Aggregate stock by warehouse name
    wh_agg = {}
    for item in stocks_detail:
        name_wh = item['wh_name']
        wh_agg[name_wh] = wh_agg.get(name_wh, 0) + item['qty']
    
    wh_summary_str = ", ".join([f"{wh}: {qty}" for wh, qty in wh_agg.items()]) if wh_agg else "Нет в наличии"
    
    # Compute image URL
    image_url = ""
    if product_id:
        def get_basket_static(v: int) -> str:
            if v <= 143: return "01"
            elif v <= 287: return "02"
            elif v <= 431: return "03"
            elif v <= 719: return "04"
            elif v <= 1007: return "05"
            elif v <= 1061: return "06"
            elif v <= 1115: return "07"
            elif v <= 1169: return "08"
            elif v <= 1313: return "09"
            elif v <= 1601: return "10"
            elif v <= 1655: return "11"
            elif v <= 1919: return "12"
            elif v <= 2045: return "13"
            elif v <= 2189: return "14"
            elif v <= 2405: return "15"
            return f"{(16 + (v - 2406) // 216):02d}"
            
        vol = product_id // 100000
        part = product_id // 1000
        basket = get_basket_static(vol)
        image_url = f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{product_id}/images/big/1.webp"
        
    return {
        'Артикул': product_id,
        'Название': name,
        'Бренд': brand,
        'Продавец': supplier,
        'ID Продавца': supplier_id,
        'Цена без скидки': price_original,
        'Цена со скидкой': price_discounted,
        'Рейтинг': rating,
        'Отзывы': feedbacks,
        'Остатки (всего)': total_stock,
        'Склады детализация': wh_summary_str,
        'Ссылка на товар': f"https://www.wildberries.ru/catalog/{product_id}/detail.aspx",
        'Ссылка на изображение': image_url
    }

def fetch_products_by_skus(skus: List[int], warehouse_map: Dict[int, str], proxies: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    Fetches details for a list of product articles (SKUs) in batches of 100.
    """
    parsed_products = []
    url = "https://card.wb.ru/cards/v4/detail"
    
    # WB API allows up to 100 SKUs per request
    batch_size = 100
    batches = [skus[i:i + batch_size] for i in range(0, len(skus), batch_size)]
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        console=console
    ) as progress:
        task = progress.add_task("[cyan]Сбор данных по артикулам...", total=len(batches))
        
        for batch in batches:
            nm_param = ";".join(map(str, batch))
            params = {
                'appType': '1',
                'curr': DEFAULT_CURR,
                'dest': DEFAULT_DEST,
                'spp': '30',
                'nm': nm_param
            }
            
            try:
                # Add slight random delay to mimic user behavior and avoid blockages
                time.sleep(random.uniform(0.5, 1.5))
                
                response = make_request(url, params=params, headers=DEFAULT_HEADERS, proxies=proxies, timeout=15)
                data = response.json()
                
                products_list = data.get('products') or data.get('data', {}).get('products', [])
                
                # Make a second request with Moscow region to fetch stock data,
                # because WB omits warehouse stocks for non-Moscow regions.
                msk_prods_by_id = {}
                try:
                    time.sleep(random.uniform(0.3, 0.8))
                    msk_params = {
                        'appType': '1',
                        'curr': 'rub',
                        'dest': MOSCOW_DEST,
                        'spp': '30',
                        'nm': nm_param
                    }
                    msk_response = make_request(url, params=msk_params, headers=DEFAULT_HEADERS, proxies=proxies, timeout=15)
                    msk_data = msk_response.json()
                    msk_products = msk_data.get('products') or msk_data.get('data', {}).get('products', [])
                    for msk_prod in msk_products:
                        pid = msk_prod.get('id')
                        if pid:
                            msk_prods_by_id[pid] = msk_prod
                except Exception as msk_err:
                    console.print(f"[yellow]Предупреждение: Московский запрос для складов не удался ({msk_err}). Продолжаем без остатков.[/yellow]")
                
                for prod in products_list:
                    prod_id = prod.get('id')
                    if prod_id and prod_id in msk_prods_by_id:
                        msk_sizes = msk_prods_by_id[prod_id].get('sizes', [])
                        gro_sizes = prod.get('sizes', [])
                        if msk_sizes:
                            if gro_sizes:
                                for gi in range(min(len(gro_sizes), len(msk_sizes))):
                                    msk_stocks = msk_sizes[gi].get('stocks', [])
                                    if msk_stocks:
                                        gro_sizes[gi]['stocks'] = msk_stocks
                            else:
                                prod['sizes'] = msk_sizes
                    
                    parsed_prod = parse_single_product(prod, warehouse_map)
                    parsed_products.append(parsed_prod)
                    
            except Exception as e:
                console.print(f"[red]Ошибка при обработке пакета артикулов: {e}[/red]")
            
            progress.advance(task)
            
    return parsed_products

def get_wbaas_token(proxies: Optional[List[str]] = None) -> Optional[str]:
    """
    Attempts to obtain the `x_wbaas_token` cookie required for the internal search API.
    First tries SeleniumBase's undetected‑chromedriver. If the driver executable is missing
    or cannot be executed (PermissionError), falls back to a plain HTTP request and extracts
    the token from the page source using a regular expression.
    """
    
    # Path to the driver used by SeleniumBase (undetected‑chromedriver)
    driver_path = os.path.join(os.path.dirname(__file__), '.venv', 'Lib', 'site-packages', 'seleniumbase', 'drivers', 'uc_driver.exe')
    
    # Helper: simple HTTP fallback
    def _fallback_http() -> Optional[str]:
        try:
            response = make_request(
                "https://www.wildberries.ru/",
                headers=DEFAULT_HEADERS,
                proxies=proxies,
                timeout=15
            )
            import re
            # Look for the token in JavaScript variable assignment
            match = re.search(r"x_wbaas_token\s*=\s*'([^']+)'", response.text)
            if match:
                token = match.group(1)
                console.print("[green]Сессионный токен получен через HTTP‑запрос.[/green]")
                return token
        except Exception as e_http:
            console.print(f"[yellow]  Предупреждение: Не удалось получить токен через HTTP ({e_http})[/yellow]")
        return None
    
    # Try SeleniumBase only if the driver file exists and is executable
    try:
        from seleniumbase import Driver
    except Exception as e_import:
        console.print(f"[yellow]  Предупреждение: SeleniumBase недоступен ({e_import}). Переходим к HTTP‑fallback.[/yellow]")
        return _fallback_http()
    
    # Check driver executable permissions
    if not os.path.isfile(driver_path):
        console.print("[yellow]  Предупреждение: uc_driver.exe не найден, используем HTTP‑fallback.[/yellow]")
        return _fallback_http()
    
    # Ensure the file is executable for the current user
    try:
        # Attempt to add execute permission if missing
        import subprocess
        subprocess.run(["icacls", driver_path, "/grant", f"%USERNAME%:RX"], capture_output=True, text=True, check=False)
    except Exception:
        pass
    
    try:
        proxy_str = None
        if proxies:
            raw_proxy = random.choice(proxies)
            proxy_str = raw_proxy.replace("http://", "").replace("https://", "")
        console.print("[cyan]Получение токена через SeleniumBase...[/cyan]")
        driver = Driver(
            uc=True,
            headed=False,
            headless=True,
            agent=DEFAULT_HEADERS['User-Agent'],
            proxy=proxy_str
        )
        try:
            driver.open("https://www.wildberries.ru/")
            for _ in range(15):
                time.sleep(1.0)
                cookies = driver.execute_cdp_cmd("Network.getAllCookies", {})
                for cookie in cookies.get("cookies", []):
                    if cookie.get("name") == "x_wbaas_token":
                        token = cookie.get("value")
                        console.print("[green]Сессионный токен получен через браузер.[/green]")
                        return token
            console.print("[yellow]  Предупреждение: Токен не найден в браузере, пробуем HTTP‑fallback.[/yellow]")
        finally:
            driver.quit()
    except PermissionError as perm_err:
        console.print(f"[yellow]  Предупреждение: Недостаточно прав для uc_driver.exe ({perm_err}). Переходим к HTTP‑fallback.[/yellow]")
    except Exception as e_selenium:
        console.print(f"[yellow]  Предупреждение: Ошибка SeleniumBase ({e_selenium}). Переходим к HTTP‑fallback.[/yellow]")
    
    # If we reach here, fallback to HTTP method
    return _fallback_http()


def search_products_by_query(query: str, limit_pages: int, warehouse_map: Dict[int, str], proxies: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    Searches products by keyword search query up to a limit of pages using internal search API.
    Then enriches them with full warehouse details via detailed SKU cards API.
    """
    token = get_wbaas_token(proxies)
    
    # Internal search API URL
    url = "https://www.wildberries.ru/__internal/u-search/exactmatch/ru/common/v18/search"
    
    headers = DEFAULT_HEADERS.copy()
    headers.update({
        'Accept': '*/*',
        'Accept-Language': 'ru-RU,ru;q=0.9',
        'X-Requested-With': 'XMLHttpRequest',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
    })
    
    import urllib.parse
    encoded_query = urllib.parse.quote(query)
    headers['Referer'] = f"https://www.wildberries.ru/catalog/0/search.aspx?search={encoded_query}"
    
    cookies = {}
    if token:
        cookies['x_wbaas_token'] = token
        
    skus_to_fetch = []
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        console=console
    ) as progress:
        task = progress.add_task(f"[cyan]Поиск артикулов по запросу '{query}'...", total=limit_pages)
        
        for page in range(1, limit_pages + 1):
            params = {
                'ab_testid': 'promo_mask_test_1',
                'appType': '1',
                'autoselectFilters': 'false',
                'curr': DEFAULT_CURR,
                'dest': DEFAULT_DEST,
                'hide_dtype': '9',
                'hide_vflags': '4294967296',
                'inheritFilters': 'false',
                'lang': 'ru',
                'query': query,
                'resultset': 'catalog',
                'spp': '30',
                'suppressSpellcheck': 'false',
                'page': str(page)
            }
            
            try:
                time.sleep(random.uniform(0.5, 1.5))
                
                response = make_request(url, params=params, headers=headers, cookies=cookies, proxies=proxies, timeout=15)
                data = response.json()
                
                products_list = data.get('products') or data.get('data', {}).get('products', [])
                if not products_list:
                    console.print(f"[yellow]Страница {page}: Товары закончились или не найдены.[/yellow]")
                    progress.advance(task, advance=limit_pages - page + 1)
                    break
                    
                for prod in products_list:
                    prod_id = prod.get('id')
                    if prod_id:
                        skus_to_fetch.append(prod_id)
                        
            except Exception as e:
                console.print(f"[red]Ошибка при парсинге страницы {page}: {e}[/red]")
            
            progress.advance(task)
            
    if not skus_to_fetch:
        console.print("[yellow]Не найдено ни одного артикула по запросу.[/yellow]")
        return []
        
    console.print(f"[cyan]Найдено артикулов по поиску: {len(skus_to_fetch)}. Загружаем полную детальную информацию о товарах (склады, цены, бренды)...[/cyan]")
    
    return fetch_products_by_skus(skus_to_fetch, warehouse_map, proxies)

def load_proxies(filepath: str) -> List[str]:
    """
    Loads list of proxy servers from a file.
    Expected formats (one per line):
    - ip:port
    - http://ip:port
    - http://user:password@ip:port
    """
    proxies = []
    if not os.path.exists(filepath):
        console.print(f"[red]Файл с прокси {filepath} не найден.[/red]")
        return proxies
        
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            # Ensure protocol prefix
            if not line.startswith('http://') and not line.startswith('https://'):
                # Default to http proxy
                proxies.append(f"http://{line}")
            else:
                proxies.append(line)
                
    console.print(f"[green]Успешно загружено {len(proxies)} прокси.[/green]")
    return proxies

def display_results_in_table(products: List[Dict[str, Any]], limit: int = 10) -> None:
    """
    Displays the top parsed products in a beautiful terminal table.
    """
    if not products:
        console.print("[yellow]Нет данных для отображения.[/yellow]")
        return
        
    table = Table(
        title=f"Результаты парсинга (Показано первые {min(limit, len(products))} из {len(products)} товаров)",
        box=box.DOUBLE_EDGE,
        header_style="bold magenta",
        title_style="bold cyan"
    )
    
    table.add_column("Артикул", style="dim", width=12)
    table.add_column("Бренд", style="green", width=15)
    table.add_column("Название", style="white", max_width=30, overflow="ellipsis")
    table.add_column("Цена (без/со ск.)", justify="right", style="cyan")
    table.add_column("Рейт./Отзывы", justify="center")
    table.add_column("Всего в наличии", justify="right", style="yellow")
    table.add_column("Продавец", style="magenta", max_width=20, overflow="ellipsis")

    for prod in products[:limit]:
        price_str = f"{prod.get('Цена без скидки', 0):.2f} / {prod.get('Цена со скидкой', 0):.2f} BYN"
        rating_str = f"★{prod.get('Рейтинг', 0)} ({prod.get('Отзывы', 0)})"
        table.add_row(
            str(prod.get('Артикул', '')),
            str(prod.get('Бренд', '')),
            str(prod.get('Название', '')),
            price_str,
            rating_str,
            f"{prod.get('Остатки (всего)', 0)} шт",
            str(prod.get('Продавец', ''))
        )

    console.print(table)


def fetch_product_details(nm_id: int, proxies: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Fetches product description and characteristics directly from wbbasket.ru CDN.
    Falls back to v5 endpoint if CDN fails.
    """
    vol = nm_id // 100000
    part = nm_id // 1000
    basket = get_basket_dynamically(vol, part, nm_id)
    
    url = f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/info/ru/card.json"
    
    # 1. Try CDN first (fastest and most reliable for static details)
    try:
        time.sleep(random.uniform(0.1, 0.3))
        response = make_request(url, proxies=proxies, timeout=15)
        if response.status_code == 200:
            product = response.json()
            description = product.get('description', '') or ''
            raw_chars = product.get('options', []) or []
            characteristics = []
            
            for item in raw_chars:
                name = item.get('name', '')
                value = item.get('value', '')
                if name:
                    characteristics.append({'Характеристика': name, 'Значение': str(value)})
                    
            for group in product.get('grouped_options', []) or []:
                group_name = group.get('name', '')
                for opt in group.get('options', []):
                    opt_name = opt.get('name', '')
                    opt_value = opt.get('value', '')
                    char_name = f"{group_name} / {opt_name}" if group_name else opt_name
                    if opt_name:
                        characteristics.append({'Характеристика': char_name, 'Значение': str(opt_value)})
                        
            return {'description': description, 'characteristics': characteristics}
    except Exception:
        pass

    # 2. Fallback to older v5/v4 endpoints if CDN fails
    global _WBAAS_TOKEN
    if not _WBAAS_TOKEN:
        _WBAAS_TOKEN = get_wbaas_token(proxies)
    token_cookies = {'x_wbaas_token': _WBAAS_TOKEN} if _WBAAS_TOKEN else {}
    
    params = {
        'appType': '1',
        'curr': DEFAULT_CURR,
        'dest': DEFAULT_DEST,
        'spp': '30',
        'nm': str(nm_id)
    }
    
    for fallback_url in ["https://card.wb.ru/cards/v5/detail", "https://card.wb.ru/cards/v4/detail"]:
        try:
            time.sleep(random.uniform(0.3, 0.8))
            response = make_request(fallback_url, params=params, headers=DEFAULT_HEADERS, proxies=proxies, timeout=15, cookies=token_cookies)
            if response.status_code == 200:
                data = response.json()
                products_list = data.get('data', {}).get('products', []) or data.get('products', [])
                if products_list:
                    product = products_list[0]
                    description = product.get('description', '') or ''
                    raw_chars = product.get('options', []) or product.get('characteristics', []) or []
                    characteristics = []
                    for item in raw_chars:
                        name = item.get('name', '')
                        value = item.get('value', '')
                        if name:
                            characteristics.append({'Характеристика': name, 'Значение': str(value)})
                    for group in product.get('grouped_options', []) or []:
                        group_name = group.get('name', '')
                        for opt in group.get('options', []):
                            opt_name = opt.get('name', '')
                            opt_value = opt.get('value', '')
                            char_name = f"{group_name} / {opt_name}" if group_name else opt_name
                            if opt_name:
                                characteristics.append({'Характеристика': char_name, 'Значение': str(opt_value)})
                    return {'description': description, 'characteristics': characteristics}
        except Exception:
            continue

    # If everything fails, return empty
    console.print(f"[yellow]  Предупреждение: Не удалось получить детали товара {nm_id} (все попытки провалились).[/yellow]")
    return {'description': '', 'characteristics': []}




def fetch_all_details_parallel(products: List[Dict[str, Any]], proxies: Optional[List[str]] = None, max_workers: int = 10) -> Dict[int, Dict]:
    """Fetches product details (description, characteristics) in parallel."""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    details_map = {}
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        console=console
    ) as progress:
        task = progress.add_task("[cyan]Загрузка характеристик...", total=len(products))

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(fetch_product_details, p.get('Артикул'), proxies): p
                for p in products if p.get('Артикул')
            }
            for future in as_completed(futures):
                p = futures[future]
                nm_id = p.get('Артикул')
                try:
                    details_map[nm_id] = future.result()
                except Exception:
                    details_map[nm_id] = {'description': '', 'characteristics': []}
                progress.advance(task)

    return details_map


def export_data(
    products: List[Dict[str, Any]],
    base_filename: str,
    include_details: bool = False,
    proxies: Optional[List[str]] = None
) -> None:
    """
    Exports the list of parsed products to Excel (.xlsx) and CSV formats.
    If include_details=True, fetches and exports description + characteristics
    for each product into separate sheets (named 1, 2, 3...).
    """
    if not products:
        console.print("[yellow]Экспорт отменен: нет данных для сохранения.[/yellow]")
        return

    df = pd.DataFrame(products)

    excel_file = f"{base_filename}.xlsx"
    csv_file = f"{base_filename}.csv"

    # Pre-fetch details in parallel if needed
    details_map = {}
    if include_details:
        console.print(f"\n[cyan]Загрузка характеристик и описания для {len(products)} товаров...[/cyan]")
        details_map = fetch_all_details_parallel(products, proxies)

    try:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Товары', index=False)

            if include_details:
                for idx, prod in enumerate(products, start=1):
                    nm_id = prod.get('Артикул')
                    sheet_name = str(idx)

                    details = details_map.get(nm_id, {'description': '', 'characteristics': []})

                    rows = []
                    rows.append({'Поле': 'Артикул', 'Значение': str(nm_id or '')})
                    rows.append({'Поле': 'Название', 'Значение': str(prod.get('Название', ''))})
                    rows.append({'Поле': 'Бренд', 'Значение': str(prod.get('Бренд', ''))})
                    rows.append({'Поле': 'Продавец', 'Значение': str(prod.get('Продавец', ''))})
                    rows.append({'Поле': 'Ссылка', 'Значение': str(prod.get('Ссылка на товар', ''))})
                    rows.append({'Поле': '', 'Значение': ''})
                    rows.append({'Поле': 'ОПИСАНИЕ', 'Значение': details['description']})
                    rows.append({'Поле': '', 'Значение': ''})

                    if details['characteristics']:
                        rows.append({'Поле': '--- ХАРАКТЕРИСТИКИ ---', 'Значение': ''})
                        for char in details['characteristics']:
                            rows.append({
                                'Поле': char.get('Характеристика', ''),
                                'Значение': char.get('Значение', '')
                            })
                    else:
                        rows.append({'Поле': 'Характеристики', 'Значение': 'Не найдены'})

                    sheet_df = pd.DataFrame(rows)
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    if nm_id:
                        vol = nm_id // 100000
                        part = nm_id // 1000
                        basket = get_basket_dynamically(vol, part, nm_id)
                        image_url = f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/images/big/1.webp"
                    else:
                        image_url = prod.get('Ссылка на изображение')

                    if image_url:
                        try:
                            import io
                            from openpyxl.drawing.image import Image as OpenpyxlImage

                            img_resp = _SESSION.get(image_url, timeout=5)
                            if img_resp.status_code == 200:
                                img_data = io.BytesIO(img_resp.content)
                                img = OpenpyxlImage(img_data)

                                max_height = 300
                                if img.height > max_height:
                                    ratio = max_height / img.height
                                    img.height = max_height
                                    img.width = int(img.width * ratio)

                                ws = writer.sheets[sheet_name]
                                insert_row = len(rows) + 3
                                ws.add_image(img, f'A{insert_row}')
                        except Exception:
                            pass

        console.print(f"[green]✔ Данные успешно экспортированы в Excel: {excel_file}[/green]")
        if include_details:
            console.print(f"[green]  Дополнительно: создано {len(products)} листов с характеристиками (листы 1–{len(products)}).[/green]")
    except Exception as e:
        console.print(f"[red]Критическая ошибка при экспорте: {e}[/red]")
        emergency_file = f"{base_filename}_emergency.csv"
        try:
            df.to_csv(emergency_file, index=False, encoding='utf-8-sig')
            console.print(f"[yellow]⚠ Аварийное сохранение: {emergency_file} ({len(products)} товаров)[/yellow]")
        except Exception:
            console.print("[red]✘ Не удалось выполнить аварийное сохранение.[/red]")
        raise

    try:
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        console.print(f"[green]✔ Данные успешно экспортированы в CSV: {csv_file}[/green]")
    except Exception as e:
        console.print(f"[red]Ошибка при записи CSV-файла: {e}[/red]")

def main():
    parser = argparse.ArgumentParser(description="Парсер товаров Wildberries на Python")
    parser.add_argument("-m", "--mode", choices=["sku", "search"], help="Режим работы: 'sku' (по списку артикулов) или 'search' (по ключевому слову)")
    parser.add_argument("-s", "--skus", help="Список артикулов через запятую (для режима sku)")
    parser.add_argument("-q", "--query", help="Поисковый запрос (для режима search)")
    parser.add_argument("-p", "--pages", type=int, default=5, help="Количество страниц для поиска (по умолчанию 5)")
    parser.add_argument("--proxy-file", help="Путь к файлу со списком прокси")
    parser.add_argument("-o", "--output", default="wb_results", help="Базовое имя файла для сохранения результатов (без расширения)")
    parser.add_argument("--supplier-id", type=int, help="ID продавца (поставщика) для фильтрации результатов")
    
    args = parser.parse_args()
    
    console.print(Panel.fit(
        "[bold cyan]Парсер товаров Wildberries (WB)[/bold cyan]\n"
        "[dim]Инструмент для сбора базовой и расширенной информации о товарах[/dim]",
        border_style="cyan"
    ))
    
    # Load proxies if provided
    proxies = None
    if args.proxy_file:
        proxies = load_proxies(args.proxy_file)
    elif os.path.exists("proxies.txt"):
        # Auto-load if file exists in the directory
        proxies = load_proxies("proxies.txt")
        
    # Fetch warehouse mapping
    console.print("[cyan]Загрузка справочника складов WB для декодирования названий складов...[/cyan]")
    warehouse_map = fetch_warehouse_map(proxies)
    if warehouse_map:
        console.print(f"[green]Справочник складов успешно загружен ({len(warehouse_map)} записей).[/green]")
        
    mode = args.mode
    
    # Interactive CLI mode selection if not specified in arguments
    if not mode:
        console.rule("[bold yellow]Выбор режима работы[/bold yellow]")
        console.print("[bold cyan]1[/bold cyan] — Сбор данных по артикулам (SKU)")
        console.print("[bold cyan]2[/bold cyan] — Поиск товаров по ключевому слову")
        mode_choice = Prompt.ask(
            "Выберите режим (введите 1 или 2)",
            choices=["1", "2"],
            default="1",
            show_choices=False
        )
        if mode_choice == "1":
            mode = "sku"
            console.print("[cyan]Выбран режим: Сбор данных по артикулам (SKU)[/cyan]")
        else:
            mode = "search"
            console.print("[cyan]Выбран режим: Поиск товаров по ключевому слову[/cyan]")
            
    products_data = []
    
    if mode == "sku":
        skus_str = args.skus
        if not skus_str:
            skus_str = Prompt.ask("Введите артикулы товаров (через запятую или пробел)")
            
        # Clean and parse SKUs list
        skus_list = []
        for val in skus_str.replace(',', ' ').split():
            try:
                skus_list.append(int(val.strip()))
            except ValueError:
                continue
                
        if not skus_list:
            console.print("[red]Ошибка: Не указано ни одного корректного артикула.[/red]")
            return
            
        console.print(f"[cyan]Начинаем сбор по {len(skus_list)} артикулам...[/cyan]")
        products_data = fetch_products_by_skus(skus_list, warehouse_map, proxies)
        
    elif mode == "search":
        query = args.query
        if not query:
            query = Prompt.ask("Введите ключевое слово для поиска")
            
        pages = args.pages
        if not args.pages or args.pages == 5:
            pages_str = Prompt.ask("Укажите количество страниц для парсинга (1 страница = 100 товаров)", default="3")
            try:
                pages = int(pages_str)
            except ValueError:
                pages = 3
                
        console.print(f"[cyan]Запуск поиска по запросу: '{query}' ({pages} стр.)...[/cyan]")
        products_data = search_products_by_query(query, pages, warehouse_map, proxies)
        
    # Check if we successfully gathered any data
    if products_data:
        console.print(f"\n[bold green]Успешно собрано товаров: {len(products_data)}[/bold green]")
        
        # Filter by supplier ID if requested or dynamically chosen
        supplier_id = args.supplier_id
        if not supplier_id:
            if Confirm.ask("Хотите сделать выборку по ID продавца (поставщика)?"):
                supplier_id_str = Prompt.ask("Введите ID продавца (например, 682757)")
                try:
                    supplier_id = int(supplier_id_str.strip())
                except ValueError:
                    console.print("[red]Некорректный ID продавца. Фильтрация отменена.[/red]")
                    
        if supplier_id:
            filtered_data = [p for p in products_data if p.get('ID Продавца') == supplier_id]
            console.print(f"[cyan]Применена выборка по ID продавца {supplier_id}. Было товаров: {len(products_data)}, Стало: {len(filtered_data)}.[/cyan]")
            products_data = filtered_data
            
        if not products_data:
            console.print("[yellow]В результате выборки не осталось ни одного товара. Экспорт отменен.[/yellow]")
            return
            
        # Display preview in a gorgeous Rich table
        display_results_in_table(products_data, limit=10)
        
        # Ask about exporting characteristics and descriptions
        console.rule("[bold yellow]Выгрузка характеристик и описания[/bold yellow]")
        console.print("[bold cyan]1[/bold cyan] — Да, выгрузить Характеристики и Описание для каждого товара (отдельный лист в Excel)")
        console.print("[bold cyan]2[/bold cyan] — Нет, только основные данные")
        details_choice = Prompt.ask(
            "Нужна ли выгрузка Характеристик и описания по выбранным товарам?",
            choices=["1", "2"],
            default="2",
            show_choices=False
        )
        include_details = (details_choice == "1")
        
        # Determine output file name
        output_name = args.output
        if not args.output or args.output == "wb_results":
            output_name = Prompt.ask("Введите имя файла для сохранения результатов", default="wb_results")
            
        export_data(products_data, output_name, include_details=include_details, proxies=proxies)
        
        extra_sheets_msg = f"\n - [cyan]{output_name}.xlsx[/cyan] содержит листы 1–{len(products_data)} с характеристиками" if include_details else ""
        console.print(Panel(
            "[bold green]Парсинг успешно завершен![/bold green]\n"
            f"Созданы файлы:\n"
            f" - [cyan]{output_name}.xlsx[/cyan] (лист 'Товары' — основные данные){extra_sheets_msg}\n"
            f" - [cyan]{output_name}.csv[/cyan] (формат CSV UTF-8)\n"
            "Вы можете открыть их для дальнейшего анализа.",
            title="Результат",
            border_style="green"
        ))
    else:
        console.print("[red]К сожалению, не удалось собрать данные. Проверьте подключение к сети или настройки прокси.[/red]")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n[yellow]Работа парсера прервана пользователем.[/yellow]")
        sys.exit(0)
