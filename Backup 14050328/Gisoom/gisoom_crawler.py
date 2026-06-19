import requests
import re
import json
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class GisoomCrawler:
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/120.0.0.0 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
        }
        # session با retry برای رفع timeout های گاه‌به‌گاه
        self.session = requests.Session()
        retries = Retry(total=3, backoff_factor=1.5,
                        status_forcelist=[500, 502, 503, 504])
        self.session.mount('https://', HTTPAdapter(max_retries=retries))

    def normalize_isbn(self, isbn):
        if not isbn:
            return ""
        return re.sub(r'[^0-9]', '', str(isbn))

    def find_book_page(self, isbn):
        isbn = self.normalize_isbn(isbn)
        if not isbn:
            return None

        search_url = f"https://www.gisoom.com/search/book/isbn-{isbn}/"
        payload = {
            'data[param][recent][id]': 'recent',
            'data[param][recent][val]': '0',
            'data[act]': 'setparam',
            'data[darsad]': '',
            'data[page]': '0'
        }
        try:
            response = self.session.post(search_url, data=payload,
                                         headers=self.headers, timeout=30)
            if response.status_code != 200:
                return None

            match = re.search(r"<div class='hide searchresult'>(.*?)</div>",
                              response.text, re.DOTALL)
            if not match:
                return None

            items = json.loads(match.group(1).strip())
            if not items:
                return None

            item = items[0]
            gid = item.get('gid')
            if not gid:
                return None

            return {
                "url": f"https://www.gisoom.com/book/{gid}/",
                "gid": gid,
                "isbn_found": item.get('isbn'),
                # کلیدها مطابق schema نهایی (سازگار با excel_handler)
                "title": (item.get('name') or "").strip(),
                "author": (item.get('author') or "").strip(),
                "publisher": (item.get('nasher') or "").strip(),
            }
        except Exception as e:
            print(f"Error in find_book_page: {e}")
            return None

    def parse_book_page(self, url):
        try:
            response = self.session.get(url, headers=self.headers, timeout=30)
            if response.status_code != 200:
                return {}

            soup = BeautifulSoup(response.text, 'html.parser')
            details = {}

            # --- meta tags ---
            og_title = soup.find("meta", property="og:title")
            if og_title and og_title.get("content"):
                details['title'] = og_title["content"].replace("کتاب ", "").strip()

            og_image = soup.find("meta", property="og:image")
            if og_image and og_image.get("content"):
                details['image_url'] = og_image["content"].strip()

            # --- متن ساده‌ی صفحه (مقاوم در برابر ساختار تگ‌ها) ---
            text = soup.get_text("\n", strip=True)

            def grab(label):
                m = re.search(rf'{label}\s*[:：]\s*(.+)', text)
                return m.group(1).strip() if m else ""

            def to_digits(s):
                # تبدیل ارقام فارسی به انگلیسی و حذف بقیه
                fa = "۰۱۲۳۴۵۶۷۸۹"
                en = "0123456789"
                s = s.translate(str.maketrans(fa, en))
                return re.sub(r'[^\d]', '', s)

            author = grab('مؤلف')
            if author:
                details['author'] = author

            publisher = grab('ناشر')
            if publisher:
                details['publisher'] = publisher

            translator = grab('مترجم')
            if translator:
                details['translator'] = translator

            pages = grab('تعداد صفحات')
            if pages:
                pages = to_digits(pages)
                if pages:
                    details['pages'] = pages

            year = grab('سال چاپ')
            if year:
                year = to_digits(year)
                if year:
                    details['year'] = year

            language = grab('زبان')
            if language:
                details['language'] = language

            return details

        except Exception as e:
            print(f"Error in parse_book_page: {e}")
            return {}
