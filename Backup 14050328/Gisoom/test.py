import requests, json, re
from urllib.parse import quote

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/plain, */*; q=0.01",
    "Accept-Language": "fa,en;q=0.8",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://www.gisoom.com",
}

POST_BODY = ("data%5Bparam%5D%5Brecent%5D%5Bid%5D=recent"
             "&data%5Bparam%5D%5Brecent%5D%5Bval%5D=0"
             "&data%5Bact%5D=setparam&data%5Bdarsad%5D=&data%5Bpage%5D=0")

def find_book_page(isbn, session=None):
    """
    با شابک، صفحه‌ی کتاب در گیسوم رو پیدا می‌کند.
    خروجی: dict شامل url, gid, urlname و داده‌های اولیه‌ی سرچ — یا None.
    """
    sess = session or requests.Session()
    isbn = str(isbn).strip().replace("-", "")
    search_url = f"https://www.gisoom.com/search/book/isbn-{isbn}/"

    headers = dict(HEADERS)
    headers["Referer"] = search_url

    try:
        r = sess.post(search_url, headers=headers, data=POST_BODY, timeout=20)
        r.raise_for_status()
    except Exception as e:
        print(f"   [خطا] شابک {isbn}: {e}")
        return None

    # استخراج JSON از داخل <div class='hide searchresult'>[...]</div>
    m = re.search(r"searchresult'?\"?\s*>\s*(\[.*?\])\s*</div>", r.text, re.S)
    if not m:
        print(f"   [یافت نشد] شابک {isbn} نتیجه‌ای نداشت.")
        return None

    try:
        items = json.loads(m.group(1))
    except json.JSONDecodeError:
        print(f"   [خطا JSON] شابک {isbn}")
        return None

    if not items:
        return None

    bk = items[0]                      # اولین (و معمولاً تنها) نتیجه
    gid = bk.get("gid")
    urlname = bk.get("urlname", "")
    book_url = f"https://www.gisoom.com/book/{gid}/{quote(urlname)}/"

    return {
        "isbn":    isbn,
        "url":     book_url,
        "gid":     gid,
        "urlname": urlname,
        "name":    bk.get("name"),
        "author":  bk.get("author"),
        "nasher":  bk.get("nasher"),
        "sal":     bk.get("sal"),
        "nobat":   bk.get("nobat"),
    }


if __name__ == "__main__":
    res = find_book_page("9786008869870")
    print(json.dumps(res, ensure_ascii=False, indent=2))
