# -*- coding: utf-8 -*-
"""
Book_Crawler_v2.4
- جداشدن منطق: get_book_div_from_page و extract_details_from_div
- دانلود تصویر با نام شابک (بدون -) به پوشه "Books Images"
- درج تصویر در اکسل به صورت LinkToFile و Placement = MoveAndSize
- استفاده از html.parser (بدون lxml)
"""
import os
import re
import time
import random
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import sys

# تنظیمات
EXCEL_FILE = "Parsa Library.xlsx"
IMAGE_DIR = "Books Images"
BASE = "https://www.iranketab.ir"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
DELAY_RANGE = (0.8, 2.0)
UPDATE_ALL = False
EXCEL_VISIBLE = True

os.makedirs(IMAGE_DIR, exist_ok=True)

def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")

def remove_old_images(ws, row, col):
    """
    تمام تصاویر یا Shapeهایی که در سلول مشخص شده (row, col) قرار دارند را حذف می‌کند.
    برای جلوگیری از بهم ریختن ایندکس‌ها هنگام حذف، از انتهای لیست شروع می‌کند.
    """
    try:
        # تعداد کل Shapeها
        count = ws.Shapes.Count
        # حلقه معکوس (از آخر به اول)
        for i in range(count, 0, -1):
            shp = ws.Shapes.Item(i)
            # بررسی اینکه آیا "سلولِ لنگر" تصویر، همان سلول مورد نظر ماست؟
            if shp.TopLeftCell.Row == row and shp.TopLeftCell.Column == col:
                shp.Delete()
    except Exception as e:
        print(f"Error removing old images: {e}")
        
def safe_get(url, allow_redirects=True, timeout=25):
    """
    درخواست HTTP ایمن با fallback خودکار برای خطای SSL.
    ابتدا با verify=True تلاش می‌کند و در صورت SSLError مجدد با verify=False امتحان می‌شود.
    """
    try:
        time.sleep(random.uniform(*DELAY_RANGE))
        # تلاش اول با تأیید SSL
        r = requests.get(url, headers=HEADERS, allow_redirects=allow_redirects, timeout=timeout, verify=True)
        r.raise_for_status()
        r.encoding = "utf-8"
        return r
    except requests.exceptions.SSLError as e:
        log(f"⚠️ هشدار SSL در GET {url}: {e} → تلاش مجدد بدون تأیید SSL ...")
        try:
            r = requests.get(url, headers=HEADERS, allow_redirects=allow_redirects, timeout=timeout, verify=False)
            r.raise_for_status()
            r.encoding = "utf-8"
            return r
        except Exception as e2:
            log(f"❌ خطا در تلاش دوم GET بدون SSL ({url}): {e2}")
            return None
    except requests.exceptions.RequestException as e:
        log(f"⚠️ خطا در GET {url}: {e}")
        return None


# ---------- تابعی که URL نهایی صفحه کتاب را برمی‌گرداند و HTML آن را ----------
def get_final_book_url_and_html(isbn):
    isbn_clean = str(isbn).replace("-", "").strip()
    if not isbn_clean:
        return None, None
    search_url = f"{BASE}/result/{isbn_clean}?t=کتاب&s=0"
    r = safe_get(search_url)
    if not r:
        return None, None
    final_url = r.url
    html = r.text

    # اگر ریدایرکت مستقیم به /book/ شده، خوب است
    if "/book/" in final_url:
        return final_url, html

    # وگرنه از صفحه نتایج اولین لینک /book/ را بردار
    soup = BeautifulSoup(html, "html.parser")
    link_tag = soup.find("a", href=re.compile(r"^/book/"))
    if link_tag:
        href = link_tag.get("href")
        book_page_url = (BASE + href) if href.startswith("/") else href
        r2 = safe_get(book_page_url)
        if r2:
            return book_page_url, r2.text
    return None, None

# ---------- تابعی که از HTML صفحه دقیقاً همان div نسخه مورد نظر را پیدا می‌کند ----------
def get_book_div_from_page(page_url, html, isbn):
    """
    ورودی: page_url (ممکن است شامل fragment مثل #p-114766)، html صفحه، و شابک مورد نظر
    خروجی: bs4 Tag مربوط به همان نسخه (div) یا None اگر صفحه تک‌کتاب یا پیدا نشد.
    رفتار:
      - اگر fragment (#p-XXXX) وجود داشت سعی می‌کند div با همان id را برگرداند.
      - در غیر این صورت، لیست divهای id^="p-" را می‌گرداند و در هر کدام دنبال span با 'شابک:' می‌گردد و مقایسه می‌کند.
      - اگر هیچی پیدا نشد، None بازمی‌گرداند (حالت تک‌کتاب یا نامشخص).
    """
    if not html:
        return None
    soup = BeautifulSoup(html, "html.parser")

    # بررسی fragment در URL (مثلاً #p-30809)
    frag = urlparse(page_url).fragment if page_url else ""
    if frag:
        # تلاش برای پیدا کردن div با همان id
        candidate = soup.find(id=frag)
        if candidate:
            log(f" -> یافتن div با fragment: #{frag}")
            return candidate

    # لیست کاندیدها: div هایی که id شان با p- شروع می‌شود یا کلاس wrapper-product-acts داخلشان data-id دارد
    candidates = soup.select('div[id^="p-"], div.flex.gap-2.mb-3, div.wrapper-product-acts')
    # تحویل فیلتر شده برای پیدا کردن کارت‌های واقعی: اغلب structure دارای id="p-XXXX" است و در کنار آن بخش متا موجود است
    # اما safest approach: پیدا کردن تمام divهایی که id شروع با p- دارند
    candidates = soup.select('div[id^="p-"]')
    if not candidates:
        # صفحه ممکن است تک‌کتاب باشد
        log(" -> صفحه حاوی div[id^='p-'] نیست — احتمالاً صفحه تک‌کتاب.")
        return None

    clean_isbn = str(isbn).replace("-", "").strip()
    log(f" -> {len(candidates)} div با id^='p-' پیدا شد؛ در جستجوی تطابق شابک {clean_isbn} ...")
    for div in candidates:
        # در داخل هر div معمولاً یک span با متن "شابک:" وجود دارد و sibling آن مقدار شابک است
        isbn_span = div.find(lambda t: t.name in ["span","div"] and isinstance(t.string, str) and "شابک" in t.string)
        if isbn_span:
            # سعی کنیم مقدار بعدی یا sibling را بخوانیم
            # حالت‌های مختلف: <span>شابک:</span><span>978-...</span>
            val = None
            nxt = isbn_span.find_next_sibling()
            if nxt and isinstance(nxt.get_text(strip=True), str) and nxt.get_text(strip=True):
                val = nxt.get_text(strip=True)
            else:
                # یا ساختار دیگر: isbn_span.parent.find(...) 
                possible = isbn_span.parent.find_all("span")
                for s in possible:
                    txt = s.get_text(strip=True)
                    if re.search(r"\d{9,13}", txt):
                        val = txt
                        break
            if val:
                if str(val).replace("-", "").strip() == clean_isbn:
                    log(f" -> تطابق شابک در div id='{div.get('id')}' یافت شد ({val}).")
                    return div
    log(" -> هیچ div منطبق با شابک پیدا نشد؛ بازگشت None (احتمالاً صفحه تک‌کتاب یا ساختار متفاوت).")
    return None

    # --- تابع کمکی هوشمند برای استخراج شماره جلد ---
    def infer_collection_number(main_title: str, sub_title: str):
        # ترکیب برای جستجوی جامع، اما اولویت جستجو معمولاً روی sub_title است
        # نرمال‌سازی اعداد فارسی به انگلیسی
        def normalize_digits(txt):
            if not txt: return ""
            replacements = {'۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4', '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'}
            for k, v in replacements.items():
                txt = txt.replace(k, v)
            return txt

        clean_sub = normalize_digits(sub_title)
        clean_main = normalize_digits(main_title)
        
        # 1. اولویت اول: بازه عددی داخل پرانتز مثل (11-10) یا (10-11)
        # مثال: کتاب ایلیا (11-10)
        range_match = re.search(r"\((\d+)\s*-\s*(\d+)\)", clean_sub)
        if range_match:
            n1, n2 = range_match.group(1), range_match.group(2)
            # مرتب‌سازی اعداد (که همیشه 10 و 11 باشد نه 11 و 10)
            nums = sorted([int(n1), int(n2)])
            return f"{nums[0]} و {nums[1]}"

        # 2. اولویت دوم: عدد داخل پرانتز (با یا بدون متن)
        # مثال: (12) یا (نگهبانان گاهول 3) یا (5)
        # پترن: پرانتز باز -> (اختیاری: هر متنی) -> عدد -> پرانتز بسته
        paren_match = re.search(r"\((?:[^)]*?\s+)?(\d{1,3})\)", clean_sub)
        if paren_match:
            return paren_match.group(1)

        # 3. اولویت سوم: جداکننده خاص مثل "_" یا "-" قبل از عدد
        # مثال: هفت نشانه _ 4
        sep_match = re.search(r"[_\-]\s*(\d{1,3})\b", clean_sub)
        if sep_match:
            return sep_match.group(1)

        # 4. اولویت چهارم: عدد در انتهای متن (با فاصله از کلمه قبل)
        # مثال: قصه های فلیکس 5
        # شرط: عدد باید 1 تا 3 رقم باشد (تا سال انتشار مثل 1399 اشتباه نشود)
        trailing_match = re.search(r"\s+(\d{1,3})$", clean_sub.strip())
        if trailing_match:
            return trailing_match.group(1)

        # 5. اولویت پنجم: تطبیق با عنوان اصلی (روش قدیمی)
        # مثال: اگر عنوان اصلی "ته جدولی ها" باشد و عنوان فرعی "ته جدولی ها 2"
        if clean_main:
            stripped_main = re.escape(clean_main.strip())
            main_pattern = rf"{stripped_main}\s*(\d+)"
            m_main = re.search(main_pattern, clean_sub)
            if m_main:
                return m_main.group(1)

        return ""


# ---------- تابع استخراج جزئیات فقط از همان div (یا در صورت None از صفحه کلی) ----------
def extract_details_from_div(div, soup, isbn):
    info = {}

    # --- تابع کمکی: فقط شماره را پیدا می‌کند (متن را تغییر نمی‌دهد) ---
    def detect_collection_number(text):
        if not text: return ""
        
        # نرمال‌سازی موقت برای جستجو
        temp_text = text.strip()
        replacements = {'۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4', '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'}
        for k, v in replacements.items():
            temp_text = temp_text.replace(k, v)

        # 1. اولویت اول: الگوی "عدد و عدد" (مثل: 6 و 7)
        m_and = re.search(r"(\d+)\s*و\s*(\d+)", temp_text)
        if m_and:
            nums = sorted([int(m_and.group(1)), int(m_and.group(2))])
            return f"{nums[0]} و {nums[1]}"

        # 2. الگوی بازه: (11-10)
        m_range = re.search(r"\((\d+)\s*-\s*(\d+)\)", temp_text)
        if m_range:
            nums = sorted([int(m_range.group(1)), int(m_range.group(2))])
            return f"{nums[0]} و {nums[1]}"

        # 3. الگوی پرانتز تکی: (12)
        m_paren = re.search(r"\((?:[^)]*?\s+)?(\d{1,3})\)", temp_text)
        if m_paren: return m_paren.group(1)

        # 4. الگوی جداکننده انتهای خط: _ 4 یا - 4
        m_sep = re.search(r"[_\-]\s*(\d{1,3})$", temp_text)
        if m_sep: return m_sep.group(1)
        
        # 5. الگوی دو نقطه: " : 1"
        m_colon = re.search(r":\s*(\d{1,3})", temp_text)
        if m_colon: return m_colon.group(1)

        # 6. عدد خالی در انتهای متن (ریسک دارد ولی برای مجموعه‌ها معمول است)
        m_trail = re.search(r"\s+(\d{1,3})$", temp_text)
        if m_trail: return m_trail.group(1)

        return ""

    # --- تابع کمکی خواندن جفت کلید/مقدار ---
    def read_kv_pairs(block):
        kv = {}
        for row in block.select("div.flex.gap-1, div.flex.gap-2, div.flex.gap-3, div.flex.gap-4, div.flex.gap-0"):
            spans = row.find_all("span")
            if len(spans) >= 2:
                key = spans[0].get_text(strip=True).replace(":", "")
                val = spans[1].get_text(strip=True)
                kv[key] = val
        return kv

    if div is not None:
        try:
            # 1. استخراج عناوین خام
            h2_tag = div.find(lambda t: t.name in ["h2","h3"] and t.get_text(strip=True))
            if not h2_tag: h2_tag = div.find(itemprop="name")
            raw_h2 = h2_tag.get_text(strip=True) if h2_tag else ""

            sub_tag = h2_tag.find_next_sibling('div') if h2_tag else None
            raw_sub = ""
            if sub_tag and 'ltr' not in sub_tag.get('class', []):
                raw_sub = sub_tag.get_text(strip=True)

            # 2. ثبت در دیکشنری (طبق دستور: فقط حذف کلمه "کتاب" از عنوان اصلی)
            # عنوان اصلی
            final_main_title = re.sub(r"^کتاب\s+", "", raw_h2).strip()
            info["عنوان اصلی"] = final_main_title
            
            # عنوان فرعی (بدون تغییر)
            info["عنوان فرعی"] = raw_sub

            # 3. استخراج شماره (جستجو در زیرنویس، اگر نبود در عنوان اصلی)
            found_num = detect_collection_number(raw_sub)
            if not found_num:
                found_num = detect_collection_number(raw_h2)
            
            info["شماره در مجموعه"] = found_num

        except Exception as e:
            print(f"Error extracting title: {e}")
            info["عنوان اصلی"] = ""
            info["عنوان فرعی"] = ""
            info["شماره در مجموعه"] = ""

        # 4. تصویر
        a_img = div.find("a", href=re.compile(r"/Images/ProductImages/|/Files/AttachFiles/"))
        img_url = ""
        if a_img and a_img.get("href"):
            img_url = a_img.get("href")
        else:
            img_tag = div.find("img", itemprop="image")
            if img_tag and img_tag.get("src"):
                img_url = img_tag.get("src")

        if img_url:
            full_url = img_url if img_url.startswith("http") else (BASE + img_url)
            info["image_url"] = full_url
            info["iranketabImageName"] = os.path.basename(urlparse(full_url).path)
        else:
            info["image_url"] = ""
            info["iranketabImageName"] = ""

        # 5. قیمت
        final_price = ""
        old_price_tag = div.find("s", class_=lambda c: c and "price" in c)
        if old_price_tag:
            final_price = re.sub(r'\D', '', old_price_tag.get_text(strip=True))
        if not final_price:
            current_price_tag = div.find(class_="toman")
            if current_price_tag:
                final_price = re.sub(r'\D', '', current_price_tag.get_text(strip=True))
            else:
                alt = div.select_one(".price, .product-price")
                if alt: final_price = re.sub(r'\D', '', alt.get_text(strip=True))
        info["قیمت"] = final_price

        # 6. تاریخ‌ها و سایر فیلدها
        def get_val(cont, lbl):
            t = cont.find(lambda x: x.name=="span" and lbl in x.get_text(strip=True))
            return t.find_next_sibling("span").get_text(strip=True) if t and t.find_next_sibling("span") else ""

        info["سال انتشار شمسی"] = get_val(div, "سال انتشار شمسی")
        info["سال انتشار میلادی"] = get_val(div, "سال انتشار میلادی")

        def find_lbl(d, l):
            s = d.find(lambda t: t.name in ["span","div"] and isinstance(t.string, str) and l in t.get_text())
            if s:
                a = s.find_next("a")
                if a: return a.get_text(strip=True)
                sib = s.find_next_sibling()
                if sib: return sib.get_text(strip=True)
            return ""

        info["نویسنده"] = find_lbl(div, "نویسنده:")
        info["مترجم"] = find_lbl(div, "مترجم:")
        info["ناشر"] = find_lbl(div, "انتشارات:")

        orig = soup.find("div", class_=lambda c: c and "ltr" in c)
        info["عنوان انگلیسی"] = orig.get_text(strip=True) if orig else ""

        kv = {}
        for block in div.select("div.card, div.absolute, div.flex"): kv.update(read_kv_pairs(block))
        if not kv: kv.update(read_kv_pairs(div))
        for k, v in kv.items(): info[k] = v

        if "شابک" not in info:
            isbn_span = div.find(lambda t: t.name in ["span","div"] and isinstance(t.string, str) and "شابک" in t.get_text())
            if isbn_span:
                nxt = isbn_span.find_next_sibling()
                if nxt: info["شابک"] = nxt.get_text(strip=True)

        # -------------------------------------------------------
        pages = ""
        # پیدا کردن تگی که متن "تعداد صفحه" دارد
        # استفاده از re.compile برای پیدا کردن متن حتی اگر فاصله داشته باشد
        page_label = soup.find("span", string=re.compile("تعداد صفحه"))
        if page_label:
            # پیدا کردن تگ بعدی (همسایه) که حاوی عدد است
            page_val_span = page_label.find_next_sibling("span")
            if page_val_span:
                pages = page_val_span.get_text(strip=True)
        
        info["صفحات"] = pages
        # -------------------------------------------------------
        
        return info
    return {}



# ---------- pywin32 helper ----------
def ensure_pywin32():
    try:
        import win32com.client as win32
        return win32
    except Exception:
        raise SystemExit("❌ نیاز به pywin32 (python -m pip install pywin32)")

# ---------- دانلود تصویر (با اصلاح SSL) ----------
def download_image(img_url, isbn):
    if not img_url:
        return None
    isbn_clean = str(isbn).replace("-", "").strip()
    filename = f"{isbn_clean}.jpg"
    path = os.path.join(IMAGE_DIR, filename)
    if os.path.exists(path):
        return path
    
    try:
        # تلاش اول: دانلود استاندارد
        r = requests.get(img_url, headers=HEADERS, stream=True, timeout=30, verify=True)
    except requests.exceptions.SSLError:
        # تلاش دوم: اگر خطای SSL داد، بدون بررسی امنیتی دانلود کن
        try:
            r = requests.get(img_url, headers=HEADERS, stream=True, timeout=30, verify=False)
        except Exception:
            return None
    except Exception:
        return None

    try:
        r.raise_for_status()
        with open(path, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        return path
    except Exception:
        return None

# ---------------------------------------------------------
# تابع اصلی (MAIN)
# ---------------------------------------------------------
def main():
    win32 = ensure_pywin32()
    try:
        # فقط برای خواندن تعداد رکوردها و ایندکس‌ها از پانداس استفاده می‌کنیم
        df_source = pd.read_excel(EXCEL_FILE, sheet_name=0, dtype=str).fillna("")
    except Exception as e:
        raise SystemExit(f"❌ خطا در خواندن فایل اکسل: {e}")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True  # مشاهده روند کار
    wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE))
    ws = wb.Worksheets(1)

    # 1. نقشه‌برداری از هدرها (Header Map)
    used_cols = ws.UsedRange.Columns.Count
    header_map = {}
    # ردیف 1 را می‌خوانیم
    for c in range(1, used_cols + 20): # کمی بیشتر می‌خوانیم تا مطمئن شویم
        val = ws.Cells(1, c).Value
        if val:
            header_map[str(val).strip()] = c
        else:
            # اگر به سلول خالی رسیدیم و بعدش هم خالی بود، فرض می‌کنیم تمام شده
            # (اما چون می‌خواهیم ستون اضافه کنیم، دقیق‌ترین راه UsedRange بود که بالا زدیم)
            pass

    # 2. لیست ستون‌های اجباری که باید باشند (اگر نباشند می‌سازیم)
    required_cols = [
        "تصویر", 
        "قیمت", 
        "iranketabImageName", 
        "عنوان اصلی", 
        "عنوان فرعی", 
        "شماره در مجموعه",
        "سال انتشار شمسی",
        "سال انتشار میلادی",
        "صفحات"  # <--- این مورد اضافه شد
    ]

    # پیدا کردن آخرین ستون پر شده
    last_col_idx = 0
    if header_map:
        last_col_idx = max(header_map.values())
    
    for req in required_cols:
        if req not in header_map:
            last_col_idx += 1
            ws.Cells(1, last_col_idx).Value = req
            header_map[req] = last_col_idx
            log(f" + ستون جدید ایجاد شد: {req}")

    if "شابک" not in header_map:
        wb.Close(SaveChanges=False)
        excel.Quit()
        raise SystemExit("❌ ستون 'شابک' پیدا نشد.")

    total_rows = len(df_source)
    log(f"شروع پردازش {total_rows} رکورد...")

    # 3. حلقه روی داده‌ها
    for idx, row_data in df_source.iterrows():
        excel_row = int(idx) + 2  # چون هدر ردیف 1 است و پانداس از 0 شروع می‌کند
        isbn = str(row_data.get("شابک", "")).strip()

        if not isbn:
            continue

        # --- شرط پرش (Skip Logic) ---
        # اگر UPDATE_ALL خاموش است، چک کن آیا "عنوان اصلی" پر است؟
        if not UPDATE_ALL:
            # مقدار فعلی سلول عنوان اصلی را از اکسل زنده بخوان (نه از پانداس که شاید قدیمی باشد)
            current_main_title = ws.Cells(excel_row, header_map["عنوان اصلی"]).Value
            # اگر سلول پر بود (None نبود و رشته خالی نبود)
            if current_main_title and str(current_main_title).strip():
                # لاگ نمی‌کنیم یا کوتاه لاگ می‌کنیم تا شلوغ نشود
                # log(f"ردیف {excel_row} رد شد (عنوان دارد).")
                continue
        # -----------------------------

        log(f"\n[ردیف {excel_row}] 🔎 شابک: {isbn}")

        # دریافت لینک و HTML
        book_url, html = get_final_book_url_and_html(isbn)
        if not book_url or not html:
            log(" -> ❌ صفحه کتاب پیدا نشد.")
            continue

        # استخراج اطلاعات
        book_div = get_book_div_from_page(book_url, html, isbn)
        soup = BeautifulSoup(html, "html.parser")
        details = extract_details_from_div(book_div, soup, isbn)

        if not details:
            log(" -> ❌ اطلاعات استخراج نشد.")
            continue

        # 4. درج اطلاعات متنی در اکسل
        for key, val in details.items():
            if key == "تصویر": continue # تصویر جداگانه هندل می‌شود
            
            if key in header_map:
                col_idx = header_map[key]
                # نوشتن مقدار در سلول
                ws.Cells(excel_row, col_idx).Value = val

        # 5. مدیریت تصویر (دانلود، حذف قبلی، درج جدید)
        img_url = details.get("image_url", "")
        if img_url:
            local_img_path = download_image(img_url, isbn)
            if local_img_path:
                img_col = header_map["تصویر"]
                
                # الف) حذف تصاویر قدیمی این سلول
                remove_old_images(ws, excel_row, img_col)

                # ب) درج تصویر جدید
                try:
                    cell_target = ws.Cells(excel_row, img_col)
                    left = cell_target.Left
                    top = cell_target.Top
                    width = cell_target.Width
                    height = cell_target.Height
                    
                    # AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                    pic = ws.Shapes.AddPicture(
                        os.path.abspath(local_img_path), 
                        False, 
                        True, 
                        left, top, width, height
                    )
                    pic.Placement = 1 # Move and Size
                except Exception as e:
                    log(f"⚠️ خطا در درج تصویر: {e}")
            else:
                log(" -> دانلود تصویر ناموفق بود.")

        # ذخیره موقت هر 10 رکورد (اختیاری، برای امنیت بیشتر)
        if idx % 10 == 0:
            wb.Save()

        time.sleep(random.uniform(*DELAY_RANGE))

    # پایان کار
    wb.Save()
    log("✅ پایان عملیات. فایل ذخیره شد.")
    # wb.Close(SaveChanges=True) # اگر می‌خواهید باز بماند این را کامنت کنید
    # excel.Quit()

if __name__ == "__main__":
    main()