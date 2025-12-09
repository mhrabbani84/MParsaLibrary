import pandas as pd
import json
import os

# --- Configuration ---
EXCEL_FILE = 'Parsa Library.xlsx'
OUTPUT_FILE = 'index.html'
IMAGE_BASE_PATH_LOCAL = 'Books Images'
# تصویر پیش‌فرض (مسیر فیزیکی)
DEFAULT_COVER_PATH = 'Books Images/default_cover.png' 
IRANKETAB_URL_PREFIX = "https://img.iranketab.ir/img/225x330?pic=www.iranketab.ir/Images/ProductImages"
IRANKETAB_IMAGE = True

def generate_html():
    print("Reading Excel file...")
    # 1. Read Excel
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name='کتابخانه')
    except Exception as e:
        print(f"Error reading sheet 'کتابخانه': {e}")
        try:
            df = pd.read_excel(EXCEL_FILE)
        except Exception as e2:
            print(f"Critical Error: {e2}")
            return

    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # 2. Column Mapping
    cols_map = {
        'isbn': ['شابک', 'ISBN'],
        'title_main': ['عنوان اصلی', 'عنوان کتاب', 'عنوان'],
        'title_sub': ['عنوان فرعی'],
        'author': ['نویسنده', 'پدیدآورنده'],
        'translator': ['مترجم'],
        'publisher': ['ناشر'],
        'year_shamsi': ['سال انتشار شمسی', 'سال انتشار'],
        'year_gregorian': ['سال انتشار میلادی'],
        'score': ['امتیاز'],
        'status': ['وضعیت', 'خوانده شده'],
        'iranketab_img': ['iranketabImageName', 'تصویر ایران کتاب'],
        'code': ['کد', 'code', 'Code'],
        'pages': ['صفحات', 'تعداد صفحه']
    }

    def get_val(row, keys):
        for k in keys:
            if k in row.index:
                val = str(row[k])
                if val and val.lower() != 'nan':
                    return val.strip()
        return ''

    # تابع کمکی برای اصلاح لینک‌ها در HTML (تبدیل فاصله به %20)
    def sanitize_url(path):
        # اگر لینک اینترنتی است، دست نزن
        if path.lower().startswith(('http://', 'https://')):
            return path
        # اگر فایل لوکال است، فاصله را با %20 جایگزین کن
        return path.replace(' ', '%20')

    # آماده‌سازی مسیر تصویر پیش‌فرض برای HTML
    default_cover_html = sanitize_url(DEFAULT_COVER_PATH)

    # 3. Process Books
    books = []
    print("Processing books...")
    for _, row in df.iterrows():
        title_main = get_val(row, cols_map['title_main'])
        if not title_main:
            continue
        title_sub = get_val(row, cols_map['title_sub'])

        isbn = get_val(row, cols_map['isbn'])
        cleaned_isbn = isbn.replace('-', '').replace(' ', '') if isbn else ''
        iranketab_filename = get_val(row, cols_map['iranketab_img'])

        # --- LOGIC FOR IMAGES ---
        if IRANKETAB_IMAGE and iranketab_filename:
            # اگر لینک کامل اینترنتی بود (مثل مورد shabgoonp)
            if iranketab_filename.lower().startswith(('http://', 'https://')):
                final_image_path = iranketab_filename
            else:
                # اگر فقط نام فایل بود، به سایت ایران کتاب وصل کن
                final_image_path = f"{IRANKETAB_URL_PREFIX}/{iranketab_filename}"
        else:
            # حالت فایل لوکال
            fname = f"{cleaned_isbn}.jpg" if cleaned_isbn else "default_cover.png"
            # ساخت مسیر کامل
            local_path = f"{IMAGE_BASE_PATH_LOCAL}/{fname}"
            # اصلاح فاصله برای HTML
            final_image_path = sanitize_url(local_path)
        # ------------------------

        raw_status = get_val(row, cols_map['status']).lower()
        if any(x in raw_status for x in ['خوانده شده', 'read', 'yes']):
            status = 'خوانده شده'
        elif any(x in raw_status for x in ['در حال خواندن', 'reading']):
            status = 'در حال خواندن'
        else:
            status = 'خوانده نشده'

        raw_score = get_val(row, cols_map['score'])
        if raw_score.endswith('.0'): raw_score = raw_score[:-2]

        y_sh = get_val(row, cols_map['year_shamsi'])
        y_gr = get_val(row, cols_map['year_gregorian'])
        if y_gr.endswith('.0'): y_gr = y_gr[:-2]
        if y_sh.endswith('.0'): y_sh = y_sh[:-2]

        year_display = y_sh
        if y_gr:
            year_display += f" ({y_gr})" if y_sh else y_gr

        code_val = get_val(row, cols_map['code'])
        if code_val.endswith('.0'): code_val = code_val[:-2]

        pages_val = get_val(row, cols_map['pages'])
        if pages_val.endswith('.0'): pages_val = pages_val[:-2]

        books.append({
            'title_main': title_main,
            'title_sub': title_sub,
            'author': get_val(row, cols_map['author']),
            'translator': get_val(row, cols_map['translator']),
            'publisher': get_val(row, cols_map['publisher']),
            'year': year_display,
            'status': status,
            'score': raw_score,
            'image_path': final_image_path,
            'isbn': cleaned_isbn,
            'code': code_val,
            'pages': pages_val
        })

    books_json = json.dumps(books, ensure_ascii=False)

    # 4. HTML Template
    html_template = r'''
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>کتابخانه دیجیتال محمّدپارسا ربّانی</title>
<style>
@font-face { font-family: 'IRANSansWeb'; src: url('./Fonts/IRANSansWeb(FaNum).woff2') format('woff2'); font-weight: 400; }
@font-face { font-family: 'IRANSansWeb'; src: url('./Fonts/IRANSansWeb(FaNum)_Bold.woff2') format('woff2'); font-weight: 700; }
@font-face { font-family: 'IRANSansWeb'; src: url('./Fonts/IRANSansWeb(FaNum)_Light.woff2') format('woff2'); font-weight: 300; }
@font-face { font-family: 'IRANSansWeb'; src: url('./Fonts/IRANSansWeb(FaNum)_Medium.woff2') format('woff2'); font-weight: 500; }

body { font-family: 'IRANSansWeb', Tahoma, Calibri, Arial, sans-serif; background: #fdfdfd; padding: 20px; color: #1d1d1d; margin: 0; }

.header { display: flex; flex-direction: column; gap: 20px; margin-bottom: 30px; border-bottom: 2px solid #eee; padding-bottom: 20px; }
@media(min-width: 950px) {
    .header { flex-direction: row; align-items: flex-start; justify-content: space-between; }
    .header h1 { white-space: nowrap; margin-top: 8px; }
    .left-column { flex-grow: 1; max-width: 1920px; display: flex; flex-direction: column; gap: 10px; }
    .controls-row { display: flex; gap: 10px; align-items: center; width: 100%; }
    input[type="search"] { flex-grow: 1; }
    .buttons-group { flex-shrink: 0; display: flex; gap: 5px; }
}
.header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; color: #1d1d1d; }
.controls-row { display: flex; gap: 10px; flex-wrap: wrap; }
input[type="search"] { padding: 10px 15px; border-radius: 8px; border: 1px solid #ccc; width: 100%; font-family: inherit; font-size: 0.95rem; outline: none; transition: border 0.3s; }
input[type="search"]:focus { border-color: #888; }
.buttons-group { display: flex; gap: 5px; flex-wrap: wrap; }
.btn { font-family: inherit; background: #fff; border: 1px solid #e0e0e0; padding: 8px 15px; border-radius: 8px; cursor: pointer; transition: all 0.2s; color: #555; font-size: 0.9rem; white-space: nowrap; }
.btn:hover { background: #f9f9f9; }
.btn.active { background: #D00400; color: #fff; border-color: #D00400; }
.stats-bar { font-size: 0.85rem; color: #777; font-weight: 500; padding-right: 5px; text-align: left; padding-left: 5px; }
.stats-highlight { color: #D00400; font-weight: 700; margin: 0 3px; font-size: 0.95rem; }

.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 30px; }
.card { background: #fff; border-radius: 12px; overflow: hidden; border: 1px solid #eee; box-shadow: 0 4px 12px rgba(0,0,0,0.04); display: flex; flex-direction: column; transition: transform 0.3s ease-out, box-shadow 0.3s ease-out; position: relative; top: 0; }
.card:hover { transform: translateY(-10px); box-shadow: 0 15px 30px rgba(0,0,0,0.12); z-index: 10; }
.cover-container { width: 100%; aspect-ratio: 2/3; background: #f4f4f4; position: relative; }
.cover { width: 100%; height: 100%; object-fit: cover; }
.info { padding: 15px; display: flex; flex-direction: column; gap: 8px; flex-grow: 1; }
.title-main { font-weight: 700; font-size: 1.05rem; color: #000; line-height: 1.3; }
.title-sub { font-weight: 400; font-size: 0.85rem; color: #777; }
.meta-rows { display: flex; flex-direction: column; gap: 4px; font-size: 0.9rem; margin-top: auto; }
.meta-row { display: flex; gap: 5px; align-items: baseline; }
.meta-label { color: #999; font-size: 0.85rem; min-width: 75px; }
.meta-value { color: #333; font-weight: 500; }
.footer { margin-top: 15px; padding-top: 12px; border-top: 1px solid #f0f0f0; display: flex; justify-content: space-between; align-items: center; }
.badge { padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 700; font-family: Calibri, Arial, Helvetica, sans-serif; direction: ltr; display: inline-block; letter-spacing: 0.5px; }
.badge.read { background: #e8f5e9; color: #2e7d32; }
.badge.reading { background: #fff3e0; color: #ef6c00; }
.badge.unread { background: #ffebee; color: #c62828; }
.score { font-weight: 700; color: #D00400; font-size: 0.95rem; }
.fade-in { animation: fadeIn 0.5s ease both; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: none; } }
</style>
</head>
<body>

<div class="header">
    <h1>کتابخانه دیجیتال محمّدپارسا ربّانی</h1>
    <div class="left-column">
        <div class="controls-row">
            <input type="search" id="search" placeholder="جستجو در کتاب‌ها...">
            <div class="buttons-group">
                <button class="btn active" data-filter="all">همه</button>
                <button class="btn" data-filter="خوانده شده">خوانده شده</button>
                <button class="btn" data-filter="در حال خواندن">در حال خواندن</button>
                <button class="btn" data-filter="خوانده نشده">خوانده نشده</button>
            </div>
            <div id="stats-bar" class="stats-bar"></div>
        </div>
    </div>
</div>

<div id="grid" class="grid"></div>

<script id="books-data" type="application/json">__BOOKS_JSON__</script>

<script>
const books = JSON.parse(document.getElementById('books-data').textContent || '[]');
const grid = document.getElementById('grid');
const searchInput = document.getElementById('search');
const btns = document.querySelectorAll('.btn');
const statsBar = document.getElementById('stats-bar');

let activeFilter = 'all';
// مقدار کاور پیش‌فرض هم باید بدون اسپیس باشد
const defaultCover = '__DEFAULT_COVER__';

function normalizeStatus(s) {
    if(!s) return 'unread';
    s = String(s).toLowerCase().replace(/\s+/g, '');
    if(s.includes('خواندهشده') || s.includes('read') || s.includes('yes')) return 'read';
    if(s.includes('درحال') || s.includes('reading')) return 'reading';
    return 'unread';
}

function render(list) {
    grid.innerHTML = '';
    const frag = document.createDocumentFragment();
    list.forEach(b => {
        const card = document.createElement('div');
        card.className = 'card fade-in';
        const st = normalizeStatus(b.status);
        const codeText = b.code ? ('MPR' + b.code) : 'MPR____';
        let imgPath = b.image_path || defaultCover;

        card.innerHTML = `
            <div class="cover-container">
                <img src="${imgPath}" class="cover" onerror="this.src='${defaultCover}'">
            </div>
            <div class="info">
                <div>
                    <div class="title-main">${b.title_main}</div>
                    ${b.title_sub ? `<div class="title-sub">${b.title_sub}</div>` : ''}
                </div>
                <div class="meta-rows">
                    ${metaRow('نویسنده:', b.author)}
                    ${metaRow('مترجم:', b.translator)}
                    ${metaRow('ناشر:', b.publisher)}
                    ${metaRow('تعداد صفحه:', b.pages)}
                    ${metaRow('سال انتشار:', b.year)}
                </div>
                <div class="footer">
                    <span class="badge ${st}">${codeText}</span>
                    ${b.score ? `<span class="score">${b.score} امتیاز</span>` : ''}
                </div>
            </div>
        `;
        frag.appendChild(card);
    });

    if(!list.length) grid.innerHTML = '<p style="color:#777;">موردی یافت نشد.</p>';
    else grid.appendChild(frag);
}

function metaRow(label, val) {
    if(!val) return '';
    return `<div class="meta-row"><span class="meta-label">${label}</span><span class="meta-value">${val}</span></div>`;
}

function updateStats(list) {
    let totalPages = 0;
    list.forEach(b => {
        let p = parseInt(b.pages);
        if(!isNaN(p)) totalPages += p;
    });
    const countStr = list.length.toLocaleString('fa-IR');
    const pagesStr = totalPages.toLocaleString('fa-IR');
    statsBar.innerHTML = `نمایش: <span class="stats-highlight">${countStr}</span> کتاب | مجموع صفحات: <span class="stats-highlight">${pagesStr}</span>`;
}

function updateButtonCounts(currentList) {
    const counts = { all: 0, read: 0, reading: 0, unread: 0 };
    currentList.forEach(b => {
        counts.all++;
        const st = normalizeStatus(b.status);
        counts[st]++;
    });
    btns.forEach(btn => {
        const key = btn.getAttribute('data-filter');
        let label = btn.textContent.split(' (')[0];
        let c = 0;
        if(key==='all') c = counts.all;
        if(key==='خوانده شده') c = counts.read;
        if(key==='در حال خواندن') c = counts.reading;
        if(key==='خوانده نشده') c = counts.unread;
        btn.textContent = `${label} (${c})`;
    });
}

function apply() {
    const q = searchInput.value.toLowerCase().trim();
    const searchMatches = books.filter(b => {
        if(!q) return true;
        const hay = (
            b.title_main + ' ' + b.title_sub + ' ' + b.author + ' ' +
            b.publisher + ' ' + b.translator + ' ' + b.year + ' ' +
            b.pages + ' ' + (b.code || '')
        ).toLowerCase();
        return hay.includes(q);
    });

    updateButtonCounts(searchMatches);
    const finalList = searchMatches.filter(b => {
        const st = normalizeStatus(b.status);
        if(activeFilter === 'all') return true;
        if(activeFilter === 'خوانده شده' && st !== 'read') return false;
        if(activeFilter === 'در حال خواندن' && st !== 'reading') return false;
        if(activeFilter === 'خوانده نشده' && st !== 'unread') return false;
        return true;
    });

    render(finalList);
    updateStats(finalList);
}

btns.forEach(btn => {
    btn.addEventListener('click', () => {
        btns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        activeFilter = btn.getAttribute('data-filter');
        apply();
    });
});

searchInput.addEventListener('input', apply);
apply();
</script>

</body>
</html>
'''

    final_html = html_template.replace('__BOOKS_JSON__', books_json)
    # اینجا هم باید مسیر اصلاح شده جایگزین شود
    final_html = final_html.replace('__DEFAULT_COVER__', default_cover_html)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(final_html)
    print(f"Success! Created {OUTPUT_FILE}")

if __name__ == "__main__":
    generate_html()
