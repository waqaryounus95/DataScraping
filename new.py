# import re
# import time
# import requests
# from bs4 import BeautifulSoup
# from urllib.parse import quote_plus
# from openpyxl import load_workbook

# # ------------------------------------------------------------
# # CONFIGURATION
# # ------------------------------------------------------------
# excel_path = 'Data_Capture.xlsx'
# sheet_name = 'Data'
# movie_col = 'B'
# start_row = 3
# num_movies = 6

# # Output columns:
# col_genre                = 'F'
# col_director             = 'G'
# col_cast1                = 'H'
# col_cast2                = 'I'
# col_cast3                = 'J'
# col_production_countries = 'K'
# col_finance              = 'L'

# # Territory pairs now start at M–Z
# pair_cols = [
#     ('M','N'), ('O','P'), ('Q','R'), ('S','T'),
#     ('U','V'), ('W','X'), ('Y','Z') , ('AA','AB')
# ]

# HEADERS = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# MAX_RETRIES = 3
# RETRY_DELAY = 5

# def get_soup_and_html(url):
#     for attempt in range(MAX_RETRIES):
#         try:
#             r = requests.get(url, headers=HEADERS, timeout=10)
#             print(f"GET {url} → {r.status_code}")
#             r.raise_for_status()
#             return BeautifulSoup(r.text, 'html.parser'), r.text
#         except requests.exceptions.RequestException as e:
#             print(f"  ! {e} (retry {attempt+1}/{MAX_RETRIES})")
#             time.sleep(RETRY_DELAY)
#     print(f"Failed to fetch {url}")
#     return None, None

# def slugify(name):
#     s = re.sub(r"[^A-Za-z0-9 ]+", "", str(name))
#     s = re.sub(r"\s+", " ", s).strip()
#     return quote_plus(s.replace(" ", "-"))

# def extract_genre(soup):
#     td = soup.find('td', string=lambda t: t and 'Genre:' in t)
#     return td.find_next_sibling('td').get_text(strip=True) if td else ''

# def extract_production_countries(soup):
#     td = soup.find('td', string=lambda t: t and 'Production Countries:' in t)
#     if not td:
#         return ''
#     return td.find_next_sibling('td').get_text(";", strip=True)

# def extract_finance(soup):
#     tbl = soup.find('table', id='movie_finances')
#     if not tbl:
#         return ''
#     td = tbl.find('td', class_='data')
#     return td.get_text(strip=True) if td else ''

# def extract_director(soup):
#     hdr = soup.find('h1', string=re.compile(r'Production and Technical Credits'))
#     if not hdr:
#         return ''
#     tbl = hdr.find_next('table')
#     for tr in tbl.find_all('tr'):
#         cols = tr.find_all('td')
#         if len(cols) >= 3 and cols[2].get_text(strip=True).lower() == 'director':
#             return cols[0].get_text(strip=True)
#     return ''

# def extract_cast(soup):
#     div = soup.find('div', class_='cast_new')
#     names = [b.get_text(strip=True) for b in div.find_all('b')[:3]] if div else []
#     return (names + ['', '', ''])[:3]

# def extract_international_data(slug):
#     url = f"https://www.the-numbers.com/current/cont/graphs/movie/international-iframe/{slug}"
#     r = requests.get(url, headers=HEADERS, timeout=15)
#     if r.status_code != 200:
#         return []
#     js = r.text
#     pattern = re.compile(r"google\.visualization\.arrayToDataTable\(\s*\[(.*?)\]\s*\);", re.DOTALL)
#     m = pattern.search(js)
#     if not m:
#         return []
#     body = m.group(1).replace("\n","")
#     pairs = re.findall(r"\['([^']+)'\s*,\s*([\d\.]+)\s*\]", body)
#     # skip header row
#     out = []
#     for country, num in pairs[1:]:
#         try:
#             val = int(float(num))
#             rev = f"${val:,}"
#         except:
#             rev = ''
#         out.append((country, rev))
#     return out

# # ———————————————————————————————————————————————————————————————
# # MAIN
# # ———————————————————————————————————————————————————————————————
# wb = load_workbook(excel_path)
# ws = wb[sheet_name]

# for i in range(num_movies):
#     row = start_row + i
#     title = ws[f"{movie_col}{row}"].value
#     if not title:
#         continue

#     print(f"Row {row}: {title}")
#     slug = slugify(title)
#     base = f"https://www.the-numbers.com/movie/{slug}"

#     # Summary tab
#     soup_sum, _  = get_soup_and_html(base + "#tab=summary")
#     # Cast & Crew tab
#     soup_cast, _ = get_soup_and_html(base + "#tab=cast-and-crew")

#     if not soup_sum or not soup_cast:
#         print(f"  Skipping {title}: failed to fetch pages")
#         continue

#     # Extract fields
#     genre    = extract_genre(soup_sum)
#     prod_cty = extract_production_countries(soup_sum)
#     finance  = extract_finance(soup_sum)
#     director = extract_director(soup_cast)
#     cast1, cast2, cast3 = extract_cast(soup_cast)
#     intl_data = extract_international_data(slug)

#     # Write into Excel
#     ws[f"{col_genre}{row}"]                .value = genre
#     ws[f"{col_production_countries}{row}"].value = prod_cty
#     ws[f"{col_finance}{row}"]              .value = finance
#     ws[f"{col_director}{row}"]             .value = director
#     ws[f"{col_cast1}{row}"]                .value = cast1
#     ws[f"{col_cast2}{row}"]                .value = cast2
#     ws[f"{col_cast3}{row}"]                .value = cast3

#     # Territory pairs M–Z
#     for idx, (c_col, r_col) in enumerate(pair_cols):
#         if idx < len(intl_data):
#             country, revenue = intl_data[idx]
#         else:
#             country, revenue = '', ''
#         ws[f"{c_col}{row}"].value = country
#         ws[f"{r_col}{row}"].value = revenue

#     print(f"  → Wrote genre, prod countries, finance, director, cast, {len(intl_data)} territories")

# wb.save(excel_path)
# print("Done – columns F–Z updated, A–E preserved.")





import re
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote_plus, urlparse
from openpyxl import load_workbook

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
excel_path = 'Data_Capture.xlsx'
sheet_name = 'Data'
movie_col = 'B'
start_row = 7
num_movies = 50

# Output columns:
col_genre                = 'F'
col_director             = 'G'
col_cast1                = 'H'
col_cast2                = 'I'
col_cast3                = 'J'
col_production_countries = 'K'
col_finance              = 'L'
pair_cols = [
    ('M','N'), ('O','P'), ('Q','R'),
    ('S','T'), ('U','V'), ('W','X'), ('Y','Z') , ('AA','AB')
]

HEADERS = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds

def slugify(name):
    """Turn any value into a URL‐friendly slug (e.g. "Jurassic World" → "Jurassic-World")."""
    s = re.sub(r"[^A-Za-z0-9 ]+", "", str(name))
    s = re.sub(r"\s+", " ", s).strip()
    return quote_plus(s.replace(" ", "-"))

def get_soup_and_html(url):
    """
    Fetches `url` (with retries) and returns (BeautifulSoup, raw HTML string).
    If it fails after retries, returns (None, None).
    """
    for attempt in range(MAX_RETRIES):
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            print(f"GET {url} → {r.status_code}")
            r.raise_for_status()
            return BeautifulSoup(r.text, 'html.parser'), r.text
        except requests.exceptions.RequestException as e:
            print(f"  ! Error fetching {url}: {e} (retry {attempt+1}/{MAX_RETRIES})")
            time.sleep(RETRY_DELAY)
    print(f"Failed to fetch {url} after {MAX_RETRIES} retries.")
    return None, None

def extract_canonical_slug(soup):
    """
    Look for: <meta property="og:url" content="https://www.the-numbers.com/movie/Some-Slug">
    Returns "Some-Slug". If not found, returns None.
    """
    tag = soup.find('meta', attrs={'property': 'og:url'})
    if not tag or not tag.get('content'):
        return None
    content = tag['content']  
    parsed = urlparse(content)
    # split path by '/', take last segment
    slug = parsed.path.rstrip('/').split('/')[-1]
    return slug

def extract_genre(soup):
    td = soup.find('td', string=lambda t: t and 'Genre:' in t)
    return td.find_next_sibling('td').get_text(strip=True) if td else ''

def extract_production_countries(soup):
    td = soup.find('td', string=lambda t: t and 'Production Countries:' in t)
    if not td:
        return ''
    # join with semicolons if multiple <a> tags
    return td.find_next_sibling('td').get_text(";", strip=True)

def extract_finance(soup):
    tbl = soup.find('table', id='movie_finances')
    if not tbl:
        return ''
    td = tbl.find('td', class_='data')
    return td.get_text(strip=True) if td else ''

def extract_director(soup):
    hdr = soup.find('h1', string=re.compile(r'Production and Technical Credits'))
    if not hdr:
        return ''
    tbl = hdr.find_next('table')
    for tr in tbl.find_all('tr'):
        cols = tr.find_all('td')
        if len(cols) >= 3 and cols[2].get_text(strip=True).lower() == 'director':
            return cols[0].get_text(strip=True)
    return ''

def extract_cast(soup):
    div = soup.find('div', class_='cast_new')
    if not div:
        return ('','','')
    names = [b.get_text(strip=True) for b in div.find_all('b')[:3]]
    # pad to exactly three
    while len(names) < 3:
        names.append('')
    return tuple(names[:3])

def extract_international_data(correct_slug):
    """
    Hits the iframe endpoint:
    https://www.the-numbers.com/current/cont/graphs/movie/international-iframe/{correct_slug}
    Extracts the JS array of [ ['Region','Box Office'], ['Japan',197062163.00], ... ]
    and returns a list of (country, formatted_string) tuples, skipping the header row.
    """
    url = f"https://www.the-numbers.com/current/cont/graphs/movie/international-iframe/{correct_slug}"
    r = requests.get(url, headers=HEADERS, timeout=15)
    print(f"GET {url} → {r.status_code}")
    if r.status_code != 200:
        return []
    js = r.text
    pattern = re.compile(r"google\.visualization\.arrayToDataTable\(\s*\[(.*?)\]\s*\);", re.DOTALL)
    m = pattern.search(js)
    if not m:
        return []
    body = m.group(1).replace("\n","").replace(" ", "")
    pairs = re.findall(r"\['([^']+)'\s*,\s*([\d\.]+)\s*\]", body)
    # skip the first header row
    out = []
    for country, num in pairs[1:]:
        try:
            val = int(float(num))
            rev = f"${val:,}"
        except:
            rev = ''
        out.append((country, rev))
    return out

# ———————————————————————————————————————————————————————————————
# MAIN SCRIPT
# ———————————————————————————————————————————————————————————————
wb = load_workbook(excel_path)
ws = wb[sheet_name]

for i in range(num_movies):
    row = start_row + i
    title = ws[f"{movie_col}{row}"].value
    if not title:
        continue

    print(f"\nRow {row}: {title!r}")
    guessed_slug = slugify(title)
    summary_url = f"https://www.the-numbers.com/movie/{guessed_slug}#tab=summary"

    # 1) Fetch Summary tab to get canonical slug via <meta property="og:url">
    soup_sum, _ = get_soup_and_html(summary_url)
    if not soup_sum:
        print("  → Could not fetch Summary page; skipping.")
        continue

    correct_slug = extract_canonical_slug(soup_sum) or guessed_slug
    print(f"  → Correct slug: {correct_slug}")

    # 2) Re-build all URLs using correct_slug
    base = f"https://www.the-numbers.com/movie/{correct_slug}"
    summary_tab      = base + "#tab=summary"
    cast_crew_tab    = base + "#tab=cast-and-crew"

    # 3) Re-fetch Summary (to be safe) and Cast & Crew
    soup_sum, _  = get_soup_and_html(summary_tab)
    soup_cast, _ = get_soup_and_html(cast_crew_tab)
    if not soup_sum or not soup_cast:
        print("  → Could not fetch new tabs; skipping.")
        continue

    # 4) Extract data from Summary
    genre    = extract_genre(soup_sum)
    prod_cty = extract_production_countries(soup_sum)
    finance  = extract_finance(soup_sum)

    # 5) Extract data from Cast & Crew
    director = extract_director(soup_cast)
    cast1, cast2, cast3 = extract_cast(soup_cast)

    # 6) Extract International data using the corrected slug
    intl_data = extract_international_data(correct_slug)

    # 7) Write into Excel
    ws[f"{col_genre}{row}"]                .value = genre
    ws[f"{col_production_countries}{row}"] .value = prod_cty
    ws[f"{col_finance}{row}"]              .value = finance
    ws[f"{col_director}{row}"]             .value = director
    ws[f"{col_cast1}{row}"]                .value = cast1
    ws[f"{col_cast2}{row}"]                .value = cast2
    ws[f"{col_cast3}{row}"]                .value = cast3

    # 8) Write up to 7 (country, revenue) pairs into M–Z
    for idx, (c_col, r_col) in enumerate(pair_cols):
        if idx < len(intl_data):
            country, revenue = intl_data[idx]
        else:
            country, revenue = '', ''
        ws[f"{c_col}{row}"].value = country
        ws[f"{r_col}{row}"].value = revenue

    print(f"  → Wrote: Genre, Production Countries, Finance, Director, Cast, {len(intl_data)} territories.")

# 9) Save workbook
wb.save(excel_path)
print("\nDone – columns F–Z updated, A–E preserved.")
