"""Download the official timetable workbook from a public SharePoint/OneDrive sharing link.

usage: python scripts/fetch_xlsx.py SHARE_URL out.xlsx

SharePoint answers a sharing link with a redirect chain that sets a guest-access cookie
before it hands out the file, so we keep cookies across redirects (like a browser) and
try two known download URLs. Prints what each attempt returned, so a failure is easy to read.
"""
import sys, io, re, zipfile, http.cookiejar, urllib.request, urllib.parse

UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36')


def candidates(url):
    url = url.strip()
    sep = '&' if '?' in url else '?'
    yield 'link + download=1', url + sep + 'download=1'
    # https://TENANT-my.sharepoint.com/:x:/g/personal/USER/TOKEN?e=xx
    #   -> https://TENANT-my.sharepoint.com/personal/USER/_layouts/15/download.aspx?share=TOKEN
    m = re.match(r'(https://[^/]+)/:[a-z]:/[a-z]/personal/([^/]+)/([^/?#]+)', url)
    if m:
        yield 'download.aspx?share=', f'{m.group(1)}/personal/{m.group(2)}/_layouts/15/download.aspx?share={m.group(3)}'


def is_workbook(data):
    if not data.startswith(b'PK'):
        return False
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        return 'xl/workbook.xml' in z.namelist() and z.testzip() is None
    except Exception:
        return False


def main():
    share, out = sys.argv[1], sys.argv[2]
    for label, url in candidates(share):
        jar = http.cookiejar.CookieJar()
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
        opener.addheaders = [('User-Agent', UA), ('Accept', '*/*'), ('Accept-Language', 'ro,en;q=0.8')]
        try:
            # first visit the plain link so SharePoint can set its guest cookie, then download
            try:
                opener.open(share, timeout=60).read()
            except Exception:
                pass
            with opener.open(url, timeout=60) as r:
                data = r.read()
                info = f'HTTP {r.status} | {r.headers.get("Content-Type")} | {len(data)} bytes | {r.geturl()[:120]}'
        except Exception as e:
            print(f'[{label}] eroare: {e}')
            continue
        print(f'[{label}] {info}')
        if is_workbook(data):
            open(out, 'wb').write(data)
            print(f'OK: am descărcat Excelul ({len(data)} bytes)')
            return 0
        snippet = re.sub(r'\s+', ' ', data[:300].decode('utf-8', 'replace'))
        print(f'[{label}] nu e un fișier Excel. Început: {snippet}')
    print('::error::SharePoint nu a dat un fișier Excel. Urcă manual fișierul în input/orar.xlsx.')
    return 1


if __name__ == '__main__':
    sys.exit(main())
