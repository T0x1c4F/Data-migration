"""WordPress uploader; every publish has a durable before/after journal entry."""
from decimal import Decimal, InvalidOperation
import json
import os
from pathlib import Path
import re
import sys
from urllib.parse import urlencode, urlparse, parse_qs

from contract_tracker import date, plain, contract_key
from employee_names import canonical_employee

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parents[2] / '.env', override=False)
except ImportError:
    pass

ORIGIN = os.environ.get('CONTRACT_DESK_SITE_URL', 'https://example.invalid').rstrip('/')
LOGIN = ORIGIN + '/wp-login.php'
NEW = ORIGIN + '/wp-admin/post-new.php?post_type=appraisal_case'


def money(value):
    try:
        amount = Decimal(str(value).strip())
    except InvalidOperation:
        raise ValueError('Phí thẩm định phải là số trong Excel.')
    if not amount.is_finite() or amount <= 0 or amount != amount.to_integral_value():
        raise ValueError('Phí thẩm định phải là số nguyên dương.')
    return format(amount, '.0f')


def field(page, label):
    labels = page.locator('.acf-label > label').filter(has_text=re.compile('^' + re.escape(label) + r'\s*\*?\s*$'))
    if labels.count() != 1:
        raise ValueError('Không tìm thấy duy nhất trường: ' + label)
    return labels.locator('..').locator('..')


def text_input(page, label):
    return field(page, label).locator('input:not([type="hidden"]), textarea').first


def set_date(page, label, value):
    display = date(value)
    if not display:
        raise ValueError('Ngày không hợp lệ: ' + label)
    day, month, year = map(int, display.split('/'))
    box = field(page, label)
    if box.count() != 1:
        raise ValueError('Không tìm thấy trường: ' + label)
    box.evaluate('''(el, v) => {
      const input=el.querySelector('input:not([type="hidden"])');
      if(!input) throw Error('Không có ô ngày');
      const $=window.jQuery;
      if($ && typeof $(input).datepicker==='function') {
        $(input).datepicker('setDate',new Date(v.year,v.month-1,v.day));
        $(input).datepicker('hide');
      }
      input.value=v.display;
      const hidden=el.querySelector('input[type="hidden"]');
      if(hidden) {hidden.value=v.ymd;hidden.dispatchEvent(new Event('change',{bubbles:true}));}
      input.dispatchEvent(new Event('input',{bubbles:true}));
      input.dispatchEvent(new Event('change',{bubbles:true}));
    }''', dict(day=day, month=month, year=year, display=display, ymd=f'{year:04}{month:02}{day:02}'))


def read_date(page, label):
    box = field(page, label)
    hidden = box.locator('input[type="hidden"]').first
    if hidden.count():
        value = hidden.input_value()
        if re.fullmatch(r'\d{8}', value):
            return date(f'{value[6:8]}/{value[4:6]}/{value[:4]}')
    return date(box.locator('input:not([type="hidden"])').first.input_value())


def selected_employee(page):
    names = field(page, 'Nhân viên phụ trách').locator('select option:checked').all_text_contents()
    return [re.sub(r'\s*\([^)]*\)\s*$', '', n).strip() for n in names]


def assign_employee(page, name):
    box = field(page, 'Nhân viên phụ trách')
    # Existing site uses Select2's searchable employee dropdown.
    box.locator('.select2-selection, .select2-choice').first.click()
    search = page.locator('input.select2-search__field:visible, .select2-drop input:visible').first
    search.fill(name)
    results = page.locator('.select2-results__option:not(.select2-results__message), .select2-results li:not(.select2-disabled)')
    exact = results.filter(has_text=re.compile(r'^\s*' + re.escape(name) + r'(?:\s*\([^)]*\))?\s*$'))
    exact.first.wait_for(state='visible', timeout=12000)
    if exact.count() != 1:
        raise ValueError('Nhân viên không khớp duy nhất trên website: ' + name)
    exact.click()
    if [plain(n) for n in selected_employee(page)] != [plain(name)]:
        raise ValueError('Website chưa xác nhận đúng nhân viên: ' + name)


def fill_record(page, record):
    d = record['data']
    title = f"[{d['contract']}/{d['contract_date'][-2:]}] {d['client']}"
    page.goto(NEW)
    page.locator('#title').wait_for()
    page.fill('#title', title)
    text_input(page, 'Tên khách hàng').fill(d['client'])
    text_input(page, 'Nội dung công việc').fill(d['content'])
    assign_employee(page, canonical_employee(d['employee']))
    set_date(page, 'Ngày giao việc', d['contract_date'])
    set_date(page, 'Deadline', d['certificate_date'])
    field(page, 'Trạng thái hồ sơ').locator('select').select_option(label='Đã phát hành')
    text_input(page, 'Thanh toán đợt 1').fill(money(d['fee']))
    field(page, 'Trạng thái thanh toán đợt 1').locator('select').select_option(label='Chưa thanh toán')
    field(page, 'Trạng thái hóa đơn đợt 1').locator('select').select_option(label='Chưa Xuất hóa đơn')
    validate_fields(page, record, title)
    return title


def validate_fields(page, record, title):
    d = record['data']
    if page.locator('#title').input_value().strip() != title:
        raise ValueError('Tên hồ sơ không khớp.')
    for label, value in [('Tên khách hàng', d['client']), ('Nội dung công việc', d['content'])]:
        if text_input(page, label).input_value().strip() != value.strip():
            raise ValueError('Website chưa lưu đúng ' + label)
    for label, value in [('Ngày giao việc', d['contract_date']), ('Deadline', d['certificate_date'])]:
        if read_date(page, label) != date(value):
            raise ValueError('Website chưa lưu đúng ' + label)
    actual_money = re.sub(r'[, .\u00a0]', '', text_input(page, 'Thanh toán đợt 1').input_value())
    if actual_money != money(d['fee']):
        raise ValueError('Phí trên website không khớp Excel.')
    if [plain(n) for n in selected_employee(page)] != [plain(canonical_employee(d['employee']))]:
        raise ValueError('Nhân viên trên website không khớp.')
    for label, expected_value in [('Trạng thái hồ sơ','Đã phát hành'), ('Trạng thái thanh toán đợt 1','Chưa thanh toán'), ('Trạng thái hóa đơn đợt 1','Chưa Xuất hóa đơn')]:
        value = field(page, label).locator('select option:checked').all_text_contents()
        if [v.strip() for v in value] != [expected_value]:
            raise ValueError('Website chưa lưu đúng ' + label)


def title_contract_key(title):
    patterns = [r'\[\s*([^/\.\]]+)\s*[/\.]\s*(\d{2,4})\s*\]',
                r'\[\s*([^\]]+)\s*\][\s\.]*(\d{4})\b']
    for pattern in patterns:
        match = re.search(pattern, title)
        if match:
            year = int(match[2]); year = year + 2000 if year < 100 else year
            return contract_key(match[1].strip(), year)
    return ''


def existing_posts(page, record):
    """Live duplicate check. A match is reviewed, never silently equated to a certificate."""
    d = record['data']
    query = d['contract'].lstrip('0') or d['contract']
    page.goto(ORIGIN + '/wp-admin/edit.php?' + urlencode(dict(post_type='appraisal_case', s=query)))
    page.locator('#the-list').wait_for(state='attached')
    candidates = []
    seen_pages = set()
    while page.url not in seen_pages:
        seen_pages.add(page.url)
        if '/wp-admin/edit.php' not in page.url:
            raise ValueError('Phiên đăng nhập đã hết hạn. Đăng nhập lại trước khi tải lên.')
        errors = page.locator('.notice-error, #error-page')
        if errors.count():
            raise ValueError('Website không trả về danh sách hồ sơ để kiểm tra trùng.')
        for link in page.locator('#the-list .row-title').all():
            title = link.inner_text()
            if title_contract_key(title) == record['contract_key']:
                candidates.append(link.get_attribute('href'))
        next_page = page.locator('.tablenav.top a.next-page:not(.disabled)')
        if not next_page.count():
            break
        href = next_page.first.get_attribute('href')
        if not href or href in seen_pages:
            break
        page.goto(href)
    matches = []
    for url in dict.fromkeys(candidates):
        parsed = urlparse(url)
        if parsed.scheme != 'https' or parsed.netloc != urlparse(ORIGIN).netloc:
            raise ValueError('Liên kết website không hợp lệ.')
        post = parse_qs(parsed.query).get('post', [''])[0]
        if not post.isdigit():
            raise ValueError('Không đọc được mã hồ sơ website.')
        page.goto(url)
        page.locator('#title').wait_for()
        if read_date(page, 'Deadline') == d['certificate_date']:
            matches.append(dict(post=post, url=ORIGIN + f'/wp-admin/post.php?post={post}&action=edit'))
    return matches


def verify_saved(page, record, title):
    parsed = urlparse(page.url)
    post = parse_qs(parsed.query).get('post', [''])[0]
    if parsed.netloc != urlparse(ORIGIN).netloc or parsed.path != '/wp-admin/post.php' or not post.isdigit():
        raise ValueError('Website chưa trả về mã hồ sơ đã lưu.')
    url = ORIGIN + f'/wp-admin/post.php?post={post}&action=edit'
    page.goto(url)  # Re-read server data; do not trust the pre-submit form.
    page.locator('#title').wait_for()
    status = page.locator('#original_post_status')
    if not status.count() or status.input_value() != 'publish':
        raise ValueError('Website chưa xác nhận hồ sơ đã Publish.')
    validate_fields(page, record, title)
    return url


def upload_with_page(store, page, keys, stop, notify):
    results = []
    for index, key in enumerate(keys, 1):
        if stop.is_set():
            break
        record = None
        submitted = False
        try:
            record = store.reserve(key)
            notify('progress', f"{index}/{len(keys)} · {record['data']['contract']} · {record['data']['certificate']}")
            posts = existing_posts(page, record)
            approved = store.approved_posts(record)
            unapproved = [p for p in posts if p['post'] not in approved]
            if unapproved:
                store.block_existing(record, posts)
                results.append(dict(key=key, state='blocked', detail='Có hồ sơ tương tự trên website; cần đối chiếu.'))
                notify('refresh', None)
                continue
            title = fill_record(page, record)
            if stop.is_set():
                store.result(record, 'failed', detail='Đã dừng trước khi gửi.')
                break
            publish = page.locator('#publish')
            if publish.count() != 1:
                raise ValueError('Không tìm thấy nút Publish.')
            page.locator('#submitdiv').evaluate("el=>{el.classList.remove('closed');const i=el.querySelector('.inside');if(i)i.style.display='block';}")
            store.submitting(record)
            submitted = True
            publish.click()
            page.wait_for_url('**/post.php?post=**', timeout=20000)
            url = verify_saved(page, record, title)
            store.result(record, 'success', url, 'Website đã xác nhận Publish và dữ liệu đã lưu.')
            results.append(dict(key=key, state='success', url=url, detail='Đã tải lên'))
        except Exception as error:
            state = 'uncertain' if submitted else 'failed'
            detail = str(error).split('Call log:')[0][:1500]
            if record:
                store.result(record, state, page.url if submitted else '', detail)
            results.append(dict(key=key, state=state, detail=detail))
        notify('refresh', None)
    return results


def run_upload(db_path, keys, username, password, browser_folder, stop, notify):
    from desk_store import DeskStore
    from playwright.sync_api import sync_playwright
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(browser_folder)
    store = DeskStore(db_path)
    browser = None
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=False, slow_mo=50)
            try:
                page = browser.new_page()
                page.set_default_timeout(15000)
                notify('progress', 'Đang đăng nhập website…')
                page.goto(LOGIN)
                page.fill('#user_login', username)
                page.fill('#user_pass', password)
                page.click('#wp-submit')
                page.wait_for_url('**/wp-admin/**', timeout=20000)
                if stop.is_set():
                    return []
                return upload_with_page(store, page, keys, stop, notify)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass  # Browser cleanup cannot discard the committed batch results.
    finally:
        store.db.close()
