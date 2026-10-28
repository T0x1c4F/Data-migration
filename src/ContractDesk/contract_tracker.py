"""Certificate-level tracker. Source workbooks are always opened read-only."""
import collections
from copy import deepcopy
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
import unicodedata
import uuid

import openpyxl
from formatted_export import write_formatted_export
from employee_names import canonical_employee

def plain(value):
    return unicodedata.normalize('NFD', str(value or '').lower().replace('đ', 'd')).encode('ascii', 'ignore').decode().strip()

def text(value):
    if value is None:
        return ''
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime('%d/%m/%Y')
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def date(value):
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime('%d/%m/%Y')
    value = text(value)
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d.%m.%Y', '%Y-%m-%d'):
        try:
            return dt.datetime.strptime(value, fmt).strftime('%d/%m/%Y')
        except ValueError:
            pass
    return ''

def lines(value):
    return [s.strip() for s in re.split(r'[\n;]+', text(value)) if s.strip()]

def contract_key(number, year):
    n = text(number)
    return f'{int(n) if n.isdigit() else plain(n)}|{year}'

def red(color):
    if color is None:
        return False
    if color.type == 'indexed':
        from openpyxl.styles.colors import COLOR_INDEXED
        if color.indexed >= len(COLOR_INDEXED):return False
        rgb = COLOR_INDEXED[color.indexed][-6:]
    elif color.type == 'rgb':
        rgb = color.rgb[-6:]
    else:return False
    try:
        r, g, b = (int(rgb[i:i+2], 16) for i in (0, 2, 4))
        return r >= 150 and r > g * 1.5 and r > b * 1.5
    except (ValueError, TypeError):
        return False

def cancellation(values):
    # Negations and customer explanations remain human decisions.
    for value in values:
        s = plain(value)
        if re.search(r'\b(huy|cancelled|canceled|cancel)\b', s):
            if re.search(r'\b(khong|chua|not|never|dung)\b', s):
                return 'review', 'Ghi chú có từ hủy kèm phủ định: ' + text(value)
            return 'cancel', 'Ghi chú hủy: ' + text(value)
    return '', ''

FIELDS = {
    'contract': 'so hop dong', 'contract_date': 'ngay hop dong',
    'client': 'khach hang', 'content': 'noi dung tham dinh',
    'fee': 'phi tham dinh', 'certificate': 'so chung thu / cvtt',
    'certificate_date': 'ngay chung thu', 'employee': 'cv thuc hien',
}

def read_file(path):
    book = openpyxl.load_workbook(path, data_only=True, rich_text=True)
    result = []
    try:
        sheet = book.worksheets[0]
        rows = iter(sheet.iter_rows())
        headers = next(rows)
        lookup = {plain(c.value): i for i, c in enumerate(headers) if c.value is not None}
        mapping = {key: lookup.get(label) for key, label in FIELDS.items()}
        if any(mapping[k] is None for k in ('contract', 'client', 'certificate')):
            raise ValueError('Không tìm thấy các cột SỐ HỢP ĐỒNG / KHÁCH HÀNG / SỐ CHỨNG THƯ')
        note_cols = [i for h, i in lookup.items() if 'ghi chu' in h or 'tinh trang' in h or 'ly do' in h]
        red_rules = []
        for group in sheet.conditional_formatting:
            if any(rule.dxf and rule.dxf.font and red(rule.dxf.font.color) for rule in sheet.conditional_formatting[group]):
                red_rules.append(group.sqref)
        for row in rows:
            values = {k: row[i].value if i is not None else None for k, i in mapping.items()}
            if not text(values['contract']):
                continue
            warnings = []
            missing = []
            for k, label in (('contract_date','Ngày hợp đồng'), ('client','Khách hàng'), ('content','Nội dung'), ('fee','Phí thẩm định'), ('employee','CV thực hiện'), ('certificate','Số chứng thư'), ('certificate_date','Ngày chứng thư')):
                if not text(values[k]):
                    missing.append(label)
            cd = date(values['contract_date'])
            if text(values['contract_date']) and not cd:
                warnings.append('Ngày hợp đồng không hợp lệ')
            fee = text(values['fee'])
            if fee and (not isinstance(values['fee'], (int, float)) or values['fee'] <= 0):
                warnings.append('Kiểm tra phí: phải là số lớn hơn 0')
            notes = [row[i].value for i in note_cols if row[i].value is not None]
            cancel, reason = cancellation([c.value for c in row if c.value is not None])
            if notes and not cancel:
                warnings.append('Có ghi chú/tình trạng cần kiểm tra: ' + '; '.join(text(v) for v in notes))
            for cell in row:
                if cell.value is None:
                    continue
                rich_red = isinstance(cell.value, openpyxl.cell.rich_text.CellRichText) and any(red(getattr(getattr(part, 'font', None), 'color', None)) for part in cell.value)
                if red(cell.font.color) or rich_red:
                    warnings.append('Chữ đỏ tại ' + cell.coordinate)
                elif cell.font.color and cell.font.color.type == 'theme' and cell.font.color.theme >= 4:
                    warnings.append('Chữ màu theo theme tại ' + cell.coordinate + ' (cần kiểm tra)')
                if any(cell.coordinate in ranges for ranges in red_rules):
                    warnings.append('Có quy tắc chữ đỏ tại ' + cell.coordinate + ' (cần kiểm tra Excel)')
            certs = lines(values['certificate']) or ['']
            dates = lines(values['certificate_date'])
            if len(certs) != len(dates):
                warnings.append('Số chứng thư và số ngày không khớp; không tự ghép ngày')
            for pos, cert in enumerate(certs):
                raw_date = dates[pos] if len(certs) == len(dates) else ''
                cert_date = date(raw_date)
                issues = warnings.copy()
                if raw_date and not cert_date:
                    issues.append('Ngày chứng thư không hợp lệ: ' + raw_date)
                year = cd[-4:] if cd else 'UNKNOWN'
                ck = contract_key(values['contract'], year)
                # Certificate identifiers, including year, are the stable identity.
                key = plain(cert).replace(' ', '') if cert else 'missing:' + ck
                data = {k: text(v) for k, v in values.items()}
                data.update(contract_date=cd, certificate=cert, certificate_date=cert_date)
                fingerprint = hashlib.sha256(json.dumps([data, cancel, reason, issues, [text(n) for n in notes]], sort_keys=True, ensure_ascii=False).encode()).hexdigest()
                result.append(dict(key=key, contract_key=ck, data=data, fingerprint=fingerprint,
                    source=str(Path(path).resolve()), row=row[0].row, cancel=cancel,
                    reason=reason, missing=missing.copy(), warnings=issues))
    finally:
        book.close()
    return result

WEBSITE_REFERENCE_DATE = ''
WEBSITE_REFERENCE_COUNT = 0
STATUS = {'ready':'Sẵn sàng tải lên','missing':'Thiếu thông tin','cancel':'Đã hủy',
          'review':'Cần kiểm tra','uploaded':'Đã tải lên','changed':'Đã tải lên • có thay đổi',
          'absent':'Đã lưu trữ','unknown':'Đã tải lên'}
STATUS_COLORS = {'cancel':'#fee2e2', 'ready':'#dcfce7', 'missing':'#f8cccc',
                 'review':'#fff2cc', 'unknown':'#ffffff', 'uploaded':'#ffffff', 'changed':'#ffedd5'}

class Store:
    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.execute('CREATE TABLE IF NOT EXISTS records (key TEXT PRIMARY KEY, payload TEXT, fingerprint TEXT, uploaded_fingerprint TEXT, url TEXT, decision TEXT, decision_fingerprint TEXT, seen TEXT)')
        self.db.execute('CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)')
        self.db.execute('CREATE TABLE IF NOT EXISTS events (time TEXT, key TEXT, action TEXT, detail TEXT)')
        self.db.execute('CREATE TABLE IF NOT EXISTS archive (key TEXT PRIMARY KEY,time TEXT,reason TEXT)')
        self.db.execute('CREATE TABLE IF NOT EXISTS record_history (id INTEGER PRIMARY KEY,time TEXT,key TEXT,action TEXT,payload TEXT)')
        with self.db:
            last=self.setting('last_scan')
            if last and not json.loads(self.setting('scan_failures','[]')):
                for key,payload in self.db.execute('SELECT key,payload FROM records WHERE seen!=?',(last,)).fetchall():
                    if not self.db.execute('SELECT 1 FROM archive WHERE key=?',(key,)).fetchone():
                        self._archive(key,json.loads(payload),'Không có trong lần quét thành công gần nhất.')

    def _remember(self,key,action,record,**detail):
        now=dt.datetime.now().isoformat()
        metadata=self.db.execute('SELECT uploaded_fingerprint,url,decision,decision_fingerprint FROM records WHERE key=?',(key,)).fetchone()
        snapshot=dict(record=record,tracking=metadata,detail=detail)
        self.db.execute('INSERT INTO record_history(time,key,action,payload) VALUES(?,?,?,?)',(now,key,action,json.dumps(snapshot,ensure_ascii=False)))
        self.db.execute('INSERT INTO events VALUES(?,?,?,?)',(now,key,action,json.dumps(detail,ensure_ascii=False)))

    def _archive(self,key,record,reason):
        self.db.execute('INSERT INTO archive VALUES(?,?,?)',(key,dt.datetime.now().isoformat(),reason))
        self._remember(key,'archived',record,reason=reason)

    def history_log(self,limit=500):
        return [dict(time=t,key=k,action=a,detail=d) for t,k,a,d in self.db.execute('SELECT time,key,action,detail FROM events ORDER BY rowid DESC LIMIT ?',(limit,))]

    def setting(self, key, default=''):
        r = self.db.execute('SELECT value FROM settings WHERE key=?',(key,)).fetchone()
        return r[0] if r else default

    def save_setting(self, key, value):
        self.db.execute('INSERT OR REPLACE INTO settings VALUES (?,?)',(key,value)); self.db.commit()

    def scan(self, paths):
        all_records, failures = [], []
        for path in paths:
            try:
                all_records.extend(read_file(path))
            except Exception as e:
                failures.append(f'{Path(path).name}: {e}')
        return self.scan_records(all_records, failures)

    def scan_records(self, all_records, failures):
        if failures:
            self.save_setting('scan_failures',json.dumps(failures,ensure_ascii=False))
            return failures
        with self.db:
            return self._save_scan(all_records)

    def _save_scan(self,all_records):
        grouped = collections.defaultdict(list)
        for r in all_records:
            grouped[r['key']].append(r)
        now = dt.datetime.now().isoformat(timespec='microseconds')
        scan_id = uuid.uuid4().hex
        stats=dict(archived=0,restored=0,moved=0)
        for key, group in grouped.items():
            r = deepcopy(group[0])
            r['sources'] = [dict(source=g['source'], row=g['row'], fingerprint=g['fingerprint']) for g in group]
            if len(group) > 1:
                r['warnings'].append('Trùng chứng thư trong nhiều dòng/tệp; cần kiểm tra')
                if len({g['fingerprint'] for g in group}) > 1:
                    r['warnings'].append('Dữ liệu giữa các bản không giống nhau')
                if any(g['cancel'] for g in group):
                    r['warnings'].append('Một bản có dấu hiệu hủy')
            prev = self.db.execute('SELECT fingerprint,payload FROM records WHERE key=?',(key,)).fetchone()
            if prev:
                old=json.loads(prev[1])
                old_paths={s['source'] for s in old.get('sources') or [old]}
                new_paths={s['source'] for s in r['sources']}
                if old_paths!=new_paths:
                    self._remember(key,'source_moved',old,new_sources=sorted(new_paths));stats['moved']+=1
                if self.db.execute('SELECT 1 FROM archive WHERE key=?',(key,)).fetchone():
                    self._remember(key,'restored',old,new_sources=sorted(new_paths));stats['restored']+=1
                if prev[0]!=r['fingerprint']:self._remember(key,'content_changed',old)
            r['change'] = 'Mới' if not prev else ('Có thay đổi' if prev[0] != r['fingerprint'] else 'Không đổi')
            if not prev or prev[0] != r['fingerprint']:
                self.db.execute('INSERT INTO events VALUES (?,?,?,?)',(now,key,'scan',r['change']))
            self.db.execute('INSERT INTO records(key,payload,fingerprint,seen) VALUES(?,?,?,?) ON CONFLICT(key) DO UPDATE SET payload=excluded.payload,fingerprint=excluded.fingerprint,seen=excluded.seen',
                (key,json.dumps(r,ensure_ascii=False),r['fingerprint'],scan_id))
            self.db.execute('DELETE FROM archive WHERE key=?',(key,))
        for key,payload in self.db.execute('SELECT key,payload FROM records WHERE seen!=?',(scan_id,)).fetchall():
            if not self.db.execute('SELECT 1 FROM archive WHERE key=?',(key,)).fetchone():
                self._archive(key,json.loads(payload),'Không xuất hiện trong lần quét thành công; giữ lịch sử để tự khôi phục khi có lại trong nguồn.')
                stats['archived']+=1
        self.db.executemany('INSERT OR REPLACE INTO settings VALUES(?,?)',[('last_scan',scan_id),('scan_failures','[]')])
        self.last_scan_stats=stats
        return []

    def records(self,include_archived=False):
        output = []
        baseline = {contract_key(n,y):post for n,y,post in BASELINE}
        by_number = {contract_key(n,y).split('|')[0]:post for n,y,post in BASELINE}
        for key, payload, fp, uploaded, url, decision, decision_fp, seen in self.db.execute('SELECT * FROM records'):
            archived=self.db.execute('SELECT time,reason FROM archive WHERE key=?',(key,)).fetchone()
            current=seen==self.setting('last_scan') and not archived
            if not include_archived and not current:continue
            r = json.loads(payload)
            r.update(url=url or '', decision=decision or '')
            exact_match = baseline.get(r['contract_key'])
            baseline_post = exact_match or by_number.get(r['contract_key'].split('|')[0])
            if plain(r['data']['client']) == 'cong ty tnhh dai thanh tin nghia':baseline_post=743
            allow_new = decision == 'new' and decision_fp == fp and self.setting('new_web_reference:' + key) == str(baseline_post)
            why = []
            if uploaded:
                state = 'uploaded' if fp == uploaded else 'changed'
                if r['cancel'] or decision == 'cancel':
                    state = 'changed'; why.append(r['reason'] or 'Đã đánh dấu hủy sau khi tải lên')
            elif r['cancel'] == 'cancel' or decision == 'cancel':
                state = 'cancel'; why.append(r['reason'] or 'Đánh dấu hủy thủ công')
            elif not current:
                state = 'absent'; why.append('Đã lưu trữ; không xuất hiện trong lần quét thành công gần nhất')
            elif r['missing']:
                state = 'missing'; why.extend(r['missing'])
            elif baseline_post and not allow_new:
                state = 'unknown' if exact_match or baseline_post == 743 else 'review'
                r['url'] = f"https://hoso.Contract Desk.vn/wp-admin/post.php?post={baseline_post}&action=edit"
                why.append('Hợp đồng đã có trên web; xác nhận chứng thư trước khi tải thêm' if state == 'unknown' else 'Số hợp đồng đã có trên web nhưng năm khác. Cần đối chiếu trước khi tải lên.')
            elif r['warnings'] or r['cancel'] == 'review':
                state = 'review'; why.extend(r['warnings']); why.append(r['reason'])
                if decision in ('approve','new') and decision_fp == fp:
                    # Explicit review cannot bypass missing/invalid required data.
                    hard = any('không hợp lệ' in w or 'không khớp' in w or 'Trùng' in w or 'phải là số' in w for w in r['warnings'])
                    if not hard:
                        state = 'ready'
            else:
                state = 'ready'
            r.update(state=state, detail='; '.join(w for w in why if w), status=STATUS[state], current=bool(current), employee_name=canonical_employee(r['data']['employee']),archived=not current,archive_time=archived[0] if archived else '',archive_reason=archived[1] if archived else '')
            output.append(r)
        return output

    def decide(self, key, action, url=''):
        row = self.db.execute('SELECT fingerprint,payload FROM records WHERE key=?',(key,)).fetchone()
        if not row:
            return
        if action == 'reset':
            self.db.execute('UPDATE records SET uploaded_fingerprint=NULL,url=NULL,decision=NULL,decision_fingerprint=NULL WHERE key=?',(key,))
        elif action == 'uploaded':
            self.db.execute('UPDATE records SET uploaded_fingerprint=?,url=? WHERE key=?',(row[0],url,key))
        else:
            self.db.execute('UPDATE records SET decision=?,decision_fingerprint=? WHERE key=?',(action,row[0],key))
        if action == 'new':
            record = json.loads(row[1])
            baseline = {contract_key(n,y):post for n,y,post in BASELINE}
            by_number = {contract_key(n,y).split('|')[0]:post for n,y,post in BASELINE}
            post = baseline.get(record['contract_key']) or by_number.get(record['contract_key'].split('|')[0])
            if plain(record['data']['client']) == 'cong ty tnhh dai thanh tin nghia':post=743
            self.db.execute('INSERT OR REPLACE INTO settings VALUES (?,?)',('new_web_reference:' + key,str(post)))
        self.db.execute('INSERT INTO events VALUES(?,?,?,?)',(dt.datetime.now().isoformat(),key,action,url))
        self.db.commit()

    def export(self, keys, path, mode='ready'):
        """Export one original design per workbook; multiple sources use a folder."""
        if mode not in ('ready', 'not_ready'):
            raise ValueError('Kiểu xuất không hợp lệ.')
        records = [r for r in self.records() if r['key'] in keys]
        if not records or len(records) != len(set(keys)):
            raise ValueError('Không tìm thấy đầy đủ các mục đã chọn. Quét lại trước khi xuất.')
        if any(not r['current'] for r in records):
            raise ValueError('Mục không có trong lần quét hiện tại. Quét lại trước khi xuất.')
        if mode == 'ready' and any(r['state'] != 'ready' for r in records):
            raise ValueError('Chỉ xuất chứng thư có trạng thái Sẵn sàng tải lên.')
        if mode == 'not_ready' and any(r['state'] not in ('review', 'missing') for r in records):
            raise ValueError('Chỉ xuất chứng thư Cần kiểm tra hoặc Thiếu thông tin trong danh sách chưa sẵn sàng.')
        latest = {}
        grouped = collections.defaultdict(list)
        for r in records:
            sources = r.get('sources') or [dict(source=r['source'], row=r['row'], fingerprint=r['fingerprint'])]
            for item in sources:
                source = item['source']
                if source not in latest:
                    latest[source] = {(n['key'], n['row']): n for n in read_file(source)}
                current = latest[source].get((r['key'], item['row']))
                expected = item.get('fingerprint') or (r['fingerprint'] if source == r['source'] and item['row'] == r['row'] else None)
                if not current or expected != current['fingerprint']:
                    raise ValueError('Dữ liệu nguồn đã thay đổi. Quét lại trước khi xuất.')
                detail = r['detail']
                if mode == 'not_ready':
                    local_details = ([('Thiếu: ' + ', '.join(current['missing']))] if current['missing'] else []) + current['warnings'] + ([current['reason']] if current['reason'] else [])
                    detail = '; '.join(dict.fromkeys([part for part in [detail] + local_details if part]))
                current.update(state=r['state'], status=r['status'], detail=detail)
                grouped[source].append(current)
        path = Path(path)
        if len(grouped) == 1:
            if path.suffix.lower() != '.xlsx':
                raise ValueError('Chọn tên tệp kết thúc bằng .xlsx.')
            targets = {next(iter(grouped)): path}
            staging_parent = path.parent
        else:
            if path.suffix.lower() == '.xlsx':
                raise ValueError('Nhiều tệp gốc: chọn thư mục để xuất riêng từng tệp, giữ thiết kế gốc.')
            path.mkdir(parents=True, exist_ok=True)
            staging_parent = path
            stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            prefix = 'SanSang' if mode == 'ready' else 'ChuaSanSang'
            targets = {source: path / f'{prefix}_{Path(source).stem[:100]}_{stamp}_{i}.xlsx' for i,source in enumerate(sorted(grouped),1)}
        protected = {str(Path(r['source']).resolve()).casefold() for r in self.records()}
        protected.update(str(Path(s['source']).resolve()).casefold() for r in self.records() for s in r.get('sources', []))
        if any(str(target.resolve()).casefold() in protected for target in targets.values()):
            raise ValueError('Không ghi đè tệp Excel gốc. Chọn tên hoặc thư mục xuất khác.')
        # Stage every workbook before committing any destination file.
        with tempfile.TemporaryDirectory(prefix='.hoso-export-', dir=staging_parent) as tmp:
            staged = []
            for i,(source,rows) in enumerate(sorted(grouped.items())):
                temp = Path(tmp) / f'{i}.xlsx'
                write_formatted_export(rows, temp, include_status=mode == 'not_ready')
                staged.append((temp, targets[source]))
            for temp,target in staged:
                os.replace(temp, target)
        action = 'export' if mode == 'ready' else 'export_not_ready'
        self.db.execute('INSERT INTO events VALUES(?,?,?,?)',(dt.datetime.now().isoformat(),'',action,json.dumps(dict(keys=sorted(keys), files=[str(p) for p in targets.values()]),ensure_ascii=False)))
        self.db.commit()
        return sum(len(rows) for rows in grouped.values())

def app():
    from tracker_ui import run_app
    run_app(Store, read_file, STATUS, STATUS_COLORS,
            reference_date=WEBSITE_REFERENCE_DATE, reference_count=WEBSITE_REFERENCE_COUNT)

if __name__=='__main__':
    try:app()
    except Exception as exc:
        import tkinter.messagebox
        tkinter.messagebox.showerror('Contract Desk Tracker — lỗi',str(exc))
        raise
