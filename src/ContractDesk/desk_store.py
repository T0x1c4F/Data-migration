"""Durable per-certificate upload journal shared by the UI and upload worker."""
import datetime as dt
from decimal import Decimal, InvalidOperation
import json
from pathlib import Path
import uuid

from contract_tracker import Store, read_file
from employee_names import canonical_employee, employee_aliases


class DeskStore(Store):
    def __init__(self, path):
        super().__init__(path)
        self.db.execute('PRAGMA busy_timeout=10000')
        self.db.execute('CREATE TABLE IF NOT EXISTS uploads (key TEXT PRIMARY KEY, fingerprint TEXT, state TEXT, url TEXT, detail TEXT, updated TEXT, attempt TEXT, approved_posts TEXT DEFAULT "[]")')
        self.db.commit()

    def records(self,include_archived=False):
        jobs = {r[0]: r[1:] for r in self.db.execute('SELECT key,state,url,detail,updated FROM uploads')}
        result = super().records(include_archived=include_archived)
        for r in result:
            if r['state'] == 'ready':
                known = set(employee_aliases().values())
                names = [n.strip() for n in canonical_employee(r['data']['employee']).split(',')]
                if len(names) != 1 or names[0] not in known:
                    r.update(state='review', status='Cần kiểm tra', detail='Cần một nhân viên phụ trách đã nhận diện. Kiểm tra CV thực hiện trong Excel.')
                try:
                    amount = Decimal(r['data']['fee'])
                    if not amount.is_finite() or amount <= 0 or amount != amount.to_integral_value():
                        raise ValueError()
                except (InvalidOperation, ValueError):
                    r.update(state='review', status='Cần kiểm tra', detail='Phí thẩm định phải là số nguyên dương trong Excel.')
            job = jobs.get(r['key'])
            r['upload_time'] = job[3] if job else ''
            if job and job[0] in ('preparing', 'submitting', 'uncertain', 'blocked') and r['state'] in ('ready', 'review', 'unknown'):
                r.update(state='review', status='Cần kiểm tra', detail=job[2] or 'Lần tải trước chưa có xác nhận. Đối chiếu website trước khi thử lại.')
                if job[1]:
                    r['url'] = job[1]
            elif job and job[0] == 'failed' and r['state'] == 'ready':
                r['detail'] = 'Lần tải trước thất bại trước khi gửi: ' + job[2]
        return result

    def recover(self):
        # Called once by the single app instance, never by its worker connection.
        self.db.execute("UPDATE uploads SET state='failed',detail='Ứng dụng dừng trước khi gửi; có thể thử lại.' WHERE state='preparing'")
        self.db.execute("UPDATE uploads SET state='uncertain',detail='Ứng dụng dừng khi đang gửi. Cần đối chiếu website trước khi thử lại.' WHERE state='submitting'")
        self.db.commit()

    def validate_record(self, key):
        record = next((r for r in self.records() if r['key'] == key), None)
        if not record or not record['current'] or record['state'] != 'ready':
            raise ValueError('Chỉ tải chứng thư hiện tại có trạng thái Sẵn sàng tải lên. ' + (record['detail'] if record else ''))
        self.validate_source(record)
        known = set(employee_aliases().values())
        names = [n.strip() for n in canonical_employee(record['data']['employee']).split(',')]
        if len(names) != 1 or names[0] not in known:
            raise ValueError('Cần một nhân viên phụ trách đã nhận diện. Kiểm tra CV thực hiện trong Excel.')
        return record

    def validate_source(self, record):
        for source in record.get('sources') or [record]:
            current = next((r for r in read_file(source['source']) if r['key'] == record['key'] and r['row'] == source['row']), None)
            if not current or current['fingerprint'] != source.get('fingerprint', record['fingerprint']):
                raise ValueError('Excel đã thay đổi sau lần quét. Quét lại trước khi tải lên.')

    def reserve(self, key):
        record = self.validate_record(key)
        previous = self.db.execute('SELECT approved_posts,fingerprint FROM uploads WHERE key=?', (key,)).fetchone()
        approved = previous[0] if previous and previous[1] == record['fingerprint'] else '[]'
        attempt = uuid.uuid4().hex
        # The app's named mutex prevents two upload workers from separate windows.
        with self.db:
            self.db.execute('INSERT OR REPLACE INTO uploads VALUES (?,?,?,?,?,?,?,?)',
                            (key, record['fingerprint'], 'preparing', '', '', dt.datetime.now().isoformat(), attempt, approved))
            self.db.execute('INSERT INTO events VALUES(?,?,?,?)', (dt.datetime.now().isoformat(), key, 'upload_start', attempt))
        return record

    def submitting(self, record):
        # Commit BEFORE clicking Publish. A crash after this point is uncertain.
        self.validate_source(record)
        self.result(record, 'submitting', detail='Đang chờ website xác nhận lưu.')

    def result(self, record, state, url='', detail=''):
        now = dt.datetime.now().isoformat()
        with self.db:
            self.db.execute('UPDATE uploads SET state=?,url=?,detail=?,updated=? WHERE key=? AND fingerprint=?',
                            (state, url, detail, now, record['key'], record['fingerprint']))
            if state == 'success':
                if not url.startswith('https://hoso.Contract Desk.vn/wp-admin/post.php?post='):
                    raise ValueError('Thiếu liên kết xác nhận từ website.')
                self.db.execute('UPDATE records SET uploaded_fingerprint=?,url=? WHERE key=?',
                                (record['fingerprint'], url, record['key']))
            self.db.execute('INSERT INTO events VALUES(?,?,?,?)',
                            (now, record['key'], 'upload_' + state, json.dumps({'url': url, 'detail': detail}, ensure_ascii=False)))

    def approved_posts(self, record):
        row = self.db.execute('SELECT approved_posts,fingerprint FROM uploads WHERE key=?', (record['key'],)).fetchone()
        approved = set(json.loads(row[0])) if row and row[1] == record['fingerprint'] else set()
        # Legacy explicit review approved the currently referenced website record.
        if record.get('decision') == 'new':
            reference = self.setting('new_web_reference:' + record['key'])
            if reference and reference != 'None':
                approved.add(reference)
        return approved

    def block_existing(self, record, posts):
        url = posts[0]['url']
        detail = 'Có hồ sơ cùng hợp đồng và ngày chứng thư trên website. Mở website để đối chiếu; xác nhận đã tải lên hoặc xác nhận chứng thư chưa có trên website.'
        self.result(record, 'blocked', url, detail)
        self.save_setting('candidate_posts:' + record['key'], json.dumps([p['post'] for p in posts]))

    def decide(self, key, action, url=''):
        job = self.db.execute('SELECT state FROM uploads WHERE key=?', (key,)).fetchone()
        if job and job[0] in ('preparing', 'submitting'):
            raise ValueError('Chờ lượt tải hiện tại hoàn tất trước khi xác nhận.')
        if action == 'reset' and job and job[0] == 'success':
            raise ValueError('Lịch sử tải thành công được giữ lại để tránh tải trùng. Mở website để kiểm tra hồ sơ.')
        super().decide(key, action, url)
        if action == 'uploaded' and job:
            self.db.execute("UPDATE uploads SET state='success',url=?,updated=? WHERE key=?", (url, dt.datetime.now().isoformat(), key))
        elif action == 'new' and job and job[0] in ('uncertain', 'blocked', 'failed'):
            posts = json.loads(self.setting('candidate_posts:' + key, '[]'))
            existing = self.db.execute('SELECT approved_posts FROM uploads WHERE key=?', (key,)).fetchone()
            posts = sorted(set(posts + json.loads(existing[0])))
            self.db.execute("UPDATE uploads SET state='failed',detail='',approved_posts=? WHERE key=?", (json.dumps(posts), key))
        self.db.commit()
