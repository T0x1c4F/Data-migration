import copy
import json
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'ContractDesk'))
import openpyxl
from contract_tracker import read_file
from desk_store import DeskStore
import website_upload as web


class FakeLocator:
    def count(self): return 1
    def evaluate(self, script): pass
    def click(self): pass


class FakePage:
    url = 'https://hoso.Contract Desk.vn/wp-admin/post.php?post=9001&action=edit'
    def locator(self, selector): return FakeLocator()
    def wait_for_url(self, *args, **kwargs): pass


class DeskTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1] / 'build')
        self.folder = Path(self.tmp.name)
        self.path = self.folder / 'employee.xlsx'
        self.db = self.folder / 'desk.db'
        self.store = DeskStore(self.db)
        self.write()
        self.store.scan([self.path])
        self.keys = [r['key'] for r in self.store.records()]
    def tearDown(self):
        self.store.db.close()
        self.tmp.cleanup()
    def write(self, fee=70000000, employee='Sample User', note='', client='Client'):
        book = openpyxl.Workbook(); sheet = book.active
        sheet.title = 'Contracts'
        sheet.append(['SỐ HỢP ĐỒNG','NGÀY HỢP ĐỒNG','KHÁCH HÀNG','NỘI DUNG THẨM ĐỊNH','PHÍ THẨM ĐỊNH','SỐ CHỨNG THƯ / CVTT','NGÀY CHỨNG THƯ','CV THỰC HIỆN','GHI CHÚ'])
        sheet.append(['000001','28/01/2026',client,'Nội dung',fee,'TEST/2026/001\nTEST/2026/002','09/04/2026\n16/07/2026',employee,note])
        book.save(self.path); book.close()
    def run_batch(self, verify=None, posts=None, fill=None):
        with patch.object(web, 'existing_posts', return_value=posts or []), patch.object(web, 'fill_record', side_effect=fill or (lambda *a:'title')), patch.object(web, 'verify_saved', side_effect=verify or (lambda *a:FakePage.url)):
            return web.upload_with_page(self.store, FakePage(), self.keys, threading.Event(), lambda *a:None)
    def test_success_per_certificate_persists_and_excludes_reupload(self):
        self.run_batch()
        self.assertEqual([r['state'] for r in self.store.records()], ['uploaded','uploaded'])
        self.store.db.close(); self.store = DeskStore(self.db); self.store.recover()
        self.store.scan([self.path])
        self.assertTrue(all(r['state']=='uploaded' for r in self.store.records()))
        with self.assertRaises(ValueError): self.store.reserve(self.keys[0])
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM events WHERE action='upload_success'").fetchone()[0],2)
    def test_partial_success_later_unknown_keeps_first(self):
        calls = [0]
        def verify(*args):
            calls[0] += 1
            if calls[0] == 2: raise TimeoutError('Connection lost after Publish')
            return FakePage.url
        result = self.run_batch(verify=verify)
        self.assertEqual([r['state'] for r in result], ['success','uncertain'])
        self.assertEqual([r['state'] for r in self.store.records()], ['uploaded','review'])
        with self.assertRaises(ValueError): self.store.reserve(self.keys[1])
    def test_failure_before_submit_is_retryable(self):
        result = self.run_batch(fill=lambda *a: (_ for _ in ()).throw(ValueError('employee unavailable')))
        self.assertTrue(all(r['state']=='failed' for r in result))
        self.assertTrue(all(r['state']=='ready' for r in self.store.records()))
    def test_crash_recovery_distinguishes_before_and_after_click(self):
        first = self.store.reserve(self.keys[0]); second = self.store.reserve(self.keys[1])
        self.store.submitting(second)
        self.store.db.close(); self.store = DeskStore(self.db); self.store.recover()
        self.assertEqual([r['state'] for r in self.store.records()], ['ready','review'])
    def test_source_edits_after_scan_block_upload(self):
        self.write(note='hủy')
        with self.assertRaisesRegex(ValueError, 'thay đổi'): self.store.reserve(self.keys[0])
        self.store.scan([self.path])
        with self.assertRaises(ValueError): self.store.reserve(self.keys[0])
    def test_source_edits_during_fill_block_before_publish(self):
        result = self.run_batch(fill=lambda *a:(self.write(fee=80000000) or 'title'))
        self.assertTrue(all(r['state']=='failed' for r in result))
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM events WHERE action='upload_submitting'").fetchone()[0],0)
    def test_live_duplicate_blocks_until_explicit_review(self):
        posts = [dict(post='9001', url=FakePage.url)]
        result = self.run_batch(posts=posts)
        self.assertTrue(all(r['state']=='blocked' for r in result))
        self.assertTrue(all(r['state']=='review' for r in self.store.records()))
        self.store.decide(self.keys[0], 'new')
        record = self.store.validate_record(self.keys[0])
        self.assertEqual(self.store.approved_posts(record), {'9001'})
        # A newly discovered website record is never included by old approval.
        with patch.object(web, 'existing_posts', return_value=posts+[dict(post='9002',url=FakePage.url)]), patch.object(web, 'fill_record') as fill:
            web.upload_with_page(self.store, FakePage(), [self.keys[0]], threading.Event(), lambda *a:None)
            fill.assert_not_called()
    def test_manual_confirmation_resolves_uncertain(self):
        record = self.store.reserve(self.keys[0]); self.store.submitting(record)
        self.store.result(record, 'uncertain', detail='lost')
        self.store.decide(self.keys[0], 'uploaded', FakePage.url)
        self.assertEqual(self.store.records()[0]['state'], 'uploaded')
        with self.assertRaises(ValueError): self.store.decide(self.keys[0], 'reset')
    def test_required_employee_and_missing_fields_block(self):
        self.write(employee='Unknown Name'); self.store.scan([self.path])
        with self.assertRaisesRegex(ValueError, 'nhân viên'): self.store.validate_record(self.keys[0])
        self.write(client=''); self.store.scan([self.path])
        with self.assertRaises(ValueError): self.store.reserve(self.keys[0])
    def test_stop_keeps_unattempted_rows_ready(self):
        stop = threading.Event(); stop.set()
        self.assertEqual(web.upload_with_page(self.store, FakePage(), self.keys, stop, lambda *a:None), [])
        self.assertTrue(all(r['state']=='ready' for r in self.store.records()))
    def test_money_preserves_integer_value(self):
        self.assertEqual(web.money('70000000.0'), '70000000')
        self.assertEqual(web.money('7E+7'), '70000000')
        for value in ['nan', 'inf', '', '12.5', '-1', '0']:
            with self.assertRaises(ValueError): web.money(value)
    def test_website_title_variations_and_leading_zeros(self):
        for title in ['[000001/26] Client', '[12802/2026] Client', '[000001.2026] Client', '[000001].2026.Client']:
            self.assertEqual(web.title_contract_key(title), '12802|2026')
    def test_success_followed_by_source_change_is_changed(self):
        self.run_batch(); self.write(fee=90000000); self.store.scan([self.path])
        self.assertTrue(all(r['state']=='changed' for r in self.store.records()))
    def test_invalid_confirmation_does_not_mark_uploaded(self):
        record = self.store.reserve(self.keys[0]); self.store.submitting(record)
        with self.assertRaises(ValueError): self.store.result(record, 'success', 'https://other.example')
        self.assertEqual(self.store.records()[0]['state'], 'review')
    def test_not_ready_export_allows_only_review_and_missing(self):
        book=openpyxl.load_workbook(self.path);sheet=book.active
        sheet.delete_rows(2,sheet.max_row)
        for i in range(7):
            sheet.append([str(120001+i),'28/01/2026','Client','Content',70000000,f'EX/{i}/2026','09/04/2026','Sample User',''])
        sheet['C2']=None
        sheet['C3'].font=openpyxl.styles.Font(color='FFFF0000')
        sheet['I4']='hủy'
        sheet['A7']='000009'  # Synthetic test record: displayed as already uploaded.
        book.save(self.path);book.close()
        self.store.scan([self.path])
        for key in ('ex/3/2026','ex/6/2026'):
            self.store.decide(key,'uploaded',FakePage.url)
        book=openpyxl.load_workbook(self.path);book.active['E8']=80000000;book.save(self.path);book.close()
        self.store.scan([self.path])
        rows=self.store.records()
        current=[r for r in rows if r['current']]
        self.assertEqual({r['state'] for r in current},{'missing','review','cancel','uploaded','ready','unknown','changed'})
        allowed={r['key'] for r in current if r['state'] in ('missing','review')}
        output=self.folder/'pending.xlsx'
        self.assertEqual(self.store.export(allowed,output,mode='not_ready'),2)
        book=openpyxl.load_workbook(output,data_only=True);sheet=book.active
        self.assertEqual([sheet.cell(i,1).value for i in (2,3)],['120001','120002'])
        status_col=next(c.column for c in sheet[1] if c.value=='TRẠNG THÁI THEO DÕI')
        self.assertEqual({sheet.cell(i,status_col).value for i in (2,3)},{'Thiếu thông tin','Cần kiểm tra'})
        book.close()
        for row in current:
            if row['key'] not in allowed:
                with self.subTest(state=row['state']),self.assertRaises(ValueError):
                    self.store.export(allowed|{row['key']},self.folder/'invalid.xlsx',mode='not_ready')


if __name__ == '__main__': unittest.main()
