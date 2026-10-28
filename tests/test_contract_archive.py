import copy
import json
from pathlib import Path
import shutil
import sys
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'ContractDesk'))
from contract_tracker import read_file
from desk_store import DeskStore
import test_contract_desk as fixtures

class ArchiveTests(unittest.TestCase):
    def setUp(self):
        self.f=fixtures.DeskTests();self.f.setUp();self.store=self.f.store
    def tearDown(self):self.f.tearDown()
    def success(self,key):
        r=self.store.reserve(key);self.store.submitting(r);self.store.result(r,'success',fixtures.FakePage.url)
    def test_uploaded_certificate_archived_without_losing_upload_or_other_certificate(self):
        rows=self.store.records();self.success(rows[0]['key'])
        self.store.scan_records(read_file(self.f.path)[1:],[])
        self.assertEqual([r['key'] for r in self.store.records()],[rows[1]['key']])
        old=next(r for r in self.store.records(include_archived=True) if r['key']==rows[0]['key'])
        self.assertTrue(old['archived']);self.assertFalse(old['current']);self.assertEqual(old['state'],'uploaded')
        self.assertEqual(old['url'],fixtures.FakePage.url)
        with self.assertRaises(ValueError):self.store.reserve(old['key'])
        with self.assertRaises(ValueError):self.store.export({old['key']},self.f.folder/'archived.xlsx')
        self.store.scan([self.f.path]);self.assertEqual(self.store.records()[0]['state'],'uploaded')
        self.assertEqual(self.store.last_scan_stats['restored'],1)
    def test_move_of_excel_source_keeps_upload_and_cancellation_decisions(self):
        rows=self.store.records();self.success(rows[0]['key']);self.store.decide(rows[1]['key'],'cancel')
        moved=self.f.folder/'moved.xlsx';shutil.move(self.f.path,moved);self.store.scan([moved])
        after=self.store.records();self.assertEqual([r['key'] for r in after],[r['key'] for r in rows])
        self.assertEqual([r['state'] for r in after],['uploaded','cancel']);self.assertEqual(self.store.last_scan_stats['moved'],2)
        self.store.validate_source(after[0]);self.assertEqual(len(self.store.records(include_archived=True)),2)
    def test_failed_and_partial_scans_leave_current_list_and_archive_unchanged(self):
        before=self.store.setting('last_scan');rows=self.store.records()
        self.assertEqual(self.store.scan_records([read_file(self.f.path)[0]],['cloud storage not available']),['cloud storage not available'])
        self.assertEqual(self.store.setting('last_scan'),before);self.assertEqual([r['key'] for r in self.store.records()],[r['key'] for r in rows])
        self.assertEqual(self.store.db.execute('SELECT COUNT(*) FROM archive').fetchone()[0],0)
        bad=self.f.folder/'bad.xlsx';bad.write_text('incomplete cloud storage file')
        failures=self.store.scan([self.f.path,bad]);self.assertTrue(failures);self.assertEqual(self.store.setting('last_scan'),before)
    def test_empty_successful_scan_archives_and_return_restores_uncertain_job(self):
        key=self.f.keys[0];r=self.store.reserve(key);self.store.submitting(r);self.store.result(r,'uncertain',detail='Check website')
        self.store.scan_records([],[]);self.assertEqual(self.store.records(),[])
        self.assertEqual(len(self.store.records(include_archived=True)),2)
        self.store.scan([self.f.path]);self.assertEqual(self.store.records()[0]['state'],'review')
        self.assertIn('Check website',self.store.records()[0]['detail'])
    def test_return_with_changed_data_preserves_uploaded_flag_and_marks_changed(self):
        key=self.f.keys[0];self.success(key);self.store.scan_records([],[])
        self.f.write(fee=90000000);self.store.scan([self.f.path])
        row=next(r for r in self.store.records() if r['key']==key)
        self.assertEqual(row['state'],'changed');self.assertEqual(row['url'],fixtures.FakePage.url)
    def test_existing_absent_rows_migrate_into_archive(self):
        key=self.f.keys[0];self.success(key)
        with self.store.db:self.store.db.execute('UPDATE records SET seen=? WHERE key=?',('older-scan',key))
        self.store.db.close();self.store=DeskStore(self.f.db);self.f.store=self.store
        self.assertEqual(len(self.store.records()),1)
        old=next(r for r in self.store.records(include_archived=True) if r['key']==key)
        self.assertTrue(old['archived']);self.assertEqual(old['state'],'uploaded')
        self.assertTrue(self.store.db.execute('SELECT 1 FROM record_history WHERE key=? AND action="archived"',(key,)).fetchone())
    def test_scan_transaction_rolls_back_on_archive_failure(self):
        before=self.store.setting('last_scan')
        with patch.object(self.store,'_archive',side_effect=RuntimeError('interrupted')):
            with self.assertRaises(RuntimeError):self.store.scan_records(read_file(self.f.path)[1:],[])
        self.assertEqual(self.store.setting('last_scan'),before);self.assertEqual(len(self.store.records()),2)
        self.assertEqual(self.store.db.execute('SELECT COUNT(*) FROM archive').fetchone()[0],0)
    def test_archive_history_survives_reopening(self):
        self.store.scan_records([],[]);self.store.db.close();self.store=DeskStore(self.f.db);self.f.store=self.store
        self.assertEqual(self.store.records(),[]);self.assertEqual(len(self.store.records(include_archived=True)),2)
        self.assertEqual(sum(e['action']=='archived' for e in self.store.history_log()),2)

if __name__=='__main__':unittest.main()
