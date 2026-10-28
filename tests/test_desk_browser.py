"""Real Chromium integration against an intercepted WordPress stand-in. No live writes."""
import html
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'ContractDesk'))
import openpyxl
from desk_store import DeskStore
import website_upload as web


LABELS = [('client','Tên khách hàng','text'), ('content','Nội dung công việc','textarea'),
          ('employee','Nhân viên phụ trách','user'), ('assigned','Ngày giao việc','date'),
          ('deadline','Deadline','date'), ('status','Trạng thái hồ sơ','status'),
          ('fee','Thanh toán đợt 1','number'), ('payment','Trạng thái thanh toán đợt 1','payment'),
          ('invoice','Trạng thái hóa đơn đợt 1','invoice')]


def form(saved=None):
    saved = saved or {}
    parts = ['<html><meta charset="utf-8"><body>', '<input id="title" value="' + html.escape(saved.get('title',''), quote=True) + '">',
             '<input id="original_post_status" type="hidden" value="' + saved.get('wp_status','auto-draft') + '">']
    options = {'user':[('31','Sample User')], 'status':[('received','Đang xử lý'),('issued','Đã phát hành')],
               'payment':[('unpaid','Chưa thanh toán'),('paid','Đã thanh toán')], 'invoice':[('pending','Chưa Xuất hóa đơn'),('done','Đã xuất hóa đơn')]}
    for key, label, kind in LABELS:
        value = saved.get(key,'')
        parts.append('<div class="acf-field"><div class="acf-label"><label>' + label + '</label></div><div class="acf-input">')
        if kind in options:
            parts.append('<select id="' + key + '">')
            if kind == 'user': parts.append('<option value=""></option>')
            for val, title in options[kind]:
                parts.append('<option value="' + val + '"' + (' selected' if value==val else '') + '>' + title + '</option>')
            parts.append('</select>')
            if kind == 'user':
                parts.append('''<button class="select2-selection" onclick="document.querySelector('#choices').style.display='block'">Choose</button>
                <div id="choices" style="display:none"><input class="select2-search__field"><ul><li class="select2-results__option" onclick="document.querySelector('#employee').value='31';document.querySelector('#choices').style.display='none'">Sample User</li></ul></div>''')
        elif kind == 'date':
            display = value
            ymd = ''.join(reversed(value.split('/'))) if value else ''
            parts.append(f'<input id="{key}" type="hidden" value="{ymd}"><input value="{display}">')
        elif kind == 'textarea':
            parts.append('<textarea id="' + key + '">' + html.escape(value) + '</textarea>')
        else:
            parts.append('<input id="' + key + '" type="' + kind + '" value="' + html.escape(value, quote=True) + '">')
        parts.append('</div></div>')
    parts.append('''<div id="submitdiv"><div class="inside"><button id="publish" onclick="saveRecord()">Publish</button></div></div>
    <script>async function saveRecord(){const data={};for(const e of document.querySelectorAll('[id]')){
      if(e.matches('input,textarea,select'))data[e.id]=e.value;
    }const response=await fetch('/save',{method:'POST',body:JSON.stringify(data)});const result=await response.json();location.href='/wp-admin/post.php?post='+result.id+'&action=edit';}</script></body></html>''')
    return ''.join(parts)


class BrowserTests(unittest.TestCase):
    def setUp(self):
        from playwright.sync_api import sync_playwright
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(Path(__file__).resolve().parents[1] / 'browsers')
        self.tmp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1] / 'build')
        folder = Path(self.tmp.name)
        source = folder/'employee.xlsx'
        b=openpyxl.Workbook();s=b.active;s.title='Contracts'
        s.append(['SỐ HỢP ĐỒNG','NGÀY HỢP ĐỒNG','KHÁCH HÀNG','NỘI DUNG THẨM ĐỊNH','PHÍ THẨM ĐỊNH','SỐ CHỨNG THƯ / CVTT','NGÀY CHỨNG THƯ','CV THỰC HIỆN'])
        s.append(['000001','28/01/2026','Test Client','Test content',70000000,'TEST/2026/001\nTEST/2026/002','09/04/2026\n16/07/2026','Sample User'])
        b.save(source);b.close()
        self.store=DeskStore(folder/'desk.db');self.store.scan([source])
        self.keys=[r['key'] for r in self.store.records()]
        self.saved={};self.corrupt_fee=False;self.wp_status='publish'
        self.pw=sync_playwright().start();self.browser=self.pw.chromium.launch(headless=True,channel='chromium')
        self.context=self.browser.new_context()
        self.context.route('**/*',self.route)
        self.page=self.context.new_page();self.page.set_default_timeout(2500)
    def tearDown(self):
        self.browser.close();self.pw.stop();self.store.db.close();self.tmp.cleanup()
    def route(self, route):
        from urllib.parse import urlparse,parse_qs
        request=route.request;p=urlparse(request.url)
        if p.netloc!='hoso.Contract Desk.vn':route.abort();return
        if p.path=='/save':
            item=json.loads(request.post_data)
            for name in ['assigned','deadline']:
                val=item[name];item[name]=val[6:]+'/'+val[4:6]+'/'+val[:4]
            item['wp_status']=self.wp_status
            if self.corrupt_fee:item['fee']='700000000'
            post=str(9000+len(self.saved)+1);self.saved[post]=item
            route.fulfill(content_type='application/json',body=json.dumps(dict(id=post)));return
        if p.path.endswith('/post-new.php'):
            body=form()
        elif p.path.endswith('/post.php'):
            body=form(self.saved[parse_qs(p.query)['post'][0]])
        elif p.path.endswith('/edit.php'):
            body='<table><tbody id="the-list">'+''.join(f'<tr><td><a class="row-title" href="https://hoso.Contract Desk.vn/wp-admin/post.php?post={post}&action=edit">{html.escape(item["title"])}</a></td></tr>' for post,item in self.saved.items())+'</tbody></table>'
        else:route.abort();return
        route.fulfill(content_type='text/html; charset=utf-8',body=body)
    def run_batch(self):
        return web.upload_with_page(self.store,self.page,self.keys,threading.Event(),lambda *a:None)
    def test_two_certificates_same_contract_publish_save_reload_and_duplicate_check(self):
        results=self.run_batch()
        self.assertEqual([r['state'] for r in results],['success','success'],results)
        self.assertEqual(len(self.saved),2)
        self.assertEqual([r['title'] for r in self.saved.values()],['[000001/26] Test Client']*2)
        self.assertEqual([r['deadline'] for r in self.saved.values()],['09/04/2026','16/07/2026'])
        self.assertTrue(all(r['fee']=='70000000' for r in self.saved.values()))
        # Simulate an upload performed elsewhere; matching date blocks a fresh local history.
        other=DeskStore(':memory:')
        try:
            other.scan_records([dict(r) for r in self.store.records()],[])
            result=web.upload_with_page(other,self.page,[self.keys[0]],threading.Event(),lambda *a:None)
            self.assertEqual(result[0]['state'],'blocked',result)
            self.assertEqual(len(self.saved),2)
        finally:other.db.close()
    def test_saved_wrong_fee_is_uncertain_and_never_automatic_retry(self):
        self.corrupt_fee=True
        result=self.run_batch()
        self.assertTrue(all(r['state']=='uncertain' for r in result),result)
        self.assertTrue(all(r['state']=='review' for r in self.store.records()))
        self.assertIn('Phí', result[0]['detail'])
    def test_redirect_to_saved_draft_is_not_success(self):
        self.wp_status='draft'
        result=self.run_batch()
        self.assertTrue(all(r['state']=='uncertain' for r in result),result)


if __name__=='__main__':unittest.main()
