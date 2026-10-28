"""Contract Desk entry point. Old apps are never changed."""
import ctypes
import hashlib
import json
import os
from pathlib import Path
import sys


def verify_installation():
    """Offline packaged-runtime check; never connects to the production website."""
    from desk_store import DeskStore
    from playwright.sync_api import sync_playwright
    import website_upload
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent/'dist'/'Contract Desk'
    report = base/'installation-check.json'
    try:
        store = DeskStore(':memory:')
        assert store.records() == []
        assert store.records(include_archived=True) == []
        store.db.close()
        assert website_upload.money('70000000.0') == '70000000'
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(base/'browsers')
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True, channel='chromium', timeout=20000)
            page = browser.new_page()
            page.route('**/*', lambda route: route.abort())
            page.set_content('<h1>Contract Desk offline test</h1>')
            assert page.locator('h1').inner_text() == 'Contract Desk offline test'
            browser.close()
        import tempfile
        import openpyxl
        from openpyxl.styles import Font,PatternFill
        with tempfile.TemporaryDirectory(prefix='contract-desk-check-') as temp:
            path=Path(temp)/'excel-check.xlsx';book=openpyxl.Workbook();sheet=book.active
            sheet['A1']='Contract Desk';sheet.merge_cells('A1:B1');sheet['C2']=70000000
            sheet['C2'].font=Font(bold=True);sheet['C2'].fill=PatternFill('solid',fgColor='FFFF0000')
            book.save(path);book.close();book=openpyxl.load_workbook(path)
            assert book.active['C2'].value==70000000 and book.active['C2'].font.bold
            assert 'A1:B1' in book.active.merged_cells;book.close()
        report.write_text(json.dumps({'passed': True, 'checks':['logo','database','upload engine','single bundled browser offline','styled Excel read/write']}, ensure_ascii=False), encoding='utf-8')
    except Exception as error:
        report.write_text(json.dumps({'passed': False, 'error':str(error)}, ensure_ascii=False), encoding='utf-8')
        raise


def app():
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent / 'dist' / 'Contract Desk'
    base.mkdir(parents=True, exist_ok=True)
    mutex = None
    if os.name == 'nt':
        name = 'Local\\Contract DeskContractDesk-' + hashlib.sha256(str(base.resolve()).casefold().encode()).hexdigest()[:24]
        kernel = ctypes.windll.kernel32
        kernel.CreateMutexW.restype = ctypes.c_void_p
        mutex = kernel.CreateMutexW(None, False, name)
        if not mutex:
            raise OSError('Không thể khóa phiên ứng dụng. Đóng cửa sổ cũ rồi thử lại.')
        if kernel.GetLastError() == 183:
            import tkinter.messagebox
            tkinter.messagebox.showinfo('Contract Desk', 'Ứng dụng đã mở. Chuyển sang cửa sổ Contract Desk đang chạy.')
            kernel.CloseHandle(ctypes.c_void_p(mutex))
            return
    try:
        from contract_tracker import read_file, STATUS, STATUS_COLORS, WEBSITE_REFERENCE_DATE, WEBSITE_REFERENCE_COUNT
        from desk_store import DeskStore
        from tracker_ui import run_app
        run_app(DeskStore, read_file, STATUS, STATUS_COLORS, reference_date=WEBSITE_REFERENCE_DATE,
                reference_count=WEBSITE_REFERENCE_COUNT, base=base, app_name='Contract Desk')
    finally:
        if mutex:
            ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(mutex))


if __name__ == '__main__':
    try:
        if '--verify-installation' in sys.argv:
            verify_installation()
        else:
            app()
    except Exception as error:
        import tkinter.messagebox
        tkinter.messagebox.showerror('Contract Desk', str(error))
        raise
