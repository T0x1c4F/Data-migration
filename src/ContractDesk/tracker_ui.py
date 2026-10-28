"""Contract Desk desktop workspace. Tracking and workbook export stay in the core module."""
import datetime as dt
import json
import os
from pathlib import Path
import queue
import re
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
import unicodedata
import webbrowser

BG = '#f4f6fa'
WHITE = '#ffffff'
INK = '#17233b'
MUTED = '#708096'
LINE = '#e4e9f1'
BLUE = '#2764e7'


def search_text(value):
    value = str(value).casefold().replace('đ', 'd')
    return unicodedata.normalize('NFD', value).encode('ascii', 'ignore').decode()


def run_app(Store, read_file, statuses, status_colors, reference_date='', reference_count=0, base=None, app_name='Contract Desk'):
    base = base or (Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent / 'dist' / app_name)
    base.mkdir(parents=True, exist_ok=True)
    store = Store(base / 'contract_tracker.sqlite3')
    store.recover()
    root = tk.Tk()
    root.title(app_name)
    root.geometry('1360x820')
    root.minsize(1000, 620)
    root.configure(background=BG)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('.', font=('Segoe UI', 10), foreground=INK)
    style.configure('TButton', background=WHITE, bordercolor=LINE, lightcolor=WHITE, darkcolor=WHITE,
                    relief='flat', padding=(10, 6), font=('Segoe UI', 10))
    style.map('TButton', background=[('active', '#edf2fa'), ('disabled', '#f3f5f8')],
              foreground=[('disabled', '#a2acba')], bordercolor=[('active', '#cbd7ec')])
    style.configure('Primary.TButton', background=BLUE, foreground=WHITE, bordercolor=BLUE,
                    lightcolor=BLUE, darkcolor=BLUE, font=('Segoe UI', 10, 'bold'))
    style.map('Primary.TButton', background=[('active', '#1e50be'), ('disabled', '#d7e1f5')],
              foreground=[('disabled', WHITE)], bordercolor=[('active', '#1e50be')])
    style.configure('Quiet.TButton', background='#f7f9fc', bordercolor='#f7f9fc', padding=(8, 5))
    style.configure('TMenubutton', background=WHITE, bordercolor=LINE, relief='flat', padding=(10, 6))
    style.map('TMenubutton', background=[('active', '#edf2fa')])
    style.configure('TEntry', fieldbackground=WHITE, bordercolor=LINE, lightcolor=LINE,
                    darkcolor=LINE, padding=(8, 5))
    style.map('TEntry', bordercolor=[('focus', '#8aaef6')])
    style.configure('TCombobox', fieldbackground=WHITE, background=WHITE, bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE, arrowsize=14, padding=(8, 5))
    style.map('TCombobox', fieldbackground=[('readonly', WHITE)], foreground=[('readonly', INK)],
              selectbackground=[('readonly', WHITE)], selectforeground=[('readonly', INK)])
    style.configure('Treeview', background=WHITE, fieldbackground=WHITE, foreground=INK,
                    borderwidth=0, relief='flat', rowheight=30, font=('Segoe UI', 10))
    style.configure('Treeview.Heading', background='#f4f7fb', foreground='#65758c',
                    borderwidth=0, relief='flat', padding=(8, 7), font=('Segoe UI', 9, 'bold'))
    style.map('Treeview', background=[('selected', '#d8e7ff')], foreground=[('selected', '#143970')])
    style.map('Treeview.Heading', background=[('active', '#edf2fa')])
    style.configure('Vertical.TScrollbar', background='#dbe2ed', troughcolor='#f8fafd', borderwidth=0, arrowsize=11)
    style.configure('Horizontal.TScrollbar', background='#dbe2ed', troughcolor='#f8fafd', borderwidth=0, arrowsize=11)
    style.configure('Scan.Horizontal.TProgressbar', background=BLUE, troughcolor=BG, borderwidth=0, thickness=3)

    def label(parent, value='', color=INK, font=('Segoe UI', 10), bg=WHITE, **kwargs):
        kwargs.setdefault('anchor', 'w')
        return tk.Label(parent, text=value, background=bg, foreground=color, font=font, **kwargs)

    # Brand header: the original logo stays crisp on a white surface.
    header = tk.Frame(root, background=WHITE, highlightbackground=LINE, highlightthickness=1)
    header.grid(row=0, column=0, sticky='ew')
    header.columnconfigure(0, weight=1)
    identity = tk.Frame(header, background=WHITE)
    identity.grid(row=0, column=0, sticky='w', padx=20, pady=8)
    label(identity, app_name, INK, ('Segoe UI', 18, 'bold')).pack(anchor='w')
    logo_path = Path(__file__).resolve().parent / 'assets' / 'Contract Desk-logo.png'
    try:
        from PIL import Image, ImageTk
        image = Image.open(logo_path).convert('RGBA')
        image.thumbnail((140, 44), Image.Resampling.LANCZOS)
        root.Contract Desk_logo = ImageTk.PhotoImage(image, master=root)
        tk.Label(header, image=root.Contract Desk_logo, background=WHITE).grid(row=0, column=2, sticky='e', padx=20, pady=7)
    except (OSError, ImportError):
        label(header, 'Contract Desk', '#ed1c24', ('Georgia', 32, 'bold')).grid(row=0, column=2, sticky='e', padx=24)
    history_button=ttk.Button(header,text='Lưu trữ / Lịch sử')
    history_button.grid(row=0,column=1,padx=10)

    body = tk.Frame(root, background=BG)
    body.grid(row=1, column=0, sticky='nsew', padx=14, pady=(10, 8))
    body.columnconfigure(0, weight=1)
    body.rowconfigure(2, weight=1)
    previous_source=store.setting('excel_file')
    legacy_source=store.setting('folder')
    if not previous_source and Path(legacy_source).is_file():previous_source=legacy_source
    excel_file = tk.StringVar(value=previous_source)
    query = tk.StringVar()
    filter_value = tk.StringVar(value='Tất cả')
    banner = tk.StringVar(value='Chọn tệp Excel, sau đó quét dữ liệu để bắt đầu.')
    table_meta = tk.StringVar()
    selection_title = tk.StringVar(value='Chọn một hồ sơ để xem thông tin')
    selection_reason = tk.StringVar(value='Nhấp đúp vào một dòng để xem toàn bộ chi tiết và nguồn Excel.')
    tasks = queue.Queue()
    visible = {}
    busy = False
    busy_kind = ''
    upload_stop = threading.Event()
    upload_context = []
    upload_started = ''
    closing = False
    sort_state = {}

    # Clickable counters also act as shortcuts into the list.
    counters = {}
    metrics = tk.Frame(body, background=BG)
    metrics.grid(row=0, column=0, sticky='ew', pady=(0, 8))
    for i, (key, title, color, selected_filter) in enumerate([
        ('all', 'TỔNG CHỨNG THƯ', '#334569', 'Tất cả'),
        ('ready', 'SẴN SÀNG TẢI LÊN', '#188252', statuses['ready']),
        ('pending', 'CHƯA SẴN SÀNG', '#b57a16', 'Chưa sẵn sàng'),
        ('uploaded', 'ĐÃ TẢI LÊN', '#66758d', statuses['uploaded']),
    ]):
        metrics.columnconfigure(i, weight=1, uniform='counter')
        card = tk.Frame(metrics, background=WHITE, highlightbackground=LINE, highlightthickness=1, cursor='hand2')
        card.grid(row=0, column=i, sticky='ew', padx=(0 if i==0 else 6, 0 if i==3 else 6))
        tk.Frame(card, width=3, background=color).pack(side='left', fill='y')
        content = tk.Frame(card, background=WHITE)
        content.pack(fill='both', expand=True, padx=10, pady=5)
        heading = label(content, title, MUTED, ('Segoe UI', 8, 'bold'))
        heading.pack(side='left')
        counters[key] = tk.StringVar(value='0')
        amount = tk.Label(content, textvariable=counters[key], background=WHITE, foreground=color,
                          font=('Segoe UI', 15, 'bold'), anchor='w')
        amount.pack(side='right', padx=(8, 0))
        def jump(event=None, value=selected_filter):
            query.set('')
            filter_value.set(value)
        for widget in (card, content, heading, amount):
            widget.bind('<Button-1>', jump)

    # Source picker is separate from list filtering and export actions.
    source = tk.Frame(body, background=WHITE, highlightbackground=LINE, highlightthickness=1)
    source.grid(row=1, column=0, sticky='ew', pady=(0, 8))
    source.columnconfigure(1, weight=1)
    label(source, 'NGUỒN EXCEL', MUTED, ('Segoe UI', 9, 'bold')).grid(row=0, column=0, padx=(10, 10), pady=7)
    source_entry = ttk.Entry(source, textvariable=excel_file)
    source_entry.grid(row=0, column=1, sticky='ew', pady=7)
    def choose():
        if busy:
            return
        current=Path(excel_file.get())
        initial_dir=current.parent if excel_file.get() else Path(legacy_source)
        path = filedialog.askopenfilename(parent=root, title='Chọn tệp Excel / cloud storage',
            initialdir=str(initial_dir) if initial_dir.is_dir() else None,filetypes=[('Excel workbook','*.xlsx')])
        if path:
            excel_file.set(path)
            store.save_setting('excel_file', path)
    choose_button = ttk.Button(source, text='Chọn Excel', command=choose)
    choose_button.grid(row=0, column=2, padx=8)
    scan_button = ttk.Button(source, text='Quét dữ liệu', style='Primary.TButton')
    scan_button.grid(row=0, column=3, padx=(0, 12))

    panel = tk.Frame(body, background=WHITE, highlightbackground=LINE, highlightthickness=1)
    panel.grid(row=2, column=0, sticky='nsew')
    panel.columnconfigure(0, weight=1)
    panel.rowconfigure(2, weight=1)
    toolbar = tk.Frame(panel, background=WHITE)
    toolbar.grid(row=0, column=0, sticky='ew', padx=12, pady=(8, 5))
    toolbar.columnconfigure(0, weight=1)
    title_group = tk.Frame(toolbar, background=WHITE)
    title_group.grid(row=0, column=0, sticky='w')
    label(title_group, 'Danh sách hợp đồng', INK, ('Segoe UI', 12, 'bold')).pack(side='left')
    tk.Label(title_group, textvariable=table_meta, background=WHITE, foreground=MUTED,
             font=('Segoe UI', 9)).pack(side='left', padx=12)
    focus_button = ttk.Button(toolbar, text='Mở rộng danh sách', style='Quiet.TButton')
    focus_button.grid(row=0, column=1, padx=(8, 6))
    ready_button = ttk.Button(toolbar, text='Xuất Excel sẵn sàng', style='Primary.TButton')
    ready_button.grid(row=0, column=2, padx=(0, 6))
    pending_button = ttk.Button(toolbar, text='Xuất Excel chưa sẵn sàng')
    pending_button.grid(row=0, column=3)
    focus_scan = ttk.Button(toolbar, text='Quét', style='Quiet.TButton')
    focus_scan.grid(row=0, column=4, padx=(6, 0))
    focus_scan.grid_remove()
    filters = tk.Frame(panel, background=WHITE)
    filters.grid(row=1, column=0, sticky='ew', padx=12, pady=(0, 6))
    filters.columnconfigure(1, weight=1)
    label(filters, 'Tìm kiếm', MUTED, ('Segoe UI', 9)).grid(row=0, column=0, padx=(0, 10))
    search_entry = ttk.Entry(filters, textvariable=query)
    search_entry.grid(row=0, column=1, sticky='ew', padx=(0, 16))
    label(filters, 'Trạng thái', MUTED, ('Segoe UI', 9)).grid(row=0, column=2, padx=(0, 10))
    status_filter = ttk.Combobox(filters, textvariable=filter_value, state='readonly', width=24,
                                values=['Tất cả', 'Chưa sẵn sàng'] + list(dict.fromkeys(v for k,v in statuses.items() if k!='absent')))
    status_filter.grid(row=0, column=3, sticky='e')
    ttk.Button(filters, text='Xóa lọc', style='Quiet.TButton', command=lambda: (query.set(''), filter_value.set('Tất cả'))).grid(row=0, column=4, padx=(8, 0))

    table = tk.Frame(panel, background=WHITE)
    table.grid(row=2, column=0, sticky='nsew')
    table.columnconfigure(0, weight=1)
    table.rowconfigure(0, weight=1)
    columns = ('contract', 'certificate', 'client', 'employee', 'status', 'change', 'file')
    tree = ttk.Treeview(table, columns=columns, show='headings', selectmode='extended', height=4)
    names = ['HỢP ĐỒNG', 'CHỨNG THƯ', 'KHÁCH HÀNG', 'NHÂN VIÊN', 'TRẠNG THÁI', 'CẬP NHẬT', 'TỆP NGUỒN']
    widths = [100, 135, 260, 115, 200, 110, 150]
    for col, title, width in zip(columns, names, widths):
        tree.heading(col, text=title, anchor='w')
        tree.column(col, width=width, minwidth=80, stretch=col in ('client', 'file'), anchor='w')
    tree.grid(row=0, column=0, sticky='nsew')
    vertical = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    vertical.grid(row=0, column=1, sticky='ns')
    horizontal = ttk.Scrollbar(table, orient='horizontal', command=tree.xview)
    horizontal.grid(row=1, column=0, sticky='ew')
    tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    for tag, color in status_colors.items():
        tree.tag_configure(tag, background=color)
    empty = label(table, 'Chưa có hồ sơ\nChọn tệp Excel và quét dữ liệu, hoặc thay đổi bộ lọc.', MUTED,
                  ('Segoe UI', 11), justify='center', anchor='center')

    # A small context area replaces the previous wall of buttons.
    dock = tk.Frame(panel, background='#fafbfd', highlightbackground=LINE, highlightthickness=1)
    dock.grid(row=3, column=0, sticky='ew')
    dock.columnconfigure(0, weight=1)
    selected_info = tk.Frame(dock, background='#fafbfd')
    selected_info.grid(row=0, column=0, sticky='ew', padx=12, pady=7)
    selection_heading = tk.Label(selected_info, textvariable=selection_title, background='#fafbfd',
                                  foreground=INK, font=('Segoe UI', 10, 'bold'), anchor='w')
    selection_heading.pack(fill='x')
    selection_description = tk.Label(selected_info, textvariable=selection_reason, background='#fafbfd',
                                     foreground=MUTED, font=('Segoe UI', 9), anchor='w', justify='left', wraplength=620)
    # Full reasons remain in the details dialog; reserve this space for the list.
    selected_info.bind('<Configure>', lambda e: selection_description.configure(wraplength=max(220, e.width-8)))
    context = tk.Frame(dock, background='#fafbfd')
    context.grid(row=0, column=1, sticky='e', padx=(0, 10), pady=7)
    details_button = ttk.Button(context, text='Chi tiết', style='Quiet.TButton')
    details_button.pack(side='left', padx=3)
    upload_button = ttk.Button(context, text='Tải lên website…', style='Primary.TButton')
    upload_button.pack(side='left', padx=3)
    stop_button = ttk.Button(context, text='Dừng sau hồ sơ này')
    more_button = ttk.Menubutton(context, text='Thao tác khác')
    more_button.pack(side='left', padx=3)
    menu = tk.Menu(more_button, tearoff=False, background=WHITE, foreground=INK,
                    activebackground='#e6efff', activeforeground=INK, font=('Segoe UI', 10))
    more_button['menu'] = menu

    footer = tk.Frame(root, background=BG)
    footer.grid(row=2, column=0, sticky='ew', padx=14, pady=(0, 6))
    footer.columnconfigure(1, weight=1)
    status_dot = tk.Label(footer, text='●', background=BG, foreground='#26976c', font=('Segoe UI', 9))
    status_dot.grid(row=0, column=0, padx=(0, 6))
    tk.Label(footer, textvariable=banner, background=BG, foreground=MUTED, font=('Segoe UI', 9),
             anchor='w').grid(row=0, column=1, sticky='ew')
    if reference_date:
        label(footer, f'Đối chiếu cũ: {reference_count} hồ sơ · {reference_date}', MUTED,
              ('Segoe UI', 9), bg=BG).grid(row=0, column=2, sticky='e', padx=(16, 0))
    progress = ttk.Progressbar(root, mode='indeterminate', style='Scan.Horizontal.TProgressbar')

    def selected():
        return [visible[key] for key in tree.selection() if key in visible]

    def set_busy(value, kind='scan'):
        nonlocal busy, busy_kind
        busy = value
        busy_kind = kind if value else ''
        scan_button.configure(text=('Đang tải…' if kind=='upload' else 'Đang quét…') if value else 'Quét dữ liệu', state='disabled' if value else 'normal')
        focus_scan.configure(state='disabled' if value else 'normal')
        choose_button.configure(state='disabled' if value else 'normal')
        source_entry.configure(state='disabled' if value else 'normal')
        ready_button.configure(state='disabled' if value else 'normal')
        pending_button.configure(state='disabled' if value else 'normal')
        if value and kind == 'upload':
            stop_button.pack(side='left', padx=3)
            stop_button.configure(state='normal', text='Dừng sau hồ sơ này')
        else:
            stop_button.pack_forget()
        if value:
            status_dot.configure(foreground=BLUE)
            progress.place(relx=0, rely=0, relwidth=1, height=3)
            progress.start(12)
        else:
            status_dot.configure(foreground='#26976c')
            progress.stop()
            progress.place_forget()
        show_selection()

    def show_selection(event=None):
        rows = selected()
        state = 'normal' if rows and not busy else 'disabled'
        for button in (details_button, more_button):
            button.configure(state=state)
        upload_button.configure(state='normal' if visible and not busy else 'disabled')
        if not rows:
            selection_title.set('Chọn một hồ sơ để xem thông tin')
            selection_reason.set('Nhấp đúp vào một dòng để xem toàn bộ chi tiết và nguồn Excel.')
        elif len(rows) > 1:
            selection_title.set(f'Đã chọn {len(rows)} chứng thư')
            selection_reason.set('Các nút xuất Excel chỉ xuất mục phù hợp trong những dòng đã chọn.')
        else:
            record = rows[0]
            selection_title.set(f"{record['data']['contract']}  /  {record['data']['certificate'] or 'Chưa có số chứng thư'}  ·  {record['status']}")
            reason = record['detail'] or f"{record['employee_name']}  ·  {Path(record['source']).name}  ·  dòng {record['row']}"
            selection_reason.set(reason[:190] + ('…' if len(reason)>190 else ''))

    def refresh(*args):
        previous = set(tree.selection())
        tree.delete(*tree.get_children())
        visible.clear()
        records = store.records()
        counters['all'].set(str(len(records)))
        counters['ready'].set(str(sum(r['state']=='ready' for r in records)))
        counters['pending'].set(str(sum(r['state']!='ready' for r in records)))
        counters['uploaded'].set(str(sum(r['state'] in ('uploaded', 'unknown', 'changed') for r in records)))
        needle = search_text(query.get().strip())
        for record in records:
            f = filter_value.get()
            if f == 'Chưa sẵn sàng' and record['state']=='ready':
                continue
            if f not in ('Tất cả', 'Chưa sẵn sàng') and record['status']!=f:
                continue
            if needle and needle not in search_text(json.dumps(record, ensure_ascii=False)):
                continue
            data = record['data']
            visible[record['key']] = record
            tree.insert('', 'end', iid=record['key'], values=(data['contract'], data['certificate'], data['client'],
                record['employee_name'], record['status'], record.get('change', ''), Path(record['source']).name), tags=(record['state'],))
        if previous & visible.keys():
            tree.selection_set(list(previous & visible.keys()))
        table_meta.set(f'{len(visible)} / {len(records)} chứng thư')
        if visible:
            empty.place_forget()
        else:
            empty.place(relx=0.5, rely=0.5, anchor='center')
        show_selection()

    def sort_column(column):
        descending = not sort_state.get(column, True)
        sort_state[column] = descending
        def natural(key):
            return [int(part) if part.isdigit() else part.casefold() for part in re.split(r'(\d+)', tree.set(key, column))]
        for i,key in enumerate(sorted(tree.get_children(), key=natural, reverse=descending)):
            tree.move(key, '', i)
        for col, name in zip(columns, names):
            tree.heading(col, text=name + (' ↓' if descending else ' ↑') if col==column else name)
    for column in columns:
        tree.heading(column, command=lambda c=column: sort_column(c))

    def scan(event=None):
        if busy:
            return
        path = Path(excel_file.get())
        if not excel_file.get() or not path.is_file() or path.suffix.lower()!='.xlsx' or path.name.startswith('~$'):
            messagebox.showerror('Chọn nguồn dữ liệu', 'Chọn một tệp Excel .xlsx hợp lệ trước khi quét.', parent=root)
            return
        files = [path]
        store.save_setting('excel_file', str(path))
        set_busy(True)
        banner.set(f'Đang đọc {path.name}…')
        def worker():
            loaded, failures = [], []
            try:
                for file in files:
                    try:
                        loaded.extend(read_file(file))
                    except Exception as error:
                        failures.append(f'{file.name}: {error}')
                tasks.put(('scan', (loaded, failures, None)))
            except Exception as error:
                tasks.put(('scan', ([], [], str(error))))
        threading.Thread(target=worker, daemon=True).start()

    def poll():
        if closing:
            return
        try:
            while True:
                kind, payload = tasks.get_nowait()
                if kind == 'scan':
                    loaded, failures, error = payload
                    try:
                        if error:
                            messagebox.showerror('Không thể quét', error, parent=root)
                        else:
                            store.scan_records(loaded, failures)
                            if failures:banner.set(f'Quét chưa hoàn tất · Giữ nguyên danh sách và lịch sử · {len(failures)} tệp chưa đọc được')
                            else:
                                stats=store.last_scan_stats
                                banner.set(f'Đã quét {len(loaded)} chứng thư · Chuyển nguồn: {stats["moved"]} · Khôi phục: {stats["restored"]} · Lưu trữ: {stats["archived"]}')
                            refresh()
                            if failures:
                                messagebox.showwarning('Một số tệp chưa đọc được', '\n'.join(failures), parent=root)
                    except Exception as error:
                        messagebox.showerror('Không thể quét',str(error),parent=root)
                    finally:
                        set_busy(False)
                elif kind == 'progress':
                    banner.set(payload)
                elif kind == 'refresh':
                    refresh()
                elif kind == 'upload_done':
                    results, skipped, error = payload
                    set_busy(False)
                    refresh()
                    finish_upload(results, skipped, error)
        except queue.Empty:
            pass
        root.after(200, poll)

    def details(event=None):
        rows = selected()
        if not rows:
            return
        dialog = tk.Toplevel(root)
        dialog.title(app_name + ' • Chi tiết chứng thư')
        dialog.geometry('760x480')
        dialog.configure(background=BG)
        label(dialog, 'Thông tin chứng thư', INK, ('Segoe UI', 17, 'bold'), bg=BG).pack(anchor='w', padx=20, pady=(20, 12))
        box = ScrolledText(dialog, font=('Segoe UI', 11), wrap='word', relief='flat', borderwidth=0,
                           background=WHITE, foreground=INK, padx=16, pady=16)
        box.pack(fill='both', expand=True, padx=20)
        for record in rows:
            data = record['data']
            box.insert('end', f"Hợp đồng: {data['contract']}\nChứng thư: {data['certificate']}\nKhách hàng: {data['client']}\nNhân viên: {record['employee_name']}\nTrạng thái: {record['status']}\nLần tải gần nhất: {record.get('upload_time','')}\n\n{record['detail']}\n\nNguồn: {record['source']}\nDòng Excel: {record['row']}\n{record['url']}\n\n")
        box.configure(state='disabled')
        ttk.Button(dialog, text='Đóng', command=dialog.destroy).pack(anchor='e', padx=20, pady=14)
        dialog.transient(root)

    def show_history():
        dialog=tk.Toplevel(root);dialog.title(app_name+' • Lưu trữ / Lịch sử');dialog.geometry('1050x650')
        label(dialog,'Chứng thư lưu trữ giữ nguyên lịch sử tải lên. Khi có lại trong nguồn và quét thành công, ứng dụng tự khôi phục.',bg=BG,wraplength=980).pack(fill='x',padx=18,pady=12)
        tabs=ttk.Notebook(dialog);tabs.pack(fill='both',expand=True,padx=18,pady=(0,16))
        archived=tk.Frame(tabs,background=WHITE);log=tk.Frame(tabs,background=WHITE)
        tabs.add(archived,text='Chứng thư lưu trữ');tabs.add(log,text='Lịch sử thao tác')
        archived.columnconfigure(0,weight=1);archived.rowconfigure(0,weight=2);archived.rowconfigure(1,weight=1)
        listing=ttk.Treeview(archived,columns=('contract','certificate','client','status'),show='headings',height=8)
        for col,title,width in [('contract','Hợp đồng',120),('certificate','Chứng thư',180),('client','Khách hàng',410),('status','Trạng thái',230)]:
            listing.heading(col,text=title);listing.column(col,width=width,minwidth=65)
        listing.grid(row=0,column=0,sticky='nsew')
        scroll=ttk.Scrollbar(archived,orient='vertical',command=listing.yview);scroll.grid(row=0,column=1,sticky='ns');listing.configure(yscrollcommand=scroll.set)
        info=ScrolledText(archived,height=9,font=('Segoe UI',10),wrap='word');info.grid(row=1,column=0,columnspan=2,sticky='nsew')
        rows={r['key']:r for r in store.records(include_archived=True) if r['archived']}
        for key,r in sorted(rows.items(),key=lambda item:(item[1]['data']['contract'],item[1]['data']['certificate'])):
            d=r['data'];listing.insert('','end',iid=key,values=(d['contract'],d['certificate'],d['client'],r['status']))
        def show(event=None):
            chosen=listing.selection()
            if not chosen:return
            r=rows[chosen[0]];info.configure(state='normal');info.delete('1.0','end')
            info.insert('end',f'{r["archive_reason"]}\n\n{r["detail"]}\n\nNguồn trước: {r["source"]}\nNgày lưu trữ: {r["archive_time"][:19].replace("T"," ")}\nLần tải gần nhất: {r.get("upload_time","")}\n{r["url"]}')
            info.configure(state='disabled')
        listing.bind('<<TreeviewSelect>>',show)
        box=ScrolledText(log,font=('Segoe UI',10),wrap='word');box.pack(fill='both',expand=True)
        actions={'source_moved':'Đã chuyển nguồn','restored':'Đã khôi phục','archived':'Đã lưu trữ','content_changed':'Excel có thay đổi','scan':'Quét nguồn','upload_start':'Bắt đầu tải','upload_submitting':'Đang chờ website xác nhận','upload_success':'Tải lên thành công','upload_uncertain':'Chưa xác nhận tải lên','upload_failed':'Tải lên chưa thành công','upload_blocked':'Cần đối chiếu website','uploaded':'Xác nhận đã tải lên','cancel':'Đánh dấu hủy','new':'Xác nhận chứng thư mới','approve':'Đã kiểm tra ghi chú','reset':'Bỏ xác nhận thủ công'}
        all_rows={r['key']:r for r in store.records(include_archived=True)}
        for item in store.history_log():
            r=all_rows.get(item['key']);name=f'{r["data"]["contract"]} · {r["data"]["certificate"]} · {r["data"]["client"]}' if r else item['key']
            box.insert('end',f'{item["time"][:19].replace("T"," ")}  {actions.get(item["action"],item["action"])}\n{name}\n{item["detail"]}\n\n')
        box.configure(state='disabled');dialog.transient(root)

    def decide(action):
        rows = selected()
        if not rows or busy:
            return
        url = ''
        if action=='uploaded':
            if len(rows)!=1:
                messagebox.showwarning('Chọn một chứng thư', 'Xác nhận từng chứng thư với đúng liên kết trên website.', parent=root)
                return
            url = simpledialog.askstring('Xác nhận đã tải lên', 'Dán liên kết hồ sơ đã lưu trên hoso.Contract Desk.vn:', initialvalue=rows[0]['url'], parent=root)
            if not url:
                return
            if not re.match(r'^https://hoso\.Contract Desk\.vn/(?:wp-admin/post\.php\?|[^\s]+)', url):
                messagebox.showerror('Kiểm tra liên kết', 'Cần liên kết hồ sơ thuộc hoso.Contract Desk.vn.', parent=root)
                return
        else:
            prompts = {'cancel':'Đánh dấu hủy các mục đã chọn?', 'approve':'Bạn đã kiểm tra chữ đỏ và ghi chú, xác nhận có thể tải lên?',
                       'new':'Bạn đã đối chiếu website và xác nhận CHỨNG THƯ này chưa được tải lên?',
                       'reset':'Bỏ xác nhận thủ công và trạng thái đã tải lên của mục đã chọn? Lịch sử thao tác vẫn được giữ.'}
            if not messagebox.askyesno('Xác nhận thao tác', prompts[action], parent=root):
                return
        try:
            for record in rows:
                store.decide(record['key'], action, url)
        except Exception as error:
            messagebox.showerror('Chưa thể xác nhận', str(error), parent=root)
        refresh()

    def export(mode='ready'):
        if busy:
            return
        def matches(record):
            return record['current'] and ((record['state']=='ready') if mode=='ready' else (record['state'] in ('review', 'missing')))
        rows = [r for r in (selected() or list(visible.values())) if matches(r)]
        if not rows:
            guidance = 'Chỉ xuất các mục Cần kiểm tra hoặc Thiếu thông tin.' if mode=='not_ready' else 'Chỉ xuất các mục Sẵn sàng tải lên.'
            messagebox.showwarning('Không có mục phù hợp', guidance + '\nKhông có mục phù hợp trong danh sách đã chọn hoặc bộ lọc hiện tại.', parent=root)
            return
        output = base/'exports'
        output.mkdir(exist_ok=True)
        sources = {item['source'] for r in rows for item in (r.get('sources') or [r])}
        prefix = 'SanSang' if mode=='ready' else 'ChuaSanSang'
        if len(sources)==1:
            path = filedialog.asksaveasfilename(parent=root, title='Xuất Excel giữ định dạng gốc', initialdir=output,
                initialfile=f'{prefix}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx', defaultextension='.xlsx', filetypes=[('Excel workbook','*.xlsx')])
        else:
            path = filedialog.askdirectory(parent=root, title='Chọn thư mục xuất: mỗi tệp gốc → một tệp Excel', initialdir=output)
        if not path:
            return
        try:
            count = store.export({r['key'] for r in rows}, path, mode=mode)
            guidance = 'Bạn có thể tải trực tiếp bằng nút Tải lên website trong ứng dụng. Tải bằng ứng dụng cũ cần xác nhận thủ công.' if mode=='ready' else 'Chỉ gồm Cần kiểm tra và Thiếu thông tin. Xem hai cột Trạng thái theo dõi và Lý do chưa sẵn sàng ở bên phải.'
            messagebox.showinfo('Xuất Excel thành công', f'{count} dòng chứng thư → {path}\n\nGiữ màu, thiết kế và các cột gốc.\n{guidance}\n\nXuất Excel không thay đổi trạng thái tải lên.', parent=root)
            banner.set(f'Đã xuất {count} dòng chứng thư  ·  Giữ định dạng Excel gốc')
        except Exception as error:
            messagebox.showerror('Không xuất được', str(error), parent=root)

    def open_source():
        rows = selected()
        if rows:
            try:
                os.startfile(rows[0]['source'])
            except OSError as error:
                messagebox.showerror('Không mở được Excel', str(error), parent=root)
    def open_web():
        rows = selected()
        webbrowser.open(rows[0]['url'] if rows and rows[0]['url'] else 'https://hoso.Contract Desk.vn/wp-admin/edit.php?post_type=appraisal_case')

    def stop_upload():
        upload_stop.set()
        stop_button.configure(state='disabled', text='Đang dừng…')
        banner.set('Sẽ dừng sau hồ sơ đang xử lý. Các kết quả đã lưu vẫn được giữ.')

    def show_report(content, path=None):
        dialog = tk.Toplevel(root)
        dialog.title(app_name + ' • Kết quả tải lên')
        dialog.geometry('880x600')
        box = ScrolledText(dialog, font=('Segoe UI', 10), wrap='word', padx=14, pady=14)
        box.pack(fill='both', expand=True, padx=12, pady=12)
        box.insert('1.0', content)
        box.configure(state='disabled')
        buttons = ttk.Frame(dialog, padding=12)
        buttons.pack(fill='x')
        if path:
            ttk.Button(buttons, text='Mở báo cáo', command=lambda: os.startfile(str(path))).pack(side='left')
        ttk.Button(buttons, text='Đóng', command=dialog.destroy).pack(side='right')
        dialog.transient(root)

    def finish_upload(results, skipped, error):
        names = {r['key']: r for r in upload_context}
        # Recover any committed results if an unexpected worker error interrupted reporting.
        received = {r['key'] for r in results}
        for key, state, url, detail, updated in store.db.execute('SELECT key,state,url,detail,updated FROM uploads'):
            if key in names and key not in received and updated >= upload_started and state in ('success', 'failed', 'uncertain', 'blocked'):
                results.append(dict(key=key, state=state, url=url, detail=detail))
        success = sum(r['state']=='success' for r in results)
        attempted = {r['key'] for r in results}
        remaining = [r for r in upload_context if r['state']=='ready' and r['key'] not in attempted]
        labels = {'success':'ĐÃ TẢI LÊN', 'blocked':'CẦN ĐỐI CHIẾU', 'uncertain':'CHƯA XÁC NHẬN — CẦN KIỂM TRA', 'failed':'THẤT BẠI TRƯỚC KHI GỬI'}
        lines = [app_name, f'{dt.datetime.now():%d/%m/%Y %H:%M:%S}',
                 f'Thành công: {success} | Cần kiểm tra / lỗi: {len(results)-success} | Bỏ qua: {len(skipped)} | Chưa thực hiện: {len(remaining)}', '']
        if error:
            lines.extend(['Lỗi lượt chạy: ' + error, ''])
        for item in results:
            record = names[item['key']]
            d = record['data']
            lines.append(f"{labels[item['state']]} | HĐ {d['contract']} | CT {d['certificate']} | {d['client']}\n{item.get('detail','')}\n{item.get('url','')}\n")
        for record in skipped + remaining:
            d = record['data']
            reason = (record['status'] + ': ' + record['detail']) if record in skipped else 'Đã dừng / chưa thực hiện.'
            lines.append(f"BỎ QUA | HĐ {d['contract']} | CT {d['certificate']} | {d['client']}\n{reason}\n")
        lines.append('Kết quả từng hồ sơ được lưu ngay trong lịch sử của ứng dụng. Chỉ hồ sơ được website xác nhận mới tự chuyển Đã tải lên.')
        content = '\n'.join(lines)
        path = None
        try:
            reports = base / 'reports'
            reports.mkdir(exist_ok=True)
            path = reports / f'UploadReport_{dt.datetime.now():%Y%m%d_%H%M%S_%f}.txt'
            path.write_text(content, encoding='utf-8-sig')
            store.save_setting('last_upload_report', str(path))
        except OSError as report_error:
            content += '\nKhông lưu được tệp báo cáo: ' + str(report_error)
            path = None
        banner.set(f'Đã hoàn tất lượt tải · {success} thành công · {len(results)-success} cần kiểm tra / lỗi')
        root.bell()
        show_report(content, path)

    def upload():
        nonlocal upload_context
        if busy:
            return
        rows = selected() or list(visible.values())
        ready = [r for r in rows if r['current'] and r['state']=='ready']
        skipped = [r for r in rows if r not in ready]
        if not ready:
            messagebox.showwarning('Chưa có mục sẵn sàng', 'Chọn chứng thư sẵn sàng hoặc lọc danh sách Sẵn sàng tải lên.', parent=root)
            return
        browser_folder = base / 'browsers'
        if not browser_folder.is_dir():
            messagebox.showerror('Thiếu trình duyệt', 'Giữ thư mục browsers bên cạnh Contract Desk.exe.', parent=root)
            return
        dialog = tk.Toplevel(root)
        dialog.title(app_name + ' • Xem lại và tải lên')
        dialog.geometry('820x650')
        dialog.minsize(660, 540)
        dialog.configure(background=BG)
        label(dialog, f'Tải {len(ready)} chứng thư lên website', INK, ('Segoe UI', 16, 'bold'), bg=BG).pack(anchor='w', padx=20, pady=(18, 6))
        label(dialog, f'{len(skipped)} mục khác được bỏ qua. Đóng cửa sổ này để thay đổi lựa chọn.', MUTED, bg=BG).pack(anchor='w', padx=20, pady=(0, 8))
        box = ScrolledText(dialog, height=10, font=('Segoe UI', 10), wrap='word', padx=10, pady=10)
        box.pack(fill='both', expand=True, padx=20)
        for r in ready:
            d = r['data']
            box.insert('end', f"HĐ {d['contract']} | CT {d['certificate']} | {d['certificate_date']}\n{d['client']} | {r['employee_name']} | Phí {d['fee']}\n\n")
        if skipped:
            box.insert('end', '\nCÁC MỤC BỎ QUA\n')
            for r in skipped:
                box.insert('end', f"HĐ {r['data']['contract']} | CT {r['data']['certificate']} | {r['status']}: {r['detail']}\n")
        box.configure(state='disabled')
        account = tk.Frame(dialog, background=BG)
        account.pack(fill='x', padx=20, pady=12)
        account.columnconfigure(1, weight=1)
        username = tk.StringVar(value=store.setting('username'))
        password = tk.StringVar()
        label(account, 'Tài khoản website', MUTED, bg=BG).grid(row=0, column=0, sticky='w', padx=(0, 16), pady=5)
        user_entry = ttk.Entry(account, textvariable=username)
        user_entry.grid(row=0, column=1, sticky='ew', pady=5)
        label(account, 'Mật khẩu', MUTED, bg=BG).grid(row=1, column=0, sticky='w', padx=(0, 16), pady=5)
        pass_entry = ttk.Entry(account, textvariable=password, show='●')
        pass_entry.grid(row=1, column=1, sticky='ew', pady=5)
        reveal = tk.BooleanVar()
        ttk.Checkbutton(account, text='Hiện mật khẩu', variable=reveal, command=lambda: pass_entry.configure(show='' if reveal.get() else '●')).grid(row=2, column=1, sticky='w')
        label(dialog, 'Mật khẩu chỉ dùng cho lượt tải này. Ứng dụng kiểm tra website trước mỗi hồ sơ.', MUTED, ('Segoe UI', 9), bg=BG).pack(anchor='w', padx=20)
        buttons = ttk.Frame(dialog, padding=16)
        buttons.pack(fill='x')
        def begin():
            nonlocal upload_context, upload_started
            user, secret = username.get().strip(), password.get()
            if not user or not secret:
                messagebox.showwarning('Thiếu tài khoản', 'Nhập tài khoản và mật khẩu website.', parent=dialog)
                return
            # Check all selected source rows before opening the browser.
            try:
                for record in ready:
                    store.validate_record(record['key'])
                    from website_upload import money
                    money(record['data']['fee'])
            except Exception as error:
                messagebox.showerror('Quét / kiểm tra lại dữ liệu', str(error), parent=dialog)
                return
            store.save_setting('username', user)
            upload_context = rows
            upload_started = dt.datetime.now().isoformat()
            password.set('')
            dialog.destroy()
            upload_stop.clear()
            set_busy(True, 'upload')
            banner.set('Đang mở trình duyệt để tải lên…')
            def worker():
                results, error = [], ''
                try:
                    from website_upload import run_upload
                    results = run_upload(base/'contract_tracker.sqlite3', [r['key'] for r in ready], user, secret,
                                         browser_folder, upload_stop, lambda kind, payload: tasks.put((kind, payload)))
                except Exception as exc:
                    error = str(exc).split('Call log:')[0][:1500]
                tasks.put(('upload_done', (results, skipped, error)))
            threading.Thread(target=worker, daemon=True).start()
        ttk.Button(buttons, text='Quay lại', command=dialog.destroy).pack(side='left')
        ttk.Button(buttons, text=f'Bắt đầu tải {len(ready)} chứng thư', style='Primary.TButton', command=begin).pack(side='right')
        dialog.transient(root)
        dialog.grab_set()
        (pass_entry if username.get() else user_entry).focus_set()

    scan_button.configure(command=scan)
    focus_scan.configure(command=scan)
    ready_button.configure(command=lambda: export('ready'))
    pending_button.configure(command=lambda: export('not_ready'))
    details_button.configure(command=details)
    upload_button.configure(command=upload)
    stop_button.configure(command=stop_upload)
    history_button.configure(command=show_history)
    menu.add_command(label='Xác nhận đã tải bằng ứng dụng khác', command=lambda: decide('uploaded'))
    menu.add_command(label='Mở báo cáo tải gần nhất', command=lambda: os.startfile(store.setting('last_upload_report')) if Path(store.setting('last_upload_report')).is_file() else messagebox.showinfo('Báo cáo', 'Chưa có báo cáo tải lên.', parent=root))
    menu.add_separator()
    menu.add_command(label='Mở Excel gốc', command=open_source)
    menu.add_command(label='Mở hồ sơ trên website', command=open_web)
    menu.add_separator()
    menu.add_command(label='Đánh dấu hủy', command=lambda: decide('cancel'))
    menu.add_command(label='Đã kiểm tra ghi chú', command=lambda: decide('approve'))
    menu.add_command(label='Chứng thư chưa có trên website', command=lambda: decide('new'))
    menu.add_separator()
    menu.add_command(label='Bỏ xác nhận thủ công', command=lambda: decide('reset'))
    tree.bind('<<TreeviewSelect>>', show_selection)
    tree.bind('<Double-1>', details)
    query.trace_add('write', refresh)
    filter_value.trace_add('write', refresh)
    list_focus = store.setting('list_focus', '0') == '1'
    def apply_list_focus(save=True):
        if list_focus:
            metrics.grid_remove()
            source.grid_remove()
            focus_scan.grid()
            focus_button.configure(text='Hiện tổng quan')
        else:
            metrics.grid()
            source.grid()
            focus_scan.grid_remove()
            focus_button.configure(text='Mở rộng danh sách')
        if save:
            store.save_setting('list_focus', '1' if list_focus else '0')
    def toggle_list_focus(event=None):
        nonlocal list_focus
        list_focus = not list_focus
        apply_list_focus()
    focus_button.configure(command=toggle_list_focus)
    root.bind('<F11>', toggle_list_focus)
    root.bind('<F5>', scan)
    root.bind('<Control-f>', lambda e: (search_entry.focus_set(), search_entry.selection_range(0, 'end')))
    def close():
        nonlocal closing
        if busy_kind == 'upload':
            stop_upload()
            messagebox.showinfo('Đang hoàn tất hồ sơ hiện tại', 'Ứng dụng sẽ dừng sau hồ sơ đang xử lý. Khi có báo cáo kết quả, bạn có thể đóng cửa sổ.', parent=root)
            return
        closing = True
        root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    apply_list_focus(save=False)
    refresh()
    poll()
    try:
        root.mainloop()
    finally:
        store.db.close()
