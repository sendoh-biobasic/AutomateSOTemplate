import pdfplumber
import openpyxl
import shutil
import re
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


def log(widget, message):
    widget.insert(tk.END, message + "\n")
    widget.see(tk.END)


def process_biobasic(pdf_path, log_widget):
    """处理 Bio Basic Sales Order 格式"""
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

        po_match = re.search(r'P\.O\.No\..*\n(\S+)', text)
        if not po_match:
            po_match = re.search(r'P[.\s]*O[.\s]*No[.\s]*.*\n(\S+)', text)
        po_no = po_match.group(1) if po_match else ""
        if not po_no:
            raise ValueError("PO No. not found in PDF")
        log(log_widget, f"PO No.: {po_no}")

        items = []
        for line in text.split('\n'):
            code_match = re.search(r'A\d{6,}-\d{4}', line)
            if code_match:
                code = code_match.group(0)
                nums = re.findall(r'[\d\.]+', line)
                if len(nums) >= 3:
                    rate = nums[-2]
                    qty = nums[-3]
                    after_code = line[code_match.end():].strip()
                    desc = re.sub(r'[\d\.]+\s+[\d\.]+\s+[\d\.]+\s*$', '', after_code).strip()
                    items.append({'sap_id': code, 'desc': desc, 'qty': qty, 'rate': rate,
                                  'po_no': po_no, 'due_date': '', 'address': '',
                                  'order_date': '', 'cat_no': '', 'spec_sheet': ''})

    if not items:
        raise ValueError("No valid item lines found in PDF")
    log(log_widget, f"Successfully parsed {len(items)} items")
    return items


def extract_ship_to_address(lines):
    """从文本行中提取 SHIP TO 地址，只返回路名（数字后的部分）"""
    ship_to_idx = None
    vendor_idx = None
    for i, line in enumerate(lines):
        if line.startswith('SHIP TO') and ship_to_idx is None:
            ship_to_idx = i
        if line.startswith('VENDOR') and ship_to_idx is not None:
            vendor_idx = i
            break

    if ship_to_idx is None:
        return ''

    end_idx = vendor_idx if vendor_idx else ship_to_idx + 6
    noise = ['CORRESPONDENCE', 'F.O.B.', 'EST. DELIVERY DATE', 'FREIGHT', 'TERMS', 'NET 30']
    address_lines = []
    for line in lines[ship_to_idx + 1:end_idx]:
        clean = line
        for n in noise:
            clean = clean.split(n)[0].strip()
        clean = re.sub(r'\s+\d+/\d+/\d+\s*$', '', clean).strip()
        if clean:
            address_lines.append(clean)

    return ', '.join(address_lines)


def extract_road_name(full_address_lines):
    """从地址行中提取路名（去掉前面的数字，只保留路名）
    例: '220 NECK ROAD' -> 'Neck Road'
    """
    for line in full_address_lines:
        # 匹配 "数字 路名" 格式的行
        m = re.match(r'^\d+\s+(.+)$', line.strip())
        if m:
            road = m.group(1).strip()
            # 转为 Title Case
            return road.title()
    return ''


def parse_order_date(raw_date):
    """将 m/dd/yy 格式转换为 yyyy.mm.dd
    例: '4/01/26' -> '2026.04.01'
    """
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})$', raw_date.strip())
    if not m:
        return raw_date
    month, day, year = m.group(1), m.group(2), m.group(3)
    if len(year) == 2:
        year = '20' + year
    return f"{year}.{month.zfill(2)}.{day.zfill(2)}"


def parse_due_date_long(raw_date):
    """将 m/dd/yy 格式转换为 mm/dd/yyyy (用于 2025 Orders M列)
    例: '5/08/26' -> '5/08/2026'
    """
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})$', raw_date.strip())
    if not m:
        return raw_date
    month, day, year = m.group(1), m.group(2), m.group(3)
    if len(year) == 2:
        year = '20' + year
    return f"{month}/{day}/{year}"


def parse_due_date_av(raw_date):
    """将 m/dd/yy 格式转换为 Due date: m/dd/yyyy (用于第一个Excel AV列)
    例: '5/08/26' -> 'Due date: 5/08/2026'
    """
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})$', raw_date.strip())
    if not m:
        return f"Due date: {raw_date}"
    month, day, year = m.group(1), m.group(2), m.group(3)
    if len(year) == 2:
        year = '20' + year
    return f"Due date: {month}/{day}/{year}"


def process_thermofisher(pdf_path, db_path, log_widget):
    """处理 Thermo Fisher Purchase Order 格式"""
    wb_db = openpyxl.load_workbook(db_path)
    # 兼容两种 sheet 名称
    sap_sheet = 'SAP Database ' if 'SAP Database ' in wb_db.sheetnames else 'Product Databse'
    ws_db = wb_db[sap_sheet]
    db = {}
    for row in ws_db.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            db[str(row[0]).strip()] = str(row[1]).strip()
    log(log_widget, f"Database loaded: {len(db)} entries")

    with pdfplumber.open(pdf_path) as pdf:
        text0 = pdf.pages[0].extract_text()
        lines0 = text0.split('\n')

        # 提取 Order Number
        order_no = ''
        for i, line in enumerate(lines0):
            if 'PURCHASE ORDER' in line and 'ORDER NUMBER' in line and i + 1 < len(lines0):
                m = re.search(r'\d{1,2}/\d{2}/\d{2}\s+(\d+)', lines0[i + 1])
                if m:
                    order_no = m.group(1)
                    break
        if not order_no:
            raise ValueError("Order Number not found in PDF")
        log(log_widget, f"Order No.: {order_no}")

        # 提取右上角订单日期（格式 m/dd/yy）
        raw_order_date = ''
        for i, line in enumerate(lines0):
            if 'PURCHASE ORDER' in line and 'ORDER NUMBER' in line and i + 1 < len(lines0):
                date_m = re.search(r'(\d{1,2}/\d{2}/\d{2})\s+\d+', lines0[i + 1])
                if date_m:
                    raw_order_date = date_m.group(1)
                    break
        order_date = parse_order_date(raw_order_date) if raw_order_date else ''
        log(log_widget, f"Order Date: {order_date}" if order_date else "WARNING: Order date not found")

        # 提取 EST. DELIVERY DATE
        raw_due_date = ''
        for i, line in enumerate(lines0):
            if 'EST. DELIVERY DATE' in line and i + 1 < len(lines0):
                date_match = re.search(r'\d+/\d+/\d+', lines0[i + 1])
                if date_match:
                    raw_due_date = date_match.group(0)
                    break
        due_date_av = parse_due_date_av(raw_due_date) if raw_due_date else ''
        due_date_m = parse_due_date_long(raw_due_date) if raw_due_date else ''
        log(log_widget, f"Due Date: {raw_due_date}" if raw_due_date else "WARNING: Due date not found")

        # 提取 SHIP TO 地址行（原始行，用于路名提取）
        ship_to_idx = None
        vendor_idx = None
        for i, line in enumerate(lines0):
            if line.startswith('SHIP TO') and ship_to_idx is None:
                ship_to_idx = i
            if line.startswith('VENDOR') and ship_to_idx is not None:
                vendor_idx = i
                break
        ship_to_raw_lines = []
        if ship_to_idx is not None:
            end_idx = vendor_idx if vendor_idx else ship_to_idx + 6
            noise = ['CORRESPONDENCE', 'F.O.B.', 'EST. DELIVERY DATE', 'FREIGHT', 'TERMS', 'NET 30']
            for line in lines0[ship_to_idx + 1:end_idx]:
                clean = line
                for n in noise:
                    clean = clean.split(n)[0].strip()
                clean = re.sub(r'\s+\d+/\d+/\d+\s*$', '', clean).strip()
                if clean:
                    ship_to_raw_lines.append(clean)

        address = ', '.join(ship_to_raw_lines)
        road_name = extract_road_name(ship_to_raw_lines)
        log(log_widget, f"Ship To: {address}" if address else "WARNING: Address not found")
        log(log_widget, f"Road Name: {road_name}" if road_name else "WARNING: Road name not found")

        # 遍历所有页提取物料，同时检测 spec sheet 链接
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            all_lines.extend(text.split('\n'))

        # 页面噪声行关键词——遇到这些行时跳过，不中断 spec sheet 搜索
        # 页面噪声关键词：跨页时这些行不应中断 spec sheet 搜索
        PAGE_NOISE = [
            'REFER ALL COMMUNICATIONS',
            'Our General Purchase Terms',
            'www.thermofisher.com/PO-Terms',
            'writing. We do not accept',
            'submit, use or refer to',
            'expressly agree to additional',
            'of our Order. These Terms',
            'AN EQUAL OPPORTUNITY',
            'NOTIFICATION OF ANY PRICE',
            'ISO 14001',
            'A part of:',
            'ORDER ENTERED UNDER',
            'COMPLETE SHIPMENT',
            'ACKNOWLEDGED BY',
            'PHONE NUMBER',
            'DATE OF THIS ACKNOWLEDGMENT',
            'THERMO FISHER SCIENTIFIC CHEMICALS INC.',
            'THERMO FISHER SCIENTIFIC',
            'ALFA AESAR',
            'PURCHASE ORDER',
            'FEDERAL ID NO.',
            'THIS ORDER IS SUBJECT',
            '** REPRINT **',
            'NOTE: THE ABOVE',
            'ALL INVOICES, BILLS',
            'CORRESPONDENCE',
            'F.O.B.',
            'EST. DELIVERY DATE',
            'FREIGHT',
            'TERMS',
            'NET 30',
            'SHIP VIA',
            'D-U-N-S',
            'SPECIAL INSTRUCTIONS',
            'SHIP TO',
            'VENDOR',
            'QUANTITY CHG',
            'BIO BASIC INC.',
            'BAILEY AVENUE',
            'AMHERST',
            'Please send electronic invoice',
            'Please confirm price',
            'Please send CoA',
            'Please reference PO',
            'Must include Country',
            'alfaaesar.accountspay',
            'Radcliff Road',
            'Phone: 978',
            'PAGE:',
            '3. VIA',
            '5. PHONE',
            '6. DATE',
        ]

        def is_noise_line(line):
            return any(kw in line for kw in PAGE_NOISE)

        # J 物料码正则：后缀允许字母、数字、# 等符号
        J_CODE_RE = re.compile(r'\b(J[A-Z0-9]+-[A-Z0-9#@.]+)\b')

        items = []
        for idx, line in enumerate(all_lines):
            code_match = J_CODE_RE.search(line)
            if code_match:
                cat_no = code_match.group(1)
                qty_match = re.match(r'(\d+)/', line.strip())
                qty = qty_match.group(1) if qty_match else ''
                nums = re.findall(r'\d+\.\d+', line)
                rate = nums[-2] if len(nums) >= 2 else (nums[0] if nums else '')
                after_code = line[code_match.end():].strip()
                desc = re.sub(r'\s+\d+\.\d+\s+.*$', '', after_code).strip()
                sap_id = db.get(cat_no, '')
                if not sap_id:
                    log(log_widget, f"  WARNING: {cat_no} not found in Database, AL will be empty")

                # 检测后续行是否有 spec sheet 链接（跳过噪声行，不提前终止）
                spec_sheet = ''
                for next_idx in range(idx + 1, min(idx + 60, len(all_lines))):
                    next_line = all_lines[next_idx].strip()
                    if not next_line:
                        continue
                    # 遇到下一个实际物料行才停止
                    if J_CODE_RE.search(next_line) and re.match(r'\d+/', next_line.split(J_CODE_RE.pattern)[0].strip() + 'x'):
                        break
                    if J_CODE_RE.search(next_line) and re.match(r'\d+/', next_line):
                        break
                    # 跳过页面噪声，继续向后找
                    if is_noise_line(next_line):
                        continue
                    if next_line.startswith('https://assets.thermofisher.com'):
                        spec_sheet = 'Spec sheet'
                        break
                    if 'Spec Sheet for this item is not available' in next_line:
                        spec_sheet = ''
                        break

                items.append({
                    'sap_id': sap_id,
                    'desc': desc,
                    'qty': qty,
                    'rate': rate,
                    'po_no': order_no,
                    'due_date_av': due_date_av,
                    'due_date_m': due_date_m,
                    'address': address,
                    'road_name': road_name,
                    'order_date': order_date,
                    'cat_no': cat_no,
                    'spec_sheet': spec_sheet,
                })

    if not items:
        raise ValueError("No valid item lines found in PDF")
    log(log_widget, f"Successfully parsed {len(items)} items")
    return items


def write_to_excel(items, template_path, output_path, log_widget):
    """复制模板并写入数据到新文件"""
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    for i, item in enumerate(items):
        r = 3 + i
        data = {'A': 1, 'B': 3220, 'C': 10, 'E': "00", 'G': "ZOR1", 'Q': "B014", 'R': 3130,
                'S': "East China-2", 'T': 118, 'U': "HQ-Domestic", 'AC': "01", 'AF': "D002",
                'AP': "PR00", 'AS': 3200}
        for col, val in data.items():
            ws[f"{col}{r}"] = val

        for col in ['I', 'K', 'M', 'O']:
            ws[f"{col}{r}"] = 5148715

        ws[f"AB{r}"] = ws[f"AR{r}"] = "USD"
        ws[f"AD{r}"] = ws[f"AE{r}"] = item['po_no']
        ws[f"AK{r}"] = (i + 1) * 10
        ws[f"AL{r}"] = item['sap_id']
        ws[f"AM{r}"] = item['desc']
        ws[f"AO{r}"] = item['qty']
        ws[f"AQ{r}"] = item['rate']

        # 使用新格式: Due date: m/dd/yyyy
        due_date_val = item.get('due_date_av') or item.get('due_date', '')
        if due_date_val:
            ws[f"AV{r}"] = due_date_val

        if item.get('address'):
            ws[f"AG{r}"] = item['address']

    wb.save(output_path)
    wb.close()
    log(log_widget, f"Done. New file saved to: {output_path}")
    messagebox.showinfo("Success", f"New Excel file created:\n{output_path}")


def write_to_orders_excel(items, db_path, log_widget):
    """直接写入并保存到原始 2025 Orders 文件"""
    from openpyxl.styles import Alignment
    log(log_widget, "Writing to 2025 Orders sheet...")
    wb = openpyxl.load_workbook(db_path)
    ws = wb['2025 Orders']

    # 找到最后一行有数据的行（B列，因为A列是Sales Order，可能为空）
    last_row = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 2).value is not None or ws.cell(r, 1).value is not None:
            last_row = r

    next_row = last_row + 1
    log(log_widget, f"Appending from row {next_row}")

    center = Alignment(horizontal='center', vertical='center', wrap_text=False)

    def set_cell(ws, r, c, value):
        cell = ws.cell(r, c)
        cell.value = value
        cell.alignment = center

    for item in items:
        r = next_row

        # H列: 保留 yyyy.mm.dd 格式
        order_date_fmt = item.get('order_date', '')
        # M列: 保留 m/dd/yyyy 格式
        due_date_fmt = item.get('due_date_m', '')

        # B: Customer PO
        set_cell(ws, r, 2, item['po_no'])
        # C: Material (SAP ID)
        set_cell(ws, r, 3, item['sap_id'])
        # D: Product Code (J-code)
        set_cell(ws, r, 4, item.get('cat_no', ''))
        # E: Description
        set_cell(ws, r, 5, item['desc'])
        # F: QTY
        try:
            qty_val = int(item['qty']) if item['qty'] else None
        except (ValueError, TypeError):
            qty_val = item['qty']
        set_cell(ws, r, 6, qty_val)
        # G: VLOOKUP公式（从 Product Databse 获取产品信息）
        set_cell(ws, r, 7, f"=VLOOKUP(D{r},'Product Databse'!B:D,3,)")
        # H: S.O Date (yyyy/mm/dd)
        set_cell(ws, r, 8, order_date_fmt)
        # I: Product Spec Sheet
        spec = item.get('spec_sheet', '')
        set_cell(ws, r, 9, spec if spec else None)
        # M: Original DELIVERY DATE (yyyy/mm/dd)
        set_cell(ws, r, 13, due_date_fmt if due_date_fmt else None)
        # P: Shipping Location
        road = item.get('road_name', '')
        set_cell(ws, r, 16, road if road else None)

        next_row += 1

    wb.save(db_path)
    wb.close()
    log(log_widget, f"2025 Orders updated and saved: {db_path}")
    messagebox.showinfo("Success", f"2025 Orders updated successfully:\n{db_path}")


def process_data(pdf_path, template_path, output_path, db_path, pdf_type,
                 orders_db_path, expected_qty_str, log_widget):
    try:
        if not pdf_path or not template_path or not output_path:
            raise ValueError("Please select PDF, Template and Output files")
        if pdf_type == "Thermo Fisher" and not db_path:
            raise ValueError("Please select Database file for Thermo Fisher PO")

        log(log_widget, f"Reading PDF ({pdf_type})...")

        if pdf_type == "Bio Basic":
            items = process_biobasic(pdf_path, log_widget)
        else:
            items = process_thermofisher(pdf_path, db_path, log_widget)

        # 数量验证
        actual_count = len(items)
        if expected_qty_str.strip():
            try:
                expected_count = int(expected_qty_str.strip())
                if actual_count != expected_count:
                    msg = (f"Item count mismatch!\n\n"
                           f"Expected:  {expected_count} items\n"
                           f"Parsed:    {actual_count} items\n\n"
                           f"Please check the log and verify the PDF before continuing.\n"
                           f"Do you want to continue anyway?")
                    log(log_widget, f"WARNING: Expected {expected_count} items but parsed {actual_count}!")
                    if not messagebox.askyesno("Count Mismatch", msg, icon="warning"):
                        log(log_widget, "Processing cancelled by user.")
                        return
                else:
                    log(log_widget, f"Item count verified: {actual_count} items ✓")
            except ValueError:
                log(log_widget, "WARNING: Invalid expected quantity input, skipping count check.")

        write_to_excel(items, template_path, output_path, log_widget)

        # 如果指定了 2025 Orders 数据库，直接写入原文件
        if orders_db_path:
            write_to_orders_excel(items, orders_db_path, log_widget)

    except Exception as e:
        log(log_widget, f"Error: {str(e)}")
        messagebox.showerror("Error", str(e))


# ── GUI ──────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("SAP Data Processing Tool")
root.geometry("560x820")
frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

# PDF Type Selection
ttk.Label(frame, text="PDF Type:").pack(anchor=tk.W)
pdf_type_var = tk.StringVar(value="Bio Basic")
type_frame = ttk.Frame(frame)
type_frame.pack(anchor=tk.W, pady=2)
ttk.Radiobutton(type_frame, text="Bio Basic Sales Order", variable=pdf_type_var, value="Bio Basic").pack(side=tk.LEFT)
ttk.Radiobutton(type_frame, text="Thermo Fisher PO", variable=pdf_type_var, value="Thermo Fisher").pack(side=tk.LEFT, padx=10)

pdf_var = tk.StringVar()
template_var = tk.StringVar()
output_var = tk.StringVar()
db_var = tk.StringVar()
orders_db_var = tk.StringVar()

ttk.Label(frame, text="PDF File:").pack(anchor=tk.W, pady=(8, 0))
ttk.Entry(frame, textvariable=pdf_var).pack(fill=tk.X)
ttk.Button(frame, text="Select PDF", command=lambda: pdf_var.set(filedialog.askopenfilename(
    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]))).pack(pady=3)

ttk.Label(frame, text="Excel Template File:").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=template_var).pack(fill=tk.X)
ttk.Button(frame, text="Select Template", command=lambda: template_var.set(filedialog.askopenfilename(
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=3)

ttk.Label(frame, text="Output File (save new Excel as):").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=output_var).pack(fill=tk.X)
ttk.Button(frame, text="Select Output Path", command=lambda: output_var.set(filedialog.asksaveasfilename(
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=3)

ttk.Label(frame, text="SAP Database File (Thermo Fisher only):").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=db_var).pack(fill=tk.X)
ttk.Button(frame, text="Select SAP Database", command=lambda: db_var.set(filedialog.askopenfilename(
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=3)

ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=6)

# 物料数量验证
qty_frame = ttk.Frame(frame)
qty_frame.pack(fill=tk.X, pady=(0, 4))
ttk.Label(qty_frame, text="Expected Item Count (optional):").pack(side=tk.LEFT)
expected_qty_var = tk.StringVar()
ttk.Entry(qty_frame, textvariable=expected_qty_var, width=8).pack(side=tk.LEFT, padx=6)
ttk.Label(qty_frame, text="(leave blank to skip check)", foreground="gray").pack(side=tk.LEFT)

ttk.Label(frame, text="2025 Orders Database (optional):").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=orders_db_var).pack(fill=tk.X)
ttk.Button(frame, text="Select Orders Database", command=lambda: orders_db_var.set(filedialog.askopenfilename(
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=3)

ttk.Button(frame, text="Start Processing",
           command=lambda: process_data(
               pdf_var.get(), template_var.get(), output_var.get(),
               db_var.get(), pdf_type_var.get(),
               orders_db_var.get(),
               expected_qty_var.get(),
               log_area)
           ).pack(pady=8)

log_area = scrolledtext.ScrolledText(frame, height=10, width=60)
log_area.pack(pady=5)

root.mainloop()