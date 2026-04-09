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
                                  'po_no': po_no, 'due_date': '', 'address': ''})

    if not items:
        raise ValueError("No valid item lines found in PDF")
    log(log_widget, f"Successfully parsed {len(items)} items")
    return items


def extract_ship_to_address(lines):
    """从文本行中提取 SHIP TO 地址"""
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


def process_thermofisher(pdf_path, db_path, log_widget):
    """处理 Thermo Fisher Purchase Order 格式"""
    wb_db = openpyxl.load_workbook(db_path)
    ws_db = wb_db['SAP Database ']
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

        # 提取 EST. DELIVERY DATE
        due_date = ''
        for i, line in enumerate(lines0):
            if 'EST. DELIVERY DATE' in line and i + 1 < len(lines0):
                date_match = re.search(r'\d+/\d+/\d+', lines0[i + 1])
                if date_match:
                    due_date = date_match.group(0)
                    break
        log(log_widget, f"Due Date: {due_date}" if due_date else "WARNING: Due date not found")

        # 提取 SHIP TO 地址
        address = extract_ship_to_address(lines0)
        log(log_widget, f"Ship To: {address}" if address else "WARNING: Address not found")

        # 遍历所有页提取物料
        items = []
        for page in pdf.pages:
            text = page.extract_text()
            for line in text.split('\n'):
                code_match = re.search(r'\b(J[A-Z0-9]+-[A-Z0-9]+)\b', line)
                if code_match:
                    code = code_match.group(1)
                    qty_match = re.match(r'(\d+)/', line.strip())
                    qty = qty_match.group(1) if qty_match else ''
                    nums = re.findall(r'\d+\.\d+', line)
                    rate = nums[-2] if len(nums) >= 2 else (nums[0] if nums else '')
                    after_code = line[code_match.end():].strip()
                    desc = re.sub(r'\s+\d+\.\d+\s+\$?\d+\.\d+\s*$', '', after_code).strip()
                    sap_id = db.get(code, '')
                    if not sap_id:
                        log(log_widget, f"  WARNING: {code} not found in Database, AL will be empty")
                    items.append({'sap_id': sap_id, 'desc': desc, 'qty': qty, 'rate': rate,
                                  'po_no': order_no, 'due_date': due_date, 'address': address})

    if not items:
        raise ValueError("No valid item lines found in PDF")
    log(log_widget, f"Successfully parsed {len(items)} items")
    return items


def write_to_excel(items, template_path, output_path, log_widget):
    """复制模板并写入数据到新文件"""
    # 复制模板到输出路径
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    # 清除模板中第3行以后的旧数据
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

        if item['due_date']:
            ws[f"AV{r}"] = f"Due date: {item['due_date']}"
        if item['address']:
            ws[f"AG{r}"] = item['address']

    wb.save(output_path)
    wb.close()
    log(log_widget, f"Done. New file saved to: {output_path}")
    messagebox.showinfo("Success", f"New Excel file created:\n{output_path}")


def process_data(pdf_path, template_path, output_path, db_path, pdf_type, log_widget):
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

        write_to_excel(items, template_path, output_path, log_widget)

    except Exception as e:
        log(log_widget, f"Error: {str(e)}")
        messagebox.showerror("Error", str(e))


# GUI Layout
root = tk.Tk()
root.title("SAP Data Processing Tool")
root.geometry("540x700")
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

ttk.Label(frame, text="Database File (Thermo Fisher only):").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=db_var).pack(fill=tk.X)
ttk.Button(frame, text="Select Database", command=lambda: db_var.set(filedialog.askopenfilename(
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]))).pack(pady=3)

ttk.Button(frame, text="Start Processing",
           command=lambda: process_data(pdf_var.get(), template_var.get(), output_var.get(),
                                        db_var.get(), pdf_type_var.get(), log_area)
           ).pack(pady=8)

log_area = scrolledtext.ScrolledText(frame, height=12, width=55)
log_area.pack(pady=5)

root.mainloop()