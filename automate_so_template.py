import pdfplumber
import openpyxl
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


def log(widget, message):
    widget.insert(tk.END, message + "\n")
    widget.see(tk.END)


def process_data(pdf_path, excel_path, log_widget):
    try:
        if not pdf_path or not excel_path:
            raise ValueError("Please select files")

        log(log_widget, "Reading PDF...")
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text()
            cid_match = re.search(r"ID:\s*(\d+)", text)
            customer_id = cid_match.group(1) if cid_match else "1001364"

            # 动态提取 PO 号，兼容两种格式
            po_match = re.search(r'P\.O\.No\..*\n(\S+)', text)
            if not po_match:
                po_match = re.search(r'P[.\s]*O[.\s]*No[.\s]*.*\n(\S+)', text)
            po_no = po_match.group(1) if po_match else ""
            if not po_no: raise ValueError("PO No. not found in PDF")
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
                        items.append({'code': code, 'desc': desc, 'qty': qty, 'rate': rate})

            if not items: raise ValueError("No valid item lines found in PDF")
            log(log_widget, f"Successfully parsed {len(items)} items")

        # Write to Excel
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        start_row = 3
        while ws[f"A{start_row}"].value is not None:
            start_row += 1

        for i, item in enumerate(items):
            r = start_row + i
            data = {'A': 1, 'B': 3220, 'C': 10, 'E': "00", 'G': "ZOR1", 'Q': "B014", 'R': 3130,
                    'S': "East China-2", 'T': 118, 'U': "HQ-Domestic", 'AC': "01", 'AF': "D002",
                    'AP': "PR00", 'AS': 3200}
            for col, val in data.items(): ws[f"{col}{r}"] = val

            for col in ['I', 'K', 'M', 'O']: ws[f"{col}{r}"] = customer_id
            ws[f"AB{r}"] = ws[f"AR{r}"] = "USD"
            ws[f"AD{r}"] = ws[f"AE{r}"] = po_no
            ws[f"AK{r}"] = (i + 1) * 10
            ws[f"AL{r}"] = item['code']
            ws[f"AM{r}"] = item['desc']
            ws[f"AO{r}"] = item['qty']
            ws[f"AQ{r}"] = item['rate']

        wb.save(excel_path)
        wb.close()
        log(log_widget, f"Done. Data written starting from row {start_row}")
        messagebox.showinfo("Success", "Data has been updated in Excel")

    except Exception as e:
        log(log_widget, f"Error: {str(e)}")
        messagebox.showerror("Error", str(e))


# GUI Layout
root = tk.Tk()
root.title("SAP Data Processing Tool")
root.geometry("500x550")
frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

pdf_var = tk.StringVar()
excel_var = tk.StringVar()

ttk.Label(frame, text="PDF File:").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=pdf_var).pack(fill=tk.X)
ttk.Button(frame, text="Select PDF", command=lambda: pdf_var.set(filedialog.askopenfilename())).pack(pady=5)

ttk.Label(frame, text="Excel File:").pack(anchor=tk.W)
ttk.Entry(frame, textvariable=excel_var).pack(fill=tk.X)
ttk.Button(frame, text="Select Excel", command=lambda: excel_var.set(filedialog.askopenfilename())).pack(pady=5)

ttk.Button(frame, text="Start Processing", command=lambda: process_data(pdf_var.get(), excel_var.get(), log_area)).pack(pady=10)

log_area = scrolledtext.ScrolledText(frame, height=12, width=50)
log_area.pack(pady=10)

root.mainloop()