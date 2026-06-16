from __future__ import annotations

import argparse
import json
import threading
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from gst_extractor import (
    scan_folder,
    save_gst_workbook,
    InvoiceDraft,
    calculate_taxes,
    generate_invoice_docx,
)

APP_TITLE = "GST Word to Excel Converter Demo"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


def feature_enabled(name: str) -> bool:
    config = load_config()
    return bool(config.get("enabled_features", {}).get(name, False))


def run_conversion(input_folder: str, output_file: str, record_type: str) -> tuple[int, int, int, Path, Path | None]:
    config = load_config()
    records, items = scan_folder(input_folder, record_type=record_type, config=config)
    out, clean_out = save_gst_workbook(output_file, records, items, create_clean_upload=True)
    check_count = sum(1 for r in records if r.status == "CHECK")
    error_count = sum(1 for r in records if r.status == "ERROR")
    return len(records), check_count, error_count, out, clean_out


class GSTSuiteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_data = load_config()
        self.title(APP_TITLE)
        self.geometry("1060x720")
        self.minsize(980, 650)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(BASE_DIR / "gst_register.xlsx"))
        self.type_var = tk.StringVar(value=self.config_data.get("default_record_type", "Sales"))
        self.status_var = tk.StringVar(value="Ready. Select invoice folder and output Excel file.")
        self.stats_var = tk.StringVar(value="No conversion run yet.")

        self.inv_no = tk.StringVar()
        self.inv_date = tk.StringVar()
        self.party_name = tk.StringVar()
        self.party_gstin = tk.StringVar()
        self.party_address = tk.StringVar()
        self.hsn = tk.StringVar(value="998346")
        self.description = tk.StringVar(value="FABRIC TESTING / ANALYSIS CHARGES")
        self.taxable = tk.StringVar(value="0")
        self.freight = tk.StringVar(value="0")
        self.special_rate = tk.StringVar(value="18")
        self.invoice_output = tk.StringVar(value=str(BASE_DIR / "generated_invoice.docx"))
        self.invoice_status = tk.StringVar(value="Premium feature enabled in Pro Trial: generate Demo-style Word invoices.")

        self._setup_theme()
        self._build_ui()

    def _setup_theme(self):
        self.configure(bg="#eef3f8")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background="#eef3f8", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 9))
        style.configure("TFrame", background="#eef3f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#eef3f8", font=("Segoe UI", 24, "bold"), foreground="#0B2545")
        style.configure("Sub.TLabel", background="#eef3f8", font=("Segoe UI", 10), foreground="#486581")
        style.configure("TLabel", background="#eef3f8", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="#ffffff", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=10)
        style.configure("Horizontal.TProgressbar", troughcolor="#d9e2ec", background="#0B5CAD")

    def _card(self, parent):
        frame = tk.Frame(parent, bg="#ffffff", highlightbackground="#d9e2ec", highlightthickness=1)
        return frame

    def _build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill="x", padx=28, pady=(20, 8))
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        plan = self.config_data.get("app_plan", "Pro Trial")
        ttk.Label(header, text=f"Desktop GST automation product shell | Current plan: {plan}", style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=24, pady=12)
        self.convert_tab = ttk.Frame(self.tabs)
        self.invoice_tab = ttk.Frame(self.tabs)
        self.plan_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.convert_tab, text="GST Converter")
        self.tabs.add(self.invoice_tab, text="Invoice Generator (Pro)")
        self.tabs.add(self.plan_tab, text="Startup Plans")
        self._build_converter_tab()
        self._build_invoice_tab()
        self._build_plan_tab()

    def _label(self, parent, text, r, c=0):
        tk.Label(parent, text=text, bg="#ffffff", fg="#0B2545", font=("Segoe UI", 10, "bold")).grid(row=r, column=c, sticky="w", padx=(20, 12), pady=12)

    def _build_converter_tab(self):
        card = self._card(self.convert_tab)
        card.pack(fill="x", padx=6, pady=12)
        card.grid_columnconfigure(1, weight=1)
        self._label(card, "Invoice Folder", 0)
        ttk.Entry(card, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", pady=12)
        ttk.Button(card, text="Browse", command=self.pick_input).grid(row=0, column=2, padx=14, pady=12)
        self._label(card, "Output Excel", 1)
        ttk.Entry(card, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", pady=12)
        ttk.Button(card, text="Save As", command=self.pick_output).grid(row=1, column=2, padx=14, pady=12)
        self._label(card, "Record Type", 2)
        ttk.Combobox(card, textvariable=self.type_var, values=["Sales", "Purchase"], width=18, state="readonly").grid(row=2, column=1, sticky="w", pady=12)

        info = self._card(self.convert_tab)
        info.pack(fill="x", padx=6, pady=8)
        txt = (
            "Outputs two files: master workbook + clean READY_UPLOAD workbook. "
            "Validation covers GSTIN, duplicate bills, GST math, freight, state-wise IGST/CGST/SGST logic and totals."
        )
        tk.Label(info, text=txt, justify="left", bg="#ffffff", fg="#334E68", wraplength=920, font=("Segoe UI", 10)).pack(anchor="w", padx=20, pady=14)

        self.progress = ttk.Progressbar(self.convert_tab, mode="indeterminate", length=680)
        self.progress.pack(anchor="w", padx=10, pady=(22, 8))
        tk.Label(self.convert_tab, textvariable=self.status_var, bg="#eef3f8", fg="#0B2545", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10)
        tk.Label(self.convert_tab, textvariable=self.stats_var, bg="#eef3f8", fg="#486581", font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=(4, 8))

        bar = ttk.Frame(self.convert_tab)
        bar.pack(fill="x", padx=6, pady=20)
        ttk.Button(bar, text="Convert to Excel", style="Accent.TButton", command=self.convert).pack(side="right")
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="right", padx=12)

    def _build_invoice_tab(self):
        card = self._card(self.invoice_tab)
        card.pack(fill="x", padx=6, pady=12)
        for col in (1, 3):
            card.grid_columnconfigure(col, weight=1)

        fields = [
            ("Invoice No", self.inv_no), ("Invoice Date", self.inv_date),
            ("Party Name", self.party_name), ("Party GSTIN", self.party_gstin),
            ("Address", self.party_address), ("HSN", self.hsn),
            ("Description", self.description), ("Taxable", self.taxable),
            ("Freight", self.freight), ("GST Rate %", self.special_rate),
        ]
        for idx, (label, var) in enumerate(fields):
            r = idx // 2
            c = 0 if idx % 2 == 0 else 2
            self._label(card, label, r, c)
            ttk.Entry(card, textvariable=var).grid(row=r, column=c + 1, sticky="ew", pady=12, padx=(0, 16))

        self._label(card, "Save Invoice As", 5, 0)
        ttk.Entry(card, textvariable=self.invoice_output).grid(row=5, column=1, columnspan=2, sticky="ew", pady=12)
        ttk.Button(card, text="Save As", command=self.pick_invoice_output).grid(row=5, column=3, padx=14, pady=12)

        tips = self._card(self.invoice_tab)
        tips.pack(fill="x", padx=6, pady=8)
        tk.Label(tips, text="Invoice Generator is placed as a Pro/Plus feature. For now, it is enabled in local Pro Trial so you can test the product idea before adding payments.", justify="left", bg="#ffffff", fg="#334E68", wraplength=920, font=("Segoe UI", 10)).pack(anchor="w", padx=20, pady=14)
        tk.Label(self.invoice_tab, textvariable=self.invoice_status, bg="#eef3f8", fg="#0B2545", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=10)

        bar = ttk.Frame(self.invoice_tab)
        bar.pack(fill="x", padx=6, pady=20)
        ttk.Button(bar, text="Generate Word Invoice", style="Accent.TButton", command=self.generate_invoice).pack(side="right")

    def _build_plan_tab(self):
        row = ttk.Frame(self.plan_tab)
        row.pack(fill="both", expand=True, padx=6, pady=12)
        plans = self.config_data.get("plans", {})
        colors = {"Basic": "#E8F4FD", "Pro": "#FFF4E5", "Plus": "#EDE7F6"}
        for name in ["Basic", "Pro", "Plus"]:
            card = tk.Frame(row, bg="#ffffff", highlightbackground="#d9e2ec", highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=8)
            tk.Label(card, text=name, bg=colors.get(name, "#ffffff"), fg="#0B2545", font=("Segoe UI", 18, "bold")).pack(fill="x", pady=(0, 12), ipady=10)
            for feature in plans.get(name, []):
                tk.Label(card, text="✓ " + feature, bg="#ffffff", fg="#334E68", anchor="w", font=("Segoe UI", 10)).pack(fill="x", padx=18, pady=6)
            if name == "Basic":
                price = "Good first free/trial plan"
            elif name == "Pro":
                price = "Invoice generation + office features"
            else:
                price = "For accountants / multi-business use"
            tk.Label(card, text=price, bg="#ffffff", fg="#0B5CAD", font=("Segoe UI", 10, "bold"), wraplength=260).pack(anchor="w", padx=18, pady=18)

        note = ttk.Frame(self.plan_tab)
        note.pack(fill="x", padx=8, pady=(0, 12))
        tk.Label(note, bg="#eef3f8", fg="#486581", justify="left", font=("Segoe UI", 10), text="Payment/login is not added yet. This version only creates a clean feature structure so the product can later become a real paid desktop app.").pack(anchor="w")

    def pick_input(self):
        folder = filedialog.askdirectory(title="Select folder containing GST invoices")
        if folder:
            self.input_var.set(folder)

    def pick_output(self):
        file = filedialog.asksaveasfilename(title="Save GST Excel Register", defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
        if file:
            self.output_var.set(file)

    def pick_invoice_output(self):
        file = filedialog.asksaveasfilename(title="Save generated invoice", defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file:
            self.invoice_output.set(file)

    def clear(self):
        self.input_var.set("")
        self.output_var.set(str(BASE_DIR / "gst_register.xlsx"))
        self.status_var.set("Ready. Select invoice folder and output Excel file.")
        self.stats_var.set("No conversion run yet.")

    def convert(self):
        input_folder = self.input_var.get().strip()
        output_file = self.output_var.get().strip()
        record_type = self.type_var.get().strip()
        if not input_folder or not Path(input_folder).exists():
            messagebox.showerror("Missing folder", "Please select a valid invoice folder.")
            return
        if not output_file.lower().endswith(".xlsx"):
            messagebox.showerror("Invalid output", "Output file must end with .xlsx")
            return
        self.progress.start(12)
        self.status_var.set("Processing invoices... Please do not close this window.")
        self.stats_var.set("Reading files, validating GST math and preparing Excel...")
        def worker():
            try:
                count, checks, errors, out, clean_out = run_conversion(input_folder, output_file, record_type)
                msg = f"Completed. Master: {out.name} | Ready upload: {clean_out.name if clean_out else 'Not created'}"
                stats = f"Files: {count}   |   Needs checking: {checks}   |   Errors: {errors}"
                self.after(0, lambda: self.status_var.set(msg))
                self.after(0, lambda: self.stats_var.set(stats))
                self.after(0, lambda: messagebox.showinfo("Completed", f"{stats}\n\nMaster file:\n{out}\n\nReady-upload file:\n{clean_out}"))
            except Exception as exc:
                error_msg = str(exc)
                self.after(0, lambda: self.status_var.set("Failed. Check error message."))
                self.after(0, lambda: self.stats_var.set("Conversion failed."))
                self.after(0, lambda: messagebox.showerror("Error", error_msg))
            finally:
                self.after(0, self.progress.stop)
        threading.Thread(target=worker, daemon=True).start()

    def generate_invoice(self):
        if not feature_enabled("invoice_generator"):
            messagebox.showinfo("Premium feature", "Invoice Generator is a Pro/Plus feature.")
            return
        try:
            taxable = float(self.taxable.get() or 0)
            freight = float(self.freight.get() or 0)
            rate = float(self.special_rate.get() or 18)
            taxable_total, cgst, sgst, igst = calculate_taxes(taxable, freight, self.party_gstin.get(), self.config_data.get("business_gstin", "00ABCDE1234A1Z0"), rate)
            draft = InvoiceDraft(
                invoice_no=self.inv_no.get().strip(),
                invoice_date=self.inv_date.get().strip(),
                party_name=self.party_name.get().strip(),
                party_gstin=self.party_gstin.get().strip().upper(),
                party_address=self.party_address.get().strip(),
                hsn=self.hsn.get().strip() or "998346",
                description=self.description.get().strip(),
                taxable=taxable,
                freight=freight,
                cgst=cgst,
                sgst=sgst,
                igst=igst,
                business_name=self.config_data.get("business_name", "DEMO TEXTILE RESEARCH LAB"),
                business_gstin=self.config_data.get("business_gstin", "00ABCDE1234A1Z0"),
            )
            out = generate_invoice_docx(draft, self.invoice_output.get().strip())
            self.invoice_status.set(f"Invoice generated: {out}")
            messagebox.showinfo("Invoice Generated", f"Word invoice created successfully:\n{out}")
        except Exception as exc:
            messagebox.showerror("Invoice generation failed", str(exc))


def main():
    parser = argparse.ArgumentParser(description="GST Word to Excel Converter Demo Startup Edition")
    parser.add_argument("--input", help="Folder containing .docx/.pdf invoices")
    parser.add_argument("--output", help="Output .xlsx path")
    parser.add_argument("--type", choices=["Sales", "Purchase"], default="Sales", help="Record type")
    args = parser.parse_args()
    if args.input and args.output:
        count, checks, errors, out, clean_out = run_conversion(args.input, args.output, args.type)
        print(f"Done: {count} files processed. {checks} need checking, {errors} errors. Master: {out}. Ready upload: {clean_out}")
    else:
        GSTSuiteApp().mainloop()


if __name__ == "__main__":
    main()
