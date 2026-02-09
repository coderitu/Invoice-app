import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.ttk import Progressbar

from excel_import import load_excel
from outlook_sender import send_test_email
from pdf_generator import generate_invoice_pdf


class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Desktop App")
        self.root.geometry("950x550")

        self.data = None
        self.shared_mailbox = tk.StringVar(value="Membership.accounts@wdclubsafrica.com")
        self.discount_percent = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        top = tk.Frame(self.root)
        top.pack(pady=10)

        tk.Button(top, text="Import Excel", command=self.import_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="Send Test Email", command=lambda: self.send_email(True)).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="Send All Emails", command=lambda: self.send_email(False, True)).pack(side=tk.LEFT, padx=5)

        tk.Label(top, text="Shared Mailbox:").pack(side=tk.LEFT)
        tk.Entry(top, textvariable=self.shared_mailbox, width=30).pack(side=tk.LEFT)

        tk.Label(top, text="Discount (%):").pack(side=tk.LEFT, padx=10)
        tk.Entry(top, textvariable=self.discount_percent, width=8).pack(side=tk.LEFT)

        tk.Label(self.root, text="Email Template:").pack(anchor="w", padx=10)
        self.email_template = tk.Text(self.root, height=6)
        self.email_template.pack(fill="x", padx=10)
        self.email_template.insert(
            "1.0",
            "Dear {name},\n\nPlease find your invoice {invoice_no} attached.\n\nRegards,\nMembership Team"
        )

        columns = ("Member", "Name", "Email", "Renewal", "Invoice no.", "Pending 2025", "Pending 2024", "Pending 2023")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120)
        self.tree.pack(expand=True, fill="both", padx=10, pady=10)

        self.progress = Progressbar(self.root, length=400)
        self.progress.pack(pady=5)

    def import_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        self.data = load_excel(path)
        self.tree.delete(*self.tree.get_children())
        for _, row in self.data.iterrows():
            self.tree.insert("", tk.END, values=list(row))

    def send_email(self, test_mode=True, bulk=False):
        if self.data is None:
            return

        template = self.email_template.get("1.0", tk.END).strip()
        discount = float(self.discount_percent.get()) if self.discount_percent.get() else 0

    # -------- TEST MODE (single selection only) --------
        if test_mode and not bulk:
            selected = self.tree.selection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a member first.")
                return

            values = self.tree.item(selected[0], "values")
            row = dict(zip(self.data.columns, values))
            rows = [row]

    # -------- BULK MODE --------
        else:
            rows = self.data.to_dict(orient="records")

    # -------- PROCESS ROWS --------
        for row in rows:
            pdf_path = generate_invoice_pdf(
                row["Member"],
                row["Name"],
                row["Invoice no."],
                row["Renewal"],
                row["Pending 2025"],
                row["Pending 2024"],
                row["Pending 2023"],
                discount
        )

            send_test_email(
                row["Email"],
                row["Name"],
                row["Invoice no."],
                self.shared_mailbox.get(),
                pdf_path,
                template,
                test_mode=False if bulk else test_mode  # ✅ KEY FIX
        )


if __name__ == "__main__":
    root = tk.Tk()
    InvoiceApp(root)
    root.mainloop()
