# main.py
import sys
import json
import os
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QDateEdit, QInputDialog
)
from PySide6.QtCore import Qt, QDate
from calculations import calculate_totals
from generator import generate_docx
from num2words import num2words

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_MAP_FILE = os.path.join(PROJECT_ROOT, "data", "template_map.json")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")

UOM_OPTIONS = ["Nos", "Kg", "Ltr", "Meter", "Box", "Packet", "Dozen", "Other..."]

def load_template_map():
    with open(TEMPLATE_MAP_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

class QuotationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quotation Automation - PySide6")
        self.resize(1000, 700)
        self.template_map = load_template_map()
        self.init_ui()

    # ---------- Safe getters ----------
    def safe_float(self, r, c):
        item = self.table.item(r, c)
        if item is None:
            return 0.0
        try:
            return float(item.text().strip() or 0)
        except:
            return 0.0

    def safe_text(self, r, c):
        # For cells which may be using a widget (UOM), check widget first
        widget = self.table.cellWidget(r, c)
        if widget is not None:
            # QComboBox or QLineEdit possible
            try:
                from PySide6.QtWidgets import QComboBox, QLineEdit
                if isinstance(widget, QComboBox):
                    return widget.currentText().strip()
                if isinstance(widget, QLineEdit):
                    return widget.text().strip()
            except:
                pass
        item = self.table.item(r, c)
        return item.text().strip() if item else ""

    # ---------- UI ----------
    def init_ui(self):
        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.customer_input = QLineEdit()
        self.gstin_input = QLineEdit()
        self.phone_input = QLineEdit()

        self.template_combo = QComboBox()
        self.template_combo.addItems(list(self.template_map.keys()))

        form_layout.addRow("Date:", self.date_edit)
        form_layout.addRow("Customer Name:", self.customer_input)
        form_layout.addRow("GSTIN:", self.gstin_input)
        form_layout.addRow("Phone:", self.phone_input)
        form_layout.addRow("Template:", self.template_combo)

        layout.addLayout(form_layout)

        # Items table (Description | UOM | Qty | Rate | Tax % | Line Total)
        items_layout = QVBoxLayout()
        h = QHBoxLayout()
        add_row_btn = QPushButton("Add Row")
        remove_row_btn = QPushButton("Remove Selected Row")
        add_row_btn.clicked.connect(self.add_row)
        remove_row_btn.clicked.connect(self.remove_selected_row)
        h.addWidget(add_row_btn)
        h.addWidget(remove_row_btn)
        h.addStretch()
        items_layout.addLayout(h)

        # 6 columns now
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Description", "UOM", "Qty", "Rate", "Tax %", "Line Total"])
        self.table.horizontalHeader().setStretchLastSection(True)
        items_layout.addWidget(self.table)
        layout.addLayout(items_layout)

        # Totals
        totals_layout = QFormLayout()
        self.subtotal_label = QLabel("0.00")
        self.total_tax_label = QLabel("0.00")
        self.grand_total_label = QLabel("0.00")
        totals_layout.addRow("Subtotal:", self.subtotal_label)
        totals_layout.addRow("Total Tax:", self.total_tax_label)
        totals_layout.addRow("Grand Total:", self.grand_total_label)
        layout.addLayout(totals_layout)

        # Buttons
        btn_layout = QHBoxLayout()
        generate_btn = QPushButton("Generate Quotation")
        save_template_btn = QPushButton("Upload Template")
        generate_btn.clicked.connect(self.on_generate)
        save_template_btn.clicked.connect(self.on_upload_template)
        btn_layout.addWidget(generate_btn)
        btn_layout.addWidget(save_template_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Connect signal
        self.table.cellChanged.connect(self.on_table_cell_changed)

        self.setLayout(layout)
        self.add_row()

    # ---------- row helpers ----------
    def make_uom_widget(self, initial=""):
        # Returns a QComboBox set as cell widget; selecting "Other..." opens input
        combo = QComboBox()
        combo.addItems(UOM_OPTIONS)
        if initial and initial in UOM_OPTIONS:
            combo.setCurrentText(initial)
        elif initial:
            # if initial is custom, add it
            combo.addItem(initial)
            combo.setCurrentText(initial)
        # connect
        combo.currentTextChanged.connect(lambda val, c=combo: self.on_uom_changed(c, val))
        return combo

    def on_uom_changed(self, combo_widget, val):
        # If user selects Other..., open text input and replace the widget with a QLineEdit containing custom text
        if val == "Other...":
            text, ok = QInputDialog.getText(self, "Custom UOM", "Enter UOM (e.g., Packs, Sets):")
            if ok and text:
                # replace combo with a line edit (so user can edit content later)
                row, col = self.find_widget_cell(combo_widget)
                if row is not None:
                    line = QLineEdit(text)
                    line.setAlignment(Qt.AlignCenter)
                    # when line edit text changes, we want to recalc if needed
                    line.textChanged.connect(lambda _: self.recalculate())
                    self.table.setCellWidget(row, col, line)
            else:
                # revert selection to first option if canceled
                combo_widget.setCurrentIndex(0)

    def find_widget_cell(self, widget):
        # find the (row, col) containing the widget
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                w = self.table.cellWidget(r, c)
                if w is widget:
                    return r, c
        return None, None

    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)

        desc_item = QTableWidgetItem("")
        desc_item.setFlags(desc_item.flags() | Qt.ItemIsEditable)
        qty_item = QTableWidgetItem("1")
        qty_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        rate_item = QTableWidgetItem("0")
        rate_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        tax_item = QTableWidgetItem("0")
        tax_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        total_item = QTableWidgetItem("0.00")
        total_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        total_item.setFlags(total_item.flags() & ~Qt.ItemIsEditable)

        # place items/widgets
        self.table.setItem(row, 0, desc_item)
        self.table.setCellWidget(row, 1, self.make_uom_widget(""))
        self.table.setItem(row, 2, qty_item)
        self.table.setItem(row, 3, rate_item)
        self.table.setItem(row, 4, tax_item)
        self.table.setItem(row, 5, total_item)

    def remove_selected_row(self):
        selected = self.table.selectionModel().selectedRows()
        for index in sorted(selected, reverse=True):
            self.table.removeRow(index.row())
        self.recalculate()

    # ---------- table change handler ----------
    def on_table_cell_changed(self, row, column):
        # Recalculate only when qty/rate/tax/description changed (we ignore uom widget events here).
        if column not in (2, 3, 4):
            return

        qty = self.safe_float(row, 2)
        rate = self.safe_float(row, 3)
        tax = self.safe_float(row, 4)

        product_total = qty * rate
        tax_amount = product_total * (tax / 100.0)
        final_total = round(product_total + tax_amount, 2)

        self.table.blockSignals(True)

        total_item = self.table.item(row, 5)
        if total_item is None:
            total_item = QTableWidgetItem("0.00")
            total_item.setFlags(total_item.flags() & ~Qt.ItemIsEditable)
            total_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 5, total_item)
        total_item.setText(f"{final_total:.2f}")

        self.table.blockSignals(False)

        self.recalculate()

    # ---------- recalc ----------
    def recalculate(self):
        items = []
        for r in range(self.table.rowCount()):
            # description
            desc = self.safe_text(r, 0)
            # uom: either a widget or a cell text
            uom = ""
            widget = self.table.cellWidget(r, 1)
            if widget is not None:
                try:
                    if isinstance(widget, QComboBox):
                        uom = widget.currentText().strip()
                    else:
                        # QLineEdit or other widget
                        uom = widget.text().strip()
                except:
                    uom = ""
            else:
                # fallback to table item
                uom = self.safe_text(r, 1)

            qty = self.safe_float(r, 2)
            rate = self.safe_float(r, 3)
            tax = self.safe_float(r, 4)

            items.append({
                "description": desc,
                "uom": uom,
                "qty": qty,
                "rate": rate,
                "tax": tax
            })

        calc = calculate_totals(items)
        self.subtotal_label.setText(f"{calc['subtotal']:.2f}")
        self.total_tax_label.setText(f"{calc['total_tax']:.2f}")
        self.grand_total_label.setText(f"{calc['grand_total']:.2f}")

    # ---------- generate ----------
    def on_generate(self):
        date_str = self.date_edit.date().toPython().strftime("%Y-%m-%d")
        quotation_no = f"QT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        customer = self.customer_input.text().strip()
        gstin = self.gstin_input.text().strip()
        phone = self.phone_input.text().strip()

        template_id = self.template_combo.currentText()
        template_path = self.template_map.get(template_id)
        if not template_path or not os.path.exists(template_path):
            QMessageBox.critical(self, "Error", f"Template not found: {template_id}")
            return

        items = []
        for r in range(self.table.rowCount()):
            desc = self.safe_text(r, 0)

            # UOM extraction (widget or item)
            uom = ""
            widget = self.table.cellWidget(r, 1)
            if widget is not None:
                try:
                    if isinstance(widget, QComboBox):
                        uom = widget.currentText().strip()
                    else:
                        uom = widget.text().strip()
                except:
                    uom = ""
            else:
                uom = self.safe_text(r, 1)

            qty = self.safe_float(r, 2)
            rate = self.safe_float(r, 3)
            tax = self.safe_float(r, 4)

            items.append({
                "description": desc,
                "uom": uom,
                "qty": qty,
                "rate": rate,
                "tax": tax
            })

        calc = calculate_totals(items)

        total_words = num2words(calc["grand_total"], to="cardinal").title() + " Rupees Only"

        context = {
            "date": date_str,
            "quotation_no": quotation_no,
            "customer": customer,
            "gstin": gstin,
            "phone": phone,
            "items": calc["items"],
            "subtotal": f"{calc['subtotal']:.2f}",
            "total_tax": f"{calc['total_tax']:.2f}",
            "grand_total": f"{calc['grand_total']:.2f}",
            "rounding": f"{calc['rounding']:.2f}",
            "total_in_words": total_words,
            "generated_on": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        default_name = f"Quotation_{customer.replace(' ', '_')}_{date_str}.docx"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Quotation As", os.path.join(OUTPUT_DIR, default_name),
            "Word Document (*.docx)"
        )
        if not save_path:
            return

        try:
            res = generate_docx(template_path, save_path, context, convert_pdf=False)
            QMessageBox.information(self, "Success", f"Quotation generated:\n{res.get('docx')}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate document: {e}")

    # ---------- upload template ----------
    def on_upload_template(self):
        fpath, _ = QFileDialog.getOpenFileName(self, "Select Template", "", "Word Document (*.docx)")
        if not fpath:
            return
        tid, ok = QInputDialog.getText(self, "Template ID", "Enter Template ID (e.g., PSP):")
        if ok and tid:
            dest_dir = os.path.join(PROJECT_ROOT, "templates")
            os.makedirs(dest_dir, exist_ok=True)
            import shutil
            dest = os.path.join(dest_dir, os.path.basename(fpath))
            shutil.copyfile(fpath, dest)
            self.template_map[tid] = dest
            with open(TEMPLATE_MAP_FILE, "w", encoding="utf-8") as f:
                json.dump(self.template_map, f, indent=2)
            self.template_combo.clear()
            self.template_combo.addItems(list(self.template_map.keys()))
            QMessageBox.information(self, "Success", f"Template added as: {tid}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = QuotationApp()
    window.show()
    sys.exit(app.exec())
