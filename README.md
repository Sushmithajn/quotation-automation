## Quotation Automation Tool
A desktop application built with Python, PySide6, and DocxTemplate for automatically generating GST-compliant quotation documents in DOCX format.
The tool allows users to input customer details, add multiple items, auto-calculate totals, apply taxes, and instantly generate a formatted quotation using customizable Word templates.

## 🚀 Features
## ✔️ User-Friendly Desktop Application

Built using PySide6 (Qt for Python)

Clean UI with form inputs and dynamic table rows

## ✔️ Automatic Calculations

The app calculates:

Line Total

Subtotal

Total Tax

Grand Total

Rounding Difference

Total amount in words

## ✔️ Supports Multiple UOM (Unit of Measurement)

Nos, Kg, Ltr, Meter, Box, Packet, Dozen

Custom UOM option

## ✔️ Template-Based Quotation Generation

Uses DOCX template (.docx)

Generates a filled quotation using placeholders

Allows uploading your own templates


## ✔️ Export Options

Save as Word Document (.docx)


🛠️ Installation (Developer Mode)
1️⃣ Clone the Repository
git clone https://github.com/YOUR_USERNAME/quotation-automation.git
cd quotation-automation

2️⃣ Create Virtual Environment
python -m venv venv


Activate:

Windows:

venv\Scripts\activate

3️⃣ Install Dependencies
pip install -r requirements.txt

4️⃣ Run the Application
python main.py

📦 Build Executable (.exe)

To create a standalone desktop application:

pyinstaller --noconsole --onefile --add-data "templates;templates" --add-data "data;data" main.py


The .exe file will be inside:

/dist/main.exe


Share this executable with users.

🧮 How Calculations Work

For each item:

Product Total = Qty × Rate
Tax Amount = Product Total × (Tax% / 100)
Line Total = Product Total + Tax Amount


Overall:

Subtotal = Sum of all product totals
Total Tax = Sum of all tax amounts
Grand Total = Subtotal + Total Tax
Rounding = round(Grand Total) - Grand Total
