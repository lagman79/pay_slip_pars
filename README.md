# Bank Payment Extractor Pro 🏦🤖

**Bank Payment Extractor Pro** is a Python desktop application that leverages **Google Gemini AI (2.0 Flash)** to automatically extract, analyze, and organize data from bank payment slips (PDFs and Images) into structured Excel reports.

It is designed to handle various Greek bank formats (Alpha Bank, Eurobank, Piraeus, NBG, etc.), perform validation checks, and streamline accounting workflows.

## 🚀 Features

* **Multi-Format Support**: Processes both **PDF** documents and **Image** files (JPG, PNG).
* **AI-Powered Extraction**: Uses Google's **Gemini 2.0 Flash** model to intelligently identify fields like Transaction IDs, Amounts, Dates, and Beneficiaries.
* **Smart IBAN Handling**:
    * Automatic detection and cleaning of IBANs (removes spaces/dashes, truncates to standard length).
    * **Bank Identification**: Automatically detects the bank name based on the IBAN digits.
    * **Same-Bank Check**: validaters if the transaction is internal (between the same bank).
* **Cross-Check Logic**: Verifies the Debit Bank against the sender's IBAN for accuracy.
* **Excel Export**: Generates a formatted Excel file (`.xlsx`) with a specific column order tailored for accounting needs.
* **User-Friendly GUI**: Built with `tkinter`, featuring:
    * File selection dialogs.
    * Live progress bar and logging.
    * "New Job" and "Exit" controls.
    * Secure API Key input (saved locally).

## 🛠️ Installation

1.  **Clone the repository**:
    ```bash
    git clone [https://github.com/YOUR_USERNAME/bank-payment-extractor.git](https://github.com/YOUR_USERNAME/bank-payment-extractor.git)
    cd bank-payment-extractor
    ```

2.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the application**:
    ```bash
    python payment_slip_extractor.py
    ```

## 📋 Requirements

* Python 3.x
* A **Google Gemini API Key** (Get one at [Google AI Studio](https://aistudio.google.com/))

## 📦 Dependencies

* `google-generativeai`
* `pandas`
* `openpyxl`
* `tk` (Standard Python library)

## 🔒 Security Note

This application stores your API Key locally in a `settings_payments.json` file. **This file is automatically added to `.gitignore` to prevent it from being uploaded to GitHub.** Never share your API keys publicly.

## 📝 Usage

1.  Paste your **Gemini API Key**.
2.  Select the **Input Folder** containing your PDFs or Images.
3.  (Optional) Select a specific **Output Excel File** path.
4.  Check **"Full Extract"** if you want to capture dynamic fields (notes, timestamps, etc.).
5.  Click **"Start Extraction"** (🚀 ΕΚΚΙΝΗΣΗ ΕΞΑΓΩΓΗΣ).
6.  Once finished, use **"New Job"** (🧹 ΝΕΑ ΕΡΓΑΣΙΑ) to clear the form or **"Exit"** (🚪 ΕΞΟΔΟΣ) to close.

## 📄 License

This project is open-source. Feel free to modify and use it for your own workflows.