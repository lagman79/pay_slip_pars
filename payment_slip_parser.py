import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import threading
import time
import pandas as pd
import google.generativeai as genai
from glob import glob
from datetime import datetime
import re

# --- ΡΥΘΜΙΣΕΙΣ & CONSTANTS ---
SETTINGS_FILE = "settings_payments.json"

# Λεξικό Κωδικών Τραπεζών (HEBIC Code - Digits 5-8 of IBAN)
BANK_CODES = {
    "0110": "ΕΘΝΙΚΗ ΤΡΑΠΕΖΑ",
    "0140": "ALPHA BANK",
    "0172": "ΤΡΑΠΕΖΑ ΠΕΙΡΑΙΩΣ",
    "0171": "ΤΡΑΠΕΖΑ ΠΕΙΡΑΙΩΣ",
    "0260": "EUROBANK",
    "0870": "ATTICA BANK",
    "0710": "ΠΑΓΚΡΗΤΙΑ ΤΡΑΠΕΖΑ",
    "0690": "VIVA WALLET",
    "0026": "OPTIMA BANK"
}

class DataProcessor:
    """Κλάση που διαχειρίζεται την επικοινωνία με το AI και την επεξεργασία δεδομένων"""
    
    @staticmethod
    def format_currency(val):
        if not val: return ""
        try:
            clean_val = str(val).replace("-", "").strip()
            return "{:,.2f}".format(float(clean_val)).replace(",", "X").replace(".", ",").replace("X", ".") + " €"
        except: return str(val)

    @staticmethod
    def clean_iban(raw_iban):
        """Καθαρίζει το IBAN, κρατάει μόνο 27 χαρακτήρες ξεκινώντας από GR"""
        if not raw_iban: return ""
        
        # 1. Αφαίρεση κενών και μη αλφαριθμητικών
        clean = "".join(c for c in str(raw_iban) if c.isalnum()).upper()
        
        # 2. Εντοπισμός 'GR'
        idx = clean.find("GR")
        if idx == -1:
            return "" # Δεν είναι ελληνικό IBAN ή είναι λάθος
        
        # 3. Κράτημα μόνο του έγκυρου τμήματος (από το GR και μετά)
        iban_candidate = clean[idx:]
        
        # 4. Truncate στους 27 χαρακτήρες (Ελληνικό πρότυπο)
        if len(iban_candidate) > 27:
            iban_candidate = iban_candidate[:27]
            
        return iban_candidate

    @staticmethod
    def get_bank_from_iban(iban):
        """Εξάγει την τράπεζα από τα ψηφία 5-8 του IBAN"""
        if not iban or len(iban) < 8: return ""
        code = iban[4:8]
        return BANK_CODES.get(code, f"ΑΓΝΩΣΤΗ ({code})")

    @staticmethod
    def check_same_bank(iban_from, iban_to):
        """Ελέγχει αν τα IBAN ανήκουν στην ίδια τράπεζα"""
        if not iban_from or not iban_to: return "Άγνωστο"
        if len(iban_from) < 8 or len(iban_to) < 8: return "Άγνωστο"
        
        bank_code_1 = iban_from[4:8]
        bank_code_2 = iban_to[4:8]
        return "ΝΑΙ" if bank_code_1 == bank_code_2 else "ΟΧΙ"

    @staticmethod
    def analyze_file(file_path, api_key, full_extract):
        genai.configure(api_key=api_key, transport='rest')
        
        # Determine mime type based on extension
        ext = os.path.splitext(file_path)[1].lower()
        mime_type = "application/pdf"
        if ext in ['.jpg', '.jpeg']: mime_type = "image/jpeg"
        elif ext == '.png': mime_type = "image/png"
        
        # Upload
        sample_file = genai.upload_file(path=file_path, mime_type=mime_type, display_name="PaymentDoc")
        
        # Wait for processing
        timeout = 60 
        start_time = time.time()
        while sample_file.state.name == "PROCESSING":
            if time.time() - start_time > timeout:
                raise TimeoutError("Timeout processing file.")
            time.sleep(1)
            sample_file = genai.get_file(sample_file.name)
        
        if sample_file.state.name == "FAILED":
            raise ValueError("File processing failed by Google.")

        model = genai.GenerativeModel("models/gemini-2.0-flash", generation_config={"response_mime_type": "application/json"})

        extra_instruction = ""
        if full_extract:
            extra_instruction = """
            FULL EXTRACT MODE:
            Ψάξε για ΟΠΟΙΟΔΗΠΟΤΕ άλλο πεδίο υπάρχει (π.χ. Κατάστημα, Ώρα, User ID, Αιτιολογία, Λεπτομέρειες, Έγκριση).
            Βάλτα σε ένα αντικείμενο 'dynamic_fields' με τα ακριβή ονόματα που βλέπεις στο έγγραφο (π.χ. "Ώρα καταχωρήσεως", "Valeur").
            """

        prompt = f"""
        Είσαι ειδικός τραπεζικών συναλλαγών. Ανάλυσε το παραστατικό πληρωμής (PDF ή Εικόνα) και δώσε JSON.
        
        ΟΔΗΓΙΕΣ ΠΕΔΙΩΝ:
        1. bank_name_header: Ποια τράπεζα φαίνεται στο λογότυπο/κεφαλίδα.
        2. transaction_id: Ο μοναδικός κωδικός συναλλαγής.
        3. date: Ημερομηνία εκτέλεσης ή καταχώρησης (Format: DD/MM/YYYY).
        4. amount: Το ποσό της πληρωμής (Απόλυτη τιμή).
        5. charges: Έξοδα/Προμήθειες συναλλαγής.
        6. sender_iban: Ο λογαριασμός χρέωσης (Από).
        7. recipient_iban: Ο λογαριασμός πίστωσης (Προς / Σε).
        8. beneficiary_name: Το όνομα του δικαιούχου.
        
        {extra_instruction}

        JSON OUTPUT KEYS:
        - bank_name_header, transaction_id, date
        - amount, charges
        - sender_iban, recipient_iban, beneficiary_name
        {"- dynamic_fields (object)" if full_extract else ""}
        """

        try:
            response = model.generate_content([sample_file, prompt])
            genai.delete_file(sample_file.name)
            
            raw_data = json.loads(response.text)
            data = raw_data[0] if isinstance(raw_data, list) and len(raw_data) > 0 else (raw_data if isinstance(raw_data, dict) else {})

            # --- POST PROCESSING & LOGIC ---
            
            # 1. Καθαρισμός IBANs
            clean_sender = DataProcessor.clean_iban(data.get('sender_iban'))
            clean_recipient = DataProcessor.clean_iban(data.get('recipient_iban'))
            
            data['sender_iban'] = clean_sender
            data['recipient_iban'] = clean_recipient

            # 2. Εντοπισμός Τραπεζών από IBAN
            bank_from_iban = DataProcessor.get_bank_from_iban(clean_sender)
            bank_to_iban = DataProcessor.get_bank_from_iban(clean_recipient)

            # 3. Λογική Τράπεζας Χρέωσης (Cross-check)
            data['final_debit_bank'] = bank_from_iban if bank_from_iban else data.get('bank_name_header', '')

            # 4. Τράπεζα Πίστωσης
            data['final_credit_bank'] = bank_to_iban

            # 5. Έλεγχος Ίδιας Τράπεζας
            data['same_bank_check'] = DataProcessor.check_same_bank(clean_sender, clean_recipient)
            
            return data
        except Exception as e:
            try: genai.delete_file(sample_file.name)
            except: pass
            raise e

class PaymentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bank Payment Extractor Pro v3")
        self.root.geometry("750x750")
        
        self.input_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.api_key = tk.StringVar()
        self.extract_all = tk.BooleanVar(value=True)
        self.is_running = False
        
        self.load_settings() # Φορτώνει ΜΟΝΟ το API Key
        self.create_widgets()
        
    def create_widgets(self):
        # API Frame
        frame_api = tk.LabelFrame(self.root, text="🔐 Ρυθμίσεις Ασφαλείας", padx=10, pady=10)
        frame_api.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_api, text="Gemini API Key:").pack(side="left")
        self.entry_api = tk.Entry(frame_api, textvariable=self.api_key, show="*", width=50)
        self.entry_api.pack(side="left", padx=5)

        # Files Frame
        frame_files = tk.LabelFrame(self.root, text="📂 Διαχείριση Αρχείων (PDF & Εικόνες)", padx=10, pady=10)
        frame_files.pack(fill="x", padx=10, pady=5)
        tk.Button(frame_files, text="Επιλογή Φακέλου", command=self.select_input, width=20).grid(row=0, column=0, pady=2)
        tk.Entry(frame_files, textvariable=self.input_folder, width=50, state="readonly").grid(row=0, column=1, padx=5)
        tk.Button(frame_files, text="Αποθήκευση Excel", command=self.select_output, width=20).grid(row=1, column=0, pady=2)
        tk.Entry(frame_files, textvariable=self.output_file, width=50, state="readonly").grid(row=1, column=1, padx=5)

        # Options Frame
        frame_opts = tk.LabelFrame(self.root, text="⚙️ Παράμετροι", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)
        tk.Checkbutton(frame_opts, text="Full Extract (Όλα τα πεδία)", variable=self.extract_all).pack(anchor="w")

        # Progress & Log
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(fill="x", padx=15, pady=10)
        self.log_text = tk.Text(self.root, height=12, state="disabled", bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

        # Action Buttons Frame
        frame_actions = tk.Frame(self.root)
        frame_actions.pack(fill="x", padx=10, pady=10)

        # Start Button
        self.btn_start = tk.Button(frame_actions, text="🚀 Εκκίνηση", command=self.start_thread, bg="#3498db", fg="white", font=("Arial", 11, "bold"), height=2)
        self.btn_start.pack(fill="x", pady=(0, 10))

        # Control Buttons (New Job / Exit)
        btn_new = tk.Button(frame_actions, text="🧹 Νέα εργασία", command=self.reset_app, bg="#f39c12", fg="white", font=("Arial", 10, "bold"))
        btn_new.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_exit = tk.Button(frame_actions, text="🚪 Έξοδος", command=self.root.quit, bg="#e74c3c", fg="white", font=("Arial", 10, "bold"))
        btn_exit.pack(side="right", fill="x", expand=True, padx=(5, 0))

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def load_settings(self):
        """Φορτώνει ΜΟΝΟ το API Key κατά την εκκίνηση"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
                    self.api_key.set(data.get("api_key", ""))
                    # Δεν φορτώνουμε φάκελο/αρχείο για να είναι καθαρό το UI
            except: pass

    def save_api_key(self):
        """Αποθηκεύει ΜΟΝΟ το API Key"""
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump({"api_key": self.api_key.get()}, f)
        except: pass

    def reset_app(self):
        """Καθαρίζει τα πεδία για νέα εργασία"""
        if self.is_running:
            messagebox.showwarning("Προσοχή", "Η διαδικασία εκτελείται ακόμη!")
            return
        
        self.input_folder.set("")
        self.output_file.set("")
        self.log_text.config(state="normal"); self.log_text.delete(1.0, "end"); self.log_text.config(state="disabled")
        self.progress["value"] = 0
        self.log("✅ Έτοιμο για νέα εργασία.")

    def select_input(self):
        f = filedialog.askdirectory()
        if f: self.input_folder.set(f)

    def select_output(self):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if f: self.output_file.set(f)

    def start_thread(self):
        if not self.api_key.get() or not self.input_folder.get():
            messagebox.showwarning("Προσοχή", "Λείπουν στοιχεία!")
            return
        
        self.save_api_key() # Save only API Key
        
        self.is_running = True
        self.btn_start.config(state="disabled", text="⏳ ΣΕ ΕΞΕΛΙΞΗ...")
        self.log_text.config(state="normal"); self.log_text.delete(1.0, "end"); self.log_text.config(state="disabled")
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        try:
            input_dir = self.input_folder.get()
            output_path = self.output_file.get()
            if not output_path: output_path = os.path.join(input_dir, "payment_report_v3.xlsx")
            
            # --- UPDATE: Search for Images AND PDFs ---
            extensions = ['*.pdf', '*.PDF', '*.jpg', '*.JPG', '*.jpeg', '*.JPEG', '*.png', '*.PNG']
            files = []
            for ext in extensions:
                files.extend(glob(os.path.join(input_dir, ext)))
            
            # Remove duplicates if any (case sensitivity issues on some OS)
            files = sorted(list(set(files)))

            if not files:
                self.log("❌ Δεν βρέθηκαν αρχεία (PDF ή Εικόνες).")
                return

            self.progress["maximum"] = len(files)
            all_data = []

            for i, f in enumerate(files, 1):
                if not self.is_running: break
                filename = os.path.basename(f)
                self.log(f"Επεξεργασία: {filename}")
                
                try:
                    data = DataProcessor.analyze_file(f, self.api_key.get().strip(), self.extract_all.get())
                    data['filename'] = filename
                    all_data.append(data)
                    self.log("✅ Επιτυχία")
                except Exception as e:
                    self.log(f"❌ Σφάλμα: {str(e)}")
                
                self.progress["value"] = i
                time.sleep(1.5)

            if all_data:
                self.generate_excel(all_data, output_path)
            else:
                self.log("⚠️ Δεν προέκυψαν δεδομένα.")

        except Exception as e:
            self.log(f"CRITICAL ERROR: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.is_running = False
            self.btn_start.config(state="normal", text="🚀 ΕΚΚΙΝΗΣΗ ΕΞΑΓΩΓΗΣ")

    def generate_excel(self, all_data, path):
        df = pd.DataFrame(all_data)
        
        for col in ['amount', 'charges']:
            if col in df.columns:
                df[col] = df[col].apply(DataProcessor.format_currency)

        mapping = {
            'final_debit_bank': 'ΤΡΑΠΕΖΑ (ΧΡΕΩΣΗ)',
            'final_credit_bank': 'ΤΡΑΠΕΖΑ (ΠΙΣΤΩΣΗ)',
            'transaction_id': 'ΚΩΔΙΚΟΣ ΣΥΝΑΛΛΑΓΗΣ',
            'date': 'ΗΜ/ΝΙΑ',
            'amount': 'ΠΟΣΟ',
            'charges': 'ΕΞΟΔΑ',
            'sender_iban': 'ΙΒΑΝ ΧΡΕΩΣΗΣ (ΑΠΟ)',
            'recipient_iban': 'ΙΒΑΝ ΠΙΣΤΩΣΗΣ (ΠΡΟΣ)',
            'beneficiary_name': 'ΔΙΚΑΙΟΥΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ ΠΙΣΤΩΣΗΣ',
            'same_bank_check': 'ΙΔΙΑ ΤΡΑΠΕΖΑ;',
            'filename': 'ΟΝΟΜΑ ΑΡΧΕΙΟΥ'
        }
        
        if self.extract_all.get() and 'dynamic_fields' in df.columns:
            dynamic_df = df['dynamic_fields'].apply(pd.Series)
            df = pd.concat([df.drop(['dynamic_fields'], axis=1), dynamic_df], axis=1)

        df.rename(columns=mapping, inplace=True)

        target_order = [
            "ΤΡΑΠΕΖΑ (ΧΡΕΩΣΗ)", "ΚΩΔΙΚΟΣ ΣΥΝΑΛΛΑΓΗΣ", "ΗΜ/ΝΙΑ", "ΠΟΣΟ", "ΕΞΟΔΑ",
            "ΙΒΑΝ ΧΡΕΩΣΗΣ (ΑΠΟ)", "ΙΒΑΝ ΠΙΣΤΩΣΗΣ (ΠΡΟΣ)", "ΤΡΑΠΕΖΑ (ΠΙΣΤΩΣΗ)",
            "ΔΙΚΑΙΟΥΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ ΠΙΣΤΩΣΗΣ", "ΙΔΙΑ ΤΡΑΠΕΖΑ;", "ΟΝΟΜΑ ΑΡΧΕΙΟΥ",
            "Ώρα καταχωρήσεως", "Όνομα παραλήπτριας τραπέζης", "Έξοδα",
            "Επιβάρυνση για τραπεζικά έξοδα δικαιούχου", "Συνολικό ποσό αγορών/χρεώσεων",
            "Ημ/νία μεταφοράς", "Στοιχεία εντολέα", "Α.Φ.Μ.", "Αιτιολογία για καταθέτη",
            "Αιτιολογία προς δικαιούχο", "Κατάσταση συναλλαγής", "Καταχώρηση μέσω",
            "Από Λογαριασμό", "Νόμισμα", "Τρόπος Εκτελέσεως", "Σε λογαριασμό",
            "Δικαιούχος", "Μήνυμα προς δικαιούχο", "ΑΙΤΙΟΛΟΓΙΑ ΑΠΟΣΤΟΛΕΑ",
            "ΑΙΤΙΟΛΟΓΙΑ ΠΑΡΑΛΗΠΤΗ", "ΗΜΕΡΟΜΗΝΙΑ ΚΑΤΑΧΩΡΗΣΗΣ", "ΠΛΗΡΟΦΟΡΙΕΣ",
            "ΧΩΡΑ", "BIC", "ΟΝΟΜΑ ΤΡΑΠΕΖΑΣ", "ΔΙΕΥΘΥΝΣΗ", "ΠΟΛΗ",
            "ΔΙΚΑΙΟΥΧΟΙ ΛΟΓΑΡΙΑΣΜΟΥ", "ΕΝΤΟΛΟΔΟΧΟΣ ΤΡΑΠΕΖΑ", "ΕΝΤΟΛΕΑΣ",
            "ΤΡΟΠΟΣ ΧΡΕΩΣΗΣ ΠΡΟΜΗΘΕΙΩΝ/ΕΞΟΔΩΝ", "ΚΑΤΑΣΤΗΜΑ", "ΚΩΔΙΚΟΣ ΑΝΑΦΟΡΑΣ ΕΝΤΟΛΕΑ",
            "ΛΟΓΑΡΙΑΣΜΟΣ ΓΙΑ ΤΑ ΕΞΟΔΑ ΜΕΤΑΦΟΡΑΣ", "ΤΟΚΟΦΟΡΟΣ ΗΜΕΡΟΜΗΝΙΑ", "Κατάσταση",
            "Κανάλι", "Κύριος Δικαιούχος", "Πληροφορίες για το δικαιούχο", "Εκτέλεση",
            "Ημερομηνία Ενημέρωσης", "Αριθμός Αίτησης", "Κωδικός Συναλλαγής",
            "Χώρα Αποστολής", "Τιμή μετατροπής", "Ημερομηνία Αξίας", "Beneficiary's Bank",
            "Value Date / Amount / Currency", "Details of Payment", "Details of Charges",
            "Ημερομηνία Καταχώρησης", "Έγκριση", "Τράπεζα πληρωμής", "Επωνυμία εντολέα",
            "Λογαριασμός εντολέα", "Λογαριασμός δικαιούχου", "Valeur",
            "Λεπτομέρειες πληρωμής", "Κατάσταση Εμβάσματος"
        ]

        existing_target_cols = [c for c in target_order if c in df.columns]
        remaining_cols = [c for c in df.columns if c not in existing_target_cols]
        final_cols = existing_target_cols + remaining_cols
        
        df = df[final_cols]
        df = df.fillna("")
        
        try:
            df.to_excel(path, index=False)
            self.log(f"🎉 Το Excel αποθηκεύτηκε: {path}")
            messagebox.showinfo("Ολοκληρώθηκε", f"Το αρχείο δημιουργήθηκε:\n{path}")
        except PermissionError:
            messagebox.showerror("Σφάλμα", "Κλείσε το αρχείο Excel! Είναι ανοιχτό.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PaymentApp(root)
    root.mainloop()