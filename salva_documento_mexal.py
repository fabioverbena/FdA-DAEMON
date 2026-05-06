
import tkinter as tk
from tkinter import messagebox
import os
import json
import getpass
import shutil
from PyPDF2 import PdfReader, PdfWriter


def _fattura_prefix_numero(numero: str) -> str:
    numero = (numero or "").strip()
    if not numero:
        return numero
    if numero.startswith("13-"):
        return numero
    return f"13-{numero}"


def _load_dotenv(env_path: str) -> None:
    try:
        if not os.path.isfile(env_path):
            return
        with open(env_path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if not k:
                    continue
                if os.environ.get("DOTENV_OVERRIDE", "").strip().lower() in {"1", "true", "yes", "y"}:
                    os.environ[k] = v
                else:
                    os.environ.setdefault(k, v)
    except Exception:
        return

# Configurazioni
USER = getpass.getuser()

_load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
os.environ["DOTENV_OVERRIDE"] = "1"
_load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), "local.env"))
os.environ.pop("DOTENV_OVERRIDE", None)

BASE_PATH = os.path.join("C:/Users", USER, "Desktop", "AMMINISTRAZIONE_2025")
_DEFAULT_FATTURE_DIR = os.path.join(BASE_PATH, "FATTURE_2025")
_DEFAULT_BOLLE_DIR = os.path.join(BASE_PATH, "BOLLE_2025")

FATTURE_DIR = os.environ.get("FATTURE_DIR", "").strip() or _DEFAULT_FATTURE_DIR
BOLLE_DIR = os.environ.get("BOLLE_DIR", "").strip() or _DEFAULT_BOLLE_DIR

PATHS = {
    "Fattura": FATTURE_DIR,
    "Bolla": BOLLE_DIR,
}

MEXAL_TEMP = os.environ.get("MEXAL_TEMP", "").strip() or r"C:\Passepartout\PassClient\mxdesk1205143000\temp"
TRACK_FILE = "progressivo_documenti.json"

def load_tracking():
    if not os.path.exists(TRACK_FILE):
        with open(TRACK_FILE, "w") as f:
            json.dump({"Fattura": "", "Bolla": ""}, f)
    with open(TRACK_FILE, "r") as f:
        return json.load(f)

def save_tracking(data):
    with open(TRACK_FILE, "w") as f:
        json.dump(data, f)

def print_pdf(path):
    try:
        os.startfile(path, "print")
    except Exception as e:
        messagebox.showerror("Errore stampa", f"Errore nella stampa del file:\n{e}")

def save_first_page_only(input_path, output_path):
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, "wb") as f_out:
            writer.write(f_out)
    except Exception as e:
        messagebox.showerror("Errore PDF", f"Errore nel salvataggio prima pagina:\n{e}")

def trova_gruppo_pdf_recente(directory, soglia_secondi=10):
    pdfs = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                try:
                    mtime = os.path.getmtime(full_path)
                    pdfs.append((full_path, mtime))
                except:
                    pass
    if len(pdfs) < 3:
        return []
    pdfs.sort(key=lambda x: x[1], reverse=True)
    for i in range(len(pdfs) - 2):
        t1, t2, t3 = pdfs[i][1], pdfs[i+1][1], pdfs[i+2][1]
        if abs(t1 - t2) <= soglia_secondi and abs(t2 - t3) <= soglia_secondi:
            return [pdfs[i][0], pdfs[i+1][0], pdfs[i+2][0]]
    return []

class App:
    def __init__(self, master):
        self.master = master
        master.title("Salvataggio documento PDF")

        self.tracking = load_tracking()

        tk.Label(master, text="Tipo documento:").grid(row=0, column=0, sticky="w")
        self.tipo_var = tk.StringVar(value="Fattura")
        self.tipo_menu = tk.OptionMenu(master, self.tipo_var, "Fattura", "Bolla", command=self.update_numero)
        self.tipo_menu.grid(row=0, column=1, sticky="ew")

        tk.Label(master, text="Numero documento:").grid(row=1, column=0, sticky="w")
        self.numero_entry = tk.Entry(master)
        self.numero_entry.grid(row=1, column=1, sticky="ew")

        tk.Label(master, text="Nome cliente:").grid(row=2, column=0, sticky="w")
        self.cliente_entry = tk.Entry(master)
        self.cliente_entry.grid(row=2, column=1, sticky="ew")

        tk.Button(master, text="Solo Salva", command=self.save_only).grid(row=3, column=0, pady=10)
        tk.Button(master, text="Salva e Stampa 1", command=self.save_and_print_first).grid(row=3, column=1, pady=10)
        tk.Button(master, text="Salva e Stampa tutto", command=self.save_and_print_all).grid(row=4, column=0, columnspan=2, pady=10)

        self.update_numero()

    def update_numero(self, *_):
        tipo = self.tipo_var.get()
        ultimo = self.tracking.get(tipo, "")
        self.numero_entry.delete(0, tk.END)
        self.numero_entry.insert(0, ultimo)

    def save_and_act(self, print_mode):
        tipo = self.tipo_var.get()
        numero = self.numero_entry.get().strip()
        cliente = self.cliente_entry.get().strip()

        if not all([numero, cliente]):
            messagebox.showwarning("Attenzione", "Compila tutti i campi.")
            return

        if tipo == "Fattura":
            pdfs = trova_gruppo_pdf_recente(MEXAL_TEMP)
            if not pdfs:
                messagebox.showerror("Errore", "Nessun gruppo di PDF trovato per la fattura.")
                return
            pdf_to_save = pdfs[-1]
            pdf_to_print = pdfs if print_mode == "all" else [pdf_to_save]
        else:
            pdf_to_save = self.get_latest_pdf()
            if not pdf_to_save:
                messagebox.showerror("Errore", "Nessun PDF trovato per la bolla.")
                return
            pdf_to_print = [pdf_to_save] if print_mode else []

        save_dir = PATHS[tipo]
        os.makedirs(save_dir, exist_ok=True)

        if tipo == "Fattura":
            numero = _fattura_prefix_numero(numero)

        filename = f"{numero} {cliente}.pdf"
        dest_path = os.path.join(save_dir, filename)

        if tipo == "Fattura":
            save_first_page_only(pdf_to_save, dest_path)
        else:
            shutil.copy2(pdf_to_save, dest_path)

        for pdf in pdf_to_print:
            print_pdf(pdf)

        self.tracking[tipo] = numero
        save_tracking(self.tracking)
        self.numero_entry.delete(0, tk.END)
        self.cliente_entry.delete(0, tk.END)

        messagebox.showinfo("Completato", f"{tipo} salvata e {'stampata' if print_mode else 'non stampata'}.")

    def save_only(self):
        self.save_and_act(print_mode=False)

    def save_and_print_first(self):
        self.save_and_act(print_mode="first")

    def save_and_print_all(self):
        self.save_and_act(print_mode="all")

    def get_latest_pdf(self):
        all_pdfs = []
        for root, dirs, files in os.walk(MEXAL_TEMP):
            for file in files:
                if file.lower().endswith(".pdf"):
                    full_path = os.path.join(root, file)
                    try:
                        mtime = os.path.getmtime(full_path)
                        all_pdfs.append((full_path, mtime))
                    except:
                        pass
        if not all_pdfs:
            return None
        all_pdfs.sort(key=lambda x: x[1], reverse=True)
        return all_pdfs[0][0]

root = tk.Tk()
app = App(root)
root.mainloop()
