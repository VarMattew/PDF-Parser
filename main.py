import multiprocessing
from pypdf import PdfReader
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
import csv
from  openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import os
import time
import os
import csv
from concurrent.futures import ProcessPoolExecutor, as_completed
from pypdf import PdfReader
import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def egy_fajl_feldolgozasa(file_path):
    filename = os.path.basename(file_path)
    try:
        reader = PdfReader(file_path)
        first_page = reader.pages[0].extract_text()
        
        last_page_text = reader.pages[-1].extract_text() if len(reader.pages) > 0 else ""

        try:
            Nummers_raw = ' '.join(first_page.split()).split('Datum ')[1].split('Bei Zahlung')[0].split()
            Nummer = Nummers_raw[0].strip()
            Date = Nummers_raw[2].strip()
        except Exception:
            Nummer = 'ERROR! Nummer Not Found'
            Date = 'ERROR! Date Not Found'
        
        try:
            USt_ID_Nr = ' '.join(first_page.split()).split('USt-ID-Nr.:')[1].split('Commerzbank')[0].strip()
        except Exception:
            USt_ID_Nr = 'N/A'

        try:
            Attention = ' '.join(last_page_text.split()).split('ATTENTION:VAT')[1].split('Ursprungsland')[0].strip()
        except Exception:
            Attention = 'N/A'

        try:
            Versendet_von = ' '.join(first_page.split()).split('Versendet von:')[1].split('LS-Nr')[0].strip()
        except Exception:
            Versendet_von = 'N/A'

        return [Date, Nummer, USt_ID_Nr, Attention, Versendet_von]

    except Exception as e:
        return [f"ERROR ({filename})", "ERROR", "ERROR", "ERROR", str(e)]


def on_futtatas():
    mappa = path_var.get()
    
    try:
        pdf_files = [f for f in os.listdir(mappa) if f.lower().endswith('.pdf')]
    except FileNotFoundError:
        print("Hiba: A mappa nem található.")
        return

    osszes_fajl = len(pdf_files)
    if osszes_fajl == 0:
        print("Nem található PDF fájl a mappában!")
        return

    full_paths = [os.path.join(mappa, f) for f in pdf_files]

    # --- GUI Setup ---
    progress_window = ctk.CTkToplevel(root)
    progress_window.title("Feldolgozás...")
    progress_window.geometry("300x150")
    
    x = root.winfo_x() + (root.winfo_width() // 2) - 150
    y = root.winfo_y() + (root.winfo_height() // 2) - 75
    progress_window.geometry(f"+{x}+{y}")
    progress_window.transient(root)
    progress_window.grab_set()

    lbl_status = ctk.CTkLabel(progress_window, text=f"Indítás: {osszes_fajl} fájl...", font=("Roboto", 13))
    lbl_status.pack(pady=(25, 10))

    progress_bar = ctk.CTkProgressBar(progress_window, width=220)
    progress_bar.set(0)
    progress_bar.pack(pady=5)
    root.update()

    final_data = [['Date', 'No.', 'USt-ID-Nr.', 'ATTENTION', 'Versendet von']]

    # --- PÁRHUZAMOSÍTÁS ---
    try:
        with ProcessPoolExecutor(max_workers=os.cpu_count()) as executor:
            # Ez egy szótárt hoz létre: {FutureObjektum: fájlnév}
            future_to_file = {executor.submit(egy_fajl_feldolgozasa, path): path for path in full_paths}
            
            completed_count = 0
            
            for future in as_completed(future_to_file):
                result_row = future.result() 
                final_data.append(result_row)
                
                completed_count += 1
                szazalek = completed_count / osszes_fajl
                progress_bar.set(szazalek)
                lbl_status.configure(text=f"Kész: {completed_count} / {osszes_fajl}")
                root.update()

        # --- Mentés ---
        lbl_status.configure(text="Mentés fájlba...")
        root.update()

        if combo.get() == 'csv':
            with open(f'{mezo.get()}.csv', mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=",")
                writer.writerows(final_data)
        else:
            wb = Workbook()
            ws = wb.active
            for sor in final_data:
                ws.append(sor)
            
            # Formázás
            for cell in ws[1]:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.freeze_panes = "A2"
            wb.save(f'{mezo.get()}.xlsx')

    except Exception as e:
        print(f"Kritikus hiba: {e}")
        lbl_status.configure(text="Hiba történt!")
    finally:
        progress_window.grab_release()
        progress_window.destroy()
        root.destroy()


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

if __name__ == "__main__":

    multiprocessing.freeze_support()

    root = ctk.CTk()
    root.geometry("360x400") 
    root.title("PDF Adat Kinyerő")
    root.resizable(False, False)

    my_font_head = ("Roboto", 14)
    my_font_btn = ("Roboto", 13)
    grey_btn_color = "gray30" 
    grey_hover_color = "gray40"

    path_var = ctk.StringVar()
    filename_var = ctk.StringVar(value="output")

    def gomb_allapot_ellenorzes(*args):
        mappa_van = len(path_var.get()) > 0
        fajlnev_van = len(filename_var.get()) > 0

        if mappa_van and fajlnev_van:
            btn_run.configure(state="normal", fg_color=grey_btn_color)
        else:
            btn_run.configure(state="disabled", fg_color="gray20")
    

    input_frame = ctk.CTkFrame(root, fg_color="transparent")
    input_frame.pack(pady=(0, 2), fill="x")

    ctk.CTkLabel(input_frame, text="Bemeneti mappa:", font=my_font_head, text_color="gray80").pack(pady=(20, 5))
    btn_browse = ctk.CTkButton(input_frame, text="Mappa választás",
                               command=lambda: path_var.set(filedialog.askdirectory()),
                               font=my_font_btn,
                               fg_color=grey_btn_color, hover_color=grey_hover_color)
    btn_browse.pack(pady=5)

    path_label = ctk.CTkLabel(input_frame, textvariable=path_var, font=("Arial", 11), text_color="gray60", wraplength=340)

    def on_path_change(*args):
        if path_var.get():
             path_label.pack(padx=10, pady=(0))
        else:
             path_label.pack_forget()
        gomb_allapot_ellenorzes()

    path_var.trace_add("write", on_path_change)

    ctk.CTkLabel(root, text="Kimenet formátuma:", font=my_font_head, text_color="gray80").pack(pady=(10, 5))
    combo = ctk.CTkOptionMenu(root, values=["excel", "csv"], width=150,
                              font=my_font_btn,
                              fg_color=grey_btn_color, button_color=grey_hover_color,
                              button_hover_color="gray50")
    combo.set("excel")
    combo.pack(padx=10, pady=5)


    ctk.CTkLabel(root, text="Kimeneti fájl neve:", font=my_font_head, text_color="gray80").pack(pady=(15, 5))
    mezo = ctk.CTkEntry(root, width=200, justify='center', font=my_font_btn, textvariable=filename_var)
    filename_var.trace_add("write", gomb_allapot_ellenorzes)
    mezo.pack(padx=10, pady=5)

    btn_run = ctk.CTkButton(root, text="Futtatás", command=on_futtatas,
                            font=("Roboto", 16, "bold"), height=40,
                            state="disabled",
                            fg_color="gray20",
                            hover_color=grey_hover_color)
    btn_run.pack(pady=(30, 20))

    gomb_allapot_ellenorzes()
    root.mainloop()