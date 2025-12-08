from pypdf import PdfReader
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
import csv
from  openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import os
import time

data_list = [['Date', 'No.', 'USt-ID-Nr.', 'ATTENTION', 'Versendet von']]

def on_futtatas():
    mappa = path_var.get()
    
    try:
        pdf_fajlok = [f for f in os.listdir(mappa) if f.lower().endswith(('.pdf', '.PDF'))]
    except FileNotFoundError:
        print("Hiba: A mappa nem található.")
        return

    osszes_fajl = len(pdf_fajlok)

    if osszes_fajl == 0:
        print("Nem található PDF fájl a mappában!")
        return

    progress_window = ctk.CTkToplevel(root)
    progress_window.title("Feldolgozás")
    progress_window.geometry("300x150")
    progress_window.resizable(False, False)
    progress_window.attributes("-topmost", True)

    x = root.winfo_x() + (root.winfo_width() // 2) - 150
    y = root.winfo_y() + (root.winfo_height() // 2) - 75
    progress_window.geometry(f"+{x}+{y}")

    lbl_status = ctk.CTkLabel(progress_window, text=f"Talált fájlok: {osszes_fajl} db", 
                              font=("Roboto", 13))
    lbl_status.pack(pady=(25, 10))

    progress_bar = ctk.CTkProgressBar(progress_window, width=220, height=12)
    progress_bar.set(0)
    progress_bar.pack(pady=5)
    progress_bar.configure(progress_color="gray80") 
    
    progress_window.grab_set()
    root.update()

    try:
        for i, file in enumerate(os.listdir(mappa)):
            print(file)

            if not (file.endswith('.pdf') or file.endswith('.PDF')):
                continue

            print(f'Processing file: {mappa}/{file}')
            reader = PdfReader(f'{mappa}/{file}')

            try:
                Nummers = ' '.join(reader.pages[0].extract_text().split()).split('Datum ')[1].split('Bei Zahlung')[0].split()
                Nummer = Nummers[0].strip()
                Date = Nummers[2].strip()
            except IndexError:
                Nummer = 'ERROR! Nummer Not Found'
                Date = 'ERROR! Date Not Found'
            
            try:
                USt_ID_Nr = ' '.join(reader.pages[0].extract_text().split()).split('USt-ID-Nr.:')[1].split('Commerzbank')[0].strip()
            except IndexError:
                USt_ID_Nr = 'N/A'

            try:
                Attention = ' '.join(reader.pages[-1].extract_text().split()).split('ATTENTION:VAT')[1].split('Ursprungsland')[0].strip()
            except IndexError:
                Attention = 'N/A'

            try:
                Versendet_von = ' '.join(reader.pages[0].extract_text().split()).split('Versendet von:')[1].split('LS-Nr')[0].strip()
            except IndexError:
                Versendet_von = 'N/A'

            data_list.append([Date, Nummer, USt_ID_Nr, Attention, Versendet_von])

            
            szazalek = (i + 1) / osszes_fajl 
            
            progress_bar.set(szazalek)
            lbl_status.configure(text=f"Feldolgozva: {i + 1} / {osszes_fajl} ({file})")
            
            root.update()

        lbl_status.configure(text="Kész! Mentés folyamatban...")
        root.update()

        if combo.get() == 'csv':
            with open(f'{mezo.get()}.csv', mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=",")
                writer.writerows(data_list)
        else:
            wb = Workbook()
            ws = wb.active

            for sor in data_list:
                ws.append(sor)

            for cell in ws[1]:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
            ws.freeze_panes = "A2"
            
            wb.save(f'{mezo.get()}.xlsx')
        

    except Exception as e:
        print(f"Hiba történt: {e}")
        lbl_status.configure(text="Hiba történt!")

    finally:
        progress_window.grab_release()
        progress_window.destroy()
        root.grab_release()
        root.destroy()


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

if __name__ == "__main__":
    root = ctk.CTk()
    root.geometry("360x400") 
    root.title("PDF Adat Kinyerő")

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