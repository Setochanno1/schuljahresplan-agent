from pathlib import Path
from datetime import datetime
from docx import Document
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry

VORLAGE = Path("vorlagen/Jahresarbeitsplan 25-26.docx")
AUSGABE = Path("ausgabe/Jahresarbeitsplan_Test_mit_Ereignis.docx")

SPALTEN = {
    0: 4,  # Montag
    1: 5,  # Dienstag
    2: 6,  # Mittwoch
    3: 7,  # Donnerstag
    4: 8,  # Freitag
}

def add_event(datum, ereignis):
    if not VORLAGE.exists():
        messagebox.showerror("Fehler", f"Vorlage nicht gefunden:\n{VORLAGE}")
        return

    doc = Document(VORLAGE)
    table = doc.tables[0]

    if datum.weekday() > 4:
        messagebox.showwarning("Hinweis", "Das Datum liegt am Wochenende.")
        return

    ziel_spalte = SPALTEN[datum.weekday()]

    gefunden = False

    for row in table.rows:
        zeitraum = row.cells[3].text.strip()

        if "–" not in zeitraum:
            continue

        teile = [x.strip() for x in zeitraum.split("–")]

        if len(teile) != 2:
            continue

        try:
            start = datetime.strptime(teile[0], "%d.%m.").date().replace(year=2025)
            ende = datetime.strptime(teile[1], "%d.%m.").date().replace(year=2025)
        except:
            continue

        vergleich = datum.replace(year=2025)

        if start <= vergleich <= ende:
            zelle = row.cells[ziel_spalte]

            if zelle.text.strip():
                zelle.text = zelle.text + "\n" + ereignis
            else:
                zelle.text = ereignis

            gefunden = True
            break

    if not gefunden:
        messagebox.showwarning("Nicht gefunden", "Keine passende Woche gefunden.")
        return

    AUSGABE.parent.mkdir(exist_ok=True)
    doc.save(AUSGABE)

    messagebox.showinfo("Fertig", f"Ereignis eingetragen:\n{AUSGABE}")

def button_klick():
    datum = kalender.get_date()
    ereignis = eingabe_ereignis.get().strip()

    if not ereignis:
        messagebox.showwarning("Fehlt", "Bitte ein Ereignis eingeben.")
        return

    add_event(datum, ereignis)

root = tk.Tk()
root.title("Schuljahresplan-Agent")
root.geometry("420x220")

tk.Label(root, text="Datum auswählen:").pack(pady=5)

kalender = DateEntry(root, date_pattern="dd.mm.yyyy")
kalender.pack(pady=5)

tk.Label(root, text="Ereignis:").pack(pady=5)

eingabe_ereignis = tk.Entry(root, width=45)
eingabe_ereignis.pack(pady=5)

button = tk.Button(root, text="Ereignis hinzufügen", command=button_klick)
button.pack(pady=20)

root.mainloop()
