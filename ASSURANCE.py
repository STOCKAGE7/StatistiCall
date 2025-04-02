import datetime
import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
from PIL import Image, ImageTk
import os

def countname(name_list):
    name_count = {name: name_list.count(name) for name in set(name_list)}
    return list(name_count.keys()), list(name_count.values())

def removespaces(strings):
    return [string.replace(" ", "") for string in strings]

def capitalizefirst(strings):
    return [string.capitalize() for string in strings]

def generate_report(file_path):
    try:
        file = pd.read_excel(rf'{file_path}')
        file2 = pd.read_excel(rf'{file_path}', sheet_name='Déplacements ')
        file3 = pd.read_excel(rf'{file_path}', sheet_name='Annulations')
        file4 = pd.read_excel(rf'{file_path}', sheet_name='Transfert')
        file5 = pd.read_excel(rf'{file_path}', sheet_name=' Autre détail ')

        dr = file['Docteur'].tolist()
        list_dr = capitalizefirst(removespaces(dr))

        message, x = "", 0
        unique_names, count_values = countname(list_dr)
        for i, k in zip(unique_names, count_values):
            message += f"   - {i} : {k:02}\n"
            x += k

        deplacements = file2['Nom et prénon '].count()
        annulations = file3['Le Docteur'].count()
        transfert = file4['Qui'].count()
        autres = file5['Autres détail '].count()
        total = deplacements + annulations + transfert + autres + x

        date = datetime.date.today().strftime("%d %m %Y")
        email = f"""
        Subject: Rapport journalier du {date}

        Bonsoir Docteur,

        Vous trouverez ci-joint les statistiques détaillées d'aujourd'hui.

        - Les Rdvs pris : {x}
        {message}
        - Les Rdvs annulés: {annulations}
        - Les Rdvs déplacés : {deplacements}
        - Les Transferts : {transfert}
        - Les autres détails : {autres}
        - Total d'appels : {total}

        Meilleures salutations.
        """
        output_text.delete("1.0", ctk.END)
        output_text.insert(ctk.END, email)
        copy_text_button.configure(state=ctk.NORMAL)
    except Exception as e:
        output_text.delete("1.0", ctk.END)
        output_text.insert(ctk.END, f"Erreur: {str(e)}")
        copy_text_button.configure(state=ctk.DISABLED)

def drop(event):
    file_path = event.data.strip('{}').strip()
    generate_report(file_path)
    file_path_label.configure(text=f"Fichier: {file_path}")

def copy_text():
    window.clipboard_clear()
    window.clipboard_append(output_text.get("1.0", ctk.END))
    window.update()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

window = TkinterDnD.Tk()
window.title("Email Report Generator")
window.geometry("800x600")

frame = ctk.CTkFrame(window, corner_radius=10)
frame.pack(pady=20, padx=20, fill="both", expand=True)

font_style = ("Sofia Pro", 16)
file_label = ctk.CTkLabel(frame, text="Déposez le fichier Excel ici:", font=font_style)
file_label.pack(pady=10)

icon_path = "drag_icon.png"
if os.path.exists(icon_path):
    drag_icon = Image.open(icon_path)
    drag_icon = drag_icon.resize((50, 50))
    drag_icon = ImageTk.PhotoImage(drag_icon)
else:
    drag_icon = None
    print("Erreur : L'image drag_icon.png est introuvable.")

drag_label = ctk.CTkLabel(frame, image=drag_icon, text="")
drag_label.pack(pady=5)

file_path_label = ctk.CTkLabel(frame, text="", font=("Sofia Pro", 14))
file_path_label.pack()

output_text = ctk.CTkTextbox(frame, height=300, width=700, wrap="word", font=("Sofia Pro", 12))
output_text.pack(pady=10, padx=10)

copy_text_button = ctk.CTkButton(frame, text="Copier le texte", command=copy_text, state=ctk.DISABLED)
copy_text_button.pack(pady=10)

window.drop_target_register(DND_FILES)
window.dnd_bind('<<Drop>>', drop)

window.mainloop()
