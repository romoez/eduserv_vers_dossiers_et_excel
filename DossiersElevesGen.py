import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter import filedialog
import xml.etree.ElementTree as ET
import xlsxwriter
import ctypes
import os


# ================================================================================================
# Constantes
EMAIL = "moez.romdhane@tarbia.tn"
VERSION = "1.0"


# ================================================================================================
# Définition des fonctions
def select_destination():
    var_rapport.set("")
    dest = filedialog.askdirectory()
    if dest != "":
        var_destination.set(dest)
        if is_dir_writable(dest):
            afficher_exmple_de_dossier()
        else:
            var_destination.set("Vous n'avez pas d'accès en écriture pour ce chemin !")
            var_exemple_dossier.set("-")


def select_fichier_xml():
    var_rapport.set("")
    fichier_xml = filedialog.askopenfilename(title="Sélectionner le fichier XML", filetypes=(
        ("Fichiers xml", "*.xml"), ("Tous les fichiers", "*.*")))
    if fichier_xml != "":
        var_fichier_xml.set(fichier_xml)
        liste_eleves.clear()
        mise_à_jour_liste(fichier_xml)


def mise_à_jour_liste(fichier_xml):
    liste_eleves.clear()
    if os.path.splitext(fichier_xml)[1].lower() != ".xml":
        var_fichier_xml.set("Format de ficher non valide !")
        return

    try:
        tree = ET.parse(fichier_xml)
    except Exception:
        var_fichier_xml.set("Format de ficher non valide !")
        return

    classe = tree.find("libeclass")
    if classe is None:
        var_fichier_xml.set("Format de ficher non valide !")
        var_classe.set("-")
        var_exemple_dossier.set("-")
        return

    var_classe.set(classe.text)
    for e in tree.findall("noteelev"):
        num = nettoyer(e.find("numOrdre").text)
        nom = nettoyer(e.find("prenomnom").text)
        liste_eleves.add(f"{num:0>2}-{nom}")

    if liste_eleves:
        afficher_exmple_de_dossier()
    else:
        var_fichier_xml.set("Format de ficher non valide !")
        var_classe.set("-")
        var_exemple_dossier.set("-")


def nettoyer(ch: str) -> str:
    ch = ch.strip()
    p = ch.find("  ")
    while p >= 0:
        ch = ch[:p] + ch[p + 1:]
        p = ch.find("  ")
    return ch


def destination_valide():
    destination = var_destination.get()
    a = is_dir_writable(destination)
    print(destination)
    print(a)
    return a


def afficher_exmple_de_dossier():
    destination = var_destination.get()
    if len(liste_eleves) > 0 and os.path.exists(destination):
        if var_sous_dossier.get():
            exemple_de_dossier = os.path.join(var_destination.get(), var_classe.get(),
                                              list(liste_eleves)[0]).replace('\\', '/')
        else:
            exemple_de_dossier = os.path.join(var_destination.get(), list(liste_eleves)[0]).replace('\\', '/')
        var_exemple_dossier.set(exemple_de_dossier)


def is_dir_writable(dir: str) -> bool:
    if dir == "":
        return False

    test_file = os.path.join(dir, "pyschool_dummy_file.txt")
    try:
        with open(test_file, "w") as file:
            file.write("Hello!")
        if os.path.exists(test_file):
            os.remove(test_file)
    except IOError:
        return False
    return True


def créer_dossier():
    var_rapport.set("")
    destination = var_destination.get()
    if len(liste_eleves) > 0 and os.path.exists(destination):
        if var_sous_dossier.get():
            destination = os.path.join(destination, var_classe.get())
        nb_problèmes = 0
        for d in liste_eleves:
            path = os.path.join(destination, d)
            try:
                os.makedirs(path)
            except FileExistsError:
                nb_problèmes += 1
        var_rapport.set(f"{len(liste_eleves) - nb_problèmes} dossiers créés ■ {nb_problèmes} erreur(s).")
    elif len(liste_eleves) == 0:
        var_rapport.set("Veuillez vérifier le fichier XML.")
    elif not os.path.exists(destination):
        var_rapport.set("Veuillez vérifier le dossier de destination.")


def ouvrir_destination(e, type):
    if type == "dossier":
        destination = var_destination.get()
        if destination != "" and os.path.exists(destination):
            os.startfile(destination)
    elif type == "fichier":
        destination = var_fichier_xml.get()
        if destination != "" and os.path.exists(destination):
            os.startfile(os.path.dirname(destination))


def créer_excel():
    var_rapport.set("")
    destination = var_destination.get()
    if len(liste_eleves) > 0 and os.path.exists(destination):
        fichier_xl = os.path.join(destination, var_classe.get() + ".xlsx")
        workbook = xlsxwriter.Workbook(fichier_xl, {'strings_to_numbers': True})
        worksheet = workbook.add_worksheet(var_classe.get())

        bold = workbook.add_format({'bold': True})
        nombre = workbook.add_format({'num_format': '00'})
        worksheet.write('A1', 'رقم', bold)
        worksheet.write('B1', 'الإسم', bold)

        row = 1
        col = 0
        lliste_eleves = list(liste_eleves)
        lliste_eleves.sort()
        largeur_col_1 = 0
        for e in lliste_eleves:
            worksheet.write(row, col, e[:2], nombre)
            worksheet.write(row, col + 1, e[3:])
            if len(e[3:]) > largeur_col_1:
                largeur_col_1 = len(e[3:])
            row += 1
        worksheet.set_column(1, 1, largeur_col_1 + 2)
        worksheet.set_column(0, 0, 4)
        try:
            workbook.close()
            var_rapport.set("Fichier excel crée avec succés.")
        except xlsxwriter.exceptions.FileCreateError:
            var_rapport.set("Erreur: fermer le fichier Excel et réessayer.")
    elif len(liste_eleves) == 0:
        var_rapport.set("Veuillez vérifier le fichier XML.")
    elif not os.path.exists(destination):
        var_rapport.set("Veuillez vérifier le dossier de destination.")


def à_propos():
    message = "(Eduserv XML) vers Dossiers | Excel\n\n"
    message += f"version: {VERSION}\n"
    message += f"contact: {EMAIL}\n"
    messagebox.showinfo("À Propos", message)


# ================================================================================================
# Variables Globales
liste_eleves = set()

# ================================================================================================
# ================================================================================================
# fix taskbar icon: https://stackoverflow.com/a/1552105
myappid = "(Eduserv_XML)_vers_Dossiers_|_Excel"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


root = tk.Tk()


root.resizable(1, 0)
icon = os.path.join(os.path.dirname(os.path.realpath(__file__)), "DossiersEleves.ico")
if os.path.exists(icon):
    root.iconbitmap(icon)
root.grid_columnconfigure(1, weight=10)
# root.grid_columnconfigure((0, 2), weight=1)
root.title("(Eduserv XML) vers Dossiers | Excel")

# -------------------------------------------------------
row = 0
btn_fichier_xml = ttk.Button(root,
                             text="Sél. fichier XML",
                             compound="center",
                             command=select_fichier_xml)
btn_fichier_xml.grid(row=row, column=0, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=12)

var_fichier_xml = tk.StringVar()
ed_fichier_xml = ttk.Entry(root, width=75, state="readonly", textvariable=var_fichier_xml)


ed_fichier_xml.grid(row=row, column=1, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=12, columnspan=2)
ed_fichier_xml.bind("<Double-Button>", lambda event: ouvrir_destination(event, "fichier"))
# -------------------------------------------------------
row += 1
btn_destination = ttk.Button(root,
                             text="Dossier distination",
                             compound="center",
                             command=select_destination)
btn_destination.grid(row=row, column=0, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=12)

var_destination = tk.StringVar()
ed_destination = ttk.Entry(root, width=75, state="readonly", textvariable=var_destination)
ed_destination.grid(row=row, column=1, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=12, columnspan=2)
ed_destination.bind("<Double-Button>", lambda event: ouvrir_destination(event, "dossier"))
# -------------------------------------------------------
row += 1
ttk.Label(text="Classe : ", anchor="e").grid(
    row=row, column=0, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=5)


var_classe = tk.StringVar()
var_classe.set("-")
lbl_classe = ttk.Label(root, textvariable=var_classe, foreground="Green",
                       font=("bold"))
lbl_classe.grid(row=row, column=1, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=5, columnspan=2)
# -------------------------------------------------------
row += 1
ttk.Label(text="Exemple de dossier : ", anchor="e").grid(
    row=row, column=0, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=5)

var_exemple_dossier = tk.StringVar()
var_exemple_dossier.set("-")
lbl_exemple_dossier = ttk.Label(root, textvariable=var_exemple_dossier, foreground="Blue")
lbl_exemple_dossier.grid(row=row, column=1, ipady=5, ipadx=5, sticky="NSEW", padx=5, pady=5, columnspan=2)
# -------------------------------------------------------
row += 1
var_sous_dossier = tk.IntVar()
chk_sous_dossier = tk.Checkbutton(root, text='Sous-dossier', variable=var_sous_dossier,
                                  onvalue=1, offvalue=0, command=afficher_exmple_de_dossier)
chk_sous_dossier.grid(row=row, column=0, ipady=5, ipadx=5, sticky="W", padx=5, pady=12)

# ttk.Label(root, text="", foreground="#888888", font=("Consolas", 9),
#           background="#ee1111").grid(row=row, column=1, ipadx=5, ipady=5, sticky="NSEW")

btn_dossiers = ttk.Button(root,
                          text="Créer les dossiers des élèves",
                          compound="center",
                          command=créer_dossier)
btn_dossiers.grid(row=row, column=1, ipady=5, ipadx=5, sticky="NSEW", padx=(0, 0), pady=12)


# ttk.Label(root, text="", foreground="#888888", font=("Consolas", 9),
#           background="#11ee11").grid(row=row, column=2, ipadx=5, ipady=5, sticky="NSEW")

btn_excel = ttk.Button(root,
                       text="Générer Fichier Excel",
                       compound="center",
                       command=créer_excel)
btn_excel.grid(row=row, column=2, ipady=5, ipadx=5, sticky="NSEW", padx=(5, 5), pady=12)
# -------------------------------------------------------
row += 1
ttk.Label(root, text="background de la barrre d'état", foreground="#323232", background="#323232",
          anchor="e").grid(row=row, column=0, columnspan=3, ipadx=5, ipady=5, sticky="NSEW")
var_rapport = tk.StringVar()
var_rapport.set("")
lbl_rapport = ttk.Label(root, textvariable=var_rapport, foreground="#00BFFF",
                        background="#323232", anchor="w")
lbl_rapport.grid(row=row, column=0, ipadx=5, ipady=5, columnspan=2, sticky="NSEW", padx=5, )

btn_à_propos = tk.Button(root,
                         text="À Propos...",
                         compound="center",
                         bg='#323232',
                         fg='#AAAAAA',
                         relief='flat',
                         command=à_propos, )
btn_à_propos.grid(row=row, column=2, ipady=5, ipadx=5, sticky="NSEW", padx=(5, 5), pady=12)
# -------------------------------------------------------
root.update_idletasks()
largeur_fenêtre = root.winfo_width()
hauteur_fenêtre = root.winfo_height()

largeur_écran = root.winfo_screenwidth()
hauteur_écran = root.winfo_screenheight()

position_horizontale = int(largeur_écran / 2 - largeur_fenêtre / 2)
position_verticale = int(hauteur_écran / 2 - hauteur_fenêtre / 2)

root.geometry(f"+{position_horizontale}+{position_verticale}")
tk.mainloop()
