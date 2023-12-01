import os
import xlwings as xw
from tkinter import *
from tkinter.messagebox import *
import pyautogui

root = Tk()
root.title("Programme d'etiquette")
label = Label(root, text="Mettez le nom de fiche CSE sans '.xlsx'")
label_vide = Label(root, text="", width=5, height=10)
entry = Entry(root, width=15)
btn = Button(root, width=10, text="Lancer")
label.grid(row=0, column=0, columnspan = 3)
label_vide.grid(row=1, column=0)
label_vide.grid(row=1, column=1)
label_vide.grid(row=1, column=2)
entry.grid(row=2, column=0)
label_vide.grid(row=2, column=1)
btn.grid(row=2, column=2)

entry.insert(0, "")

def colis():
    dossier_isabelle = '//srvlabreche\Dossier semaine commun\Doc. Isabelle\JIN'
    doc = f'{entry.get()}.xlsx'
    with xw.App(visible=False) as app:
        wb = xw.Book(os.path.join(dossier_isabelle, doc))
        ws = []
        etiquette = wb.sheets['Etiquettes']
        for i in range(5, 19 + 1):
            ws.append(wb.sheets[i])
        et_rng = etiquette['C2:F1500']
        ls = []
        numero = 1
        nb_colis = 0
        nom = None
        prenom = None
        for row in et_rng.rows:
            dic = []
            if not row.columns[0].value:
                dic.extend([numero, nb_colis, nom, prenom])
                ls.append(dic)
                break
            if row.columns[0].value != numero:
                dic.extend([numero, nb_colis, nom, prenom])
                ls.append(dic)
                nb_colis = 0
            if row.columns[0].value == numero and row.columns[2].value != nom:
                dic.extend([numero, nb_colis, nom, prenom])
                ls.append(dic)
                nb_colis = 0
            if row.columns[0].value == numero and row.columns[2].value == nom and row.columns[3].value != prenom:
                dic.extend([numero, nb_colis, nom, prenom])
                ls.append(dic)
                nb_colis = 0
            numero = row.columns[0].value
            nb_colis += row.columns[1].value
            nom = row.columns[2].value
            prenom = row.columns[3].value
            print(ls)
        for w in ws:
            if not w['B6'].value:
                break
            name_rng = w['B6:D105']
            nom = None
            prenom = None
            for name_row in name_rng.rows:
                if name_row.columns[0].value == nom and name_row.columns[1].value == prenom:
                    name_row.columns[2].value = 0
                else:
                    nom = name_row.columns[0].value
                    if not name_row.columns[1].value is None:
                        prenom = name_row.columns[1].value
                    else:
                        prenom = None
                    if not nom:
                        break
                    for dc in ls:
                        if dc[2] == nom.upper() and dc[3] == prenom.capitalize():
                            name_row.columns[2].value = dc[1]
                            break
        wb.save(os.path.join(dossier_isabelle, doc))
        wb.close()
    print('parfait !')
    pyautogui.alert("parfait !")

def check_colis():
    if askyesno(title="confirmation", message="Vous voulez lancer le programme ?"):
        try:
            colis()
        except Exception as e:
            print(e)
            pyautogui.alert("Appelez Jin hyeong !")

btn.bind('<Button-1>', check_colis())
root.mainloop()


