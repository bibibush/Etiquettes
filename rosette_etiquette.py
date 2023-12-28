import os.path
from tkinter import *
from tkinter.messagebox import *
import pyautogui
import xlwings as xw

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

def etiquetter():
    dossier_isabelle = '//srvlabreche/Dossier semaine commun/EXPEDITIONS/Gestion CSE/Tableau pour Emballage + Récap Général Emballage'
    doc = f'{entry.get()}.xlsx'

    with xw.App(visible=False) as app:
        wb = xw.Book(os.path.join(dossier_isabelle, doc))
        sheets = []
        for i in range(5, 19 + 1):
            sheets.append(wb.sheets[i])
        ws = []
        for sh in sheets:
            if sh['U1'].value == 'fin':
                ws.append(sh)
                break
            ws.append(sh)

        for w in ws:
            et_list = []
            for i in range(6, 105 + 1):
                cli = w[f"B{i}:U{i}"]
                if not w[f"B{i}"].value:
                    break
                def clier(index):
                    if cli.columns[index].value:
                        return int(cli.columns[index].value)
                    else:
                        return 0
                hm = clier(3)
                hg = clier(4)
                arti = clier(5)
                ros_lyon = clier(6)
                fromage = clier(7)
                gibier = clier(8)
                spe = clier(9)
                mign = clier(10)
                cuire = clier(11)
                sabodet = clier(12)
                barq = clier(13)
                demi = clier(14)
                jam_os = clier(15)
                jam_6 = clier(16)
                jam_9 = clier(17)
                terrine = clier(18)

                total = {
                    "HM": hm, "HG": hg, "ARTI": arti, "FROMAGE": fromage, "GIBIER": gibier,
                    "SPECIALITE": spe, "MIGNONETTE": mign, "Á CUIRE": cuire, "SABODET": sabodet, "BARQUETTE": barq,
                    "DEMI-JAMBON": demi, "JAMBON-OS": jam_os, "JAMBON-6M": jam_6, "JAMBON-9M": jam_9,
                    "TERRINE": terrine, "ROSETTE": ros_lyon
                }
                dict_vide = {}

                lots = {"HM": total['HM'], "HG": total['HG'], "ARTI": total['ARTI'], "FROMAGE": total['FROMAGE'],
                        "GIBIER": total['GIBIER'], "SPECIALITE": total['SPECIALITE'],
                        "Á CUIRE": total['Á CUIRE'], "SABODET": total['SABODET']}

                lots_dic = {
                            "category": "lots", "count": sum(list(lots.values()))
                        }
                lots_vide = {}

                terrine_dic = {"category": "terrines", "count": total['TERRINE']}
                terrine_vide = {"TERRINE": total['TERRINE']}

                jambon = {"JAMBON-6M": total['JAMBON-6M'], "JAMBON-9M": total["JAMBON-9M"]}

                jambon_dic = {"category": "jambon", "count": sum(list(jambon.values()))}
                jambon_vide = {}

                side = {
                    "MIGNONETTE": total['MIGNONETTE'], "BARQUETTE": total['BARQUETTE'],
                    "DEMI-JAMBON": total['DEMI-JAMBON']
                }
                side_dic = {
                    "category": "side", "count": sum(list(side.values()))
                }
                side_vide = {}

                def vvc():
                    v = 0
                    for key, value in lots.items():
                        if value >= 4:
                            lots_vide[key] = 4 - v
                            dict_vide[key] = 4 - v
                            break
                        v += value
                        if v > 4:
                            lots_vide[key] = 4 - (v - value)
                            dict_vide[key] = 4 - (v - value)
                            break
                        lots_vide[key] = value
                        dict_vide[key] = value

                def lots_minus():
                    for keys in lots_vide:
                        total[keys] -= lots_vide[keys]

                def renew():
                    lots.update({"HM": total['HM'], "HG": total['HG'], "ARTI": total['ARTI'], "FROMAGE": total['FROMAGE'],
                        "GIBIER": total['GIBIER'], "SPECIALITE": total['SPECIALITE'],
                        "Á CUIRE": total['Á CUIRE'], "SABODET": total['SABODET']})
                    jambon.update({"JAMBON-6M": total['JAMBON-6M'], "JAMBON-9M": total['JAMBON-9M']})
                    side.update({
                    "MIGNONETTE": total['MIGNONETTE'], "BARQUETTE": total['BARQUETTE'],
                    "DEMI-JAMBON": total['DEMI-JAMBON']
                })

                    lots_dic["count"] = sum(list(lots.values()))
                    terrine_dic['count'] = total['TERRINE']
                    jambon_dic['count'] = sum(list(jambon.values()))
                    side_dic['count'] = sum(list(side.values()))

                    dict_vide.clear()
                    lots_vide.clear()
                    jambon_vide.clear()
                    side_vide.clear()

                def run_func():
                    def lots_func():
                        if sum(list(lots_vide.values())) == 2 and terrine_dic["count"] == 1:
                            dict_vide['TERRINE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['TERRINE'] -= 1
                        elif sum(list(lots_vide.values())) == 3 and total['MIGNONETTE'] == 1:
                            dict_vide['MIGNONETTE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['MIGNONETTE'] -= 1
                        elif sum(list(lots_vide.values())) == 3 and total['BARQUETTE'] == 1:
                            dict_vide['BARQUETTE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['BARQUETTE'] -= 1
                        elif sum(list(lots_vide.values())) == 3 and total['DEMI-JAMBON'] == 1:
                            dict_vide['DEMI-JAMBON'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['DEMI-JAMBON'] -= 1
                        elif sum(list(lots_vide.values())) == 1 and terrine_dic['count'] == 1 and total['DEMI-JAMBON'] == 1:
                            dict_vide['TERRINE'] = 1
                            dict_vide['DEMI-JAMBON'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['TERRINE'] -= 1
                            total['DEMI-JAMBON'] -= 1
                        elif sum(list(lots_vide.values())) == 1 and terrine_dic['count'] == 1 and total['BARQUETTE'] == 1:
                            dict_vide['TERRINE'] = 1
                            dict_vide['BARQUETTE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['TERRINE'] -= 1
                            total['BARQUETTE'] -= 1
                        elif sum(list(lots_vide.values())) == 1 and terrine_dic['count'] == 1 and total['MIGNONETTE'] == 1:
                            dict_vide['TERRINE'] = 1
                            dict_vide['MIGNONETTE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['TERRINE'] -= 1
                            total['MIGNONETTE'] -= 1
                        elif sum(list(lots_vide.values())) == 1 and jambon_dic['count'] == 1:
                            x = 0
                            for key, value in jambon.items():
                                if value >= 2:
                                    jambon_vide[key] = 2 - x
                                    break
                                x += value
                                if x > 2:
                                    break
                                jambon_vide[key] = value
                            dict_vide.update(jambon_vide)
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                            total['JAMBON-9M'] = 0
                            total['JAMBON-6M'] = 0
                        else:
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                                "ART": list(dict_vide.items())
                            }
                        return data

                    if lots_dic['count'] >= 4:
                        vvc()
                        et_list.append(lots_func())
                        lots_minus()
                        renew()
                    if lots_dic['count'] == 2 and terrine_dic["count"] == 1:
                        vvc()
                        dict_vide['TERRINE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['TERRINE'] -= 1
                        renew()
                    if lots_dic['count'] == 3 and total['MIGNONETTE'] == 1:
                        vvc()
                        dict_vide['MIGNONETTE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['MIGNONETTE'] -= 1
                        renew()
                    if lots_dic['count'] == 3 and total['BARQUETTE'] == 1:
                        vvc()
                        dict_vide['BARQUETTE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['BARQUETTE'] -= 1
                        renew()
                    if lots_dic['count'] == 3 and total['DEMI-JAMBON'] == 1:
                        vvc()
                        dict_vide['DEMI-JAMBON'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['DEMI-JAMBON'] -= 1
                        renew()
                    if lots_dic['count'] == 1 and terrine_dic['count'] == 1 and total['DEMI-JAMBON'] == 1:
                        vvc()
                        dict_vide['TERRINE'] = 1
                        dict_vide['DEMI-JAMBON'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['TERRINE'] -= 1
                        total['DEMI-JAMBON'] -= 1
                        renew()
                    if lots_dic['count'] == 1 and terrine_dic['count'] == 1 and total['BARQUETTE'] == 1:
                        vvc()
                        dict_vide['TERRINE'] = 1
                        dict_vide['BARQUETTE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['TERRINE'] -= 1
                        total['BARQUETTE'] -= 1
                        renew()
                    if lots_dic['count'] == 1 and terrine_dic['count'] == 1 and total['MIGNONETTE'] == 1:
                        vvc()
                        dict_vide['TERRINE'] = 1
                        dict_vide['MIGNONETTE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['TERRINE'] -= 1
                        total['MIGNONETTE'] -= 1
                        renew()
                    if lots_dic['count'] == 1 and jambon_dic['count'] == 1:
                        vvc()
                        x = 0
                        for key, value in jambon.items():
                            if value >= 2:
                                jambon_vide[key] = 2 - x
                                break
                            x += value
                            if x > 2:
                                break
                            jambon_vide[key] = value
                        dict_vide.update(jambon_vide)
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value, "nb_colis": 1,
                            "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['JAMBON-9M'] = 0
                        total['JAMBON-6M'] = 0
                        renew()
                    if terrine_dic["count"] >= 2:
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": [("TERRINE", 2)]
                        }
                        et_list.append(data)
                        total['TERRINE'] -= 2
                        renew()
                    if side_dic['count'] >= 4:
                        v = 0
                        for key, value in side.items():
                            if value >= 4:
                                side_vide[key] = 4 - v
                                dict_vide[key] = 4 - v
                                break
                            v += value
                            if v > 4:
                                side_vide[key] = 4 - (v - value)
                                dict_vide[key] = 4 - (v - value)
                                break
                            side_vide[key] = value
                            dict_vide[key] = value
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        for keys in side_vide:
                            total[keys] -= side_vide[keys]
                        renew()

                    if total['JAMBON-OS'] >= 2:
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": [("JAMBON-OS", 2)]
                        }
                        et_list.append(data)
                        total['JAMBON-OS'] -= 2
                        renew()
                    if total['JAMBON-OS'] == 1:
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": [("JAMBON-OS", 1)]
                        }
                        et_list.append(data)
                        total['JAMBON-OS'] -= 1
                        renew()
                    if jambon_dic['count'] >= 2:
                        x = 0
                        for key, value in jambon.items():
                            if value >= 2:
                                jambon_vide[key] = 2 - x
                                break
                            x += value
                            if x > 2:
                                break
                            jambon_vide[key] = value
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": list(jambon_vide.items())
                        }
                        et_list.append(data)
                        for key in jambon_vide:
                            total[key] -= jambon_vide[key]
                        renew()
                    if total['ROSETTE'] >= 2:
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": [("ROSETTE", 2)]
                        }
                        et_list.append(data)
                        total['ROSETTE'] -= 2
                        renew()
                    if total['ROSETTE'] == 1:
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": [("ROSETTE", 1)]
                        }
                        et_list.append(data)
                        total['ROSETTE'] -= 1
                        renew()
                    if lots_dic['count'] == 3 and total['TERRINE'] == 1:
                        v = 0
                        for key, value in lots.items():
                            if value >= 2:
                                lots_vide[key] = 2 - v
                                dict_vide[key] = 2 - v
                                break
                            v += value
                            if v > 2:
                                break
                            lots_vide[key] = value
                            dict_vide[key] = value
                        dict_vide['TERRINE'] = 1
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        lots_minus()
                        total['TERRINE'] -= 1
                        renew()
                    if side_dic['count'] >= 2:
                        v = 0
                        for key, value in side.items():
                            if value >= 4:
                                side_vide[key] = 4 - v
                                dict_vide[key] = 4 - v
                                break
                            v += value
                            if v > 4:
                                side_vide[key] = 4 - (v - value)
                                dict_vide[key] = 4 - (v - value)
                                break
                            side_vide[key] = value
                            dict_vide[key] = value
                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": list(dict_vide.items())
                        }
                        et_list.append(data)
                        for keys in side_vide:
                            total[keys] -= side_vide[keys]
                        renew()

                while sum(list(total.values())) != 0:
                    run_func()
                    if 0 < sum(list(total.values())) <= 4 and total['ROSETTE'] == 0 and total['JAMBON-OS'] == 0 and total['TERRINE'] <= 1 and jambon_dic['count'] <= 1:
                        if lots_dic['count'] + side_dic['count'] == 3 and total['TERRINE'] == 1:
                            dict_vide.update(lots)
                            dict_vide.update(side)
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": list(dict_vide.items())
                            }
                            data2 = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": [("TERRINE", 1)]
                            }
                            et_list.append(data)
                            et_list.append(data2)
                            for key in total:
                                total[key] = 0
                            break

                        if lots_dic['count'] + side_dic['count'] == 2 and total['TERRINE'] == 1 and jambon_dic['count'] == 1:
                            dict_vide.update(lots)
                            dict_vide.update(side)
                            dict_vide['TERRINE'] = 1
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": list(dict_vide.items())
                            }
                            data2 = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": list(jambon.items())
                            }
                            et_list.append(data)
                            et_list.append(data2)
                            for key in total:
                                total[key] = 0
                            break

                        if lots_dic['count'] + side_dic['count'] >= 2 and jambon_dic['count'] == 1:
                            dict_vide.update(lots)
                            dict_vide.update(side)
                            data = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": list(dict_vide.items())
                            }
                            data2 = {
                                "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                                "nb_colis": 1, "ART": list(jambon.items())
                            }
                            et_list.append(data)
                            et_list.append(data2)
                            for key in total:
                                total[key] = 0
                            break

                        data = {
                            "nom": cli.columns[0].value, "prenom": cli.columns[1].value,
                            "nb_colis": 1, "ART": list(total.items())
                        }
                        et_list.append(data)
                        for key in total:
                            total[key] = 0
                        break
            for et in et_list:
                et['nom'] = et['nom'].upper()
                if not et['prenom'] is None:
                    et['prenom'] = et['prenom'].capitalize()
            et_list.sort(key=lambda x:x["nom"])
            for et in et_list:
                et['ART'] = [e for e in et['ART'] if e[1] != 0]
                et['ART'] = dict(et['ART'])
                print(et)

            s = wb.sheets['Etiquettes']
            for et_num in range(len(et_list)):
                if et_num - 1 != -1 and et_list[et_num]['nom'] == et_list[et_num - 1]['nom'] and et_list[et_num]['prenom'] == et_list[et_num - 1]['prenom']:
                    et_list[et_num]['number'] = et_list[et_num - 1]['number']
                elif et_num == 0:
                    et_list[et_num]['number'] = 1
                else:
                    et_list[et_num]['number'] = et_list[et_num - 1]['number'] + 1

                s_col = s['A1: H1500']
                for s_c in s_col.columns[0]:
                    if not s_c.value:
                        s_c.value = [
                            w['A1'].value, w['A2'].value, et_list[et_num]['number'],et_list[et_num]['nb_colis'] ,et_list[et_num]['nom'],
                            et_list[et_num]['prenom']
                        ]
                        break
                    else:
                        pass
                for s_a in s_col.columns[6]:
                    if not s_a.value:
                            s_a.value = [f'{value} {key}' for key, value in list(et_list[et_num]['ART'].items())]
                            break
                    else:
                        pass
        wb.save(os.path.join(dossier_isabelle, doc))
        wb.close()
    print("parfait !")
    pyautogui.alert("Parfait !")

def check_etiquetter(event):
    if askyesno(title="Confirmation", message="Vous voulez lancer le programme ?"):
        try:
            etiquetter()
        except Exception as e:
            print(e)
            pyautogui.alert("Appelez Jin hyeong !")

btn.bind('<Button-1>', check_etiquetter)
root.mainloop()









