# Salaisons de la Brèche 라벨 제작 자동화

## 소개
회사는 CSE라는 B2B 시스템을 가지고 있습니다.  

간단히 말하자면, 어떤 다른 회사에서 Saucissons(쏘시쏭)을 구매하고 싶은 직원들의 명단과 품목을
저희 회사에 제출하면 포장박스마다 각 명단과 품목이 적혀져있는 라벨을 붙여 제품들을 넣고 파렛트에 실어 배달하는 방식입니다.

문제는 포장박스의 크기가 정해져있기 때문에 한 고객이 많은 품목을 원하면 하나의 포장박스에 다 넣을 수 없습니다.

그래서 여러개의 박스가 하나의 고객을 위한 경우가 많습니다.

원래 문서 담당 비서가 명단과 품목이 적혀져 있는 엑셀 파일을 읽고 일일이 포장박스의 개수를 계산해 한 시트에 표를 만들고,  
워드파일의 편지 레이블 기능을 사용해 라벨을 만들었습니다.

하지만 주문이 많은 시즌엔 모든 주문을 이렇게 처리하기가 너무 번거롭다는 문제가 생겼습니다.

이런 번거러움을 덜기 위해 포장박스의 사이즈에 맞게 명단과 품목들이 나열된 표를 자동으로 만드는 프로그램을 개발했습니다.

<br />

## 사용한 기술

<br />

- Python:3.11


- xlwings


- pyautogui


- tkinter

<br />

## 핵심 코드

포장 박스를 채울 수 있는 경우의 수가 많기 때문에 코드가 굉장히 길어졌습니다.  
모든 코드를 보여드리는 것이 아닌 어떤 방식으로 코드를 작성했는지 간략한 핵심 코드만 보여드리겠습니다.

<details>
<summary><b>코드보기</b></summary>

한 사람이 주문한 모든 제품의 개수를 total 이라는 딕셔너리 객체를 만들어 품목을 key로 개수를 value로 만들어 정리했습니다.
```python
total = {
            "HM": hm, "HG": hg, "ARTI": arti, "FROMAGE": fromage, "GIBIER": gibier,
            "SPECIALITE": spe, "MIGNONETTE": mign, "Á CUIRE": cuire, "SABODET": sabodet, "BARQUETTE": barq,
            "DEMI-JAMBON": demi, "JAMBON-OS": jam_os, "JAMBON-6M": jam_6, "JAMBON-9M": jam_9,
            "TERRINE": terrine,
        }
```

또한 품목 별로 나눈 객체와 품목의 개수를 총합한 객체도 만들었습니다.
```python
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
```

<br />

이제 한 포장박스를 담을 수 있는 경우의 수 예를 들면, 
```python
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
```
side라는 품목들이 최대 4개까지만 되도록 하여 성, 이름, 상자 개수, 품목 개수를 나타내는 data 딕셔너리 객체를 만들어 리스트에 넣었습니다.  
그리고 total 딕셔너리 객체에서 data 딕셔너리 객체에 있던 개수를 빼줘 다시 코드가 반복할 때 남은 개수가 data 객체에 담아질 것입니다.

<br />

각 품목들의 총 개수의 정보를 업데이트 해주는 renew 함수를 정의했습니다.
```python
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
```

<br />

이런 식으로 total객체의 value들의 총 합이 0이 될때 까지 반복할 수 있도록 while문을 사용했습니다.
```python
while sum(list(total.values())) != 0:
```
이렇게 리스트에 넣은 데이터들을 lambda 함수를 사용해 성을 기준으로 정렬해주었습니다.
```python
for et in et_list:
    et['nom'] = et['nom'].upper()
    if not et['prenom'] is None:
        et['prenom'] = et['prenom'].capitalize()
et_list.sort(key=lambda x:x["nom"])
```

<br />

이렇게 정리한 데이터들을 Etiquettes 이름을 가진 시트에 보여주는 코드를 작성했습니다.
```python
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
```

이 시트를 워드파일에서 불러와 편지/레이블를 사용해 라벨을 만들 수 있습니다.
</details>

<br />

## 개선할 점

이 코드를 작성한 시점에는 Pandas 라이브러리를 배우기 전 이였기 때문에, 코드가 복잡하고 조건문이 너무 많다는 단점이 있습니다.  

Pandas 데이터프레임 기능을 사용하면 코드를 더울 깔끔하게 작성할 수 있어, 프로그램의 실행속도을 향상시킬 수 있을 것같습니다.

아직 Pandas 공부를 끝내지 못했고 숙련도가 낮아, 다른 프로젝트로 연습을 한 뒤 코드를 Pandas 라이브러리를 사용해 수정할 예정입니다.