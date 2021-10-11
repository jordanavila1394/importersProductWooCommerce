
# import pandas as pd
# data = pd.read_excel(
#     r'C:\Users\Avila\Desktop\GiemmeHats\ExelToCsvGiemmeHats\inputExcel\databaseGiemme.xlsx')
# df = pd.DataFrame(data, columns=['MODELLO', 'CAPPELLO', 'GRADI'])
# df.to_csv(index=False)
# print(df)
import csv
from datetime import date
today = date.today()
day = today.strftime("%m-%d-%y")
print("Generate date =", day)

header = ['ID', 'Tipo', 'Nome', 'Pubblicato', '\"In primo piano?\"', '\"Visibilità nel catalogo\"', '\"Stato delle imposte\"', '\"Aliquota di imposta\"', '\"In stock?\"', '\"Abilita gli ordini arretrati?\"', '\"Venduto singolarmente?\"', '\"Permetti le recensioni clienti?\"', 'Genitore', 'Posizione', '\"Attributo 1 nome\"',
          '\"Attributo 1 valuta(e)\"', '\"Attributo 1 visibile\"', '\"Attributo 1 globale\"', '\"Attributo 2 nome\"', '\"Attributo 2 valuta(e)\"', '\"Attributo 2 visibile\"', '\"Attributo 2 globale\"', '\"Attributo 3 nome\"', '\"Attributo 3 valuta(e)\"', '\"Attributo 3 visibile\"', '\"Attributo 3 globale\"', '\"Attributo 4 nome\"', '\"Attributo 4 valuta(e)\"', '\"Attributo 4 visibile\"', '\"Attributo 4 globale\"', '\"Attributo 5 nome\"', '\"Attributo 5 valuta(e)\"', '\"Attributo 5 visibile\"', '\"Attributo 5 globale\"', '\"Attributo 6 nome\"', '\"Attributo 6 valuta(e)\"', '\"Attributo 6 visibile\"', '\"Attributo 6 globale"']
data = ['Afghanistan', 652090, 'AF', 'AFG']
grades = [14, 16, 18, 19, 20, 22, 23, 25, 27, 30, 40]  # gradi
misures = [55, 56, 57, 58, 59, 60, 61, 62]  # misure
brims = [4, '\"4\,5\"', 5, 6, 7, 8, 9, 10, 13, 14, 15]  # tesas
names = ['Montecristi', 'Domingo', 'Triunfo', 'Corazon', 'Portovelo', 'Duran']
row = []


def getCodeParent(name):
    return {
        'Montecristi': 2570,
        'Domingo': 2571,
        'Triunfo': 2572,
        'Corazon': 2573,
        'Portovelo': 2574,
        'Duran': 2575
    }.get(name, 2570)


with open('wc-product-export-'+day+'.csv', 'w', newline='', encoding='UTF8') as f:
    writer = csv.writer(
        f, delimiter=';', quoting=csv.QUOTE_NONE, quotechar=None)

    # write the header
    writer.writerow(header)
    for name in names:
        row.append("variation")
        print(name)
        row.append(name)
        row.append(1)
        row.append(0)
        row.append("visible")
        row.append("taxable")
        row.append("parent")
        row.append(1)
        row.append(0)
        row.append(0)
        row.append(0)
        row.append(getCodeParent(name))
        row.append("Gradi")
        for grade in grades:
            row.append(grade)
            row.append('')
            row.append(0)
            for misure in misures:
                row.append(misure)
                row.append('')
                row.append(0)
                for brim in brims:
                    row.append(brim)
                    row.append('')
                    row.append(0)
                    writer.writerow(row)
                    row.pop()
                    row.pop()
                    row.pop()
                row.pop()
                row.pop()
                row.pop()
            row.pop()
            row.pop()
            row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
        row.pop()
    # write the data
    writer.writerow(data)
    writer.writerow(data)
