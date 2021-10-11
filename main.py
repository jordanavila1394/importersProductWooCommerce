
# import pandas as pd
# data = pd.read_excel(
#     r'C:\Users\Avila\Desktop\GiemmeHats\ExelToCsvGiemmeHats\inputExcel\databaseGiemme.xlsx')
# df = pd.DataFrame(data, columns=['MODELLO', 'CAPPELLO', 'GRADI'])
# df.to_csv(index=False)
# print(df)
import csv
import os
import glob

from datetime import date
today = date.today()
day = today.strftime("%m-%d-%y")
print("Generate date =", day)

header = ['Tipo', 'COD', 'Nome', 'Pubblicato', '\"In primo piano?\"', '\"Visibilità nel catalogo\"', '\"Stato delle imposte\"', '\"Aliquota di imposta\"', '\"In stock?\"', '\"Abilita gli ordini arretrati?\"', '\"Venduto singolarmente?\"', '\"Permetti le recensioni clienti?\"', 'Genitore', 'Posizione', '\"Attributo 1 nome\"',
          '\"Attributo 1 valuta(e)\"', '\"Attributo 1 visibile\"', '\"Attributo 1 globale\"', '\"Attributo 2 nome\"', '\"Attributo 2 valuta(e)\"', '\"Attributo 2 visibile\"', '\"Attributo 2 globale\"', '\"Attributo 3 nome\"', '\"Attributo 3 valuta(e)\"', '\"Attributo 3 visibile\"', '\"Attributo 3 globale\"', '\"Attributo 4 nome\"', '\"Attributo 4 valuta(e)\"', '\"Attributo 4 visibile\"', '\"Attributo 4 globale\"', '\"Attributo 5 nome\"', '\"Attributo 5 valuta(e)\"', '\"Attributo 5 visibile\"', '\"Attributo 5 globale\"', '\"Attributo 6 nome\"', '\"Attributo 6 valuta(e)\"', '\"Attributo 6 visibile\"', '\"Attributo 6 globale"']
grades = [14, 16, 18, 19, 20, 22, 23, 25, 27, 30, 40]  # gradi
misures = [55, 56, 57, 58, 59, 60, 61, 62]  # misure
brims = [4, '\"4\,5\"', 5, '\"5\,5\"', 6, '\"6\,5\"', 7, '\"7\,5\"',
         8, '\"8\,5\"', 9, 10, '\"12\,5\"', 13, 14, 15]  # tesas
externalRibbons = ['\"VERDE BOSCO\"', '\"CALCEDONIA\"', '\"BRUN SUDAN\"', '\"ARANCIO\"',
                   '\"BLU REALE\"', '\"NERO\"', '\"AZALIA\"', '\"INDOLO NERO\"', '\"VINATTO\"', '\"ROSSO\"']
internalRibbons = ['\"GROS GRAIN AVORIO\"',
                   '\"GROS GRAIN NERO\"', '\"PELLE CHIARA\"', '\"PELLE NERA\"']

names = ['Montecristi', 'Domingo', 'Triunfo', 'Corazon', 'Portovelo',
         'Duran', 'Angel', 'Olmedo', 'Vinces', 'Isabela', 'Salinas', 'Tulcan']
#names = ['Montecristi', 'Domingo']

row = []
indexOrder = 0


def getCodeParent(name):
    return {
        'Montecristi': 2570,
        'Domingo': 2571,
        'Triunfo': 2572,
        'Corazon': 2573,
        'Portovelo': 2574,
        'Duran': 2575,
        'Angel': 2576,
        'Olmedo': 2577,
        'Vinces': 2578,
        'Isabela': 2579,
        'Salinas': 2580,
        'Tulcan': 2581,
    }.get(name, 2570)


def getBow(name):
    return {
        'Montecristi': '\"ELOY ALFARO\"',
        'Domingo': '\"DOMINGO CHOEZ\"',
        'Triunfo': '\"JULIO JARAMILLO\"',
        'Corazon': '\"MANUELITA SAEZ\"',
        'Portovelo': '\"OSWALDO GUAYASAMIN\"',
        'Duran': '\"ANA PERALTA\"',
        'Angel': '\"HERMELINDA URVINA\"',
        'Olmedo': '\"RAMIRO SANCHEZ\"',
        'Vinces': '\"KARINA GALVEZ\"',
        'Isabela': '\"ISABEL DE CASTILLA\"',
        'Salinas': '\"PADRE POLO ANTONIO\"',
        'Tulcan': '\"RICHARD CARAPAZ\"',
    }.get(name, 'ERROR')


def generateCOD(array):
    codeHat = {
        'Montecristi': 'GM01',
        'Domingo': 'GM02',
        'Triunfo': 'GM03',
        'Corazon': 'GM04',
        'Portovelo': 'GM05',
        'Duran': 'GM06',
        'Angel': 'GM07',
        'Olmedo': 'GM08',
        'Vinces': 'GM09',
        'Isabela': 'GM10',
        'Salinas': 'GM11',
        'Tulcan': 'GM012',
    }.get(array[1], 'ERROR')
    codeBrim = {
        4: 'A',
        '\"4\,5\"': 'B',
        5: 'C',
        '\"5\,5\"': 'D',
        6: 'E',
        '\"6\,5\"': 'F',
        7: 'G',
        '\"7\,5\"': 'H',
        8: 'I',
        '\"8\,5\"': 'J',
        9: 'K',
        10: 'M',
        '\"12\,5\"': 'R',
        13: 'S',
        14: 'U',
        15: 'W'
    }.get(array[21], 'ERROR')
    codeExternalRibbon = {
        '\"VERDE BOSCO\"': '351',
        '\"CALCEDONIA\"': '352',
        '\"BRUN SUDAN\"': '353',
        '\"ARANCIO\"': '354',
        '\"BLU REALE\"': '355',
        '\"NERO\"': '356',
        '\"AZALIA\"': '151',
        '\"INDOLO NERO\"': '152',
        '\"VINATTO\"': '471',
        '\"ROSSO\"': '472'
    }.get(array[25], 'ERROR')

    codeInternalRibbon = {
        '\"GROS GRAIN AVORIO\"': '321',
        '\"GROS GRAIN NERO\"': '322',
        '\"PELLE CHIARA\"': '323',
        '\"PELLE NERA\"': '324'
    }.get(array[29], 'ERROR')
    codeBow = {
        '\"ELOY ALFARO\"': 'LEA',
        '\"DOMINGO CHOEZ\"': 'LDC',
        '\"JULIO JARAMILLO\"': 'LJJ',
        '\"MANUELITA SAEZ\"': 'LMS',
        '\"OSWALDO GUAYASAMIN\"': 'LOG',
        '\"ANA PERALTA\"': 'LAP',
        '\"HERMELINDA URVINA\"': 'LHU',
        '\"RAMIRO SANCHEZ\"': 'LRS',
        '\"KARINA GALVEZ\"': 'LKG',
        '\"ISABEL DE CASTILLA\"': 'LIC',
        '\"PADRE POLO ANTONIO\"': 'LPA',
        '\"RICHARD CARAPAZ\"': 'LRC',
    }.get(array[33], 'ERROR')

    finalCode = codeHat + \
        str(array[13]) + str(array[17]) + codeBrim + \
        codeExternalRibbon + codeInternalRibbon + codeBow
    print(finalCode)
    return (finalCode)


for name in names:
    fileName = 'wc-product-export-'+day+'-'+name.lower()
    with open(fileName+'.csv', 'w', newline='', encoding='UTF8') as f:
        writer = csv.writer(
            f, delimiter=';', quoting=csv.QUOTE_NONE, quotechar=None)
        writer.writerow(header)
        row.append("variation")
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
                row.append('Misura')
                row.append(misure)
                row.append('')
                row.append(0)
                for brim in brims:
                    row.append('Tesa')
                    row.append(brim)
                    row.append('')
                    row.append(0)
                    for externalRibbon in externalRibbons:
                        row.append('\"Nastro Esterno\"')
                        row.append(externalRibbon)
                        row.append('')
                        row.append(0)
                        for internalRibbon in internalRibbons:
                            row.append('\"Nastro Interno\"')
                            row.append(internalRibbon)
                            row.append('')
                            row.append(0)
                            row.append('Fiocco')
                            row.append(getBow(name))
                            row.append('')
                            row.append(0)
                            cod = generateCOD(row)
                            finalRow = row.copy()
                            finalRow.insert(1, cod)
                            finalRow.insert(13, indexOrder)
                            writer.writerow(finalRow)
                            indexOrder = indexOrder + 1
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
            row.pop()
        indexOrder = 0
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
        print("File created successfully")
        data = ""
        with open(fileName+'.csv', 'r') as file:
            data = file.read().replace(';', ',')

        with open('outputCsv/'+fileName+'.csv', "w") as out_file:
            out_file.write(data)

for filename in glob.glob('wc-product-export-*'):
    os.remove(filename)
