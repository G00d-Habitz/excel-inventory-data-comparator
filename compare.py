import yaml
import requests
import shutil
import os
import pandas


pandas.set_option('display.max_columns', None)
pandas.set_option('display.width', None)
pandas.set_option('display.max_colwidth', None)

def download_data():
    if not os.path.exists("./download"):
        os.mkdir("download")
    if not os.path.exists("dataSource.yml"):
        with open("dataSource.yml", "w", encoding="utf-8") as dataSource:
            print("LOG: Stworzyłem pusty plik dataSource.yml")

    with open("dataSource.yml", "r", encoding="utf-8") as dataSource:
        data_source = yaml.safe_load(dataSource)
        for year in data_source:
            for link in data_source[year]:
                sheet = requests.get(link, stream=True)
                with open(f"./download/{year}.xlsx", "wb") as sheetFile:
                    shutil.copyfileobj(sheet.raw, sheetFile)


def normalize_data(sheet, sheet_name):
    # Obdzieranie komórek ze zbędnych spacji :O
    sheet = sheet.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    if sheet_name == "Intro":
        return sheet
    elif sheet_name == "Elektronarzedzia":
        sheet["Nazwa"] = sheet["Nazwa"].apply(lambda x: x.capitalize())
        sheet["Typ ostrza"] = sheet["Typ ostrza"].apply(lambda x: x.capitalize())
        sheet["Typ silnika"] = sheet["Typ silnika"].apply(lambda x: x.capitalize())
        sheet["Typ zasilania"] = sheet["Typ zasilania"].apply(lambda x: x.capitalize())

        return sheet
    elif sheet_name == "Ostrza":
        # Kapitalikowanie ;)
        sheet["Nazwa"] = sheet["Nazwa"].apply(lambda x: x.capitalize())
        sheet["Typ"] = sheet["Typ"].apply(lambda x: x.capitalize())
        sheet["Material"] = sheet["Material"].apply(lambda x: x.capitalize())
        # Kapitalizowanie każdego słowa w zastosowaniu :D
        sheet["Zastosowanie"] = sheet["Zastosowanie"].apply(lambda x: "\n".join(word.capitalize() for word in x.split("\n")) if "\n" in x else x.capitalize())
        # Sortowanie listy zastosowań :)
        sheet["Zastosowanie"] = sheet["Zastosowanie"].apply(lambda x: "\n".capitalize().join(sorted(x.split("\n"))) if "\n" in x else x)

        return sheet



def compare_sheets(older_sheet, newer_sheet):
    # https://docs.kanaries.net/topics/Pandas/pandas-search-value-column
    # Podzielić to na mniejsze funkcje..

    # Wyciągnięcie roku arkusza
    older_year = older_sheet["Intro"].columns[1]
    newer_year = newer_sheet["Intro"].columns[1]

    #Tworzenie nowych nazw kolumn dla arkuszy
    column_order_elektro = ['ID']
    column_order_ostrza = ['ID']
    elektro_attributes = [column for column in older_sheet["Elektronarzedzia"].columns if column != "ID"]
    ostrza_attributes = [column for column in older_sheet["Ostrza"].columns if column != "ID"]

    older_elektro_columns = {column: f"{column} {older_year}" for column in older_sheet["Elektronarzedzia"].columns if column != "ID"}
    older_elektro_columns["ID"] = "ID"
    newer_elektro_columns = {column: f"{column} {newer_year}" for column in newer_sheet["Elektronarzedzia"].columns if column != "ID"}
    newer_elektro_columns["ID"] = "ID"

    older_ostrza_columns = {column: f"{column} {older_year}" for column in older_sheet["Ostrza"].columns if column != "ID"}
    older_ostrza_columns["ID"] = "ID"
    newer_ostrza_columns = {column: f"{column} {newer_year}" for column in newer_sheet["Ostrza"].columns if column != "ID"}
    newer_ostrza_columns["ID"] = "ID"

    # Zmiany nazw kolumn w arkuszach
    older_sheet["Elektronarzedzia"].rename(columns=older_elektro_columns, inplace=True)
    newer_sheet["Elektronarzedzia"].rename(columns=newer_elektro_columns, inplace=True)

    older_sheet["Ostrza"].rename(columns=older_ostrza_columns, inplace=True)
    newer_sheet["Ostrza"].rename(columns=newer_ostrza_columns, inplace=True)

    # Łączenie ze sobą arkuszy starych i nowych
    merged_elektro = pandas.merge(older_sheet["Elektronarzedzia"], newer_sheet["Elektronarzedzia"], on="ID", how="outer")
    merged_ostrza = pandas.merge(older_sheet["Ostrza"], newer_sheet["Ostrza"], on="ID", how="outer")




    for attr in elektro_attributes:
        column_order_elektro.append(attr + f' {older_year}')
        column_order_elektro.append(attr + f' {newer_year}')


    for attr in ostrza_attributes:
        column_order_ostrza.append(attr + f' {older_year}')
        column_order_ostrza.append(attr + f' {newer_year}')


    # Dodanie kolumny "Status" przed ID
    merged_elektro = merged_elektro[column_order_elektro]
    merged_elektro.insert(0, "Status", " ")

    merged_ostrza = merged_ostrza[column_order_ostrza]
    merged_ostrza.insert(0, "Status", " ")

    # Ustalenie statusu elektronarzędzia
    for index, row in merged_elektro.iterrows():
        for i in range(2, len(merged_elektro.columns) - 1, 2):
            if pandas.isna(row.iloc[i]) and not pandas.isna(row.iloc[i+1]):
                merged_elektro.at[index, 'Status'] = 'Nowy'
                break
            elif pandas.isna(row.iloc[i+1]) and not pandas.isna(row.iloc[i]):
                merged_elektro.at[index, 'Status'] = 'Wycofany'
            elif row.iloc[i] != row.iloc[i + 1]:
                merged_elektro.at[index, 'Status'] = 'Zmieniony'
                break
    # Ustalenie statusu ostrza
    for index, row in merged_ostrza.iterrows():
        for i in range(2, len(merged_ostrza.columns) - 1, 2):
            if pandas.isna(row.iloc[i]) and not pandas.isna(row.iloc[i+1]):
                merged_ostrza.at[index, 'Status'] = 'Nowy'
                break
            elif pandas.isna(row.iloc[i+1]) and not pandas.isna(row.iloc[i]):
                merged_ostrza.at[index, 'Status'] = 'Wycofany'
            elif row.iloc[i] != row.iloc[i + 1]:
                merged_ostrza.at[index, 'Status'] = 'Zmieniony'
                break

    # Tworzenie pythonowych JSONów/słowników ze zmianami
    zmiany_elektro = {column: [] for column in elektro_attributes if column != "ID"}
    zmiany_ostrza = {column: [] for column in ostrza_attributes if column != "ID"}
    zmiany_elektro["Nowe"] = []
    zmiany_ostrza["Nowe"] = []
    zmiany_elektro["Wycofane"] = []
    zmiany_ostrza["Wycofane"] = []

    # Dodawanie zmian elektronarzędzi do słownika
    for index, row in merged_elektro.iterrows():
        if row.iloc[0] == "Zmieniony":
            for i in range(2, len(merged_elektro.columns) - 1, 2):
                if row.iloc[i] != row.iloc[i + 1]:
                    zmiany_elektro[merged_elektro.columns[i][:-5]].append(row.iloc[1])
        elif row.iloc[0] == "Nowy":
            zmiany_elektro["Nowe"].append(row.iloc[1])
        elif row.iloc[0] == "Wycofany":
            zmiany_elektro["Wycofane"].append(row.iloc[1])

    # Dodawanie zmian ostrzy do słownika
    for index, row in merged_ostrza.iterrows():
        if row.iloc[0] == "Zmieniony":
            for i in range(2, len(merged_ostrza.columns) - 1, 2):
                if row.iloc[i] != row.iloc[i + 1]:
                    zmiany_ostrza[merged_ostrza.columns[i][:-5]].append(row.iloc[1])
        elif row.iloc[0] == "Nowy":
            zmiany_ostrza["Nowe"].append(row.iloc[1])
        elif row.iloc[0] == "Wycofany":
            zmiany_ostrza["Wycofane"].append(row.iloc[1])


    # Tworzenie arkusza z listą zmian
    lista_zmian_sheet = []

    if zmiany_ostrza['Nowe']:
        lista_zmian_sheet.append(f"Nowe ostrza w ofercie: {', '.join(zmiany_ostrza['Nowe'])}")

    if zmiany_ostrza['Wycofane']:
        lista_zmian_sheet.append(f"Ostrza wycofane z oferty: {', '.join(zmiany_ostrza['Wycofane'])}")


    for key, values in zmiany_ostrza.items():
        if key not in ['Nowe', 'Wycofane'] and values:
            lista_zmian_sheet.append(f"Zmieniona wartość '{key}' dla ostrzy: {', '.join(values)}")

    if zmiany_elektro['Nowe']:
        lista_zmian_sheet.append(f"Nowe elektronarzędzia w ofercie: {', '.join(zmiany_elektro['Nowe'])}")

    if zmiany_elektro['Wycofane']:
        lista_zmian_sheet.append(f"Elektronarzędzia wycofane z oferty: {', '.join(zmiany_elektro['Wycofane'])}")


    for key, values in zmiany_elektro.items():
        if key not in ['Nowe', 'Wycofane'] and values:
            lista_zmian_sheet.append(f"Zmieniona wartość '{key}' dla elektronarzędzi: {', '.join(values)}")

    lista_zmian_sheet = pandas.DataFrame(lista_zmian_sheet)


    if not os.path.exists("./completed"):
        os.mkdir("completed")

    #Zapisanie arkuszy
    with pandas.ExcelWriter(f'./completed/{older_year}_{newer_year}.xlsx', engine='xlsxwriter') as writer:
        pandas.DataFrame([[f"Raport zmian w latach {older_year} i {newer_year}"]]).to_excel(writer, sheet_name='Intro', index=False)
        lista_zmian_sheet.to_excel(writer, sheet_name='Lista zmian', index=False)
        merged_elektro.to_excel(writer, sheet_name='Elektronarzedzia', index=False)
        merged_ostrza.to_excel(writer, sheet_name='Ostrza', index=False)




def read_sheet(path):
    sheet_names = pandas.ExcelFile(path).sheet_names
    returnable_sheet = {}
    for name in sheet_names:
        if name not in "Intro, Elektronarzedzia, Ostrza":
            return f"Niewłaściwy arkusz at {path}"
        sheet = pandas.read_excel(path, sheet_name=name)
        sheet = normalize_data(sheet, name)
        returnable_sheet[name] = sheet;

    return returnable_sheet



print("INSTRUKCJA: Przed rozpoczęciem upewnij się, że w folderze download znajdują się dwa pliki .xlsx, które chcesz porównać, lub linki do nich znajdują się w pliku dataSource.yml")
print("INSTRUKCJA: AKCEPTOWANE są jedynie pliki składające się z arkuszy Intro, Elektronarzedzia i Ostrza (jako pracownicy wiecie które :))) )")
print("INSTRUKCJA: Plik .yml może zawierać dowolną ilość par: 'nazwa_pliku:' i w nowej linijce dwie spacje, myślnik i link do pobrania pliku .xlsx :)))")
print("INSTRUKCJA: Plik porównujący zostanie zapisany w folderze 'completed', jeśli już istnieje plik o takiej nazwie - zostanie nadpisany!")

download_bool = input("Czy pobrać pliki z linków w folderze dataSource.yml? (Y/N)")

if download_bool == "Y":
    download_data()

sheet_pick = input("Jakie arkusze porównać? Podaj dwie nazwy plików z folderu download wraz z rozszerzeniem, przedzielone przecinkiem, wcześniejszy arkusz pierwszy (np '2023.xlsx,2024.xlsx)'")

newer = read_sheet(f"./download/{sheet_pick.split(',')[1]}")
older = read_sheet(f"./download/{sheet_pick.split(',')[0]}")

compare_sheets(older, newer)
print("Dziękuję za współpracę!")






