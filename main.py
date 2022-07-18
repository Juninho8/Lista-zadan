# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from openpyxl import load_workbook
from datetime import date

workbook = load_workbook('Zadania.xlsx')
sheet = workbook.active
today = date.today()
def refreshdate(sheet,today):

    refreshing = 1
    i=1
    while refreshing ==1:
        if sheet['A'+str(i)].value != '#':
            until_str=str(sheet['C' + str(i)].value)
            until_str=until_str.split(" ")
            until_date = until_str[0].split("-")
            time = date(int(until_date[0]),int(until_date[1]),int(until_date[2]))- today
            sheet['E' + str(i)].value = time.days

            #max until
            until = str(sheet['D' + str(i)].value)
            until_str = until.split('-')
            until_days = until_str[2].split(" ")
            until_date = date(int(until_str[0]), int(until_str[1]), int(until_days[0]))
            time = abs(until_date - today)
            sheet['F' + str(i)].value = time.days

        elif sheet['A'+str(i)].value=='#':
            refreshing = 0
        i=i+1
    return sheet




def printingduties(sheet):
    tab=['ID','Nazwa','Do kiedy','Do kiedy maksymalnie','Czas','Max Czas','Usuniete','Zrealizwone','Priorytet']
    for row in sheet.values:
        i=0
        if row[0]!='#':
            if row[6]!=1:
                for value in row:
                        if i<=8:
                            if i==0:
                                print("------------------")
                                print(tab[0],":",value)
                            elif i==2 or i==3:
                                value_str=str(value)
                                date = value_str.split(" ")
                                print(tab[i],":",date[0])
                            elif i==6:
                                pass
                            elif i==7:
                                print(tab[i],":",str(value))
                            else:
                                print(tab[i],":",value)
                            i+=1

def printingdutiesprio(sheet):
    tab = ['ID', 'Nazwa', 'Do kiedy', 'Do kiedy maksymalnie', 'Czas', 'Max Czas', 'Usuniete', 'Zrealizwone',
           'Priorytet']
    prio = int(input("Podaj numer priorytetu od 1-10"))
    for row in sheet.values:
        i=0
        if row[8]==prio:
            if row[6]!=1:
                for value in row:
                        if i<=8:
                            if i==0:
                                print("------------------")
                                print(tab[0],":",value)
                            elif i==2 or i==3:
                                value_str=str(value)
                                date = value_str.split(" ")
                                print(tab[i],":",date[0])
                            elif i==6:
                                pass
                            elif i==7:
                                print(tab[i],":",str(value))
                            else:
                                print(tab[i],":",value)
                            i+=1

def addduties(sheet):
    i=1
    saving = 1
    while saving==1:
        if sheet['A' + str(i)].value=='#':
            sheet['A' + str(i)].value = i

            name = input('Podaj nazwe zadania:')

            sheet['B' + str(i)].value = name
            print('Podaj do kiedy zrealizowac:')
            day = int(input("Podaj dzien:"))
            month = int(input("Podaj miesiac:"))
            year = int(input("Podaj rok"))

            sheet['C' + str(i)].value = date(year,month,day)
            print('Podaj do kiedy maksymalnie:')
            day = int(input("Podaj dzien:"))
            month = int(input("Podaj miesiac:"))
            year = int(input("Podaj rok"))

            sheet['D' + str(i)].value = date(year, month, day)
            sheet['H'+str(i)].value = 'NIE'

            prio = int(input('Podaj priorytet zadania 1-10:'))
            sheet['I'+str(i)].value = prio
            saving = 0
        else: i = i+1
    return sheet

def deleteduties(sheet):
    id = int(input("Wpisz numer ID do usuniecia:"))
    for row in sheet.values:
        if row[0]==id:
            sheet['G'+str(id)].value = 1
    return sheet

def realizedduties(sheet):
    id = int(input("Wpisz numer ID do zaznaczenia "))
    sheet['H'+str(id)].value = 'TAK'
    return sheet

working=True
while working:
    print("--------------------")
    print("Wyswietl zadania-1")
    print("Dodaj zadanie-2")
    print("Usun zadanie  zadanie-3")
    print("Odswiez czas do konca-4")
    print("Oznacz jako zrealizowane-5")
    print('Wyswietl po priorytecie-6')
    print("Zakoncz-7")
    choose = input("Co chcesz zrobic?")
    if choose=='1':
        refreshdate(sheet,today)
        printingduties(sheet)

    elif choose=='2':
        addduties(sheet)
        workbook.save('Zadania.xlsx')
    elif choose=='3':
        deleteduties(sheet)
        workbook.save('Zadania.xlsx')
    elif choose=='4':
        refreshdate(sheet,today)
        workbook.save('Zadania.xlsx')
    elif choose=='5':
        realizedduties(sheet)
        workbook.save('Zadania.xlsx')

    elif choose=='6':
        printingdutiesprio(sheet)
    elif choose=='7':
        working=False



