from openpyxl import Workbook
from argparse import ArgumentParser
from os.path import abspath
from kalkulationsdaten import Kalkulationsdaten
from util import *
formatvorlage = "resourcen/Kalkulationsvorlage.xlsx"
daten_auf_konsole_ausgeben = False

#Anfangspunkt bei direkter Ausführung des Moduls
def main():
    parser = ArgumentParser()
    #Dateinamen einlesen
    parser.add_argument('dateien', nargs='+', help='Liste von Dateinamen.')
    args = parser.parse_args()
    #Kalkulation für jede Input-Datei erstellen
    dateien_abwickeln(args.dateien)
    
def dateien_abwickeln(dateinamen : list[str]):
    for dateiname in dateinamen:
        datei_abwickeln(dateiname)

def datei_abwickeln(dateiname : str):
    dateiname = abspath(dateiname)
    if workbook_pfad_verifizieren(dateiname):
        daten = lese_kalkulationsdaten(dateiname)
        if (daten_auf_konsole_ausgeben):
            daten_auf_konsole_drucken(daten)
        schreibe_formatierte_kalkulationsdaten(daten, dateiname)
    else:
        print(f"Konnte keine Datei unter dem Namen/Pfad {dateiname} finden")


#Steuert die Kalkulationserstellung für eine einzelne Inputdatei
def lese_kalkulationsdaten(dateiname : str):
    datei : Workbook = workbook_oeffnen(dateiname)
    daten : Kalkulationsdaten = daten_einlesen(datei)
    return daten

def vervielfaeltige_formatvorlage(input_pfad : str):
    pfad_formatvorlage = pfad_zu_ressource_finden(formatvorlage)
    pfad_fuer_kalkulationsdatei = neuen_dateipfad_ermitteln(input_pfad)
    kalkulationsdatei = workbook_kopieren(pfad_formatvorlage, pfad_fuer_kalkulationsdatei)
    if kalkulationsdatei is None:
        print("Fehler beim vervielfältigen der Formatvorlage, abbruch.")
        return None
    else:
        return pfad_fuer_kalkulationsdatei

def schreibe_formatierte_kalkulationsdaten(daten : Kalkulationsdaten, input_pfad : str):
    xlsx = vervielfaeltige_formatvorlage(input_pfad)
    if xlsx is None:
        return
    wb = workbook_oeffnen(xlsx)
    ws = wb["Kalkulation"]
    startpunkt = formatierungsstartpunkt_finden(ws)
    if startpunkt is not None:
        formatierte_daten = daten.gefiltert_ausgeben()
        formatierte_daten_einfuegen(ws, formatierte_daten, startpunkt)
        wb.save(xlsx)
        wb.close()

#Druckt einen Datensatz Kalkulationsdaten als Baumdiagramm auf die Konsole
def daten_auf_konsole_drucken(daten : Kalkulationsdaten):
    print("Daten erhalten: ", daten)
    for kategorie in daten.abrechnungskategorien:
        print("├─Kategorie:", kategorie.name)
        for einheit in kategorie.rechnungseinheiten:
            print("│  ├─Rechnungseinheit: ", einheit.name, " \n│  │  ├─Preis: ", einheit.preis, " \n│  │  ├─Nenner: ", einheit.nenner)

if __name__ == "__main__":
    daten_auf_konsole_ausgeben = True
    main()