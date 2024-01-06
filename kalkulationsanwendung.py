from angebotsrechner import dateien_abwickeln
from grafikinterface import eingabefenster_oeffnen

def main():
    def kalk_funktion(dateiname : list[str]):
        dateien_abwickeln(dateiname)
    eingabefenster_oeffnen(kalk_funktion)

if __name__ == "__main__":
    main()