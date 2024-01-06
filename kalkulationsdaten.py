
#Datencontainer für alle zur Kalkulation gehörenden Daten
class Kalkulationsdaten:
    def __init__(self):
        self.abrechnungskategorien : list[Kategorie] = []

    def kategorie_suchen(self, kategoriename : str):
        for kategorie in self.abrechnungskategorien:
            if kategorie.name == kategoriename:
                return kategorie
        return None
    
    def gesamtsumme_berechnen(self):
        summe = 0
        for kategorie in self.abrechnungskategorien:
            summe += kategorie.untersumme_berechnen()
        return summe
    
    def gefiltert_ausgeben(self):
        ausgabe : list[list[str]] = []
        for kategorie in self.abrechnungskategorien:
            if kategorie.ist_nicht_leer():
                ausgabe.append([kategorie.name, None, None, kategorie.untersumme_berechnen()])
                for einheit in kategorie.nichtleere_ausfiltern():
                    ausgabe.append([einheit.name, einheit.menge, str(einheit.preis) + "/" + str(einheit.nenner), einheit.summe_berechnen()])
        ausgabe.append(["Gesamt", None, None, self.gesamtsumme_berechnen()])
        return ausgabe


#Datencontainer fuer alle Daten einer Unterkategorie der Kalkulation
class Kategorie:
    def __init__(self, name : str):
        self.name = name
        self.rechnungseinheiten : list[Rechnungseinheit] = []

    #Prueft ob in der Kategorie mindestens eine nicht leere Rechnungseinheit exitiert
    def ist_nicht_leer(self):
        if self.rechnungseinheiten is None:
            return False
        for einheit in self.rechnungseinheiten:
            if einheit.ist_nicht_leer():
                return True
        return False
    
    def nichtleere_ausfiltern(self):
        nichtleere = []
        for einheit in self.rechnungseinheiten:
            if einheit.ist_nicht_leer():
                nichtleere.append(einheit)
        return nichtleere
    
    def untersumme_berechnen(self):
        untersumme = 0
        for einheit in self.rechnungseinheiten:
            untersumme += einheit.summe_berechnen()
        return untersumme

#Datencontainer für die Daten eines einzelnen Punkts der Kalkulationstabelle
class Rechnungseinheit:
    def __init__(self, name : str, preis, menge, nenner):
        self.name = name
        self.preis = preis
        self.menge = menge
        self.nenner = nenner

    #Prueft ob die Einheit leer ist (Menge == 0 oder kein Wert)
    def ist_nicht_leer(self):
        return self.menge != 0 and self.menge is not None and self.menge != ""
    
    def summe_berechnen(self):
        return self.preis * self.menge