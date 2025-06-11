import xlsxwriter
import pandas as pd
import math
from scipy.interpolate import interp1d
from multiprocessing import Pool, Manager
import multiprocessing
from datetime import datetime
import numpy as np

def Prep():
    #Uhrzeiten festlegen, an denen geheizt wird (Heizen, Warmwasser, Heizen, Heizen)
    #Heizzeiten=[0,6,6.79,16]
    Heizzeiten=[0,5,6.79,15]
    #Daten einlesen
    PV = pd.read_excel("data/PV Erzeugungsdaten Freiburg.xlsb", sheet_name=1)
    CHR05 = pd.read_excel("data/SumProfiles.HH1.Electricity.xlsx")
        
    #Temperatur einlesen
    # Öffnen und Lesen der Messdaten
    with open('data/Messdaten.txt', 'r') as file:
        lines = file.readlines()[1:]

    # Array für die Temperatur
    Temperatur = []

    # Durch jede Zeile des Textdokuments iterieren und die fünfte Spalte extrahieren
    for line in lines:
        data = line.split(';') 
        fuenfte_spalte_wert = data[4].strip()
        fuenfte_spalte_wert = float(fuenfte_spalte_wert.replace(',', '.').replace("-999","12"))
        Temperatur.append(fuenfte_spalte_wert)

    # Ausbeute in W/kWp
    Ausbeute = PV.iloc[1:round(525603), PV.columns.get_loc("Unnamed: 10")].values
    Ausbeute = Ausbeute.astype(float)

    #Ausbeute mit DC-AC-Wandler verlusten
    DCAC=0.987
    AusbeuteAC=[]
    for i in Ausbeute:
        AusbeuteAC.append(i*DCAC)
        
    # Verbrauch in kWh
    Verbrauch = CHR05.iloc[0:527040,CHR05.columns.get_loc("Sum [kWh]")].values    
    for i in range(1,len(Verbrauch)) : 
        if abs(Verbrauch[i]) > 1 : 
            Verbrauch[i] = Verbrauch[i-1]

    #Jahresverbrauch gesamt
    Gesamtverbrauch=sum(Verbrauch)

    # Hansetherm R32 Premium HT-R32P-12/8-B 55Grad Vorlauftemperatur (letzer Wert interpoliert)
    A = [ -20, -15, -10, -7, -5, 0, 2, 7, 10, 15, 20, 25, 30]
    COP14 =  [ 1.47, 1.68, 2.03, 2.08, 2.12, 2.13, 2.16, 2.64, 2.75, 3.14, 3.42, 3.48, 3.63]
    Pheiz14 = [ 6.15, 6.98, 7.82, 8.20, 8.53, 8.20, 8.40, 10.42, 11.05, 11.63, 12.79, 13.14, 14.03]

    # Quadratische Interpolation
    COP14interp = interp1d(A, COP14, kind="quadratic")
    Pheiz14interp = interp1d(A, Pheiz14, kind="quadratic")

    #Gradtagzahlmethode
    Tagesmittel=[]
    Gradzahl=[]
    Raumtemperatur=20
    Heizgrenze=15

    #Tagesdurchschnittstemperatur und Grattagzahl
    for i in range(365):
        Tagesmittel.append(sum(Temperatur[(i*144):((i+1)*144)])/144)
        if Tagesmittel[i]<Heizgrenze:
            Gradzahl.append(Raumtemperatur-Tagesmittel[i])  
        else:
            Gradzahl.append(0)

    #Jahreswärmebedarf in kWh
    Wärmebedarf=19482

    #Wärmebedarf pro Heizgradtag
    WärmeProHGT=Wärmebedarf/sum(Gradzahl)

    #Jahreswarmwasserbedarf in kWh
    Warmwasserbedarf=1836

    #Tageswärmebedarf
    Tageswärmebedarf=[]
    WPLeistung=[]

    for i in range(len(Tagesmittel)):
        Tageswärmebedarf.append(Gradzahl[i]*WärmeProHGT)

    Tageswärmeerzeugung=[]
    Tageswarmwassererzeugung=[]

    #Steuerung der Wärmepumpe
    for i in range(len(Tageswärmebedarf)):
        Tageswärmeerzeugung.append(0)
        Tageswarmwassererzeugung.append(0)
        
        for u in range(i*1440+Heizzeiten[0]*60,i*1440+Heizzeiten[1]*60): #Erste Heizphase
            if Tageswärmeerzeugung[i]<Tageswärmebedarf[i]/3:
                Tageswärmeerzeugung[i]=Tageswärmeerzeugung[i]+Pheiz14interp(Temperatur[round(u/10)])/60
                WPLeistung.append(Pheiz14interp(Temperatur[round(u/10)])/COP14interp(Temperatur[round(u/10)]))
            elif Tageswärmeerzeugung[i]>=Tageswärmebedarf[i]/3:
                Tageswärmeerzeugung[i]=Tageswärmebedarf[i]/3
                WPLeistung.append(0)
                
        for u in range(i*1440+Heizzeiten[1]*60, i*1440+round(Heizzeiten[2]*60)): #Warmwasserphase
            if Tageswarmwassererzeugung[i]<Warmwasserbedarf/365:
                Tageswarmwassererzeugung[i]=Tageswarmwassererzeugung[i]+Pheiz14interp(Temperatur[round(u/10)])/60
                WPLeistung.append(Pheiz14interp(Temperatur[round(u/10)])/COP14interp(Temperatur[round(u/10)]))
            elif Tageswarmwassererzeugung[i]>=Warmwasserbedarf/365:
                Tageswarmwassererzeugung[i]=Warmwasserbedarf/365
                WPLeistung.append(0)
                
        for u in range(i*1440+round(Heizzeiten[2]*60),i*1440+Heizzeiten[3]*60): #Zweite Heizphase
            if Tageswärmeerzeugung[i]<Tageswärmebedarf[i]*(2/3):
                Tageswärmeerzeugung[i]=Tageswärmeerzeugung[i]+Pheiz14interp(Temperatur[round(u/10)])/60
                WPLeistung.append(Pheiz14interp(Temperatur[round(u/10)])/COP14interp(Temperatur[round(u/10)]))
            elif Tageswärmeerzeugung[i]>=Tageswärmebedarf[i]*(2/3):
                Tageswärmeerzeugung[i]=Tageswärmebedarf[i]*(2/3)
                WPLeistung.append(0)
                    
        for u in range(i*1440+Heizzeiten[3]*60,(i+1)*1440+Heizzeiten[0]*60): #Dritte Heizphase
            if Tageswärmeerzeugung[i]<Tageswärmebedarf[i]:
                Tageswärmeerzeugung[i]=Tageswärmeerzeugung[i]+Pheiz14interp(Temperatur[round(u/10)])/60
                WPLeistung.append(Pheiz14interp(Temperatur[round(u/10)])/COP14interp(Temperatur[round(u/10)]))
            elif Tageswärmeerzeugung[i]>=Tageswärmebedarf[i]:
                Tageswärmeerzeugung[i]=Tageswärmebedarf[i]
                WPLeistung.append(0)

    #Jahresverbrauch gesamt erweitern
    Gesamtverbrauch += sum(WPLeistung)/60
    return AusbeuteAC, Verbrauch, WPLeistung, Gesamtverbrauch

def Sim(nkWp, Kapazität, AusbeuteAC, Verbrauch, WPLeistung, Gesamtverbrauch,_process_done_count, _process_count, lock):
    #Überschuss
    Überschuss=[]
    for i in range(len(AusbeuteAC)): 
        Überschuss.append((AusbeuteAC[i]*nkWp)/1000-(Verbrauch[i])-(WPLeistung[i]))

    #Speicherstratategie
    Ladestand=[]
    Einspeisung=[]
    Netzbezug=[]

    nBatterie = 0.8

    #Datenstand zu Jahresbeginn
    Ladestand.append(Kapazität/2)
    Netzbezug.append(0)
    Einspeisung.append(0)

    for i in range(len(Überschuss)):
        if Überschuss[i]>0 and Ladestand[i]<Kapazität: # Akku laden
            if Ladestand[i]+(Überschuss[i]/60)*nBatterie <= Kapazität:
                Ladestand.append(Ladestand[i]+(Überschuss[i]/60)*nBatterie)
                Einspeisung.append(Einspeisung[i])
                Netzbezug.append(Netzbezug[i])
            else:
                Ladestand.append(Kapazität)
                Einspeisung.append(Einspeisung[i]+(Überschuss[i]/60)-(Kapazität-Ladestand[i])*nBatterie)
                Netzbezug.append(Netzbezug[i])
            
        elif Überschuss[i]>0 and Ladestand[i]>=Kapazität: # Einspeisen
            Ladestand.append(Kapazität)
            Einspeisung.append(Einspeisung[i]+(Überschuss[i]/60))
            Netzbezug.append(Netzbezug[i])
            
            
        elif Überschuss[i]<=0 and Ladestand[i]>0: # Akku entladen
            if Ladestand[i]+(Überschuss[i]/60)*(1/nBatterie) >= 0:
                Ladestand.append(Ladestand[i]+(Überschuss[i]/60)*(1/nBatterie))
                Einspeisung.append(Einspeisung[i])
                Netzbezug.append(Netzbezug[i])
            else:
                Ladestand.append(0)
                Einspeisung.append(Einspeisung[i])
                Netzbezug.append(Netzbezug[i]+((Überschuss[i]/60)-Ladestand[i]*(1/nBatterie)))
            
        elif Überschuss[i]<=0 and Ladestand[i]<=0: # Netzbezug
            Ladestand.append(0)
            Netzbezug.append(Netzbezug[i]-Überschuss[i]/60)
            Einspeisung.append(Einspeisung[i])

    # Die letzten Werte verwenden
    Stromkosten=0.3*Netzbezug[-1]
    Speicherkosten=750*Kapazität
    Einspeisevergütung=0.07*Einspeisung[-1]
    PVKosten=2000*nkWp
    Ersparnis = Gesamtverbrauch*0.3-Stromkosten
    Nutzungszeit=20
    Annuität = (Speicherkosten+PVKosten)/Nutzungszeit
    Kosten = Annuität+Stromkosten-Ersparnis-Einspeisevergütung
    Autakiegrad=(Gesamtverbrauch-Netzbezug[-1])/Gesamtverbrauch

    with lock:
        _process_done_count.value+=1
        print("["+datetime.strftime(datetime.now(), "%H:%M:%S")+"]", "{:.0f}".format((_process_done_count.value/_process_count)*100), "%" , end="\r")

    return nkWp, Kapazität, Kosten, Autakiegrad*100

if __name__ == "__main__":

    # Start der Simulation
    print("["+datetime.strftime(datetime.now(), "%H:%M:%S")+"]","Simulation gestartet")
    simStartTime = datetime.now()

    AusbeuteAC, Verbrauch, WPLeistung, Gesamtverbrauch = Prep()

    #Iteratives Simulieren der Werte
    workbook = xlsxwriter.Workbook("Simulation.xlsx")
    _kosten_worksheet = workbook.add_worksheet("Kosten")
    _autakie_worksheet = workbook.add_worksheet("Autakiegrad")

    #Minimal/Maximal kWh und kWp
    Minimal = 1 
    Maximal = 200
    Step = 4

    _kosten_BestPrice = math.inf
    _kosten_BestnkWp = 0
    _kosten_BestKapazität = 0

    _autakie_BestAutakiegrad = 0
    _autakie_BestnkWp = 0
    _autakie_BestKapazität = 0

    _nkWp = list(range(Minimal, Maximal + 1, Step))
    _kapazität = list(range(Minimal, Maximal + 1, Step))

    _process_count = len(_nkWp)*len(_kapazität)

    with Manager() as manager:
        _process_done_count = manager.Value("i", 0)
        lock = manager.Lock()
        with Pool(multiprocessing.cpu_count()) as p:
            results = p.starmap(Sim, [(nkWp, Kapazität, AusbeuteAC, Verbrauch, WPLeistung, Gesamtverbrauch,_process_done_count, _process_count, lock) for nkWp in _nkWp for Kapazität in _kapazität])

    for result in results:
        nkWp, Kapazität, Kosten, Autakiegrad = result

        row = _nkWp.index(nkWp) + 4
        col = _kapazität.index(Kapazität) + 1

        _kosten_worksheet.write(row, 0, str(nkWp)+" kWp")
        _kosten_worksheet.write(3, col, str(Kapazität)+" kWh")
        _kosten_worksheet.write(row, col, Kosten)
        
        _autakie_worksheet.write(row, 0, str(nkWp)+" kWp")
        _autakie_worksheet.write(3, col, str(Kapazität)+" kWh")
        _autakie_worksheet.write(row, col, int(Autakiegrad))

        if Kosten < _kosten_BestPrice:
            _kosten_BestPrice = Kosten
            _kosten_BestnkWp = nkWp
            _kosten_BestKapazität = Kapazität
        if int(Autakiegrad) > _autakie_BestAutakiegrad:
            _autakie_BestAutakiegrad = int(Autakiegrad)
            _autakie_BestnkWp = nkWp
            _autakie_BestKapazität = Kapazität

    _kosten_worksheet.write(1, 0, "Lowest:")
    _kosten_worksheet.write(1, 1, _kosten_BestKapazität)
    _kosten_worksheet.write(1, 2, "kWh")
    _kosten_worksheet.write(1, 3, _kosten_BestnkWp)
    _kosten_worksheet.write(1, 4, "kWp")
    _kosten_worksheet.write(1, 5, _kosten_BestPrice)
    _kosten_worksheet.write(1, 6, "€/a")

    _autakie_worksheet.write(1, 0, "Highest:")
    _autakie_worksheet.write(1, 1, _autakie_BestKapazität)
    _autakie_worksheet.write(1, 2, "kWh")
    _autakie_worksheet.write(1, 3, _autakie_BestnkWp)
    _autakie_worksheet.write(1, 4, "kWp")
    _autakie_worksheet.write(1, 5, _autakie_BestAutakiegrad)
    _autakie_worksheet.write(1, 6, "%")

    _kosten_worksheet.conditional_format("B5:ZZ99",{"type": "3_color_scale", "min_color": "#88ff88", "mid_color": "#ffff88", "max_color": "#ff8888"})
    _autakie_worksheet.conditional_format("B5:ZZ99",{"type": "3_color_scale", "min_color": "#ff8888", "mid_color": "#ffff88", "max_color": "#88ff88"})
    workbook.close()
    print("["+datetime.strftime(datetime.now(), "%H:%M:%S")+"]","Simulation beendet nach", datetime.now()-simStartTime)