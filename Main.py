import matplotlib.pyplot as plt
import pandas as pd
import os
import numpy as np
from scipy.interpolate import interp1d

if os.path.exists('plots') == False:
    os.mkdir('plots')

#Einstellen der variablen Parameter
#Kapazität in kWh
Kapazität=18
     
#PV-Leistung in kWp
nkWp=28

#Uhrzeiten festlegen, an denen geheizt wird (Heizen, Warmwasser, Heizen, Heizen)
Heizzeiten=[0,5,6.79,15]
#Daten einlesen
PV = pd.read_excel("PV Erzeugungsdaten Freiburg.xlsb", sheet_name=1)
CHR05 = pd.read_excel("CHR05 Family, 3 children, both with work/SumProfiles.HH1.Electricity.xlsx")
    
#Temperatur einlesen
# Öffnen und Lesen der Messdaten
with open('Messdaten.txt', 'r') as file:
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

# Durchschnitt pro Minute pro Tag berechnen
täglicher_durchschnitt = Verbrauch.reshape(366, 1440)

# X-Achse: Minuten im Tag
x = np.arange(1440)

# Y-Achse: Durchschnittlicher Verbrauch pro Minute
y = täglicher_durchschnitt.mean(axis=0)  # Transponieren für die Darstellung
for i in range(1,len(y)) : 
    if abs(y[i]-y[i-1]) > 0.03 :
        y[i] = y[i-1]

#Jahresverbrauch gesamt
Gesamtverbrauch=sum(Verbrauch)

# Hansetherm R32 Premium HT-R32P-12/8-B 55Grad Vorlauftemperatur (letzer Wert interpoliert)
A = [ -20, -15, -10, -7, -5, 0, 2, 7, 10, 15, 20, 25, 30]
COP14 =  [ 1.47, 1.68, 2.03, 2.08, 2.12, 2.13, 2.16, 2.64, 2.75, 3.14, 3.42, 3.48, 3.63]
Pheiz14 = [ 6.15, 6.98, 7.82, 8.20, 8.53, 8.20, 8.40, 10.42, 11.05, 11.63, 12.79, 13.14, 14.03]

# Quadratische Interpolation
COP14interp = interp1d(A, COP14, kind="quadratic")
Pheiz14interp = interp1d(A, Pheiz14, kind="quadratic")

# Neue x-Werte für die Interpolation
AirTemp = np.linspace(-20, 30, 5500)

# Berechne die interpolierten y-Werte
Kennlinie14 = COP14interp(AirTemp)
PKennlinie14 = Pheiz14interp(AirTemp)

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
Gesamtverbrauch+=sum(WPLeistung)/60

#Überschuss
Überschuss=[]
for i in range(len(AusbeuteAC)): 
    Überschuss.append((AusbeuteAC[i]*nkWp)/1000-(Verbrauch[i]*60)-(WPLeistung[i]))

# Datenpunkt liste
data = np.array(WPLeistung).reshape(365, 1440)

#Speicherstratategie
Ladestand=[]
Einspeisung=[]
Netzbezug=[]

#Datenstand zu Jahresbeginn
Ladestand.append(Kapazität/2)
Netzbezug.append(0)
Einspeisung.append(0)

#Definiere Batteriewirkungsgrad
nBatterie=0.8

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
        if (Überschuss[i]/60)+Ladestand[i]*(1/nBatterie) >= 0:
            Ladestand.append(Ladestand[i]+(Überschuss[i]/60)*(1/nBatterie))
            Einspeisung.append(Einspeisung[i])
            Netzbezug.append(Netzbezug[i])
        else:
            Ladestand.append(0)
            Einspeisung.append(Einspeisung[i])
            Netzbezug.append(Netzbezug[i]-((Überschuss[i]/60)+Ladestand[i]*(1/nBatterie)))
        
    elif Überschuss[i]<=0 and Ladestand[i]<=0: # Netzbezug
        Ladestand.append(0)
        Netzbezug.append(Netzbezug[i]-Überschuss[i]/60)
        Einspeisung.append(Einspeisung[i])

Gesamtausbeute = [(AusbeuteAC[0]*nkWp)/1000]
for i in range(1, len(AusbeuteAC)):
    Gesamtausbeute.append(Gesamtausbeute[i-1]+(AusbeuteAC[i]*nkWp)/1000/60)

# Prints

#Wärmebedarf pro Heizgradzahl
print("Wärme pro Heizgradzahl",WärmeProHGT)

#Jahresarbeitszahl
print("Die Jahresarbeitszahl beträgt",(Wärmebedarf+Warmwasserbedarf)/(sum(WPLeistung)/60))

#Gesamtverbrauch
print("Der Gesamtverbrauch ist: ",Gesamtverbrauch,"kWh")

# Jahresbilanz kWh
Jahresbilanz=sum(Überschuss)/60
print("Die Jahresbilanz beträgt bei",nkWp,"kWp:",Jahresbilanz,"kWh")

# Autakiegrad in kWh
Autakiegrad=(Gesamtverbrauch-Netzbezug[-1])/Gesamtverbrauch
print("Der Autarkiegrad beträgt",Autakiegrad*100,"%")

# Eigenverbrauchsanteil
Eigenverbrauchsanteil=(Gesamtausbeute[-1]-Einspeisung[-1])/Gesamtausbeute[-1]
print("Der Eigenverbrauchsanteil beträgt:",Eigenverbrauchsanteil*100,"%")

# Kostenrechnung
# Festlegung der spezifischen Kosten
Stromkosten=0.3*Netzbezug[-1] # €/a
Speicherkosten=750*Kapazität #€
Einspeisevergütung=0.07*Einspeisung[-1] #€/a
PVKosten=2000*nkWp #€
Nutzungszeit=20 #Nutzungszeit in Jahren

# Ausgabe der Kosten
print("Stromkosten pro Jahr",Stromkosten,"€")
print("Anschaffungskosten Speicher",Speicherkosten,"€")
print("Einspeisevergütung pro Jahr",Einspeisevergütung,"€")
print("Anschaffungskosten Photovoltaikanlage",PVKosten,"€")

# Bilanz mit Solar über n Jahre
Annuität = (Speicherkosten+PVKosten)/Nutzungszeit
print("Anschaffungskosten pro Jahr",Annuität,"€")
print("Anschaffungskosten gesamt",Annuität*Nutzungszeit,"€")
Ersparnis = (Gesamtverbrauch*0.3-Stromkosten)*Nutzungszeit
print("Die Ersparnis pro Jahr:",Ersparnis/Nutzungszeit,"€")
print("Die Ersparnis nach der Nutzungszeit ist:",Ersparnis,"€")
Kosten = (Annuität+Stromkosten-Ersparnis/Nutzungszeit-Einspeisevergütung)*Nutzungszeit
print("Die Kosten pro Jahr:",Kosten/Nutzungszeit,"€/a")
print("Die Kosten nach der Nutzungszeit:",Kosten,"€")

# Plots erstellen
# Diagramm Durchschnitt Stromverbrauch
plt.figure(figsize=(10, 6))
plt.plot(x, y, marker='.', linestyle='-', color='b')
plt.xlabel('Minute am Tag')
plt.ylabel('Durchschnittlicher Verbrauch (kWh)')
plt.title('Durchschnittlicher täglicher Stromverbrauch')
plt.grid(True) 
plt.savefig('plots/Durchschnittlicher täglicher Stromverbrauch.png')
#plt.show()

# Diagramm Originaldaten und Interpolation
plt.figure(figsize=(10,6))
plt.scatter(A, COP14, label='Herstellerangaben')
plt.plot(AirTemp, Kennlinie14, label='Interpolation ', color='red')
plt.xlabel('Lufttemperatur Wärmequelle [°C]')
plt.ylabel('COP')
plt.title('Wirkungsgradkennlinie Hansetherm R32 Premium HT-R32P-8-B')
plt.legend()
plt.savefig('plots/Wirkungsgradkennlinie Hansetherm R32 Premium HT-R32P-8-B.png')
#plt.show()

# Diagramm für Gradtagzahl
plt.figure(figsize=(12, 6))
plt.plot(range(len(Gradzahl)), Gradzahl, color='blue', marker='.', linestyle='-')
plt.title('Gradtagzahl')
plt.xlabel('Zeitpunkt')
plt.ylabel('Gradtagzahl')
plt.grid(True)
plt.savefig('plots/Gradtagzahl.png')
#plt.show()

# Diagramm mit pcolormesh
plt.figure(figsize=(10, 6))
plt.pcolormesh(data.T, cmap='viridis')
plt.colorbar(label='Elektrische Leistung [kW]')
plt.xlabel('Tag')
plt.ylabel("Zeitpunkt am Tag [min]") 
plt.title('Darstellung der WP-Leistung')
plt.savefig('plots/Darstellung der WP-Leistung.png')
#plt.show()

# Diagramm für WPLeistung
plt.figure(figsize=(12, 6))
plt.plot(range(len(WPLeistung)), WPLeistung, color='blue', marker='', linestyle='-', linewidth="0.2")
plt.title('WP Leistung')
plt.xlabel('Zeitpunkt')
plt.ylabel('Leistung (kW)')
plt.grid(True)
plt.savefig('plots/WP Leistung.png')
#plt.show()

# Diagramm für Überschuss
plt.figure(figsize=(12, 6))
plt.fill_between(range(len(Überschuss)), Überschuss, where=np.array(Überschuss) >= 0, color='green', alpha=0.5)
plt.fill_between(range(len(Überschuss)), Überschuss, where=np.array(Überschuss) < 0, color='red', alpha=0.5)
plt.title('Energiebilanz')
plt.xlabel('Zeitpunkt')
plt.ylabel('Überschuss (kW)')
plt.grid(True)
plt.savefig('plots/Energiebilanz.png')
#plt.show()

# Diagramm für Ladestand
plt.figure(figsize=(12, 6))
plt.plot(range(len(Ladestand)), Ladestand, color='green', marker='', linestyle='-', linewidth="0.5")
plt.title('Ladestand Akku')
plt.xlabel('Zeitpunkt [min]')
plt.ylabel('Ladesstand (kWh)')
plt.grid(True)
plt.savefig('plots/Ladestand Akku.png')
#plt.show()

# Diagramm für Einspeisung und Netzbezug und Gesamtausbeute
plt.figure(figsize=(12, 6))
plt.plot(range(len(Gesamtausbeute)), Gesamtausbeute, color='green', marker='.', linestyle='-', label="Gesamtausbeute")
plt.plot(range(len(Einspeisung)), Einspeisung, color='grey', marker='.', linestyle='-', label="Einspeisung")
plt.plot(range(len(Netzbezug)), Netzbezug, color='brown', marker='.', linestyle='-', label="Netzbezug")
plt.title('Netzbezug und Einspeisung')
plt.xlabel('Zeitpunkt')
plt.ylabel('Netzbezug und Einspeisung (kWh)')
plt.grid(True)
plt.legend()
plt.savefig('plots/Netzbezug und Einspeisung.png')
#plt.show()

# Diagramm für PV-Ausbeute
plt.figure(figsize=(12, 6))
plt.plot(range(len(Ausbeute)), Ausbeute, color='yellow', marker='', linestyle='-', linewidth="0.5")
plt.title('PV-Ausbeute')
plt.xlabel('Zeitpunkt [min]')
plt.ylabel('Ausbeute (W/kwp)')
plt.grid(True)
plt.savefig('plots/PV-Ausbeute.png')
#plt.show()
