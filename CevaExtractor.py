#CEVAEXTRACTOR REV 2.0 
#Estrae i seriali degli articoli INFRATEL dalle packing list CEVA per outputtare un file .xls e poter fare un caricamento massivo.
#PROGRAMMA SCRITTO "AS IS" senza nessuna garanzia di funzionamento, in un momento di divertimento.
#
import re #importazione modulo per le regexp
import xlwt  #modulo che mi permette di gestire e manipolare fogli di calcolo excel 
from xlwt import Workbook
from playsound import playsound
#########################FUNZIONI#######################################################################################################
def LeggiFile(): #funzione che mi legge il file sorgente
    test = 0 #variabile booleana che prova a leggere il file finchè l'utente non inserisce il nome del file corretto 
    while not test:
        try:
            nomefilesorgente = input("Inserisci il nome del file sorgente (senza .txt): ")
            sorgente =open(nomefilesorgente+".txt","r") #apre il file della bolla in lettura 
            test = 1
        except:
            test =0
            print("File non trovato!\n")
    return sorgente

#il seriale dei prodotti INFRATEL è composto nel seguente modo:
#3 caratteri dalla a alla z maiuscoli ; 7 cifre da 0 a 9 e 10 cifre che son lettere o cifre (è presente il codiarti) e 6 cifre da 0 e 9 per un tot di 26 caratteri
def TrovaSeriale(riga):   #funzione che, data la riga in input, mi estrae il seriale 
     regexpseriali = "[A-Z]{3}[0-9]{7}[A-Z0-9]{10}[0-9]{6}" 
     matcher = re.search(regexpseriali,str(riga))
     if matcher is not None:
         seriale = matcher.group()
         return seriale


def CodiceArticoloDaSeriale(seriale): #mi estrae il codice articolo dal seriale ("è sottinteso nel seriale dal carattere 10 al 19")
    codicearticolo = []
    for i in range(10,20):
        codicearticolo.append(seriale[i])
    return (''.join(codicearticolo))

def PrintaDisclaimer():
   print("ATTENZIONE! Il tool è stato scritto in base alla struttura dei seriali che avevo in possesso del materiale C&D,percui è possibile che i nuovi articoli che verranno codificati possono non essere riconosciuti se il seriale ha struttura diversa!\n")
   print("TESTATO E FUNZIONANTE SUI SEGUENTI PRODOTTI: ROE E ROI DA 12 / 24 / 48 AMARRI DA DIM. 24 ALLA 396\n")
   print("POZZETTO-SOPRALZI H10 E 20 - ANELLI PORTACHIUSINO 125*80 POZZETTO 40*76 E SOPRALZO H20 40*76\n")
   print("MUFFOLA GTL DA 400 FO E 144 FO CASSETTO G/T 48 FO,ARMADI CNO E DISP SIST. SCORTA FO, TUTTI I TIPI DI SOSPENSIONI ATTACCO PER SOSTEGNI A TRALICCIO\n")
   print(" CABINET FWA OUTDOOR Se il seriale ha 26 lettere \n")


####################################################
#programma principale

print("**** **** *       * ***** ****  *    * ******  ***** *****  ***** ***** ***** *****    ")
print("*    *     *     *  *   * *      *  *    *     *   * *   *  *       *   *   * *   *   ")
print("*    ****   *   *   ***** ****    *      *     ***** *****  *       *   *   * *****   ")
print("*    *       * *    *   * *      * *     *     **    *   *  *       *   *   * * *    ") 
print("**** ****     *     *   * ****  *   *    *     * *   *   *  *****   *   ***** *  * \n ")

print("\nBy Arnaldo Lucchini SITTEL - GR")

PrintaDisclaimer()
sorgente = LeggiFile()


nomefoglio = input("Inserisci il nome del file che vuoi salvare (senza .xls)")
wb = Workbook()
sheet1 = wb.add_sheet('Foglio 1')
sheet1.write(0,0, 'codiarti')
sheet1.write(0,1, 'serialno')    
contapezzi = 0

for line in sorgente:
    stringapulita = TrovaSeriale(line)
    if stringapulita is not None:
        contapezzi = contapezzi + 1 
        #print(CodiceArticoloDaSeriale(stringapulita)+ "  "+stringapulita)
        sheet1.write(contapezzi,0,CodiceArticoloDaSeriale(stringapulita)) #scrive nella prima colonna il codicearticolo
        sheet1.write(contapezzi,1,stringapulita) #scrive nella seconda colonna il seriale trovato.

wb.save(''+nomefoglio+' N.PEZZI '+str(contapezzi)+'.xls')
print("Il totale dei pezzi è: "+str(contapezzi))    
sorgente.close()

