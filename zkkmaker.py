from tkinter import *
from tkinter import filedialog as fd
from pathlib import Path
import docx2txt
import os
import re


#struktura danych - po kolei kazdy elemnt typu i ma przypisane cechy swoje jakieś, moze jako struct jakis? [{gpu}, {cpu}, {ram} {...}]
#dodawanie potem kolejno _HEADER i _PARAGPRAH (?) loopując
#struktura pliku główna:
    # słownik
    # text = {'MAIN_HEAEDER': 'text}, 'MAIN_PARAGRAPH1': 'text', [...]}
    



### TWORZENIE BAZY

# checkbox czy tylko baza czy iteracje takze
        # jesli kox to wpisac ile iteracji (jesli chcemy ograniczyc bo np zaskoczyly zmiany jakies w pliku docx i sie zmienila struktura zmian)


#1. otwórz plik Word
    #1.1 Przypisz nazwę serii do wartości
    #1.2 Zabawa w poprawianie nazwy jeśli niezbędne (ten case z 0 albo brak 0)
    #1.3 Posprzątaj indeksy kolejnych modeli z pliku (w sensie [S01] --> [S], Z WYJATKIEM HEADER 1 !!!!!!)

#2. split Word into textblocks

#3. elemnty rozdzielone splitem przypisz jako values do kazdego kolejnego itemu w dictionary 
    # warunek ze jesli item = zawiera series name to nowa seria(?) (!!! BEZ SENSU BO PARAGRAFY MAJA TO CZESTO PRZECIEZ)
    # albo warunek w loopie w pkt 5, ze jesli text = "END_SERIES", to odcinka i leci nowy pliczek jakos rekurencyjne wtedy do nowej kopii pliku?


#4. sprawdź która seria i szablon odpowiedni wybierz


#5. loop przez szablon gdzie szukasz kolejno czy tekst sie zgadza z kolejnym itemem w dict, jesli tak to podmien ten tekst na jego value
     # w teorii kazde powinno byc przypisane po kolei idealnie w szablonie, wiec nie wiem czy konieczny bedzie if

    #  warunek w loopie w pkt 5, ze jesli text = "END_SERIES", to odcinka i leci nowy pliczek jakos rekurencyjne wtedy do nowej kopii pliku?

#POTEM TO ZAIMPLEMENTWOAC - NAJPIERW BAZA
### TWORZENIE ITERACJI KOLEJNYCH MODELI

    # tworzysz nowy plik totalnie, kopiujesz tam tam zrobiona baze
    # wg schematu teraz jedziesz
        # if 2 --> doklejasz windows (i montaz jesli trzeba ew.)
        # if 3 --> doklejasz hdd
        # if 4 --> doklejasz hdd i windows
        # if = 5 lub większe bierzesz plik n-4 jako baze i zmieniasz co trzeba

class WordText():
  def __init__(self):
    self.document = ""
    self.seriesName = ""
    
    self.textBlocks = []
    self.variationBlocks = []
    
    self.pattern = ""
    self.splitter = ""
    
    self.before = ""
    self.after = ""
    
    self.variationAmount = 0
    self.variation = 1

#



  
def extractData(self, wordFile):
    self.document = docx2txt.process(wordFile)
    self.textBlocks = self.document.split('\n\n')
    self.seriesName = os.path.basename(wordFile)
    self.seriesName = re.sub(r'(\.docx|Opis )', '', self.seriesName)
    
    self.pattern = re.escape(self.seriesName[:-1]) + r'(\d+)'
    allFinds = re.findall(self.pattern, self.document)
 
    
    if allFinds:
        self.variationAmount = int(allFinds[-1])
        return self.variationAmount
    else:
        return None
#2. sprawdź która seria
#3. dopasuj odpowiedni szablon do tej serii
#4. extract data z worda do prostszego formatu
  #4.1. sprawdź ile wariacji jest
#5. wstaw data do szablonu po kolei następne iteracje
  #5.1 zapisuj je do pliku


#nbsp dodawanie?


