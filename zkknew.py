from tkinter import *
from tkinter import filedialog as fd
from pathlib import Path
import docx2txt
import os
import re

class Word():
  def __init__(self):
    self.document = ""
    self.seriesName = ""
    self.textBlocks = []
    self.variationBlocks = []
    self.pattern = ""
    self.variationAmount = 0
    self.splitter = ""
    self.before = ""
    self.after = ""
    self.variation = 1

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


  def simplifyVariation(self, after, seriesName):
    base = seriesName[:-1]
    toSimplify = re.escape(base) + ".{0,3}"
    replaced_patterns = {}

    def replacement(match):
        
        if replaced_patterns.get(match.group(), False):
            return base + "]"
        replaced_patterns[match.group()] = True
        return match.group()

    return re.sub(toSimplify, replacement, after)



  def splitText(self, before, after, splitter, variation):
    # to bedzie sie wypierdalalo jesli w docx bedzie format bez zera
    if variation > self.variationAmount:
      return


    # wersja dla formatu z 0
    if variation == 1:
        splitter = self.seriesName[:-1] + '02]'  

        parts = self.document.split(splitter)
        before = parts[0] 

        after = splitter + splitter.join(parts[1:]) if len(parts) > 1 else ""
        after = self.simplifyVariation(after, self.seriesName)
        self.variationBlocks.append(before)
    else:
      if variation < 9:
        splitter = self.seriesName[:-1] + '0' + str(variation+1) + ']'
        
      else:
        splitter = self.seriesName[:-1] + str(variation+1) + ']'
      # wersja dla formatu bez 0
      # if variation == 1:
      #     splitter = self.seriesName[:-1] + '2]'  
      #     parts = self.document.split(splitter)
      #     before = parts[0] 
      #     after = splitter + splitter.join(parts[1:]) if len(parts) > 1 else ""
      #     self.variationBlocks.append(before)
      # else:
      #     splitter = self.seriesName[:-1] + str(variation+1) + ']'

        # KONCEPCJA MECHANIZMU USUWAJACEGO TE NUMERKI W TEKSCIE
        # print(self.pattern)
        # matches = list(re.finditer(self.pattern , after))
        # print(matches)
        # for match in matches[:1]:
        #   after = after.replace(match.group(), self.seriesName[:-1])

      parts = after.split(splitter)
      before = parts[0]
      after = splitter + splitter.join(parts[1:]) if len(parts) > 1 else ""
      self.variationBlocks.append(before)  # Dodajemy do listy
        
    # Rekursywne wywołanie z nowymi wartościami
    return self.splitText(before, after, splitter, variation+1)

class Template():

  # 1 analiza co róźni  Ultimate, Pro, Infinity w szablonach -> czyli znalezienie odpowiedzi na pytanie czy mozna doklejac cssa na koncu czy trzeba miec oddzielne szablony

  # pre - 2: wykrywanie serii na podstawie nazwy pliku po to aby dopasować css albo cały szablon

  # 2 kompozycja bloków szablonowych 
  #   2a jeśli trzeba kaźdą serię rozdzielać na swoje własne bloki to template_blocks/nazwa_serii/hdd-blok | grafika-blok itp
  #   2b jeśli moźna ten sam szkielet to jedno do kaźdego i potem nakładać na to cssa

  # 3 ustalenie kolejności bloków dla bazy

  # 4 mechanizm wklejania danych do bloków na podstawie danych z worda (kazdy blok oprocz pierwszego tego banerka ma header i p, to rozdzielenie tego moze jakos na dwie czesci, ale za to jak wariant ma kilka modyfikacji to bedzie wiecej niz 2 czesci, tylko np 4, 6, 8 itd i tam to juz ciezej bedzie jakos porozdzielac ale do przekminki to jak w tych obfitszych tekstowo wariacjach)

  # 5 mechanizm doklejania bloków dodatkowych na bazie danych z wariacji w wordzie (przydatny mechanizm if last = lewo then prawo)
  #   5a przeszukiwanie listy komponentów w poszukiwaniu słów kluczy? for block in templateBlocks: for component in componentBlocks: if blockWord in component -> mechanizm wklejania tekstu do bloku szablonowego


  infinityBase = "test"
  #ultimateBase = 
  #proBase = 


class main():

  word = Word()
  word.extractData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity X510 [I].docx")
  #word.extractData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity R550 [S].docx")
  word.splitText(word.before, word.after, word.splitter, word.variation)


  for i, block in enumerate(word.variationBlocks):
      print(f"\n---WERSJA {i+1}---")
      print(block)


#word.extractData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity R550 [S].docx")

#sprawdź ilość modeli w pliku Word
#test
# print(checkvariationAmount(r"Opis Komputronik Infinity R550 [S].docx"))

main = main()



# 0 to tytul

# 1 i 2 to pierwsze paragrafy
##### !!!!!!!!!!!!!!! sprobuj połączyć w bloki header + paragraf jakoś


#rozdziel całość tekstu Word na bloki każdy model to kolejny blok

#---------
# sprawdź ilość modeli w pliku Word
 
# rozdziel całość tekstu Word na bloki
# każdy model to kolejny blok

# bloki niech będą kolejnymi elementami tablicy (dwuwymiarowe żeby header i paragraf? czy tuple albo coś takiego?)

# każdy blok bazuje na bloku podstawie który nie ma windowsa, montażu i hdd - zmienna foundation

# sprawdź co się różni w bloku i podmień odpowiednie paragrafy foundation

# (mechanizm do opracowania)
# rozdziel szablon foundation na sekcje -> przypisz treść foundation do sekcji -> w bloku każdym kolejnym ODNAJDZ na podstawie info w bloku odpowiednią część w foundation i podmień informacje


# if windows - dodaj szablon windows (tutaj jeszcze rozdzielenie na pro i home)
# if HDD - dodaj szablon HDD
# if both - dodaj szablon both

# sprawdź czy klasy i id są takie same i czy można css dodawać na końcu też po prostu do pliku w zależności od serii 

# +na koniec:   mechanizm


#1. otwórz plik Word
#2. sprawdź która seria
#3. dopasuj odpowiedni szablon do tej serii
#4. extract data z worda do prostszego formatu
  #4.1. sprawdź ile wariacji jest
#5. wstaw data do szablonu po kolei następne iteracje
  #5.1 zapisuj je do pliku
