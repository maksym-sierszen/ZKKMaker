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

  def getData(self, wordFile):
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

    # INNA PRÓBA KONCEPCJI USUWANIA NUMERKOW Z TEKSTU
    # if(self.textBlocks[0][-3] == '0'):
    #   text = print(self.textBlocks[0][:-2])
    #   textNew = print(self.textBlocks[0][:-3] + ']')
    #   self.document = self.document.replace(text, textNew)
      

  def splitText(self, before, after, splitter, variation):
    # to bedzie sie wypierdalalo jesli w docx bedzie format bez zera
    if variation > self.variationAmount:
      return


    # wersja dla formatu z 0
    # if variation == 1:
    #     splitter = self.seriesName[:-1] + '02]'  
    #     parts = self.document.split(splitter)
    #     before = parts[0] 
    #     after = splitter + splitter.join(parts[1:]) if len(parts) > 1 else ""
    #     self.variationBlocks.append(before)
    # else:
    #   if variation < 9:
    #     splitter = self.seriesName[:-1] + '0' + str(variation+1) + ']'
    #   else:
    #     splitter = self.seriesName[:-1] + str(variation+1) + ']'

    # wersja dla formatu bez 0
    if variation == 1:
        splitter = self.seriesName[:-1] + '2]'  
        parts = self.document.split(splitter)
        before = parts[0] 
        after = splitter + splitter.join(parts[1:]) if len(parts) > 1 else ""
        self.variationBlocks.append(before)
    else:
        splitter = self.seriesName[:-1] + str(variation+1) + ']'

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

word = Word()
word.getData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity X510 [I].docx")
word.splitText(word.before, word.after, word.splitter, word.variation)


for i, block in enumerate(word.variationBlocks):
    print(f"\n---WERSJA {i+1}---")
    print(block)


#word.getData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity R550 [S].docx")

#sprawdź ilość modeli w pliku Word
#test
# print(checkvariationAmount(r"Opis Komputronik Infinity R550 [S].docx"))





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