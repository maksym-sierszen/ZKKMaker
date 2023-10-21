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

  def getData(self, wordFile):
    self.document = docx2txt.process(wordFile)
    self.textBlocks = self.document.split('\n\n')
    self.seriesName = os.path.basename(wordFile)
    self.seriesName = re.sub(r'(\.docx|Opis )', '', self.seriesName)
    #print(self.seriesName)
    # print("\n")
    # print(self.textBlocks[0])
    # print("\n")
    # print(self.textBlocks[1])
    # print("\n")
    # print(self.textBlocks[2])
    # print("\n")
    # print(self.textBlocks[3])
    print("\n")


#test i przypisane nazwy serii do zmiennej
#test i przypisane nazwy serii do zmiennej
#word = Word()
#word.getData(r"/Users/maksymsierszen/Desktop/ZKKMaker/Opis Komputronik Infinity R550 [S].docx")

#sprawdź ilość modeli w pliku Word
# 0 to tytul

# 1 i 2 to pierwsze paragrafy
##### !!!!!!!!!!!!!!! sprobuj połączyć w bloki header + paragraf jakoś


#rozdziel całość tekstu Word na bloki każdy model to kolejny blok

  
#class Text():

class fileManager:
  def __init__(self, gui):
        self.gui = gui
        self.desktopPath = Path.home() / 'OneDrive - Grupa Komputronik' / 'Pulpit'
        self.filePath = self.desktopPath / 'test.txt'
  
  def saveContent(self, readyToUse):
    try:
        with open(self.filePath, "w", encoding="utf-8") as file:
            file.write(readyToUse)
    except Exception as e:
        print(f"Error while saving the file: {e}")

  def openExplorer(self):
      try:
          filePath = fd.askopenfilename(initialdir=self.desktopPath, title="Wybierz plik", filetypes=(("docx files", "*.docx"), ("all files", "*.*")))
          if filePath:
              self.gui.fileBox.delete(0, END)
              self.gui.fileBox.insert(0, filePath)
              self.gui.imageNameBox.delete(0, END)
              word = Word()
              word.getData(filePath)
      except Exception as e:
          print(f"Error while opening the file explorer: {e}")

  def generateSEO(self):
          fileName = os.path.abspath(self.gui.fileBox.get())
          word = Word()
          word.getData(fileName) 
          text = Text()
          text.createFile() 
          readyToUse = text.processContent(word.seriesName, word.paragraphs)  
          dirPath = os.path.dirname(fileName) 
          self.filePath = os.path.join(dirPath, f'wpiszERP - {word.seriesName}.txt')
          self.saveContent(readyToUse) 


class GUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("ZKK Maker v1.0")
        ##self.root.iconbitmap("icon.ico")
        self.root.geometry("700x250")
        self.root.configure(bg='white')  
        self.root.resizable(False,False)
        self.manager = fileManager(self)

        self.pathLabel = Label(self.root, text="Ścieżka pliku:", bg='white', font=('Helvetica', 12))  
        self.fileBox = Entry(self.root, width=40, font=('Helvetica', 12))  
        self.fileBox.insert(0, "Wybierz opis")

        self.imageName = Label(self.root, text="Nazwy grafik:", bg='white', font=('Helvetica', 12)) 
        self.imageNameBox = Entry(self.root, width=40, font=('Helvetica', 12))  
        self.imageNameBox.insert(0, " ")

        #usuniete commands tu sa
        self.browseButton = Button(self.root, text="Przeglądaj pliki", command=self.manager.openExplorer, bg='#008CBA', fg='white', font=('Helvetica', 12))  
        self.createButton = Button(self.root, text = "Generuj", command=Word.test, bg='#008CBA', fg='white', font=('Helvetica', 12))
        
        self.copyNameButton = Button(self.root, text="Kopiuj",bg='#008CBA', fg='white', font=('Helvetica', 12))  





        self.pathLabel.grid(row=0, column=0, padx=20, pady=20)
        self.fileBox.grid(row=0, column=1, padx=10, pady=20)
        self.browseButton.grid(row=0, column=2, padx=20, pady=20)

        self.imageName.grid(row=1, column=0, padx=20, pady=20)
        self.imageNameBox.grid(row=1, column=1, padx=10, pady=20)
        self.copyNameButton.grid(row=1, column=2, padx=20, pady=20)

        self.createButton.grid(row=2, column=2, padx=20, pady=20)


    def copyImageName(self):
      imageName = self.imageNameBox.get()
      self.root.clipboard_clear()  
      self.root.clipboard_append(imageName)

    def run(self):
        self.root.mainloop()

# gui = GUI()
# gui.run()