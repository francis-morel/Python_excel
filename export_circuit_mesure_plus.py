import openpyxl
from os import listdir

class CircuitCount: #définition d'une classe pour mettre plusieurs informations dans un dictionnaire selon le nom du circuit
    circuit = ""
    mesure = 0
    nombre = 0

    def __init__(self, circuit, mesure):
        self.circuit = circuit
        self.mesure = mesure

    def printValues(self):
        print(self.circuit, self.mesure, "mm x", self.nombre)

def valid_sheet_name(name):
    return len(name) == 1

def get_column(sheet, col, start_row = 1, nb_lines = 0):
    array = []
    row_number = start_row

    while True:
        if nb_lines > 0 and row_number-start_row > nb_lines:
            break
        cell = get_cell(sheet, col, row_number)
        if cell == None:
            break
        array.append(cell)

    return array

def get_cell(sheet, col, row):
    return sheet.cell(row = row, column = col).value

def write_sheet_2d_list(sheet, data, start_col = 1, start_row = 1):
    pass

def write_sheet_list(sheet, data, start_col =1, start_row = 1):
    for i in range(len(data)):
        write_sheet(sheet, data[i], start_col+i, start_row+i)

def write_sheet(sheet, data, col, row):
    sheet.cell(row = row, column = col).value = data


newWb = openpyxl.Workbook()
newWs = newWb.active
writeCol = 1
returnColCircuit = []
returnColMesure = []
fileRow = 1

totalCables = {}

for file in listdir(): #Itérer tous les fichiers dans le dossier
    if file[-5:] == ".xlsm" or file[-5:] == ".xlsx": #Si le fichier est un fichier excel
        write_sheet(newWs, file, fileRow, 1)#Écrire le nom du fichier dans le nouveau fichier
        fileRow += 1
        wb = openpyxl.load_workbook(filename = file, data_only = True) #Ouvrir le fichier excel

        for sheet in wb.worksheets: #Itérer toutes les feuilles du fichier ouvert
                        
            if not valid_sheet_name(sheet.title): #Vérification du nom de la feuille et la skipper si elle ne m'intéresse pas
                continue

            ColCircuit = get_column(sheet, 4)
            ColMesure = get_column(sheet, 5)

            circuitRow = fileRow
            
            write_sheet_list(sheet, ColCircuit, start_col = writeCol, start_row = fileRow)
            writeCol += 1
            write_sheet_list(sheet, ColCircuit, start_col = writeCol+1, start_row = fileRow)
            writeCol += 2

            fileRow += len(ColCircuit) + 1



        #Fin du fichier, préparation des variables pour le prochain.
        fileRow += 2
        writeCol = 1


newWb.save("total modeles python plus.xlsx")