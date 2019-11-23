import matplotlib.pyplot as plt
import networkx as nx
from xlrd import open_workbook, cellname
import xlrd
from datetime import datetime
__author__ = "Thomaz"
__date__ = "$15/11/2019 12:11:52$"
# encode UTF-8

plan = open_workbook(r'C:\Users\estudo\Desktop\UFLA\2019-2\Grafos\Trabalho\instancias_resultados\Escola_A.xlsx')
aba = plan.sheet_by_name('Dados')
abaConfig = plan.sheet_by_name('Configuracoes')
abaRestr = plan.sheet_by_name('Restricao')
abaRestrTurma = plan.sheet_by_name('Restricoes Turma')


dias = {'Segunda':[],'Terça':[],'Quarta':[],'Quinta':[],'Sexta':[]}
cores = []
def Horarios(aba):
    global dias
    for row_index in range(1, aba.nrows):
        dias['Segunda'].append(aba.cell(row_index,0).value)
        dias['Terça'].append(aba.cell(row_index, 0).value)
        dias['Quarta'].append(aba.cell(row_index, 0).value)
        dias['Quinta'].append(aba.cell(row_index, 0).value)
        dias['Sexta'].append(aba.cell(row_index, 0).value)

grafo = nx.Graph()

#print(dias)

qtd = 0
class Vertice():
    def __init__(self):
        self.nome = 0
        self.professor = ''
        self.turma = ''
        self.materia = ''
        self.cor = None
        self.fake = False
        self.nVizinhos = 0
        self.grauSatur = 0

    def setProfessor(self, professor):
        self.professor = professor

    def setNvizinhos(self, nVizinhos):
        self.nVizinhos = nVizinhos

    def getProfessor(self):
        return self.professor

    def getNvizinhos(self):
        return self.nVizinhos

    def getNome(self):
        return self.nome

    def setTurma(self, turma):
        self.turma = turma

    def getTurma(self):
        return self.turma

    def setMateria(self, materia):
        self.materia = materia

    def getMateria(self):
        return self.materia

    def getCor(self):
        return self.cor

def inserirVertice(vert):
   global grafo
   grafo.add_node(vert)

def leituraGrafoPrincipal(aba):
    global qtd
    for row_index in range(1, aba.nrows):
        for qtd_aula in range(int(aba.cell(row_index, 3).value)):
            vert = Vertice()
            vert.setProfessor(aba.cell(row_index, 2).value)
            vert.setTurma(aba.cell(row_index, 1).value)
            vert.setMateria(aba.cell(row_index, 1).value)
            qtd += 1
            vert.nome = qtd
            inserirVertice(vert)
    inserirArestas()

def inserirArestas():
    global grafo
    for u in grafo:
        for v in grafo:
            if u.getProfessor() == v.getProfessor() or u.getTurma() == v.getTurma():
                if not grafo.has_edge(u, v) and u != v:
                    grafo.add_edge(u, v)
                    u.nVizinhos += 1
                    v.nVizinhos += 1

def insertionSort(arr):
    for i in range(1, len(arr)):
        key = arr[i]
        j = i - 1
        while j >= 0 and key.getNvizinhos() > arr[j].getNvizinhos():
            arr[j + 1] = arr[j]
            j -= 1
        arr[j + 1] = key

def procuraCor(dia,hora):
    global dias
    cor = 0
    for d in dias:
        for i in dias[d]:
            cor+=1
            if hora == i and d == dia:
                return cor
    return None

def criaCores():
    global dias, cores
    cor = 0
    for d in dias:
        for i in dias[d]:
            cor+=1
            cores.append(cor)

def lerRestricoes(aba):
    global qtd, dias
    for row_index in range(1, aba.nrows):
        vert = Vertice()
        if aba.cell(0, 0).value == 'Professor:':
            vert.setProfessor(aba.cell(row_index, 0).value)
        elif aba.cell(0, 0).value == 'Turma':
            vert.setTurma(aba.cell(row_index, 0).value)
        qtd += 1
        vert.nome = qtd
        vert.fake = True
        hora = aba.cell(row_index, 1).value
        dia = aba.cell(row_index, 2).value
        cor = procuraCor(dia,hora)
        vert.cor = cor
        inserirVertice(vert)

    inserirArestas()


ordGrau = []
grauSatur = []

leituraGrafoPrincipal(aba)
Horarios(abaConfig)
criaCores()
lerRestricoes(abaRestr)
lerRestricoes(abaRestrTurma)


print('quantidade de vertices: ',grafo.number_of_nodes())
print('quantidade de arestas: ',grafo.number_of_edges())
matrizAdj = nx.to_numpy_matrix(grafo)
listaAdj = nx.to_dict_of_lists(grafo)

for i in listaAdj:
    ordGrau.append(i)
    grauSatur.append(i)
insertionSort(ordGrau)
#Dsatur()
#for i in range(len(ordGrau)):
#   print(ordGrau[i].getNvizinhos())
#nx.draw(grafo)
#plt.show()
