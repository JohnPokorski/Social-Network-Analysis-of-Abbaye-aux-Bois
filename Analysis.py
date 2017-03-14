#Imports
import networkx as nx
from networkx import nodes
from networkx import write_graphml
from networkx import write_gexf
import openpyxl as pyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import sys
import re
import os.path

#Builds a networkx graph from the xcel File
def buildGraph(file):
		network = nx.Graph()
		dataBook = load_workbook(filename=file, read_only=True)
		dataSheet = dataBook.active
		entries = []
		labelLine = True
		for row in dataSheet.rows:
			if(labelLine == True):
				labelLine = False;
			else:
				number = 0
				mainName = ""
				location = ""
				date = ""
				for cell in row:
					if(number == 0):
						number = cell.value
					elif(mainName == ""):
						mainName = cell.value
					elif(location == ""):
						location = cell.value
					elif(date == ""):
						date = cell.value
					else:
						if(cell.value != None):
							rawcell = cell.value
							otherNamesWparnes = rawcell.split(",")
							otherNames = []
							for name in otherNamesWparnes:
								if '(' in name:
									otherNames.append(name[0:name.index('(')])
						else:
							otherNames = []
						entries.append([number,mainName,location,date,otherNames])
		for line in entries:
			if line[1] not in nodes(network):
				network.add_node(line[1])
			for otherName in line[4]:
				if otherName not in nodes(network):
					network.add_node(otherName)
				if not network.has_edge(line[1],otherName):
					network.add_edge(line[1],otherName,weight = 1)
				else:
					network[line[1]][otherName]['weight'] += 1
		return network

#Exports the networkx graph as a gexf file
def export(graph, filename):
	root = tk.Tk()
	root.withdraw()
	if os.path.isfile(filename):
		response = messagebox.askquestion("Duplicate File", "Do you wish to overwrite " + filename + "?")
		if response == "yes":
			write_gexf(graph,filename)
		else:
			i = 1;
			while os.path.isfile(filename.split('.')[0] + " (" + str(i) + ") " + filename.split('.')[1]):
				i = i + 1
			write_gexf(graph,filename.split('.')[0] + " (" + str(i) + ") " + filename.split('.')[1])
	else:
		write_gexf(graph,filename)

#get the file to be analyzed
def getData():
	root = tk.Tk()
	root.withdraw()
	filePath = filedialog.askopenfilename()
	if filePath.split(".")[1] != ".xlsx":
		messagebox.showinfo("Error","Invalid MS Excel 2010 file format")
		return ""
	return filePath

#Gets data, builds graph, then exports
if __name__ == "__main__":
	#get the path to the file
	dataPath = getData()
	if dataPath != "":
		graph = buildGraph(dataPath)
		export(graph,dataPath.split('.')[0]+'result'+'.gexf')