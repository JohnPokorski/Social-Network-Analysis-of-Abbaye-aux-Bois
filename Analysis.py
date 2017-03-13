import networkx as nx
from networkx import nodes
from networkx import write_graphml
from networkx import write_gexf
import openpyxl as pyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
import sys
import re
#reads the given file and returns it as a graph
#Build network with networkx
#	Nodes are names and places
#	Edges are made and strengthened by shared charters between nodes
def buildGraph(file):
		network = nx.Graph()
		dataBook = load_workbook(filename=file, read_only=True)
		dataSheet = dataBook.active
		entries = []
		#for each line in the file
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
							regex = re.compile(".*?\((.*?)\)")
							rawcell = cell.value
							otherNames = rawcell.split(",")
							#for name in otherNames:
								#name = (re.sub("[\(\[].*?[\)\]]", "", name)).rstrip()
						else:
							otherNames = []
						#print("Entry: ")
						#print("Number: " + str(number) +" Name: " + mainName + " Location: " + location + " Date: " + str(date) + " Other names in charter: " + str(otherNames) + "\n")
						entries.append([number,mainName,location,date,otherNames])
		for line in entries:
			if line[1] not in nodes(network):
				network.add_node(line[1])
			for otherName in line[4]:
				if otherName not in nodes(network):
					network.add_node(otherName)
				network.add_edge(line[1],otherName)

		#check if the first cell has already had a node created
			#if not, create one
		#check that the node has a date and number
			#if not, add them
		#for each name in other names mentioned
			#check if the nodes share an edge
				#if not, make one
				#else, re-enforce it
		return network
#network is exported into a raphml file to be viewed in cytoscape
def export(graph, filename):
	#print(filename)
	write_gexf(graph,filename)

#get the file to be analyzed
def getData():
	root = tk.Tk()
	root.withdraw()
	filePath = filedialog.askopenfilename()
	return filePath

if __name__ == "__main__":
	#get the path to the file
	dataPath = getData()
	graph = buildGraph(dataPath)
	export(graph,dataPath.split('.')[0]+'result'+'.gexf')

	#if the path is bad, error and end
	#else, pass the file to buildGraph() and store the resulting graph
	#export the graph to the desired format
