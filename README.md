Social-Network-Analysis-of-Abbaye-aux-Bois is a project implimented in python 3.6 to aid in the analysis 
of a corpus of charters collected from the Abbaye aux Bois in an attempt to visualize and analyse the 
social networks surrounding the Abbey.

As a specialized project, this program is not intended for widspread use, but may be used and analyzed freely.

A windows installer is provided, created with cx_freeze with the included setup.py script

Uses openpyxl (https://openpyxl.readthedocs.io/en/default/) for file parseing, and networkx (https://networkx.github.io/) for construction
and exporting the graph.

Input is in the form of a 2010 Microsoft Excel file with each line representing a charter, with the first cell containing the charter
number, the second cell containing the primary author of the charter, the third the location, the third the date, and the fourth
a list of other names mentioned on the charter, delimited by commas(","), with relations being listed following the name in parentheses,
delimited by slashed ("/").

Export is created in the GEXF Graph File format, .gexf, for importation into gephy (https://gephi.org/)

Known Bugs:
Relations are not properly handled
