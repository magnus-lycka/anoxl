# anoxl

A tool for adding an anonymous id for tabular data in Excel Workbooks.
N.B. The script does not remove the sensitive data, it just adds the
     anonymous id. Be sure to remove the sensitive id in all worksheets
     before sharing data.

## Mode of operation

You need a "mapping workbook" where the active worksheet has a column with
the sensitive id, and a column with the anonymous id. The first row in the
worksheet should contain the names of the ids, and row two and below should 
contain all the ids.

The workbook where you want the anonymous ids added need to have columns for
the sensitive ids and the anonymous ids in each worksheet where you want
anonymous ids inserted. The names of the ids must be in the first row. The
sensitive ids found in the worksheets must be present int mapping file.

The first thing you do after starting the program is to select the mapping
file. Having done that, you select which column is the sensitive id and 
anoymous id in the dropdown boxes.

Having done that, you select a data file where you want to add anonymous ids.
Progress is logged in the GUI, and you are asked where to save the changed
data. After that, you can select additional datafiles to process.

