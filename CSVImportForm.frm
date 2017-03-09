VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CSVImportForm 
   Caption         =   "CSV Import Settings"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   OleObjectBlob   =   "CSVImportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CSVImportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:    Samuel Tegenfeldt (stegenfeldt)
'Date:      20170309
'Summary:   Allows for configuration of delimiter to the CSV-imports
'Update:    201703  stegenfeldt Minor UI changes

Private Sub btnImport_Click()
    Delimiter = CSVImportForm.tbDelimiter.Text
    Unload CSVImportForm
End Sub
