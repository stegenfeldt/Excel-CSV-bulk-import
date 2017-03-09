Attribute VB_Name = "CSVImportModule"
Option Explicit
Public Delimiter As String

Sub ImportCSVs()
'Author:    Jerry Beaucaire
'Date:      8/16/2010
'Summary:   Import all CSV files from a folder into separate sheets
'           named for the CSV filenames

'Update:    2/8/2013    Macro replaces existing sheets if they already exist in master workbook
'Update:    20170309    Uses current workbook path and a delimiter set from form //stegenfeldt

Dim fPath   As String
Dim fCSV    As String
Dim wbCSV   As Workbook
Dim wbMST   As Workbook
Dim importForm As CSVImportForm

CSVImportForm.Show

Set wbMST = ThisWorkbook
fPath = ActiveWorkbook.Path & "\"        'path to CSV files, include the final \
Application.ScreenUpdating = False  'speed up macro
Application.DisplayAlerts = False   'no error messages, take default answers
fCSV = Dir(fPath & "*.csv")         'start the CSV file listing

    On Error Resume Next
    Do While Len(fCSV) > 0
        Set wbCSV = Workbooks.Open(fPath & fCSV, , , 6, , , , , Delimiter)            'open a CSV file
        wbMST.Sheets(ActiveSheet.Name).Delete                       'delete sheet if it exists
        ActiveSheet.Move After:=wbMST.Sheets(wbMST.Sheets.Count)    'move new sheet into Mstr
        Columns.AutoFit             'clean up display
        FormatAsTable               'format the content as a proper table
        fCSV = Dir                  'ready next CSV
    Loop

Application.ScreenUpdating = True
Set wbCSV = Nothing
End Sub

Sub FormatAsTable()
    ' Credit to this function goes to simoco on stackoverflow
    ' http://stackoverflow.com/a/21558003
    
    Dim tbl As ListObject
    Dim rng As Range
    
    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium15"
End Sub
