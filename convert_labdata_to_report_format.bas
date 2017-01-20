Option Explicit


Sub AccutestCopyData()
'NOTES:
'This macro takes an analytical data table from Accutests LabLink website (Only VOCs in GW at this time) and transforms it into a LoCastro Group Standard Table.
'This macro uses an empty template and copies the relevant info and data from the Accutest Table

'TO DO:
'   3. Replace all the hard coded info (noted in the code)
'   4. Additonal formatting
'       Remove unneeded footnotes
'       Add TICs
'       Page breaks (if possible)
'       Number Formata Totals
'       Change to only number format results
'       Sort samples (if possible)
'       Add depths (if possible)

    Application.ScreenUpdating = False
    
    'Copy the Template table and notes and rename the accutest sheet
    Call CopyVOCTemplate
    

'=============================================================================
'                   Define Variables
'=============================================================================

    'Sets AccutestSheet and TableSheet to worksheets
    Dim AccutestSheet, TableSheet As Worksheet
    Set AccutestSheet = ActiveWorkBook.Sheets("Accutest Table")
    Set TableSheet = ActiveWorkBook.Sheets("Table")

    Sheets("Table").Select
    'Sets TableDataRange and TableStandardsRange to the data and standards for the destination table
    '!!!!MAKE DYNAMIC
    Dim TableDataRange, TableStandardsRange  As Range
    Set TableStandardsRange = Range("C7:C63")
    Set TableDataRange = Range("E7:YL63")
    
    'Sets TableFirstRow and TableFirstColumn to first row and column of table data (results)
    Dim TableFirstRow, TableFirstCol As Integer
    TableFirstRow = TableDataRange.Cells(1, 1).Row
    TableFirstCol = TableDataRange.Cells(1, 1).Column
    
    Dim TableNumRows As Integer
    TableNumRows = TableDataRange.Rows.Count
       
    'Find last column and row in Accutest data (whole table)
    Dim AccutestLastCol, AccutestLastRow As Integer
    AccutestLastCol = AccutestSheet.Cells(9, Columns.Count).End(xlToLeft).Column + 1
    AccutestLastRow = AccutestSheet.Cells(Rows.Count, 1).End(xlUp).Row

    'Find first column and row in Accutest data (results)
    '!!!!!NEED TO FIX: caluculate instead of hard code
    '!!!!!NEED TO FIX: Assums that there are no standards i.e first column is alwasy 3 over
    Dim AccutestFirstCol, AccutestFirstRow As Integer
    AccutestFirstCol = 2
    AccutestFirstRow = 15
    
    'Find number of columns and rows in Accutest data (results)
    Dim AccutestNumCols, AccutestNumRows As Integer
    AccutestNumCols = AccutestLastCol - AccutestFirstCol + 1
    AccutestNumRows = AccutestLastRow - AccutestFirstRow
    

'========================================================================
'               Copy Results
'========================================================================
      
    'Difines the variables for the Table CAS references in the loop
    Dim TableCASRange As Range
    Dim TableCASString As String
    
    'Difine VlookupOffset as the offset from CAS number to first data entery in Accutest Table
    '!!!NEED TO MAKE DYNAMIC
    Dim VlookupOffset As Integer
    
    
   'Itterates through the selection
    TableSheet.Select
    Dim Row, Col As Integer
    
    For Row = TableFirstRow To TableFirstRow + TableNumRows - 1
    
        'Reset VlookupOffset to 2
        VlookupOffset = 2
    
        For Col = TableFirstCol To TableFirstCol + AccutestNumCols
                
            '!!!!NEED TO MAKE DYNAMIC (UNLESS THE TABLE CAS NUMBERS WILL ALWAYS BE COLUMN B)
            TableCASString = "B" & Row
            Set TableCASRange = TableSheet.Range(TableCASString)
            
            'Increment VlookupOffset by one each iteration so VLookup references correct column
            VlookupOffset = VlookupOffset + 1
            
            '!!!NEED TO FIX Replace accutestSheet.Range(XXX) with dynamic variable
            On Error Resume Next
            Cells(Row, Col).Value = Application.WorksheetFunction.VLookup(TableCASRange, AccutestSheet.Range("B15:YL200"), VlookupOffset, False)
            'Cells(Row, Col).Value = "Test"
        
        Next Col
        
    Next Row

'========================================================================
'               Copy Headers
'========================================================================


    'Set Accutest Header Ranges
    AccutestSheet.Select
    Dim AccutestSampleIDRange, AccutestLabIDRange, AccutestDateRange As Range
    Set AccutestSampleIDRange = Range("D7:YD7")
    Set AccutestLabIDRange = Range("D8:YD8")
    Set AccutestDateRange = Range("D9:YD9")
   
    'Set Table Header Ranges
    TableSheet.Select
    Dim TableSampleIDRange, TableLabIDRange, TableDateRange As Range
    Set TableSampleIDRange = Range("E1:YD1")
    Set TableLabIDRange = Range("E3:YD3")
    Set TableDateRange = Range("E4:YD4")
   
    'Copy the Headers
    TableSampleIDRange.Value = AccutestSampleIDRange.Value
    TableLabIDRange.Value = AccutestLabIDRange.Value
    TableDateRange.Value = AccutestDateRange.Value
 
'========================================================================
'               Screen Data
'========================================================================
  
    Call ScreenGWData(TableDataRange, TableStandardsRange)
    
    Call FormatNumbers(TableDataRange)
    
    Call FormatNumbersResults(Range("E65:YL65"))
        
    Call AddNA(TableDataRange)
     
    'Deleate Accutest Tab and don't show the delete worksheet popup message (Currently turned off)
    Application.DisplayAlerts = False
    'AccutestSheet.Delete
    Application.DisplayAlerts = True
    
    'Adjust the print area
    '!!!!To be fixed
    Dim TablePrintRange As Range
    TablePrintRange = Range("A1:AA:61")
    'Cells(1, 1), Cells(61, 118))
    
    TableSheet.PageSetup.PrintArea = TablePrintRange
        
    Call RemoveBlankRows(TableDataRange)
    
End Sub