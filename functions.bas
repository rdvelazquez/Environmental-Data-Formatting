
Public Function CopyTemplate(FileName) As String
    
    'Rename current sheet to "Accutest Table"
    ActiveSheet.Name = "Accutest Table"
    
    'Set AccutestWorkbook to the active (open) workbook
    Dim AccutestWorkbook As Workbook
    Set AccutestWorkbook = ActiveWorkBook
    
    'Set AccutestTable to the filename of the open workbook
    Dim AccutestTable As String
    AccutestTable = ThisWorkbook.Name
        
    'Open workbook template from file
    Dim TemplatePath As String, TemplateFile As String
    Dim TemplateWorkbook As Workbook
    TemplatePath = "\\langan.com\data\DT\other\RVelazquez\VBA\"
    TemplateFile = TemplatePath & FileName
    Set TemplateWorkbook = Workbooks.Open(TemplateFile)
    
    'Copy table and notes sheets to Data Workbook
    TemplateWorkbook.Sheets("Table").Copy After:=AccutestWorkbook.Sheets(1)
    TemplateWorkbook.Sheets("Notes").Copy After:=AccutestWorkbook.Sheets(2)
    
    TemplateWorkbook.Close
    
End Function

Public Function CopyVOCTemplate()
    
    'Rename current sheet to "Accutest Table"
    ActiveSheet.Name = "Accutest Table"
    
    'Set AccutestWorkbook to the active (open) workbook
    Dim AccutestWorkbook As Workbook
    Set AccutestWorkbook = ActiveWorkBook
    
    'Set AccutestTable to the filename of the open workbook
    Dim AccutestTable As String
    AccutestTable = ThisWorkbook.Name
        
    'Open workbook template from file
    Dim TemplatePath As String, TemplateFile As String
    Dim TemplateWorkbook As Workbook
    TemplatePath = "\\langan.com\data\DT\other\RVelazquez\VBA\"
    TemplateFile = TemplatePath & "Table X - VOC Analytical Results for Groundwater.xlsm"
    Set TemplateWorkbook = Workbooks.Open(TemplateFile)
    
    'Copy table and notes sheets to Data Workbook
    TemplateWorkbook.Sheets("Table").Copy After:=AccutestWorkbook.Sheets(1)
    TemplateWorkbook.Sheets("Notes").Copy After:=AccutestWorkbook.Sheets(2)
    
    TemplateWorkbook.Close
    
End Function
Public Function CheckIfRangeNumberFormat(Range1) As Range

    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRowsCheck, NumColsCheck As Integer
    NumRowsCheck = Range1.Rows.Count
    NumColsCheck = Range1.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = Range1.Cells(1, 1).Row
    FirstCol = Range1.Cells(1, 1).Column
    
    'Itterates through the selection
    Dim Row, Col As Integer
    Row = 1
    Col = 1
    
    For Row = FirstRow To FirstRow + NumRowsCheck - 1
        
        'Check if the standards are number-formatted-as-text and let the user know
        If Cells(Row, Col).NumberFormat = "@" Then
            MsgBox "WARNING! Standards are not formatted as numbers"
            'End
        End If
        
        For Col = FirstCol To FirstCol + NumColsCheck - 1
            
            'Check if result
            If Cells(Row, Col).NumberFormat = "@" Then
                MsgBox "WARNING! Standards are not formatted as numbers"
                'End
            End If
           
        Next Col
        
    Next Row

End Function

Public Function FormatNumbers(Range1) As Range
   
    'Sets NumRows and NumCols to the number of rows and columns in the selection
    Dim NumRows, NumCols As Integer
    NumRows = Range1.Rows.Count
    NumCols = Range1.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the selection
    Dim FirstRow, FirstCol As Integer
    FirstRow = Range1.Cells(1, 1).Row
    FirstCol = Range1.Cells(1, 1).Column
    
    'itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        For Col = FirstCol To FirstCol + NumCols - 1
            
            'add check for if cell is number
            'add check for if cell contains superscript
            
            '!!!!NEED TO FIX SO IT DOESN'T FORMAT DATES
            If Application.WorksheetFunction.IsNumber(Cells(Row, Col)) = True Then
                
                If Cells(Row, Col).Value = 0 Then
                    Cells(Row, Col).NumberFormat = "#,##0"
                    
                ElseIf Cells(Row, Col).Value > 100 Then
                    Cells(Row, Col).NumberFormat = "#,##0"
                
                ElseIf Cells(Row, Col).Value >= 1 Then
                    Cells(Row, Col).NumberFormat = "#,##0.0"
                    
                ElseIf Cells(Row, Col).Value >= 0.1 Then
                    Cells(Row, Col).NumberFormat = "#,##0.00"
                    
                ElseIf Cells(Row, Col).Value >= 0.01 Then
                    Cells(Row, Col).NumberFormat = "#,##0.000"
                    
                ElseIf Cells(Row, Col).Value >= 0.001 Then
                    Cells(Row, Col).NumberFormat = "#,##0.0000"
                    
                ElseIf Cells(Row, Col).Value >= 0.00001 Then
                    Cells(Row, Col).NumberFormat = "#,##0.000000"
                    
                End If
                
            End If
                        
        Next Col
        
    Next Row

End Function

Public Function HighlightExceedances(DataRange, StandardsRange) As Range

'**********************************************************
'               DEFINE VARIABLES
'**********************************************************

    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column
   
    'Sets StandardsRow and StandardsCol to the first row and column of the Standards
    Dim StandardsRow, StandardsCol As Integer
    StandardsRow = StandardsRange.Cells(1, 1).Row
    StandardsCol = StandardsRange.Cells(1, 1).Column
    
    'Sets NumStandardsRows and NumStandardsCols to the number of rows and columns in the StandardsRange
    Dim NumStandardsRows, NumStandardsCols As Integer
    NumStandardsRows = StandardsRange.Rows.Count
    NumStandardsCols = StandardsRange.Columns.Count
    
    'Return a Warning if the number of rows in the data and the standards didn't match
    If NumStandardsRows <> NumRows Then
        MsgBox "WARNING! The number of rows in the data and the number of rows in the standards didn't match."
    End If
    
    'Return a Warning if more than one column was selected for the standards
    If NumStandardsCols <> 1 Then
        MsgBox "WARNING! More than one Column was selected for the standards. Only the first column was used."
    End If
   
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        'Check if Standard is Number
        If IsEmpty(Cells(Row, StandardsCol)) = False And IsNumeric(Cells(Row, StandardsCol)) Then
                    
            For Col = FirstCol To FirstCol + NumCols - 1
                
                'Check if result
                If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                    
                    'Check if detection
                    If Cells(Row, Col + 1) <> "U" And Cells(Row, Col + 1) <> "UJ" And Cells(Row, Col + 1) <> "r  " And Cells(Row, Col + 1) <> "r " And Cells(Row, Col + 1) <> "r J" And Cells(Row, Col + 1) <> "s " And Cells(Row, Col + 1) <> "s J" Then
                    
                        'Check if an exceedance
                        If Cells(Row, Col) > Cells(Row, StandardsCol) Then
                            
                            '**********************************************
                            '       Formatting Exceedances
                            '**********************************************
                                
                                'Format Result Cell
                                With Cells(Row, Col).Font
                                    .FontStyle = "Bold"
                                End With
                                With Cells(Row, Col).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .ThemeColor = xlThemeColorDark1
                                    .TintAndShade = -0.149998474074526
                                    .PatternTintAndShade = 0
                                End With
                                
                                'Format Qualifier Cell
                                With Cells(Row, Col + 1).Font
                                    .FontStyle = "Bold"
                                End With
                                With Cells(Row, Col + 1).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .ThemeColor = xlThemeColorDark1
                                    .TintAndShade = -0.149998474074526
                                    .PatternTintAndShade = 0
                                End With
                                
                        End If
                        
                    End If
                    
                End If
                
            Next Col
            
        End If
        
    Next Row

End Function

Public Function ItalicsRLExceedances(DataRange, StandardsRange) As Range

    
'**********************************************************
'               DEFINE VARIABLES
'**********************************************************


    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column
   
    'Sets StandardsRow and StandardsCol to the first row and column of the Standards
    Dim StandardsRow, StandardsCol As Integer
    StandardsRow = StandardsRange.Cells(1, 1).Row
    StandardsCol = StandardsRange.Cells(1, 1).Column
    
    'Sets NumStandardsRows and NumStandardsCols to the number of rows and columns in the StandardsRange
    Dim NumStandardsRows, NumStandardsCols As Integer
    NumStandardsRows = StandardsRange.Rows.Count
    NumStandardsCols = StandardsRange.Columns.Count
    
    'Return a Warning if the number of rows in the data and the standards didn't match
    If NumStandardsRows <> NumRows Then
        MsgBox "WARNING! The number of rows in the data and the number of rows in the standards didn't match."
    End If
    
    'Return a Warning if more than one column was selected for the standards
    If NumStandardsCols <> 1 Then
        MsgBox "WARNING! More than one Column was selected for the standards. Only the first column was used."
    End If
   
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        'Check if Standard is Number
        If IsEmpty(Cells(Row, StandardsCol)) = False And IsNumeric(Cells(Row, StandardsCol)) Then
                    
            For Col = FirstCol To FirstCol + NumCols - 1
                
                'Check if result
                If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                    
                    'Check if detection
                    If Cells(Row, Col + 1) = "U" Or Cells(Row, Col + 1) = "UJ" Then
                    
                        'Check if an exceedance
                        If Cells(Row, Col) > Cells(Row, StandardsCol) Then
                            
                            '**********************************************
                            '       Formatting Exceedances
                            '**********************************************
                                
                                'Format Result Cell
                                Cells(Row, Col).Font.Italic = True
                                
                                'Format Qualifier Cell
                                Cells(Row, Col + 1).Font.Italic = True
                                
                        End If
                        
                    End If
                    
                End If
                
            Next Col
            
        End If
        
    Next Row

End Function

Public Function Rounding(DataRange, StandardsRange) As Range


    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column
   
    'Sets StandardsRow and StandardsCol to the first row and column of the Standards
    Dim StandardsRow, StandardsCol As Integer
    StandardsRow = StandardsRange.Cells(1, 1).Row
    StandardsCol = StandardsRange.Cells(1, 1).Column
    
    'Sets NumStandardsRows and NumStandardsCols to the number of rows and columns in the StandardsRange
    Dim NumStandardsRows, NumStandardsCols As Integer
    NumStandardsRows = StandardsRange.Rows.Count
    NumStandardsCols = StandardsRange.Columns.Count
    
    'Return a Warning if the number of rows in the data and the standards didn't match
    If NumStandardsRows <> NumRows Then
        MsgBox "WARNING! The number of rows in the data and the number of rows in the standards didn't match."
        End
    End If
    
    'Return a Warning if more than one column was selected for the standards
    If NumStandardsCols <> 1 Then
        MsgBox "WARNING! More than one Column was selected for the standards."
        End
    End If

    
    
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        'Check if the standards are number-formatted-as-text and let the user know
        If Cells(Row, StandardsCol).NumberFormat = "@" Then
            MsgBox "WARNING! Standards are not formatted as numbers"
            End
        End If
        
        'Check if Standard is Number
        If IsEmpty(Cells(Row, StandardsCol)) = False And IsNumeric(Cells(Row, StandardsCol)) Then
                    
            For Col = FirstCol To FirstCol + NumCols - 1
                
                'Check if result
                If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                    
                    'Check if detection
                    If Cells(Row, Col + 1) <> "U" And Cells(Row, Col + 1) <> "UJ" Then
                    
                        'Check if an exceedance
                        If Cells(Row, Col) > Cells(Row, StandardsCol) Then
                        
                            'Check if result can be rounded below standard
                            If Cells(Row, StandardsCol) >= 1 And Cells(Row, Col) - 0.499 <= Cells(Row, StandardsCol) Then
                                
                                'Add "r" qualifier
                                Cells(Row, Col + 1).Value = "r " & Cells(Row, Col + 1).Value
                                'Truncate result to number of decimal places
                                'Cells(Row, Col).NumberFormat = "#,##0"
                                
                            ElseIf Cells(Row, StandardsCol) < 1 And Cells(Row, StandardsCol) >= 0.1 And Cells(Row, Col) - 0.0499999 <= Cells(Row, StandardsCol) Then
                            
                                'Add"r" qualifier
                                Cells(Row, Col + 1).Value = "r " & Cells(Row, Col + 1).Value
                                'Truncate result to number of decimal places
                                'Cells(Row, Col).NumberFormat = "#,##0.0"
                            
                            End If
    
                        End If
                        
                    End If
                    
                End If
                
            Next Col
            
        End If
        
    Next Row


End Function

Public Function ScreenGWData(DataRange, StandardsRange) As Range

    'Check if standards and data are formatted as numbers
    'Call CheckIfRangeNumberFormat(DataRange)
    'Call CheckIfRangeNumberFormat(StandardsRange)
    
    Call ItalicsRLExceedances(DataRange, StandardsRange)
    
    Call Rounding(DataRange, StandardsRange)
    
    Call HighlightExceedances(DataRange, StandardsRange)
    
End Function
Public Function RoundingIGW(DataRange, StandardsRange, SaturationRange) As Range
    
    'Sets SaturationRow to the row with the Saturation info
    Dim SaturationRow As Integer
    SaturationRow = SaturationRange.Cells(1, 1).Row
    
    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column
   
    'Sets StandardsRow and StandardsCol to the first row and column of the Standards
    Dim StandardsRow, StandardsCol As Integer
    StandardsRow = StandardsRange.Cells(1, 1).Row
    StandardsCol = StandardsRange.Cells(1, 1).Column
    
    'Sets NumStandardsRows and NumStandardsCols to the number of rows and columns in the StandardsRange
    Dim NumStandardsRows, NumStandardsCols As Integer
    NumStandardsRows = StandardsRange.Rows.Count
    NumStandardsCols = StandardsRange.Columns.Count
    
    'Return a Warning if the number of rows in the data and the standards didn't match
    If NumStandardsRows <> NumRows Then
        MsgBox "WARNING! The number of rows in the data and the number of rows in the standards didn't match."
    End If
    
    'Return a Warning if more than one column was selected for the standards
    If NumStandardsCols <> 1 Then
        MsgBox "WARNING! More than one Column was selected for the standards. Only the first column was used."
    End If
   
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        'Check if Standard is Number
        If IsEmpty(Cells(Row, StandardsCol)) = False And IsNumeric(Cells(Row, StandardsCol)) Then
                    
            For Col = FirstCol To FirstCol + NumCols - 1
                
                'Check if result
                If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                    
                    'Check if sample is saturated
                    If Cells(SaturationRow, Col) = "Yes" Or Cells(SaturationRow, Col) = "Saturated" Or Cells(SaturationRow, Col) = "YES" Or Cells(SaturationRow, Col) = "SATURATED" Or Cells(SaturationRow, Col) = "yes" Then
                        
                    Else
                        
                        'Check if detection
                        If Cells(Row, Col + 1) <> "U" And Cells(Row, Col + 1) <> "UJ" Then
                        
                            'Check if an exceedance
                            If Cells(Row, Col) > Cells(Row, StandardsCol) Then
                            
                                'Check if result can be rounded below standard
                                If Cells(Row, StandardsCol) >= 1 And Cells(Row, Col) - 0.499 <= Cells(Row, StandardsCol) Then
                                    
                                    'Add "r" qualifier
                                    Cells(Row, Col + 1).Value = "r " & Cells(Row, Col + 1).Value
                                    'Truncate result to number of decimal places
                                    'Cells(Row, Col).NumberFormat = "#,##0"
                                    
                                ElseIf Cells(Row, StandardsCol) < 1 And Cells(Row, StandardsCol) >= 0.1 And Cells(Row, Col) - 0.0499999 <= Cells(Row, StandardsCol) Then
                                
                                    'Add"r" qualifier
                                    Cells(Row, Col + 1).Value = "r " & Cells(Row, Col + 1).Value
                                    'Truncate result to number of decimal places
                                    'Cells(Row, Col).NumberFormat = "#,##0.0"
                                    
                                End If
                                    
                            End If
                            
                        End If
                    
                    End If
                
                End If
                
            Next Col
            
        End If
        
    Next Row

End Function

Public Function HighlightExceedancesIGW(DataRange, StandardsRange, SaturationRange) As Range
    
    'Sets SaturationRow to the row with the Saturation info
    Dim SaturationRow As Integer
    SaturationRow = SaturationRange.Cells(1, 1).Row
    
    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column
   
    'Sets StandardsRow and StandardsCol to the first row and column of the Standards
    Dim StandardsRow, StandardsCol As Integer
    StandardsRow = StandardsRange.Cells(1, 1).Row
    StandardsCol = StandardsRange.Cells(1, 1).Column
    
    'Sets NumStandardsRows and NumStandardsCols to the number of rows and columns in the StandardsRange
    Dim NumStandardsRows, NumStandardsCols As Integer
    NumStandardsRows = StandardsRange.Rows.Count
    NumStandardsCols = StandardsRange.Columns.Count
    
    'Return a Warning if the number of rows in the data and the standards didn't match
    If NumStandardsRows <> NumRows Then
        MsgBox "WARNING! The number of rows in the data and the number of rows in the standards didn't match."
    End If
    
    'Return a Warning if more than one column was selected for the standards
    If NumStandardsCols <> 1 Then
        MsgBox "WARNING! More than one Column was selected for the standards. Only the first column was used."
    End If
   
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        'Check if Standard is Number
        If IsEmpty(Cells(Row, StandardsCol)) = False And IsNumeric(Cells(Row, StandardsCol)) Then
                    
            For Col = FirstCol To FirstCol + NumCols - 1
                
                'Check if result
                If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                    
                    'Check if detection
                    If Cells(Row, Col + 1) <> "U" And Cells(Row, Col + 1) <> "UJ" Then
                    
                        'Check if an exceedance
                        If Cells(Row, Col) > Cells(Row, StandardsCol) Then
                        
                            'Check if sample is saturated
                            If Cells(SaturationRow, Col) = "Yes" Or Cells(SaturationRow, Col) = "Saturated" Or Cells(SaturationRow, Col) = "YES" Or Cells(SaturationRow, Col) = "SATURATED" Or Cells(SaturationRow, Col) = "yes" Then
                                                       
                                  'Add "s" qualifier
                                  Cells(Row, Col + 1).Value = "s " & Cells(Row, Col + 1).Value
                           Else
                            
                                'Format Result Cell with bold blue text
                                With Cells(Row, Col).Font
                                    .FontStyle = "Bold"
                                    .Color = RGB(0, 176, 240)
                                End With
                                
                                'Format Qualifier Cell with bold blue text
                                With Cells(Row, Col).Font
                                    .FontStyle = "Bold"
                                    .Color = RGB(0, 176, 240)
                                End With
                                
                            End If
                                
                        End If
                        
                    End If
                    
                End If
                
            Next Col
            
        End If
        
    Next Row

End Function

Public Function AddNA(Range1) As Range
'This Function adds "NA" to the result cell for samples that weren't analyzed for a specific compound and currently have "-"
  
    'Sets NumRows and NumCols to the number of rows and columns in the selection
    Dim NumRows, NumCols As Integer
    NumRows = Range1.Rows.Count
    NumCols = Range1.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the selection
    Dim FirstRow, FirstCol As Integer
    FirstRow = Range1.Cells(1, 1).Row
    FirstCol = Range1.Cells(1, 1).Column
    
    'itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
        
        For Col = FirstCol To FirstCol + NumCols - 1
                       
            If Cells(Row, Col).Value = "-" Then
                Cells(Row, Col).Value = "NA"
                
            End If
                                      
        Next Col
        
    Next Row

End Function
Public Function RemoveBlankRows(Range1) As Range
'This function removes rows that don't have any data. It checks the first column for blanks and removes those rows

    'Sets NumRows and NumCols to the number of rows and columns in the selection
    Dim NumRows As Integer
    NumRows = Range1.Rows.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the selection
    Dim FirstRow, FirstCol As Integer
    FirstRow = Range1.Cells(1, 1).Row
    FirstCol = Range1.Cells(1, 1).Column
    
    'itterates through the selection
    Dim Row, Col, X As Integer
           
    Col = 1
    Row = 1
    X = 0
    
    While X < NumRows
        If IsEmpty(Range1.Cells(Row, Col)) Then
            'Range1.Cells(Row, Col).Value = "test"
            Range1.Rows(Row).EntireRow.Delete
            Row = Row - 1
        End If
        X = X + 1
        Row = Row + 1
    Wend


End Function

Public Function FormatNumbersResults(DataRange) As Range

'**********************************************************
'               DEFINE VARIABLES
'**********************************************************

    'Sets NumRows and NumCols to the number of rows and columns in the Data
    Dim NumRows, NumCols As Integer
    NumRows = DataRange.Rows.Count
    NumCols = DataRange.Columns.Count
    
    'Sets FirstRow and FirstCol to the first row and column of the Data
    Dim FirstRow, FirstCol As Integer
    FirstRow = DataRange.Cells(1, 1).Row
    FirstCol = DataRange.Cells(1, 1).Column

   
 '*********************************************************
 '          Main Loop for Screening Data
 '*********************************************************
    
    'Itterates through the selection
    Dim Row, Col As Integer
    
    For Row = FirstRow To FirstRow + NumRows - 1
                        
        For Col = FirstCol To FirstCol + NumCols - 1
            
            'Check if result
            If IsEmpty(Cells(Row, Col)) = False And IsNumeric(Cells(Row, Col)) Then
                
                'Check if detection
                If Cells(Row, Col + 1) <> "U" And Cells(Row, Col + 1) <> "UJ" And Cells(Row, Col + 1) <> "r " And Cells(Row, Col + 1) <> "r J" And Cells(Row, Col + 1) <> "s " And Cells(Row, Col + 1) <> "s J" Then
                
                    FormatNumbers (Cells(Row, Col))
                                        
                End If
                
            End If
            
        Next Col
        
    Next Row

End Function