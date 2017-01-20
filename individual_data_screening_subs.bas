Option Explicit


Sub HighlightExceedancesSub()

    'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedDataRange, UserDefinedStandardsRange As Range
    Set UserDefinedDataRange = Application.InputBox("", "Select the Data to Screen", Type:=8)
    Set UserDefinedStandardsRange = Application.InputBox("", "Select the Standards to Screen Against", Type:=8)

    'Check if standards and data are formatted as numbers
    Call CheckIfRangeNumberFormat(UserDefinedDataRange)
    Call CheckIfRangeNumberFormat(UserDefinedStandardsRange)
    
    Call HighlightExceedances(UserDefinedDataRange, UserDefinedStandardsRange)
    
End Sub

Sub ItalicsRLExceedancesSub()

    'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedDataRange, UserDefinedStandardsRange As Range
    Set UserDefinedDataRange = Application.InputBox("", "Select the Data to Screen", Type:=8)
    Set UserDefinedStandardsRange = Application.InputBox("", "Select the Standards to Screen Against", Type:=8)

    'Check if standards and data are formatted as numbers
    'Call CheckIfRangeNumberFormat(UserDefinedDataRange)
    Call CheckIfRangeNumberFormat(UserDefinedStandardsRange)
    
    Call ItalicsRLExceedances(UserDefinedDataRange, UserDefinedStandardsRange)
    
End Sub

Sub AddRoundingQualifier()

    'Confirm that data and standards are formatted as numbers
    
    
    'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedDataRange, UserDefinedStandardsRange As Range
    Set UserDefinedDataRange = Application.InputBox("", "Select the Data to Screen", Type:=8)
    Set UserDefinedStandardsRange = Application.InputBox("", "Select the Standards to Screen Against", Type:=8)

    'Check if standards and data are formatted as numbers
    Call CheckIfRangeNumberFormat(UserDefinedDataRange)
    Call CheckIfRangeNumberFormat(UserDefinedStandardsRange)
    
    Call Rounding(UserDefinedDataRange, UserDefinedStandardsRange)


    
End Sub

Sub ScreenData()
'Screens data against standards (rounds results that would not exceed the standard if they were rounded, italicices NDs with RLs over standard and highlights exceedances)
   
    'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedDataRange, UserDefinedStandardsRange As Range
    Set UserDefinedDataRange = Application.InputBox("", "Select the Data to Screen", Type:=8)
    Set UserDefinedStandardsRange = Application.InputBox("", "Select the Standards to Screen Against", Type:=8)
    
    Call ScreenGWData(UserDefinedDataRange, UserDefinedStandardsRange)

End Sub

Sub FullyScreenSoilData()
'Screens data against standards (rounds results that would not exceed the standard if they were rounded, italicices NDs with RLs over standard and highlights exceedances)
   
    'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedDataRange, UserDefinedStandardsRangeNonRes, UserDefinedStandardsRangeRes, UserDefinedStandardsRangeIGW, UserDefinedSaturationRange As Range
    Set UserDefinedDataRange = Application.InputBox("", "Select the Data to Screen", Type:=8)
    Set UserDefinedStandardsRangeNonRes = Application.InputBox("", "Select the Non-Res Standards to Screen Against", Type:=8)
    Set UserDefinedStandardsRangeRes = Application.InputBox("", "Select the Res Standards to Screen Against", Type:=8)
    Set UserDefinedStandardsRangeIGW = Application.InputBox("", "Select the IGW Standards to Screen Against", Type:=8)
    Set UserDefinedSaturationRange = Application.InputBox("", "Select the Saturation Information Range (The column with saturation info)", Type:=8)
    
    'Check if standards and data are formatted as numbers
'    Call CheckIfRangeNumberFormat(UserDefinedDataRange)
'    Call CheckIfRangeNumberFormat(UserDefinedStandardsRangeNonRes)
'    Call CheckIfRangeNumberFormat(UserDefinedStandardsRangeRes)
'    Call CheckIfRangeNumberFormat(UserDefinedStandardsRangeIGW)
    
    'Direct Contact Screening
    Call ItalicsRLExceedances(UserDefinedDataRange, UserDefinedStandardsRangeNonRes)
    Call ItalicsRLExceedances(UserDefinedDataRange, UserDefinedStandardsRangeRes)
    'TO DO LATER: Need to fix so it won't add two "r" if both res and non-res are rounded
    Call Rounding(UserDefinedDataRange, UserDefinedStandardsRangeNonRes)
    Call Rounding(UserDefinedDataRange, UserDefinedStandardsRangeRes)
    Call HighlightExceedances(UserDefinedDataRange, UserDefinedStandardsRangeNonRes)
    Call HighlightExceedances(UserDefinedDataRange, UserDefinedStandardsRangeNonRes)
    
    'IGW Screening
    'TO DO LATER: Need to add rounding and italics RLs for IGW
    Call HighlightExceedancesIGW(UserDefinedDataRange, UserDefinedStandardsRangeIGW, UserDefinedSaturationRange)
    Call RoundingIGW(UserDefinedDataRange, UserDefinedStandardsRangeIGW, UserDefinedSaturationRange)
End Sub


Sub NumberFormat()
   
        'Sets DataRange and StandardsRange to the data and standards from user input
    Dim UserDefinedRange As Range
    Set UserDefinedRange = Application.InputBox("", "Select the Data to Format", Type:=8)

    Call FormatNumbers(UserDefinedRange)
      
End Sub
