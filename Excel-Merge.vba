' Merge 2 MS Excel files with VBA code with exact format and formulas
' Instructions Below
' Open MS Excel, run Alt+F11, Insert -> Module, copy and paste code below


Sub Copy_Paste_With_Format_And_Formula_Without_source_append()
    Dim rngSource As Range
    Dim rngTarget As Range
    
    ' Set the range with the formula to autofill
    Set rngSource = Workbooks("Asian PaintsTF.xlsx").Sheets("GAME OF AVERAGES").Range("A:Z")
    
    ' Set the range where the formula should be autofilled
    Set rngTarget = Workbooks("Asian Paints.xlsx").Sheets("GAME OF AVERAGES").Range("A:Z")
    
    ' Clear existing conditional formatting in the target range
    ' rngTarget.FormatConditions.Delete
        
    ' Copy conditional formatting from source range to target range
    rngSource.Copy
    rngTarget.PasteSpecial Paste:=xlPasteFormats
    
    ' Autofill the formula in the target range
    rngTarget.Formula = rngSource.Formula
    
    ' Optional: Turn off the Autofill Options button
    Application.AutoCorrect.AutoFillFormulasInLists = False
    
    ' Optional: Clear the clipboard
    Application.CutCopyMode = False
    
    ' Optional: Display a message box when the autofill is complete
    MsgBox "Formulas have been autofilled in the range."
End Sub

