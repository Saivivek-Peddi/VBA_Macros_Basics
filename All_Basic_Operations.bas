Attribute VB_Name = "All_Basic_Operations"
Sub basic_operations()
' In this sub you can learn how to perform all the basic operations of excel, like cut, copy, autofilter etc... , _
    using excel VBA

' 1. Selection of a range
    ' Examples
    Range("A1").Select
    Range("A1:B20").Select
    ' Dynamic Selection - Multiple Selections
    Range("A:A,C:C,F:F").Select ' - Selects columns A, C, F
    Range("A1:" & FindLast(3)).Select '- Select all the data at once

' 2. Cut Copy Paste
    ' Examples
    Range("A1:" & FindLast(3)).Copy Sheets("Sheet2").Range("A1")
    ' Copy all the data from one sheet to another sheet, named Sheet2, to a specific location in the sheet, here A1 _
     important note - After .Copy put the destination
    Range("A1:" & FindLast(3)).Cut Sheets("Sheet2").Range("A1")
    ' Similar to copy
    
    ' PasteSpecial
    Range("A1:" & FindLast(3)).Copy
    Sheets("Sheet2").Range("A1").PasteSpecial xlPasteValues
    
    'More Examples
    ' Dynamically select few columns and paste them in another sheet
    Range("A:A, C:C, E:E").Copy Sheets("Sheet2").Range("A1")
    
    ' Copy to another workbook - I will show two ways - hard and easy
    ' Hard Way
    Range("A1:B2").Select
    Selection.Copy

    Windows("Book1.xlsx").Activate
    ActiveSheet.Paste
    
    ' Simple Way
    Range("A1:B2").Copy Workbooks("Book1.xlsx").Sheets("Sheet2").Range("A1")
    ' Similar is for Cut Operation
    
' 3. Find and Replace
    
' 4. AutoFilter

' 5. SpecialCells - Very important and very useful
    

End Sub
