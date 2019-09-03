Attribute VB_Name = "Last_Cell_Function"
' Function for finding the last column or row in excel
' Written by Ron de Bruin - for more visit - https://www.rondebruin.nl - I tweaked it a little for my usecase
' conventional xlEnd won't work as it doesn't take blanks into consideration while findint the last column or row _
    that's why we need to use function like FindLast for finding the last column or row

Function FindLast(lRowColCell As Long, _
                    Optional sSheet As String, _
                    Optional sRange As String)
'Find the last row, column, or cell using the Range.Find method
'lRowColCell: 1=Row, 2=Col, 3=Cell

Dim lRow As Long
Dim lCol As Long
Dim wsFind As Worksheet
Dim rFind As Range

    'Default to ActiveSheet if none specified
    On Error GoTo ErrExit
    
    If sSheet = "" Then
        Set wsFind = ActiveSheet
    Else
        Set wsFind = Worksheets(sSheet)
    End If

    'Default to all cells if range no specified
    If sRange = "" Then
        Set rFind = wsFind.Cells
    Else
        Set rFind = wsFind.Range(sRange)
    End If
    
    On Error GoTo 0

    Select Case lRowColCell
    
        Case 1 'Find last row
            On Error Resume Next
    
             FindLast = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row
            
            On Error GoTo 0

        Case 2 'Find last column
            On Error Resume Next
            
            Dim ColumnNumber As Integer

            ColumnNumber = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            
            ' Converting Column Number into Cloumn Name
            
            FindLast = Split(Cells(1, ColumnNumber).Address, "$")(1)
            On Error GoTo 0

        Case 3 'Find last cell by finding last row & col
            On Error Resume Next
            lRow = rFind.Find(What:="*", _
                           After:=rFind.Cells(1), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).row
            On Error GoTo 0

            On Error Resume Next
            lCol = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

            On Error Resume Next
            FindLast = wsFind.Cells(lRow, lCol).Address(False, False)
            'If lRow or lCol = 0 then entire sheet is blank, return "A1"
            If Err.Number > 0 Then
                FindLast = rFind.Cells(1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0

    End Select
    
    Exit Function
    
ErrExit:

    MsgBox "Error setting the worksheet or range."

End Function


