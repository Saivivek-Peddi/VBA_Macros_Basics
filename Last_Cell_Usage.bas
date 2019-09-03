Attribute VB_Name = "Last_Cell_Usage"
Sub usage()
Debug.Print FindLast(3)
' Prints Last Cell address
' Eg - if your last cell address is B2 the function will return B2 as a string

Debug.Print FindLast(2)
' Prints Last column address
' Eg - if your last column address is B the function will return B

Debug.Print FindLast(1)
' Prints Last row address
' Eg - if your last row address is 10 the function will return 10

'In all the three cases it will take Blanks also into consideration
End Sub
