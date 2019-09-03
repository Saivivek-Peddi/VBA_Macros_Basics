Attribute VB_Name = "Variables"
' The following includes all the basic variables, their syntax and references and instantiations, that will
' be useful for most of the VBA projects

' Varible types - Integer, String, Date, Variant, Range, Object, Collection, Dictionary

' Variable Delcaration Format.
' Dim <variable_name> as <variable_type>

' Examples
Sub quick_start()
    
    Dim i, j As Integer, l As Long, d As Double, follow As String, today_date As Date, whatever As Variant, bool As Boolean
    i = 2
    bool = True
    j = 4
    d = 1.444
    l = 1000000
    follow = "Hello World"
    today_date = Date ' This is default function which returns systems date
    whatever = 1 ' can set this long, integer, object - whatever you want
    
    Debug.Print i, j, l, d, follow, today_date, whatever, bool ' similar to print statemnet and the console will
        ' be the immediate window
    
    ' Object, Collection, Dictionary
    ' First go to Tools --> References --> set Microsoft Scritping Runtime - just like importing moudles in python
    Dim new_collection As Collection, new_dict As Dictionary, obj As Object
    
    Set new_collection = New Collection ' instantiating objects - allways complusory while using objects
    Set new_dict = New Dictionary
    
    Set obj = New Collection ' You can set the object variable to any of the aviable object types -
        ' can be collection, dictionary, file system object, range etc...
    
    ' Adding items to collections and dictionaries
    new_collection.Add "Hola"
    new_collection.Add 1        ' collection is similar to list in python and add is simlar to append
    new_collection.Add d
    
    new_dict.Add "Hola", 1
    new_dict.Add "City", "Hyderabad"    ' Dictionary is similar to dictionary
    
    Debug.Print new_dict("Hola"), new_dict("City")
    Debug.Print new_collection(2)
    
    ' print all items in dict or collection
    
    For Each Item In new_dict
        Debug.Print Item, new_dict(Item)
    Next
    
    For Each Item In new_collection
        Debug.Print Item
    Next
    
    ' Get all the open workbook names into a collection(list)
    Dim collection2 As Collection
    Set collection2 = New Collection
    
    Dim oBook As Workbook
    For Each oBook In Workbooks
        collection2.Add oBook.Name
    Next
    
    ' Rnage object - Very useful
    
    Dim rng As Range
    
    Set rng = Range("A1:B2")
    
    ' Iterating Via Range
    For Each Item In rng
        Debug.Print Item.Value ' .Value property can be used, for both, to get and to put content in a cell.
        ' Here we are getting the content of the cell
        ' items will be read row wise
        
        Item.Value = "Hola"
        ' Here we are setting the content of the cell to Hola
    Next
    
End Sub



