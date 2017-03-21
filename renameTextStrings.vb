
Public Sub rename()
'set up some vars to be used
    Dim symb As Symbol
    Dim name As String
    Dim typ As Integer
    Dim tagname As String
    Dim val As Value
'go through all symbols in a display
    For Each symb In ThisDisplay.Symbols
        name = symb.name
        typ = symb.Type
          'type 7 symbols are Values
          'type 4 symbols are text
        If typ = 4 Then
            symb.Contents = Replace(symb.Contents, "OLDString", "NEWString")
              
        End If
                        
    Next
      
End Sub
