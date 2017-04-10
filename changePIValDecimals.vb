Option Explicit  
Public Sub decimals()  
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
        If typ = 7 Then    
               'get the name of the PI point (if you only want to change certain tags)  
               'format is \\servername\PointName  
            tagname = symb.GetTagName(1)   
               'this will only change tags with the word 'temp' in the name, you can take this if statement out if you want to change all  values  
            If InStr(1, tagname, "temp") Then  
               'need to have a var of type Value to set the number format  
                Set val = symb  
               'sets the formal to 3 decimals always  
                val.NumberFormat = "0.000"  
            End If  
              
        End If  
                        
    Next  
      
End Sub 
