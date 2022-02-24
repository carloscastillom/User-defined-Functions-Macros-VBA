Attribute VB_Name = "Modul1"
Function CountCcolor(range_data As Range) As Integer
   
    Dim cellx As Range
    Dim countF As Integer
    
    countF = 0
    Set Rango = range_data
    
    For Each cellx In Rango
        If cellx.Font.Color = 14857357 Then
            countF = countF + cellx.Value
        End If
        
    Next cellx

    CountCcolor = countF
End Function


     
