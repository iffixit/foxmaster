Function FIND(searchStr As String) As Integer
Dim price As Dictionary
Set price = New Dictionary

With price
On Error Resume Next
    .Add "Послуга 1", 10
    .Add "Послуга 2", 20    
    .Add "Послуга 3", 30
End With
FIND = price(searchStr)
End Function



