Function FIND(searchStr As String) As Integer
Dim price As Dictionary
Set price = New Dictionary

With price
On Error Resume Next
    .Add "Адаптація Sim до стандарту microSim", 49
    .Add "Поклейка захисної плівки", 99
    .Add "Налаштування смартфону Minimum", 499
    .Add "Налаштування смартфону Standard", 699
    .Add "Налаштування смартфону Maximum", 1299
    .Add "Налаштування смартфону Ultra", 1699
    .Add "Налаштування планшетного ПК Minimum", 499
    .Add "Налаштування планшетного ПК Standard", 699
    .Add "Налаштування планшетного ПК Maximum", 1299
    .Add "Налаштування iPhone Minimum", 649
    .Add "Налаштування iPhone Standard", 949
    .Add "Налаштування iPhone Ultra", 2399
    .Add "Налаштування ПК Minimum", 999
    .Add "Налаштування ПК Standard", 1699
    .Add "Налаштування ПК Maximum", 2399
    .Add "Налаштування SmartTV Minimum", 1099
    .Add "Налаштування SmartTV Standard", 1699
    .Add "Налаштування SmartTV Ultra", 4999
End With

FIND = price(searchStr)
End Function



