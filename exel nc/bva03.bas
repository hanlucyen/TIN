Attribute VB_Name = "Module1"
Function CountGach(chungtu As Range) As Byte
        
For Each giatri In chungtu
    If Left(giatri, 1) = "G" Or Left(giatri, 1) = "X" Then
        CountGach = CountGach + 1
    End If
Next
        
End Function
