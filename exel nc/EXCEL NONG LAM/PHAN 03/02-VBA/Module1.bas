Attribute VB_Name = "Module1"
Function ThueGTGT(mahang As String, soluong As Double, dongia As Double) As Double
                                    
    If Left(mahang, 1) = "A" And soluong > 1000 Then
        ThueGTGT = soluong * dongia * 0.1
    ElseIf Left(mahang, 1) = "B" Then
        ThueGTGT = soluong * dongia * 0.05
    Else
        ThueGTGT = soluong * dongia * 0.03
    End If
                                    
End Function
