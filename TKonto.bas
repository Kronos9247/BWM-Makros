Attribute VB_Name = "TKonto"
'Protected under "MIT License"
'Copyright (c) 2019 Rafael Orman
'
'You will find the full license on the github repo. "https://raw.githubusercontent.com/Kronos9247/BWM-Makros/master/LICENSE"
'

Sub TKontoErstellen()
'
' TKontoErstellen Makro
'
' Tastenkombination: Strg+Umschalt+T
'
    Dim SumLine As Boolean
    SumLine = False 'Wenn die Summen-Zeile nicht benötigt wird das True zu einem False ändern!
    
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.count = 1 Then
            Dim height As Integer, restHeight As Integer
            Dim startX As Integer, startY As Integer
            Dim endX As Integer, endY As Integer
            
            height = Selection.Rows.count
            startX = Selection.Column
            startY = Selection.Row
            
            endX = startX + 1
            endY = startY + height - 1
            
            If (SumLine And height >= 2) Or (Not SumLine And height >= 1) Then
                'Select header soll
                SelectRange startX, startY, startX, startY
                ActiveCell.Value = "Soll"
                
                'Select header haben
                SelectRange endX, startY, endX, startY
                ActiveCell.Value = "Haben"
                
                
                'Select header
                SelectRange startX, startY, endX, startY
                Selection.HorizontalAlignment = xlCenter
                With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                End With
                
                
                'Select body
                SelectRange startX, startY + 1, endX, endY
                Selection.NumberFormat = "#,##0.00 $"
                
                If SumLine Then
                    For i = 0 To 1
                        ActiveSheet.Cells(endY, startX + i).Select
                        ActiveCell.FormulaR1C1 = "=SUM(R[-" & (height - 2) & "]C:R[-1]C)"
                        
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlDouble
                            .Weight = xlThick
                        End With
                    Next
                End If
                
                
                
                'Select everything
                SelectRange startX, startY, endX, endY
                With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                End With
            Else
                If SumLine Then
                    MsgBox "[BWM-Macro] Min height is 2!"
                Else
                    MsgBox "[BWM-Macro] Min height is 1!"
                End If
            End If
        End If
    Else
        MsgBox "[BWM-Macro] Nothing selected!"
    End If
End Sub
Public Function SelectRange(startX As Integer, startY As Integer, endX As Integer, endY As Integer)
    With ActiveSheet
        .Range(.Cells(startY, startX), _
            .Cells(endY, endX)).Select
    End With
End Function
