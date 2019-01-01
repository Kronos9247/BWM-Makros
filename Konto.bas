Attribute VB_Name = "Konto"
'Protected under "MIT License"
'
'Copyright (c) 2019 Rafael Orman
'
'
'You will find the full license on the github repo. "https://raw.githubusercontent.com/Kronos9247/BWM-Makros/master/LICENSE"

Public Type KontoLabel
        Label As String
        NumberFormat As String
        Width As Double
        
        Sum As Boolean
End Type
Sub KontoErstellen()
'
' KontoErstellen Makro
'
' Tastenkombination: Strg+Umschalt+N
'
    Dim SumLine As Boolean
    Dim header() As KontoLabel
    header = Labels(False)   'Wenn Erfolg nicht benötigt dann True zu False
    SumLine = False          'Wenn die Summen-Zeile nicht benötigt wird das True zu einem False ändern!
    
    If TypeName(Selection) = "Range" Then
        MsgBox Selection.Address & " item(s) selected"
        
        If Selection.Areas.count = 1 Then
            Dim height, restHeight As Integer
            Dim startX As Integer, startY As Integer
            
            height = Selection.Rows.count
            startX = Selection.Column
            startY = Selection.Row
            
            restHeight = height - 1
            
            If (SumLine And height >= 2) Or (Not SumLine And height >= 1) Then
                Dim offsetX, offsetY As Integer
                Dim curCells As Range
                Dim Item As KontoLabel
                
                
                For i = LBound(header) To UBound(header)
                    offsetX = startX + i
                    Set curCells = ActiveSheet.Cells(startY, offsetX)
                    Item = header(i)

                    curCells.Select
                    curCells.Value = Item.Label
                    
                    For j = 0 To restHeight
                        offsetY = startY + j
                        ActiveSheet.Cells(offsetY, offsetX).Select
                        
                        If Item.Width > 0 Then ActiveCell.ColumnWidth = Item.Width
                        ActiveCell.NumberFormat = Item.NumberFormat
                    Next
                    
                    If SumLine And Item.Sum Then
                        ActiveCell.FormulaR1C1 = "=SUM(R[-" & (restHeight - 1) & "]C:R[-1]C)"
                    End If
                Next
                
                
                Dim endX As Integer, endY As Integer
                endX = startX + (UBound(header) - LBound(header))
                endY = startY + height - 1
                
                'Border Thin
                SelectRange startX, startY, endX, endY
                ApplyBorderThin
                
                'Border for the header
                SelectRange startX, startY, endX, startY
                ApplyBorderHeader
                
                If SumLine Then
                    SelectRange startX, endY, endX, endY
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                    End With
                End If
                
                'Select everything
                SelectRange startX, startY, endX, endY
            Else
                If SumLine Then
                    MsgBox "[BWM-Konto] Min height is 2!"
                Else
                    MsgBox "[BWM-Konto] Min height is 1!"
                End If
            End If
        End If
    Else
        MsgBox "[BWM-Konto] Nothing selected!"
    End If
End Sub
Public Function Labels(Erfolg As Boolean) As KontoLabel()
    Dim Length As Integer
    If Erfolg Then
        Length = 4
    Else
        Length = 3
    End If
    
    Dim labs(0 To 4) As KontoLabel
    labs(0) = DatumLabel()
    labs(1) = TextLabel()
    labs(2) = SollLabel()
    labs(3) = HabenLabel()
    
    If Erfolg Then labs(4) = ErfolgLabel()
    
    Labels = labs
End Function
Public Function DatumLabel() As KontoLabel
    Dim Obj As KontoLabel
    Obj.Label = "Datum"
    Obj.NumberFormat = "d/m;@"
    
    DatumLabel = Obj
End Function
Private Function TextLabel() As KontoLabel
    Dim Obj As KontoLabel
    Obj.Label = "Konto"
    Obj.Width = 24
    Obj.NumberFormat = "@"
    
    TextLabel = Obj
End Function
Private Function SollLabel() As KontoLabel
    Dim Obj As KontoLabel
    Obj.Label = "Soll"
    Obj.Width = 10.27
    Obj.NumberFormat = "#,##0.00 $"
    
    Obj.Sum = True 'Info for SumLine
    
    SollLabel = Obj
End Function
Private Function HabenLabel() As KontoLabel
    Dim Obj As KontoLabel
    Obj.Label = "Haben"
    Obj.Width = 10.27
    Obj.NumberFormat = "#,##0.00 $"
    
    Obj.Sum = True 'Info for SumLine
    
    HabenLabel = Obj
End Function
Private Function ErfolgLabel() As KontoLabel
    Dim Obj As KontoLabel
    Obj.Label = "Erfolg"
    Obj.Width = 7.55
    
    ErfolgLabel = Obj
End Function


Public Function SelectRange(startX As Integer, startY As Integer, endX As Integer, endY As Integer)
    With ActiveSheet
        .Range(.Cells(startY, startX), _
            .Cells(endY, endX)).Select
    End With
End Function
Public Function ApplyBorderThin()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Function
Public Function ApplyBorderHeader()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function
