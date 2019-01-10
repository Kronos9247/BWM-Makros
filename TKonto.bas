Attribute VB_Name = "TKonto"
'Protected under "MIT License"
'Copyright (c) 2019 Rafael Orman
'
'You will find the full license on the github repo. "https://raw.githubusercontent.com/Kronos9247/BWM-Makros/master/LICENSE"
'


'Const ErfolgL As Boolean = True 'Wenn Erfolg nicht benötigt dann True zu False
Const SumL As Boolean = False   'Wenn die Summen-Zeile nicht benötigt wird das True zu einem False ändern!


'Undo Code
Type SaveRange
    Val As Variant
    Addr As String
    Format As String
    Borders As Borders
End Type

Option Explicit
Public OldWorkbook As Workbook
Public OldSheet As Worksheet
Public OldSelection() As SaveRange


Sub TKontoErstellen()
Attribute TKontoErstellen.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' TKontoErstellen Makro
'
' Tastenkombination: Strg+Umschalt+T
'
    Dim SumLine As Boolean
    SumLine = SumL
    
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            Dim height As Integer, restHeight As Integer
            Dim startX As Integer, startY As Integer
            Dim endX As Integer, endY As Integer
            
            height = Selection.Rows.Count
            startX = Selection.Column
            startY = Selection.Row
            
            endX = startX + 1
            endY = startY + height - 1
            
            If (SumLine And height >= 2) Or (Not SumLine And height >= 1) Then
                SelectRange startX, startY, endX, endY
                AddUndo
            
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
                
                Dim i As Integer
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
                
                Application.OnUndo "Undo Macro - TKonto", "TKontoEntfernen"
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



Private Function AddUndo()
    ReDim OldSelection(Selection.Count)
    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    
    Dim i As Integer, Cell As Range, Edge As Variant
    i = 0
    For Each Cell In Selection
        i = i + 1
        OldSelection(i).Addr = Cell.Address
        OldSelection(i).Val = Cell.Formula
        OldSelection(i).Format = Cell.NumberFormat
        
        
        Set OldSelection(i).Borders = Cell.Borders
    Next Cell
End Function
Private Sub TKontoEntfernen()
    OldWorkbook.Activate
    OldSheet.Activate
    
    Dim i As Integer, Cell As Range
    For i = 1 To UBound(OldSelection)
        Set Cell = Range(OldSelection(i).Addr)
        
        Cell.Formula = OldSelection(i).Val
        Cell.NumberFormat = OldSelection(i).Format
        
        Dim Edge As Integer, OldBorder As Borders
        If Not OldSelection(i).Borders Is Nothing Then
            Set OldBorder = OldSelection(i).Borders
            
            For Edge = XlBordersIndex.xlDiagonalDown To XlBordersIndex.xlInsideHorizontal
                If Not OldBorder(Edge) Is Nothing Then
                    If Not OldBorder(Edge).LineStyle = 1 Then
                        If OldBorder(Edge).LineStyle = xlNone Then
                            Cell.Borders(Edge).LineStyle = OldBorder(Edge).LineStyle
                        Else
                            With Cell.Borders(Edge)
                                .LineStyle = OldBorder(Edge).LineStyle
                                .ColorIndex = OldBorder(Edge).ColorIndex
                                .TintAndShade = OldBorder(Edge).TintAndShade
                                .Weight = OldBorder(Edge).Weight
                            End With
                        End If
                    Else
                        With Cell.Borders(Edge)
                            .LineStyle = xlNone
                        End With
                    End If
                Else
                    With Cell.Borders(Edge)
                        .LineStyle = xlNone
                    End With
                End If
            Next Edge
        End If
    Next i
    
    Exit Sub
End Sub


