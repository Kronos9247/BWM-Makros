Attribute VB_Name = "Konto"
'Protected under "MIT License"
'
'Copyright (c) 2019 Rafael Orman
'
'
'You will find the full license on the github repo. "https://raw.githubusercontent.com/Kronos9247/BWM-Makros/master/LICENSE"


Const ErfolgL As Boolean = True 'Wenn Erfolg nicht benötigt dann True zu False
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


Public Type KontoLabel
    Label As String
    NumberFormat As String
    Width As Double
    
    Sum As Boolean
End Type
Sub KontoErstellen()
Attribute KontoErstellen.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' KontoErstellen Makro
'
' Tastenkombination: Strg+Umschalt+M
'
    Dim SumLine As Boolean
    Dim header() As KontoLabel
    header = Labels(ErfolgL)
    SumLine = SumL
    
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            Dim height, restHeight As Integer
            Dim startX As Integer, startY As Integer
            
            height = Selection.Rows.Count
            startX = Selection.Column
            startY = Selection.Row
            
            restHeight = height - 1
            
            MsgBox height
            
            If (SumLine And height >= 2) Or (Not SumLine And height >= 1) Then
                Dim offsetX, offsetY As Integer
                Dim curCells As Range
                Dim Item As KontoLabel
                
                
                Dim endX As Integer, endY As Integer
                endX = startX + (UBound(header) - LBound(header))
                endY = startY + height - 1
                
                'Select everything
                SelectRange startX, startY, endX, endY
                AddUndo
                
                Dim j As Integer, i As Integer
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
                
                'Apply undo for everything
                Application.OnUndo "Undo Macro - Konto", "KontoEntfernen"
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
    Dim labs() As KontoLabel
    If Not Erfolg Then
        ReDim labs(0 To 3)
    Else
        ReDim labs(0 To 4)
    End If
    
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
Private Sub KontoEntfernen()
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

