Attribute VB_Name = "PetrophysicalAnalysis"

'================================================================================================================================
' CALCULATES NUMEROUS PETROPHYSICAL CALCULATIONS AND DISPLAYS THEM ON A NEW SHEET
'================================================================================================================================
Sub PetrophysicsAnalysis()

    Dim Row As Long 'Row index for the .LAS file data sheet
    Dim LASLastRow As Long 'Row index for the last row of the .LAS file data sheet
    Dim AllSumPerm As Double 'Sum of permeabilities across the entire depth interval
    Dim PaySumPerm As Double 'Sum of permeabilities from all pay readings
    Dim ResSumPerm As Double 'Sum of permeabilities from all reservoir readings
    Dim AllSumKH As Double 'Sum of thickness-weighted permeabilities across the entire depth interval
    Dim PaySumKH As Double 'Sum of thickness-weighted permeabilities from all pay readings
    Dim ResSumKH As Double 'Sum of thickness-weighted permeabilities from all reservoir readings
    Dim AllSumPorosity As Double 'Sum of porosities across the entire depth interval
    Dim PaySumPorosity As Double 'Sum of porosities from all pay readings
    Dim ResSumPorosity As Double 'Sum of porosities from all reservoir readings
    Dim AllSumWaterSat As Double 'Sum of water saturations across the entire depth interval
    Dim PaySumWaterSat As Double 'Sum of water saturations from all pay readings
    Dim ResSumWaterSat As Double 'Sum of water saturations from all reservoir readings
    Dim AllNumEntries As Double 'Number of readings across the entire depth interval
    Dim PayNumEntries As Double 'Number of pay readings
    Dim ResNumEntries As Double 'Number of reservoir readings
    Dim AllSumThickness As Double 'Sum of thicknesses across the entire depth interval
    Dim PaySumThickness As Double 'Sum of thicknesses from all pay readings
    Dim ResSumThickness As Double 'Sum of thicknesses from all reservoir readings
    Dim AllAvgPerm As Double 'Average permeability across the entire depth interval
    Dim PayAvgPerm As Double 'Average permeability from all pay readings
    Dim ResAvgPerm As Double 'Average permeability from all reservoir readings
    Dim AllAvgKH As Double 'Average thickness-weighted permeability across the entire depth interval
    Dim PayAvgKH As Double 'Average thickness-weighted permeability from all pay readings
    Dim ResAvgKH As Double 'Average thickness-weighted permeability from all reservoir readings
    Dim AllAvgPorosity As Double 'Average porosity across the entire depth interval
    Dim PayAvgPorosity As Double 'Average porosity from all pay readings
    Dim ResAvgPorosity As Double 'Average porosity from all reservoir readings
    Dim AllAvgWaterSat As Double 'Average water saturation across the entire depth interval
    Dim PayAvgWaterSat As Double 'Average water saturation from all pay readings
    Dim ResAvgWaterSat As Double 'Average water saturation from all reservoir readings
    Dim Rect As Shape 'Rectangle title shape
    Dim RectTop As Double 'Rectangle title shape top
    Dim RectLeft As Double 'Rectangle title shape left
    Dim RectHeight As Double 'Rectangle title shape height
    Dim RectWidth As Double 'Rectangle title shape width
    
    Sheets.Add
    ActiveSheet.Name = "Petrophysical Analysis"
    
    Row = 5
    LASLastRow = Sheets(2).Cells(Rows.Count, "B").End(xlUp).Row
    
    Do While Row <= LASLastRow
        If Sheets(2).Cells(Row, "E").Value <> "N/A" Then
            AllSumPerm = AllSumPerm + CDbl(Sheets(2).Cells(Row, "E").Value)
        End If
        
        If Row <> LASLastRow Then
            AllSumKH = AllSumKH + (CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)) * _
                CDbl(Sheets(2).Cells(Row, "E").Value)
            AllSumThickness = AllSumThickness + CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)
        End If
        
        AllSumPorosity = AllSumPorosity + CDbl(Sheets(2).Cells(Row, "F").Value)
        AllSumWaterSat = AllSumWaterSat + CDbl(Sheets(2).Cells(Row, "H").Value)
        
        If Sheets(2).Cells(Row, "J").Interior.Color = RGB(0, 255, 0) Then
            PaySumPerm = PaySumPerm + CDbl(Sheets(2).Cells(Row, "E").Value)
            
            If Row <> LASLastRow Then
                PaySumKH = PaySumKH + (CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)) * _
                    CDbl(Sheets(2).Cells(Row, "E").Value)
                PaySumThickness = PaySumThickness + CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)
            End If
            
            PaySumPorosity = PaySumPorosity + CDbl(Sheets(2).Cells(Row, "F").Value)
            PaySumWaterSat = PaySumWaterSat + CDbl(Sheets(2).Cells(Row, "H").Value)
            
            PayNumEntries = PayNumEntries + 1
        End If
        
        If Sheets(2).Cells(Row, "K").Interior.Color = RGB(0, 255, 0) Then
            ResSumPerm = ResSumPerm + CDbl(Sheets(2).Cells(Row, "E").Value)
            
            If Row <> LASLastRow Then
                ResSumKH = ResSumKH + (CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)) * _
                    CDbl(Sheets(2).Cells(Row, "E").Value)
                ResSumThickness = ResSumThickness + CDbl(Sheets(2).Cells(Row + 1, "C").Value) - CDbl(Sheets(2).Cells(Row, "C").Value)
            End If
            
            ResSumPorosity = ResSumPorosity + CDbl(Sheets(2).Cells(Row, "F").Value)
            ResSumWaterSat = ResSumWaterSat + CDbl(Sheets(2).Cells(Row, "H").Value)
            
            ResNumEntries = ResNumEntries + 1
        End If
        
        AllNumEntries = AllNumEntries + 1
        Row = Row + 1
    Loop
    
    AllAvgPerm = AllSumPerm / AllNumEntries
    AllAvgKH = AllSumKH / AllSumThickness
    AllAvgPorosity = AllSumPorosity / AllNumEntries
    AllAvgWaterSat = AllSumWaterSat / AllNumEntries
    
    If PayNumEntries <> 0 Then 'Pay flags existed in .LAS file
        PayAvgPerm = PaySumPerm / PayNumEntries
        PayAvgKH = PaySumKH / PaySumThickness
        PayAvgPorosity = PaySumPorosity / PayNumEntries
        PayAvgWaterSat = PaySumWaterSat / PayNumEntries
    Else 'Pay flags did not exist in .LAS file
        PayAvgPerm = 0
        PayAvgKH = 0
        PayAvgPorosity = 0
        PayAvgWaterSat = 0
    End If
    
    If ResNumEntries <> 0 Then 'Res flags existed in .LAS file
        ResAvgPerm = ResSumPerm / ResNumEntries
        ResAvgKH = ResSumKH / ResSumThickness
        ResAvgPorosity = ResSumPorosity / ResNumEntries
        ResAvgWaterSat = ResSumWaterSat / ResNumEntries
    Else 'Res flags did not exist in .LAS file
        ResAvgPerm = 0
        ResAvgKH = 0
        ResAvgPorosity = 0
        ResAvgWaterSat = 0
    End If
    
    Range("A4:C4").Merge
    Cells(4, "A").Font.Bold = True
    Cells(4, "A").Value = "Average Permeability (mD)"
    
    Range("A5:B5").Merge
    Cells(5, "A").Font.Underline = True
    Cells(5, "A").Value = "Entire Depth Interval:"
    Cells(5, "C").Value = Format(AllAvgPerm, "#.00")
    
    Range("A6:B6").Merge
    Cells(6, "A").Font.Underline = True
    Cells(6, "A").Value = "Pay Only:"
    Cells(6, "C").Value = Format(PayAvgPerm, "#.00")
    
    Range("A7:B7").Merge
    Cells(7, "A").Font.Underline = True
    Cells(7, "A").Value = "Reservoir Only:"
    Cells(7, "C").Value = Format(ResAvgPerm, "#.00")
    
    With Range("A4:C7")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A4:C4")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Range("A9:C9").Merge
    Cells(9, "A").Font.Bold = True
    Cells(9, "A").Value = "Thickness-Weighted Permeability (mD)"
    
    Range("A10:B10").Merge
    Cells(10, "A").Font.Underline = True
    Cells(10, "A").Value = "Entire Depth Interval:"
    Cells(10, "C").Value = Format(AllAvgKH, "#.00")
    
    Range("A11:B11").Merge
    Cells(11, "A").Font.Underline = True
    Cells(11, "A").Value = "Pay Only:"
    Cells(11, "C").Value = Format(PayAvgKH, "#.00")
    
    Range("A12:B12").Merge
    Cells(12, "A").Font.Underline = True
    Cells(12, "A").Value = "Reservoir Only:"
    Cells(12, "C").Value = Format(ResAvgKH, "#.00")
    
    With Range("A9:C12")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A9:C9")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Range("A14:C14").Merge
    Cells(14, "A").Font.Bold = True
    Cells(14, "A").Value = "Average Porosity (Fraction)"
    
    Range("A15:B15").Merge
    Cells(15, "A").Font.Underline = True
    Cells(15, "A").Value = "Entire Depth Interval:"
    Cells(15, "C").Value = Format(AllAvgPorosity, "#.000")
    
    Range("A16:B16").Merge
    Cells(16, "A").Font.Underline = True
    Cells(16, "A").Value = "Pay Only:"
    Cells(16, "C").Value = Format(PayAvgPorosity, "#.000")
    
    Range("A17:B17").Merge
    Cells(17, "A").Font.Underline = True
    Cells(17, "A").Value = "Reservoir Only:"
    Cells(17, "C").Value = Format(ResAvgPorosity, "#.000")
    
    With Range("A14:C17")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A14:C14")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Range("A19:C19").Merge
    Cells(19, "A").Font.Bold = True
    Cells(19, "A").Value = "Average Porosity (%)"
    
    Range("A20:B20").Merge
    Cells(20, "A").Font.Underline = True
    Cells(20, "A").Value = "Entire Depth Interval:"
    Cells(20, "C").Value = Format(AllAvgPorosity * 100, "#.0")
    
    Range("A21:B21").Merge
    Cells(21, "A").Font.Underline = True
    Cells(21, "A").Value = "Pay Only:"
    Cells(21, "C").Value = Format(PayAvgPorosity * 100, "#.0")
    
    Range("A22:B22").Merge
    Cells(22, "A").Font.Underline = True
    Cells(22, "A").Value = "Reservoir Only:"
    Cells(22, "C").Value = Format(ResAvgPorosity * 100, "#.0")
    
    With Range("A19:C22")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A19:C19")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Range("A24:C24").Merge
    Cells(24, "A").Font.Bold = True
    Cells(24, "A").Value = "Average Water Saturation (Fraction)"
    
    Range("A25:B25").Merge
    Cells(25, "A").Font.Underline = True
    Cells(25, "A").Value = "Entire Depth Interval:"
    Cells(25, "C").Value = Format(AllAvgWaterSat, "#.000")
    
    Range("A26:B26").Merge
    Cells(26, "A").Font.Underline = True
    Cells(26, "A").Value = "Pay Only:"
    Cells(26, "C").Value = Format(PayAvgWaterSat, "#.000")
    
    Range("A27:B27").Merge
    Cells(27, "A").Font.Underline = True
    Cells(27, "A").Value = "Reservoir Only:"
    Cells(27, "C").Value = Format(ResAvgWaterSat, "#.000")
    
    With Range("A24:C27")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A24:C24")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Range("A29:C29").Merge
    Cells(29, "A").Font.Bold = True
    Cells(29, "A").Value = "Average Water Saturation (%)"
    
    Range("A30:B30").Merge
    Cells(30, "A").Font.Underline = True
    Cells(30, "A").Value = "Entire Depth Interval:"
    Cells(30, "C").Value = Format(AllAvgWaterSat * 100, "#.0")
    
    Range("A31:B31").Merge
    Cells(31, "A").Font.Underline = True
    Cells(31, "A").Value = "Pay Only:"
    Cells(31, "C").Value = Format(PayAvgWaterSat * 100, "#.0")
    
    Range("A32:B32").Merge
    Cells(32, "A").Font.Underline = True
    Cells(32, "A").Value = "Reservoir Only:"
    Cells(32, "C").Value = Format(ResAvgWaterSat * 100, "#.0")
    
    With Range("A29:C32")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    Range("E4:G4").Merge
    Cells(4, "E").Font.Bold = True
    Cells(4, "E").Value = "All Thicknesses (ft)"
    
    Range("E5:F5").Merge
    Cells(5, "E").Font.Underline = True
    Cells(5, "E").Value = "Entire Depth Interval:"
    Cells(5, "G").Value = Format(AllSumThickness, "#.00")
    
    Range("E6:F6").Merge
    Cells(6, "E").Font.Underline = True
    Cells(6, "E").Value = "Pay Only:"
    Cells(6, "G").Value = Format(PaySumThickness, "#.00")
    
    Range("E7:F7").Merge
    Cells(7, "E").Font.Underline = True
    Cells(7, "E").Value = "Reservoir Only:"
    Cells(7, "G").Value = Format(ResSumThickness, "#.00")
    
    With Range("E4:G7")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("E4:G4")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    With Range("A29:C29")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    Columns("A").ColumnWidth = 11.22
    Columns("B").ColumnWidth = 11.22
    Columns("C").ColumnWidth = 11.22
    Range("A4:C32").HorizontalAlignment = xlCenter
    Columns("E").ColumnWidth = 11.22
    Columns("F").ColumnWidth = 11.22
    Columns("G").ColumnWidth = 11.22
    Range("E4:G7").HorizontalAlignment = xlCenter
    
    With Range("A1:G2")
        RectLeft = .Left
        RectTop = .Top
        RectWidth = .Width
        RectHeight = .Height
    End With
    
    Set Rect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, RectLeft, RectTop, RectWidth, RectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, RectLeft, RectTop, RectWidth, RectHeight).Name = "Petrophysical Analysis"
    
    With ActiveSheet.Shapes("Petrophysical Analysis")
        .TextFrame.Characters.Text = "PETROPHYSICAL ANALYSIS"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With

End Sub
