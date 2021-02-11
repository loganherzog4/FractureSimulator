Attribute VB_Name = "EngineeringAnalysis"

'================================================================================================================================
' PERFORMS ENGINEERING ANALYSIS CALCULATIONS AND DISPLAYS THEM IN A NEW SHEET
'================================================================================================================================
Sub EngineeringsAnalysis()

    Dim GrossThickness As Double 'Gross thickness of the depth interval in ft TVD
    Dim ResThickness As Double 'Reservoir thickness in ft TVD
    Dim PayThickness As Double 'Pay thickness in ft TVD
    Dim Ri As Double 'Radius of investigation in ft
    Dim GridAreaFt2 As Double 'Reservoir simulation grid area in ft^2
    Dim GridAreaAcres As Double 'Reservoir simulation grid area in acres
    Dim Q As Double 'Surface flowrate in BBL/Day
    Dim Drawdown As Double 'Drawdown from drawdown equation in psi
    Dim ResAvgPerm As Double 'Average reservoir permeability in mD
    Dim Viscosity As Double 'Reservoir fluid viscosity in cP
    Dim FVF As Double 'Formation volume factor in Res BBL/STB
    Dim Re As Double 'Drainage radius in ft
    Dim Rw As Double 'Wellbore radius in ft
    Dim Sf As Double 'Skin factor from user
    Dim Row As Long 'Row index for parsing .LAS File Data sheet
    Dim LastRow As Long 'Last row index from .LAS File Data sheet
    Dim PIDrawdown As Double 'PI calculated from the drawdown equation
    Dim PIDarcy As Double 'PI calculated from the Darcy formulation
    Dim ResToGross As Double 'Reservoir to gross thickness ratio
    Dim PayToGross As Double 'Pay to gross thickness ratio
    Dim PayToRes As Double 'Pay to reservoir thickness ratio
    Dim Rect As Shape 'Title rectangle for new worksheet
    
    Sheets.Add
    ActiveSheet.Name = "Engineering Analysis"
    
    GrossThickness = Sheets(".LAS File Data").Cells(Sheets(".LAS File Data").Cells(Rows.Count, "C").End(xlUp).Row, "C").Value - _
        Sheets(".LAS File Data").Cells(5, "C").Value
    ResAvgPerm = Sheets("Petrophysical Analysis").Cells(7, "C").Value
    
    If FormExpressRun.OptEqn.Value = True Then
        Ri = CDbl(FormExpressRun.LblRiValue.Caption)
        GridAreaFt2 = Ri ^ 2
        GridAreaAcres = GridAreaFt2 * (1 / 43560)
    Else
        GridAreaAcres = 600
        GridAreaFt2 = GridAreaAcres * 43560
        Ri = Sqr(GridAreaFt2)
    End If
    
    Row = 5
    LastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    
    Do While Row < LastRow
        If Sheets(".LAS File Data").Cells(Row, "K").Interior.Color = RGB(0, 255, 0) Then
            ResThickness = ResThickness + Sheets(".LAS File Data").Cells(Row + 1, "C").Value - _
                Sheets(".LAS File Data").Cells(Row, "C").Value
        End If
        
        If Sheets(".LAS File Data").Cells(Row, "J").Interior.Color = RGB(0, 255, 0) Then
            PayThickness = PayThickness + Sheets(".LAS File Data").Cells(Row + 1, "C").Value - _
                Sheets(".LAS File Data").Cells(Row, "C").Value
        End If
        
        Row = Row + 1
    Loop
    
    ResToGross = ResThickness / GrossThickness
    PayToGross = PayThickness / GrossThickness
    
    If ResThickness <> 0 Then
        PayToRes = PayThickness / ResThickness
    Else
        PayToRes = 0
    End If
    
    If FormExpressRun.ChkProductivityIndex.Value = True Then
        Viscosity = CDbl(FormPIDarcy.TxtViscosity.Text)
        FVF = CDbl(FormPIDarcy.TxtFVF.Text)
        Re = CDbl(FormPIDarcy.TxtRe.Text)
        Rw = CDbl(FormPIDarcy.TxtRw.Text)
        Sf = CDbl(FormPIDarcy.TxtSf.Text)
        
        PIDarcy = (0.0078 * ResAvgPerm * ResThickness) / (Viscosity * FVF * (Log(Re / Rw) + Sf))
    End If
    
    Range("A4:C4").Merge
    Cells(4, "A").Font.Bold = True
    Cells(4, "A").Value = "Thickness Ratios (ft/ft)"
    
    Range("A5:B5").Merge
    Cells(5, "A").Font.Underline = True
    Cells(5, "A").Value = "Reservoir-To-Gross:"
    Cells(5, "c").Value = Format(ResToGross, "#.00")
    
    Range("A6:B6").Merge
    Cells(6, "A").Font.Underline = True
    Cells(6, "A").Value = "Pay-To-Gross:"
    Cells(6, "c").Value = Format(PayToGross, "#.00")
    
    Range("A7:B7").Merge
    Cells(7, "A").Font.Underline = True
    Cells(7, "A").Value = "Pay-To-Reservoir:"
    Cells(7, "c").Value = Format(PayToRes, "#.00")
    
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
    Cells(9, "A").Value = "Radius of Investigation"
    
    Range("A10:B10").Merge
    Cells(10, "A").Font.Underline = True
    Cells(10, "A").Value = "Ri (ft):"
    Cells(10, "c").Value = Format(Ri, "#.00")
    
    Range("A11:B11").Merge
    Cells(11, "A").Font.Underline = True
    Cells(11, "A").Value = "Grid Area (ft^2):"
    Cells(11, "c").Value = Format(GridAreaFt2, "#.00")
    
    Range("A12:B12").Merge
    Cells(12, "A").Font.Underline = True
    Cells(12, "A").Value = "Grid Area (acres):"
    Cells(12, "c").Value = Format(GridAreaAcres, "#.00")
    
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
    Cells(14, "A").Value = "Productivity Index"
    
    Range("A15:B15").Merge
    Cells(15, "A").Font.Underline = True
    Cells(15, "A").Value = "J_0 (Pre-Frac):"
    
    If FormExpressRun.ChkProductivityIndex.Value = True Then
        Cells(15, "C").Value = Format(PIDarcy, "#.00")
    Else
        Cells(15, "C").Value = "-"
    End If
    
    Range("A16:B16").Merge
    Cells(16, "A").Font.Underline = True
    Cells(16, "A").Value = "J (Post-Frac):"
    Cells(16, "C").Value = "-"
    
    Range("A17:B17").Merge
    Cells(17, "A").Font.Underline = True
    Cells(17, "A").Value = "FOI (J/J_0):"
    Cells(17, "C").Value = "-"
    
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
    
    With Range("A1:C2")
        RectLeft = .Left
        RectTop = .Top
        RectWidth = .Width
        RectHeight = .Height
    End With
    
    Set Rect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, RectLeft, RectTop, RectWidth, RectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, RectLeft, RectTop, RectWidth, RectHeight).Name = "Engineering Analysis"
    
    With ActiveSheet.Shapes("Engineering Analysis")
        .TextFrame.Characters.Text = "ENGINEERING ANALYSIS"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With

    Columns("A").ColumnWidth = 11.22
    Columns("B").ColumnWidth = 11.22
    Columns("C").ColumnWidth = 11.22
    Range("A4:C17").HorizontalAlignment = xlCenter

End Sub

'================================================================================================================================
' ENABLES AND SWITCHES TO THE DIRECTIONAL DATA PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub SwitchToHydraulicFracturePage()

    With FormExpressRun.MultiPageExpressRun
        .Pages(3).Enabled = True
        .Value = .Value + 1
    End With

End Sub

'================================================================================================================================
' DISABLES ALL CONTROLS IN THE "ENGINEERING ANALYSIS" PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub DisableAllEngineeringControls()

    With FormExpressRun
        .OptDefaultArea.Enabled = False
        .OptEqn.Enabled = False
        .ChkProductivityIndex.Enabled = False
        .BtnPIDarcy.Enabled = False
        .BtnEngineeringAnalysisContinue.Enabled = False
        .LblEngineeringAnalysisErrors.Enabled = False
        .LblEngineeringAnalysisError.Visible = False
    End With

End Sub
