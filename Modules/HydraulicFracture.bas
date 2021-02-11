Attribute VB_Name = "HydraulicFracture"

'================================================================================================================================
' DISPLAYS THE HYDRAULIC FRACTURE PROPERTIES IN A NEW SHEET
'================================================================================================================================
Sub HydraulicFractures()
 
    Dim FracHL As Double 'Fracture half-length from user in ft
    Dim FracWidth As Double 'Average fracture width from user in inches
    Dim FracHeight As Double 'Fracture height from user in ft
    Dim FracTop As Double 'Fracture top depth from user in ft TVD
    Dim FracBase As Double 'Fracture base depth in ft TVD calculated from fracture top depth and fracture height
    Dim FCD As Double 'Dimensionless fracture conductivity from user
    Dim ResAvgPerm As Double 'Average reservoir permeability from petrophysical analysis
    Dim FracPerm As Double 'Fracture permeability calculated from Fcd and above fracture parameters
    Dim Rect As Shape 'Title rectangle
    
    With FormExpressRun
        FracHL = CDbl(.TxtHL.Text)
        FracWidth = CDbl(.TxtFracWidth.Text) / 12
        FracHeight = CDbl(.TxtFracHeight.Text)
        FracTop = CDbl(.TxtFracTop.Text)
        FCD = CDbl(.TxtFcd.Text)
    End With
    
    FracBase = FracTop + FracHeight
    ResAvgPerm = Sheets("Petrophysical Analysis").Cells(7, "C").Value
    FracPerm = (FCD * ResAvgPerm * FracHL) / FracWidth
    
    Sheets.Add
    ActiveSheet.Name = "Hydraulic Fracture"
    
    Range("A4:C4").Merge
    Cells(4, "A").Font.Bold = True
    Cells(4, "A").Value = "Fracture Properties"
    
    Range("A5:B5").Merge
    Cells(5, "A").Font.Underline = True
    Cells(5, "A").Value = "Half-Length (ft):"
    Cells(5, "C").Value = FracHL
    
    Range("A6:B6").Merge
    Cells(6, "A").Font.Underline = True
    Cells(6, "A").Value = "Average Width (in):"
    Cells(6, "C").Value = FracWidth * 12
    
    Range("A7:B7").Merge
    Cells(7, "A").Font.Underline = True
    Cells(7, "A").Value = "Height (ft):"
    Cells(7, "C").Value = FracHeight
    
    Range("A8:B8").Merge
    Cells(8, "A").Font.Underline = True
    Cells(8, "A").Value = "Top Depth (ft TVD):"
    Cells(8, "C").Value = FracTop
    
    Range("A9:B9").Merge
    Cells(9, "A").Font.Underline = True
    Cells(9, "A").Value = "Base Depth (ft TVD):"
    Cells(9, "C").Value = FracBase
    
    Range("A10:B10").Merge
    Cells(10, "A").Font.Underline = True
    Cells(10, "A").Value = "Permeability (mD):"
    Cells(10, "C").Value = FracPerm
    
    Range("A12:C12").Merge
    Cells(12, "A").Font.Bold = True
    Cells(12, "A").Value = "Fracture Effects"
    
    Range("A13:B13").Merge
    Cells(13, "A").Font.Underline = True
    Cells(13, "A").Value = "Fracture Skin:"
    
    If FormExpressRun.ChkProductivityIndex.Value = True Then
        Cells(13, "C").Value = ((1.65 - 0.328 * Log(FCD) + 0.116 * (Log(FCD)) ^ 2) / (1 + 0.18 * Log(FCD) + 0.064 * (Log(FCD)) ^ 2 + _
            0.005 * (Log(FCD)) ^ 3)) - Log(FracHL / CDbl(FormPIDarcy.TxtRw.Text))
    Else
        Cells(13, "C").Value = "-"
    End If
        
    Range("A14:B14").Merge
    Cells(14, "A").Font.Underline = True
    Cells(14, "A").Value = "Effective Wellbore Radius (ft):"
    
    If FormExpressRun.ChkProductivityIndex.Value = True Then
        Cells(14, "C").Value = CDbl(FormPIDarcy.TxtRw.Text) * Exp(-1 * CDbl(Cells(13, "C").Value))
    Else
        Cells(14, "C").Value = "-"
    End If
    
    With Range("A4:C10")
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
    
    With Range("A12:C14")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A12:C12")
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
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, RectLeft, RectTop, RectWidth, RectHeight).Name = "Hydraulic Fracture"
    
    With ActiveSheet.Shapes("Hydraulic Fracture")
        .TextFrame.Characters.Text = "HYDRAULIC FRACTURE"
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
    
    If FormExpressRun.ChkProductivityIndex.Value = True Then
        Sheets("Engineering Analysis").Cells(16, "C").Value = (0.0078 * CDbl(Sheets("Petrophysical Analysis").Cells(7, "C").Value) * _
            CDbl(Sheets("Petrophysical Analysis").Cells(7, "G").Value)) / (CDbl(FormPIDarcy.TxtViscosity.Text) * _
            CDbl(FormPIDarcy.TxtFVF.Text) * (Log(CDbl(FormPIDarcy.TxtRe.Text) / CDbl(FormPIDarcy.TxtRw)) + _
            CDbl(Cells(13, "C").Value)))
    
        Sheets("Engineering Analysis").Cells(17, "C").Value = CDbl(Sheets("Engineering Analysis").Cells(16, "C").Value) / _
            CDbl(Sheets("Engineering Analysis").Cells(15, "C").Value)
    End If

End Sub

'================================================================================================================================
' DISABLES ALL CONTROLS IN THE "HYDRAULIC FRACTURE" PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub DisableFracturePage()

    With FormExpressRun
        
        .LblHL.Enabled = False
        .TxtHL.Enabled = False
        .TxtHL.BackColor = vbButtonFace
        .LblFracWidth.Enabled = False
        .TxtFracWidth.Enabled = False
        .TxtFracWidth.BackColor = vbButtonFace
        .LblFracHeight.Enabled = False
        .TxtFracHeight.Enabled = False
        .TxtFracHeight.BackColor = vbButtonFace
        .LblFracTop.Enabled = False
        .TxtFracTop.Enabled = False
        .TxtFracTop.BackColor = vbButtonFace
        .LblFcd.Enabled = False
        .TxtFcd.Enabled = False
        .TxtFcd.BackColor = vbButtonFace
        .BtnFractureContinue.Enabled = False
        .LblFractureErrors.Enabled = False
        .LblFractureError.Visible = False
        
    End With

End Sub

'================================================================================================================================
' ENABLES AND SWITCHES TO THE RESERVOIR SIMULATION PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub SwitchToSimulationPage()

    With FormExpressRun.MultiPageExpressRun
        .Pages(4).Enabled = True
        .Value = .Value + 1
    End With

End Sub
