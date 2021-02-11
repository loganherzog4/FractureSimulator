Attribute VB_Name = "ReservoirSimulation"

'================================================================================================================================================
' RETURNS THE FILE PATH THE USER SELECTS FROM THE FILE EXPLORER
'================================================================================================================================================
Function GetFilePath(StartingPath As String, NonFrac As Boolean, Frac As Boolean) As String
    
    Dim Folder As FileDialog
    Dim SelectedPath As String
    
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Folder
        .Title = "Select a File Path"
        .AllowMultiSelect = False
        .InitialFileName = StartingPath
        
        If .Show = True Then
          FileName = .SelectedItems(1)
          
          If Right(FileName, 1) = "\" Then
            GetFilePath = FileName & StrLegalFileName(InputBox("Please enter your desired file name.", "Grid File Name"))
          Else
            GetFilePath = FileName & "\" & StrLegalFileName(InputBox("Please enter your desired file name.", "Grid File Name"))
          End If
          
          If NonFrac = True Then
            FormExpressRun.TxtNonFracturedGridPath.BackColor = vbButtonFace
          End If
          
          If Frac = True Then
            FormExpressRun.TxtFracturedGridPath.BackColor = vbButtonFace
          End If
        Else
          MsgBox "A file was not selected. The macro has been terminated.", vbCritical, "Reservoir Simulation"
        End If
    End With

End Function

'================================================================================================================================================
' RETURNS THE FILE PATH THE USER SELECTS FROM THE FILE EXPLORER
'================================================================================================================================================
Function StrLegalFileName(StrFileNameIn As String) As String
    
    Dim i As Integer 'Count variable for loops
    Const StrIllegals = "\/|?*<>"":" 'Illegal file name characters
    
    StrLegalFileName = StrFileNameIn
    
    For i = 1 To Len(StrIllegals)
        StrLegalFileName = Replace(StrLegalFileName, Mid$(StrIllegals, i, 1), "_")
    Next i
    
End Function

'================================================================================================================================
' GENERATES DX AND DY PROFILES FOR THE GRID AND A VISUAL TOP VIEW IN A NEW SHEET
'================================================================================================================================
Sub GenerateCrossSections()

    Dim Rw As Double 'Wellbore radius in feet
    Dim FracHL As Double 'Fracture half-length in feet
    Dim FracWidth As Double 'Average fracture width in inches
    Dim FracHeight As Double 'Fracture height in feet
    Dim FracTop As Double 'Depth value at the top of the fracture in feet
    Dim FracBase As Double 'Depth value at the bottom of the fracture in feet
    Dim FracPerm As Double 'Induced permeability along the fracture
    Dim FibonacciCoeffs() As Long 'Array of fibonacci coefficients
    Dim i As Integer 'Count variable for loops
    Dim j As Integer 'Count variable for loops
    Dim Row As Integer 'Row index for loops
    Dim dYProfileRow As Integer 'Row index for "dY Profile" text
    Dim Column As Integer 'Column index for loops
    Dim dXs() As Double 'Array of altered dX values
    Dim dYs() As Double 'Array of altered dY values
    Dim LastXIndex As Integer 'Value indicating the last index of the dXs array
    Dim LastYIndex As Integer 'Value indicating the last index of the dYs array
    Dim GridXDim As Integer 'Sim grid X dimensions
    Dim GridYDim As Integer 'Sim grid Y dimensions
    Dim Min As Double 'The minimum dX or dY value
    Dim Max As Double 'The maximum dX or dY value
    Dim TopRect As Shape 'Grid Top View title rectangle
    Dim TopRectLeft As Double 'Grid Top View title rectangle left
    Dim TopRectTop As Double 'Grid Top View title rectangle top
    Dim TopRectWidth As Double 'Grid Top View title rectangle width
    Dim TopRectHeight As Double 'Grid Top View title rectangle height
    Dim XRect As Shape 'Grid Top View X axis rectangle
    Dim XRectLeft As Double 'Grid Top View X axis rectangle left
    Dim XRectTop As Double 'Grid Top View X axis rectangle top
    Dim XRectWidth As Double 'Grid Top View X axis rectangle width
    Dim XRectHeight As Double 'Grid Top View X axis rectangle height
    Dim YRect As Shape 'Grid Top View Y axis rectangle
    Dim YRectLeft As Double 'Grid Top View Y axis rectangle left
    Dim YRectTop As Double 'Grid Top View Y axis rectangle top
    Dim YRectWidth As Double 'Grid Top View Y axis rectangle width
    Dim YRectHeight As Double 'Grid Top View Y axis rectangle height
    Dim SumDXs As Double 'Sum of the current parsed dX values
    Dim LastFracIndex As Integer 'dX index that the fracture half-length goes to
    
    Rw = CDbl(FormExpressRun.TxtRw.Text)
    
    With Sheets("Hydraulic Fracture")
        FracHL = .Cells(5, "C").Value
        FracWidth = .Cells(6, "C").Value / 12
        FracHeight = .Cells(7, "C").Value
        FracTop = .Cells(8, "C").Value
        FracBase = .Cells(9, "C").Value
        FracPerm = .Cells(10, "C").Value
    End With
    
    ReDim FibonacciCoeffs(40)
    ReDim dXs(1000)
    ReDim dYs(1000)
        
    i = 0
    
    Do While i <= UBound(FibonacciCoeffs)
        If i = 0 Then
            FibonacciCoeffs(i) = 0
        ElseIf i = 1 Then
            FibonacciCoeffs(i) = 1
        Else
            FibonacciCoeffs(i) = FibonacciCoeffs(i - 1) + FibonacciCoeffs(i - 2)
        End If
        
        i = i + 1
    Loop
    
    If FracWidth < 5 Then
        dYs(0) = 5
    Else
        dYs(0) = FracWidth
    End If
    
    i = 1
            
    Do While WorksheetFunction.Sum(dYs) <= Sheets("Engineering Analysis").Cells(10, "C").Value
        If WorksheetFunction.Sum(dYs) + FibonacciCoeffs(i + 1) * dYs(0) + FibonacciCoeffs(i + 2) * dYs(0) <= _
            Sheets("Engineering Analysis").Cells(10, "C").Value Then
            
            dYs(i) = FibonacciCoeffs(i + 1) * dYs(0)
        Else
            Dim SumY As Double
            
            SumY = WorksheetFunction.Sum(dYs)
            
            dYs(i) = Sheets("Engineering Analysis").Cells(10, "C").Value - SumY
            
            Exit Do
        End If
        
        i = i + 1
    Loop
    
    i = 0
    
    Do While i <= UBound(dYs)
        If dYs(i) = 0 Then
            LastYIndex = i - 1
            Exit Do
        End If
        
        i = i + 1
    Loop
    
    i = 0
    
    Do While i <= LastYIndex
        dXs(i) = dYs(i)
        
        i = i + 1
    Loop
    
    LastXIndex = LastYIndex
    
    FracHL = CDbl(Sheets("Hydraulic Fracture").Cells(5, "C").Value)
    SumDXs = 5
    
    i = 1
    
    Do While i <= LastXIndex
        SumDXs = SumDXs + dXs(i)
        
        If SumDXs = FracHL Then
            LastFracIndex = i
            Exit Do
        End If
        
        If SumDXs + dXs(i + 1) > FracHL Then
            If SumDXs + dXs(i + 1) - FracHL <= dXs(i + 1) / 2 Then
                LastFracIndex = i + 1
                
                Dim AmountSubtracted As Double
                
                AmountSubtracted = ((SumDXs + dXs(i + 1)) - FracHL)
                
                dXs(i + 1) = dXs(i + 1) - AmountSubtracted + dXs(0)
                dYs(i + 1) = dXs(i + 1)
                
                dXs(LastXIndex) = dXs(LastXIndex) + AmountSubtracted
                dYs(LastYIndex) = dXs(LastXIndex)
                Exit Do
            Else
                LastFracIndex = i
                
                Dim AmountAdded As Double
                
                AmountAdded = (FracHL - SumDXs)
                
                dXs(i) = dXs(i) + AmountAdded - dXs(0)
                dYs(i) = dXs(i)
                
                dXs(LastXIndex) = dXs(LastXIndex) - AmountAdded
                dYs(LastYIndex) = dXs(LastXIndex)
                Exit Do
            End If
        End If
        
        i = i + 1
    Loop

    GridXDim = LastXIndex * 2 + 1
    GridYDim = LastYIndex * 2 + 1
    
    Sheets.Add
    ActiveSheet.Name = "Grid Statistics"
    
    Range("A4:C4").Merge
    Cells(4, "A").Font.Bold = True
    Cells(4, "A").Value = "Aerial Dimensions"
    
    Range("A5:B5").Merge
    Cells(5, "A").Font.Underline = True
    Cells(5, "A").Value = "X Dimensions (Grid Blocks):"
    Cells(5, "C").Value = GridXDim
    
    Range("A6:B6").Merge
    Cells(6, "A").Font.Underline = True
    Cells(6, "A").Value = "Total X Distance (ft):"
    Cells(6, "C").Value = 2 * Sheets("Engineering Analysis").Cells(10, "C").Value
    
    Range("A7:B7").Merge
    Cells(7, "A").Font.Underline = True
    Cells(7, "A").Value = "Well X (Coordinate):"
    Cells(7, "C").Value = LastXIndex + 1
    
    Range("A8:B8").Merge
    Cells(8, "A").Font.Underline = True
    Cells(8, "A").Value = "Y Dimensions (Grid Blocks):"
    Cells(8, "C").Value = GridYDim
    
    Range("A9:B9").Merge
    Cells(9, "A").Font.Underline = True
    Cells(9, "A").Value = "Total Y Distance (ft):"
    Cells(9, "C").Value = 2 * Sheets("Engineering Analysis").Cells(10, "C").Value
    
    Range("A10:B10").Merge
    Cells(10, "A").Font.Underline = True
    Cells(10, "A").Value = "Well Y (Coordinate):"
    Cells(10, "C").Value = LastYIndex + 1
    
    Range("A12:C12").Merge
    Cells(12, "A").Font.Bold = True
    Cells(12, "A").Value = "dX Profile"
    
    Range("A13:B13").Merge
    Cells(13, "A").Font.Underline = True
    Cells(13, "A").Value = "X Coordinate"
    Cells(13, "C").Font.Underline = True
    Cells(13, "C").Value = "dX (ft)"
    
    Row = 14
    i = LastXIndex
    j = 1
    
    Do While i >= 0
        Cells(Row, "A").Value = j
        Range("A" & Row & ":B" & Row).Merge
        
'        If i = 0 Then
'            Cells(Row, "C").Value = 10
'        Else
'            Cells(Row, "C").Value = dXs(i)
'        End If

        Cells(Row, "C").Value = dXs(i)
        
        If i <= LastFracIndex And i <> 0 Then
            Cells(Row, "A").Interior.Color = RGB(0, 255, 255)
            Cells(Row, "C").Interior.Color = RGB(0, 255, 255)
        End If
    
        Row = Row + 1
        j = j + 1
        i = i - 1
    Loop
    
    Row = 14 + LastXIndex + 1
    i = 1
    j = LastXIndex + 2
    
    Do While i <= LastXIndex
        Cells(Row, "A").Value = j
        Range("A" & Row & ":B" & Row).Merge
        Cells(Row, "C").Value = dXs(i)
        
        If i <= LastFracIndex Then
            Cells(Row, "A").Interior.Color = RGB(0, 255, 255)
            Cells(Row, "C").Interior.Color = RGB(0, 255, 255)
        End If
    
        Row = Row + 1
        j = j + 1
        i = i + 1
    Loop
    
    With Range("A12:C" & Row - 1)
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
    
    Row = 14 + LastXIndex + 1 + LastXIndex + 1
    
    Range("A" & Row & ":C" & Row).Merge
    Cells(Row, "A").Font.Bold = True
    Cells(Row, "A").Value = "dY Profile"
    
    dYProfileRow = Row
    
    Row = Row + 1
    
    Range("A" & Row & ":B" & Row).Merge
    Cells(Row, "A").Font.Underline = True
    Cells(Row, "A").Value = "Y Coordinate"
    Cells(Row, "C").Font.Underline = True
    Cells(Row, "C").Value = "dY (ft)"
    
    Row = Row + 1
    i = LastYIndex
    j = 1
    
    Do While i >= 0
        Cells(Row, "A").Value = j
        Range("A" & Row & ":B" & Row).Merge
        
        If i = 0 Then
            'Cells(Row, "C").Value = 10
            Cells(Row, "C").Value = dYs(i)
            Cells(Row, "A").Interior.Color = RGB(0, 255, 255)
            Cells(Row, "C").Interior.Color = RGB(0, 255, 255)
        Else
            Cells(Row, "C").Value = dYs(i)
        End If
        
        Row = Row + 1
        j = j + 1
        i = i - 1
    Loop
    
    Row = 14 + LastXIndex + 1 + LastXIndex + 2 + LastYIndex + 2
    i = 1
    j = LastYIndex + 2
    
    Do While i <= LastYIndex
        Cells(Row, "A").Value = j
        Range("A" & Row & ":B" & Row).Merge
        Cells(Row, "C").Value = dYs(i)
        
        Row = Row + 1
        j = j + 1
        i = i + 1
    Loop
    
    With Range("A" & dYProfileRow & ":C" & Row - 1)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A" & dYProfileRow & ":C" & dYProfileRow)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
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
    
    Columns("A").ColumnWidth = 11.22
    Columns("B").ColumnWidth = 11.22
    Columns("C").ColumnWidth = 11.22
    Range("A4:C" & Row - 1).HorizontalAlignment = xlCenter
    
    ReDim dXs(LastXIndex * 2 + 1)
    ReDim dYs(LastYIndex * 2 + 1)
    
    Row = 14
    i = 1
    
    Do While Row <= 14 + UBound(dXs) - 1
        dXs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Row = Row + 3
    i = 1
    
    Do While Row <= 14 + UBound(dXs) + 2 + UBound(dYs)
        dYs(i) = Cells(Row, "C").Value
    
        Row = Row + 1
        i = i + 1
    Loop
    
    Min = dXs(1)
    Max = dXs(1)
    
    For i = 2 To UBound(dXs)
        If dXs(i) < Min Then
            Min = dXs(i)
        End If
        
        If dXs(i) > Max Then
            Max = dXs(i)
        End If
    Next i
    
    For i = 1 To UBound(dYs)
        If dYs(i) < Min Then
            Min = dYs(i)
        End If
        
        If dYs(i) > Max Then
            Max = dYs(i)
        End If
    Next i
        
    Sheets.Add
    ActiveSheet.Name = "Grid Top View"
    
    Row = 4
    Column = 1
    i = 1
    j = 1

    Do While Row <= GridYDim + 3
        Rows(Row).RowHeight = GetExcelDim(dYs(i), Min, Max)

        Do While Column <= GridXDim
            Columns(Column).ColumnWidth = GetExcelDim(dXs(j), Min, Max) * (125 / 682)

            With Cells(Row, Column)
                .Interior.Color = Sheets(".LAS File Data").Cells(5, "K").Interior.Color
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
            End With

            If j = LastXIndex * 2 + 1 Then
                Exit Do
            End If

            Column = Column + 1
            j = j + 1
        Loop

        Column = 1
        Row = Row + 1
        i = Row - 3
        j = 1
    Loop
    
    With Range(Cells(1, "A"), Cells(2, GridXDim))
        TopRectLeft = .Left
        TopRectTop = .Top
        TopRectWidth = .Width
        TopRectHeight = .Height
    End With
    
    Set TopRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, TopRectLeft, TopRectTop, TopRectWidth, TopRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, TopRectLeft, TopRectTop, TopRectWidth, TopRectHeight).Name = "Grid Top View"
    
    With ActiveSheet.Shapes("Grid Top View")
        .TextFrame.Characters.Text = "GRID TOP VIEW"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    With Range(Cells(GridYDim + 5, "A"), Cells(GridYDim + 6, GridXDim))
        XRectLeft = .Left
        XRectTop = .Top
        XRectWidth = .Width
        XRectHeight = .Height
    End With
    
    Set XRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, XRectLeft, XRectTop, XRectWidth, XRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, XRectLeft, XRectTop, XRectWidth, XRectHeight).Name = "X-Axis"
    
    With ActiveSheet.Shapes("X-Axis")
        .TextFrame.Characters.Text = "X-AXIS"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    With Range(Cells(4, GridXDim + 2), Cells(GridYDim + 3, GridXDim + 2))
        YRectLeft = .Left
        YRectTop = .Top
        YRectWidth = .Width
        YRectHeight = .Height
    End With
    
    Set YRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, YRectLeft, YRectTop, YRectWidth, YRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, YRectLeft, YRectTop, YRectWidth, YRectHeight).Name = "Y-Axis"
    
    With ActiveSheet.Shapes("Y-Axis")
        .TextFrame.Characters.Text = "Y-AXIS"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    Sheets.Add
    ActiveSheet.Name = "Cross-Section (Without Frac)"
    
    With FormExpressRun
        If .OptRoughScaleGrid.Value = True Then
            GenerateRoughScaleCrossSection GridXDim
            GenerateLayerProperties GridXDim
            GenerateRoughScaleCrossSectionFrac GridXDim, LastFracIndex
            GenerateLayerPropertiesFrac GridXDim, LastFracIndex
        ElseIf .OptFineScaleGrid.Value = True Then
            GenerateFineScaleCrossSection GridXDim
            GenerateLayerProperties GridXDim
            GenerateFineScaleCrossSectionFrac GridXDim, LastFracIndex
            GenerateLayerPropertiesFrac GridXDim, LastFracIndex
        Else
            GenerateUpScaledCrossSection GridXDim
            GenerateLayerProperties GridXDim
            GenerateUpScaledCrossSectionFrac GridXDim, LastFracIndex
            GenerateLayerPropertiesFrac GridXDim, LastFracIndex
        End If
    End With
    
    Sheets("Fracture Simulator").Move Before:=Sheets(1)
    Sheets(".LAS File Data").Move Before:=Sheets(2)
    Sheets("Petrophysical Analysis").Move Before:=Sheets(3)
    Sheets("Engineering Analysis").Move Before:=Sheets(4)
    Sheets("Hydraulic Fracture").Move Before:=Sheets(5)
    Sheets("Grid Statistics").Move Before:=Sheets(6)
    Sheets("Grid Top View").Move Before:=Sheets(7)
    Sheets("Cross-Section (Without Frac)").Move Before:=Sheets(8)
    Sheets("Cross-Section (With Frac)").Move Before:=Sheets(9)
    
    UpdateGridStats GridXDim
    GenerateNonFracInclude
    GenerateFracInclude LastFracIndex
    
End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR A ROUGH SCALE MODEL WITHOUT A FRACTURE
'================================================================================================================================
Private Sub GenerateRoughScaleCrossSection(ByVal XDim As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR A ROUGH SCALE MODEL WITH A FRACTURE
'================================================================================================================================
Private Sub GenerateRoughScaleCrossSectionFrac(ByVal XDim As Integer, ByVal LastFrac As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    Dim FracTop As Double 'Frac top depth
    Dim FracBase As Double 'Frac base depth
    Dim MiddleColumn As Integer 'Middle column of the cross-section
    
    Sheets.Add
    ActiveSheet.Name = "Cross-Section (With Frac)"
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    FracTop = CDbl(FormExpressRun.TxtFracTop.Text)
    FracBase = FracTop + CDbl(FormExpressRun.TxtFracHeight.Text)
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    ThisRow = 5
    LASRow = 6
    
    Do While ThisRow <= LASLastRow - 1
        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) > FracTop And _
            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "C").Value) < FracTop Then
            
            If Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            While CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) < FracBase
                MiddleColumn = XDim \ 2 + 1
                
                For i = 1 To LastFrac
                    Cells(ThisRow, MiddleColumn + i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn - i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn).Interior.Color = RGB(0, 255, 255)
                Next i
                
                LASRow = LASRow + 1
                ThisRow = ThisRow + 1
            Wend
            
            If Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            Exit Do
            
        End If
        
        LASRow = LASRow + 1
        ThisRow = ThisRow + 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR A FINE SCALE MODEL WITHOUT A FRACTURE
'================================================================================================================================
Private Sub GenerateFineScaleCrossSection(ByVal XDim As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        If Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR A FINE SCALE MODEL WITH A FRACTURE
'================================================================================================================================
Private Sub GenerateFineScaleCrossSectionFrac(ByVal XDim As Integer, ByVal LastFrac As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    Dim FracTop As Double 'Frac top depth
    Dim FracBase As Double 'Frac base depth
    Dim MiddleColumn As Integer 'Middle column of the cross-section
    
    Sheets.Add
    ActiveSheet.Name = "Cross-Section (With Frac)"
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    FracTop = CDbl(FormExpressRun.TxtFracTop.Text)
    FracBase = FracTop + CDbl(FormExpressRun.TxtFracHeight.Text)
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        If Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    ThisRow = 5
    LASRow = 6
    
    Do While ThisRow <= LASLastRow - 1
        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) > FracTop And _
            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "C").Value) < FracTop Then
            
            If Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            While CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) < FracBase
                MiddleColumn = XDim \ 2 + 1
                
                For i = 1 To LastFrac
                    Cells(ThisRow, MiddleColumn + i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn - i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn).Interior.Color = RGB(0, 255, 255)
                Next i
                
                LASRow = LASRow + 1
                ThisRow = ThisRow + 1
            Wend
            
            If Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            Exit Do
            
        End If
        
        LASRow = LASRow + 1
        ThisRow = ThisRow + 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR AN UP-SCALED MODEL WITHOUT A FRACTURE
'================================================================================================================================
Private Sub GenerateUpScaledCrossSection(ByVal XDim As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    Dim PermTol As Double 'Numeric or percentage permeability tolerance
    Dim PoroTol As Double 'Numeric or percentage permeability tolerance
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    
    If FormUpScaledGrid.TxtPermTolValue.Enabled = True Then
        PermTol = CDbl(FormUpScaledGrid.TxtPermTolValue.Text)
    End If
    
    If FormUpScaledGrid.TxtPoroTolValue.Enabled = True Then
        PoroTol = CDbl(FormUpScaledGrid.TxtPoroTolValue.Text)
    End If
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        If Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
            If FormUpScaledGrid.OptPermTol.Value = True Then
                If FormUpScaledGrid.OptPermNumeric.Value = True Then
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                Else
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                End If
            End If
            
            If FormUpScaledGrid.OptPoroTol.Value = True Then
                If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                Else
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                End If
            End If
            
            If FormUpScaledGrid.OptBothTol.Value = True Then
                If FormUpScaledGrid.OptPermNumeric.Value = True Then
                    If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    Else
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    End If
                Else
                    If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    Else
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                                
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    End If
                End If
            End If
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' GENERATES A VISUAL GRID CROSS-SECTION FOR AN UP-SCALED MODEL WITH A FRACTURE
'================================================================================================================================
Private Sub GenerateUpScaledCrossSectionFrac(ByVal XDim As Integer, ByVal LastFrac As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim LASLastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim ThisColumn As Integer 'Column index for the active sheet
    Dim PermTol As Double 'Numeric or percentage permeability tolerance
    Dim PoroTol As Double 'Numeric or percentage permeability tolerance
    Dim FracTop As Double 'Frac top depth
    Dim FracBase As Double 'Frac base depth
    Dim MiddleColumn As Integer 'Middle column of the cross-section
    
    Sheets.Add
    ActiveSheet.Name = "Cross-Section (With Frac)"
    
    LASLastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row
    FracTop = CDbl(FormExpressRun.TxtFracTop.Text)
    FracBase = FracTop + CDbl(FormExpressRun.TxtFracHeight.Text)
    
    If FormUpScaledGrid.TxtPermTolValue.Enabled = True Then
        PermTol = CDbl(FormUpScaledGrid.TxtPermTolValue.Text)
    End If
    
    If FormUpScaledGrid.TxtPoroTolValue.Enabled = True Then
        PoroTol = CDbl(FormUpScaledGrid.TxtPoroTolValue.Text)
    End If
    
    ThisColumn = 1
    
    Do While ThisColumn <= XDim
        Columns(ThisColumn).ColumnWidth = Sheets("Grid Top View").Columns(ThisColumn).ColumnWidth
        
        ThisColumn = ThisColumn + 1
    Loop
    
    Columns("A").Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    ThisRow = 4
    LASRow = 5
    ThisColumn = 1
    
    Do While ThisRow <= LASLastRow - 1
        Do While ThisColumn <= XDim
            
            With Cells(ThisRow, ThisColumn)
                .Interior.Color = Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        
            ThisColumn = ThisColumn + 1
        Loop
        
        If (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(255, 0, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0)) Or _
            (Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) And _
            Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(255, 0, 0)) Or _
            ThisRow = 4 Then
            
            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
        
        If Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
            If FormUpScaledGrid.OptPermTol.Value = True Then
                If FormUpScaledGrid.OptPermNumeric.Value = True Then
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                Else
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                End If
            End If
            
            If FormUpScaledGrid.OptPoroTol.Value = True Then
                If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                Else
                    If CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                        CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                        
                        Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    End If
                End If
            End If
            
            If FormUpScaledGrid.OptBothTol.Value = True Then
                If FormUpScaledGrid.OptPermNumeric.Value = True Then
                    If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    Else
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - PermTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    End If
                Else
                    If FormUpScaledGrid.OptPoroNumeric.Value = True Then
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + PoroTol Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - PoroTol Then
                            
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    Else
                        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "E").Value) * (PermTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) > _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) + _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Or _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow, "G").Value) < _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) - _
                            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "G").Value) * (PoroTol / 100) Then
                                
                            Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If
                    End If
                End If
            End If
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
        ThisColumn = 1
    Loop
    
    ThisRow = 5
    LASRow = 6
    
    Do While ThisRow <= LASLastRow - 1
        If CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) > FracTop And _
            CDbl(Sheets(".LAS File Data").Cells(LASRow - 1, "C").Value) < FracTop Then
            
            If Sheets(".LAS File Data").Cells(LASRow, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            While CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value) < FracBase
                MiddleColumn = XDim \ 2 + 1
                
                For i = 1 To LastFrac
                    Cells(ThisRow, MiddleColumn + i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn - i).Interior.Color = RGB(0, 255, 255)
                    Cells(ThisRow, MiddleColumn).Interior.Color = RGB(0, 255, 255)
                Next i
                
                LASRow = LASRow + 1
                ThisRow = ThisRow + 1
            Wend
            
            If Sheets(".LAS File Data").Cells(LASRow - 1, "K").Interior.Color = RGB(0, 255, 0) Then
                Range(Cells(ThisRow, "A"), Cells(ThisRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            Exit Do
            
        End If
        
        LASRow = LASRow + 1
        ThisRow = ThisRow + 1
    Loop
    
    Range(Cells(LASLastRow, "A"), Cells(LASLastRow, XDim)).Borders(xlEdgeTop).LineStyle = xlContinuous

End Sub

'================================================================================================================================
' DETERMINES THE NUMBER OF GRID LAYERS
'================================================================================================================================
Private Function GetNumLayers() As Integer
 
    Dim Row As Long 'Row index for loop
    Dim LastRow As Long 'Row index of last row
    Dim Output As Integer 'Variable that will eventually become the number of layers

    LastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row - 1
    
    Row = 4
    Output = 0
    
    Do While Row <= LastRow
        If Cells(Row, "A").Borders(xlEdgeTop).LineStyle = xlContinuous Then
            Output = Output + 1
        End If
    
        Row = Row + 1
    Loop
    
    GetNumLayers = Output

End Function

'================================================================================================================================
' GENERATES THE LAYER PROPERTIES NEXT TO THE VISUAL GRID CROSS-SECTION
'================================================================================================================================
Private Sub GenerateLayerProperties(ByVal XDim As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim i As Long 'Count variable for layer tops
    Dim j As Long 'Count variable for layer average properties
    Dim LastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim NumLayers As Integer 'Total number of grid layers
    Dim ActiveFlags(2000) As Integer 'Contains whether or not each layer is active in the simulation
    Dim LayerTops(2000) As Double 'Top depth value for each layer
    Dim dZs(2000) As Double 'dZ value for each layer
    Dim LayerNumReadings As Long 'Number of depth readings in the layer currently being parsed
    Dim LayerAvgPerms(2000) As Double 'Average permeability for each layer
    Dim LayerAvgPoros(2000) As Double 'Average porosity for each layer
    Dim LayerAvgSws(2000) As Double 'Average water saturation for each layer
    Dim CSRect As Shape 'Grid Cross-Section title rectangle
    Dim CSRectLeft As Double 'Grid Cross-Section title rectangle left
    Dim CSRectTop As Double 'Grid Cross-Section title rectangle top
    Dim CSRectWidth As Double 'Grid Cross-Section title rectangle width
    Dim CSRectHeight As Double 'Grid Cross-Section title rectangle height
    Dim DataRect As Shape 'Grid Data title rectangle
    Dim DataRectLeft As Double 'Grid Data title rectangle left
    Dim DataRectTop As Double 'Grid Data title rectangle top
    Dim DataRectWidth As Double 'Grid Data title rectangle width
    Dim DataRectHeight As Double 'Grid Data title rectangle height
    Dim FracTopZ As Integer 'The top Z value (layer number) for the fracture
    Dim FracBottomZ As Integer 'The bottom Z value (layer number) for the fracture
    
    LastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row - 1
    NumLayers = GetNumLayers
    
    ThisRow = 4
    LASRow = 5
    j = 1
    LayerNumReadings = 0
    
    Do While ThisRow <= LastRow
        LayerTops(1) = CDbl(Sheets(".LAS File Data").Cells(5, "C").Value)
        
        If Cells(ThisRow, "A").Borders(xlEdgeTop).LineStyle = xlContinuous And LASRow <> 5 Then
            LayerAvgPerms(j) = LayerAvgPerms(j) / LayerNumReadings
            LayerAvgPoros(j) = LayerAvgPoros(j) / LayerNumReadings
            LayerAvgSws(j) = LayerAvgSws(j) / LayerNumReadings
    
            Select Case Cells(ThisRow - 1, "A").Interior.Color
                Case Is = RGB(255, 0, 0)
                    ActiveFlags(j) = 0
                Case Is = RGB(0, 255, 0)
                    ActiveFlags(j) = 1
            End Select
            
            j = j + 1
            
            If j <= NumLayers Then
                LayerTops(j) = CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value)
            End If
            
            LayerNumReadings = 0
        End If
        
        If j <= NumLayers Then
            LayerAvgPerms(j) = LayerAvgPerms(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value)
            LayerAvgPoros(j) = LayerAvgPoros(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "F").Value)
            LayerAvgSws(j) = LayerAvgSws(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "H").Value)
        End If
        
        LayerNumReadings = LayerNumReadings + 1
        
        If ThisRow = LastRow And j <= NumLayers Then
            LayerAvgPerms(j) = LayerAvgPerms(j) / LayerNumReadings
            LayerAvgPoros(j) = LayerAvgPoros(j) / LayerNumReadings
            LayerAvgSws(j) = LayerAvgSws(j) / LayerNumReadings
            
            Select Case Sheets(".LAS File Data").Cells(ThisRow, "A").Interior.Color
                Case Is = RGB(255, 0, 0)
                    ActiveFlags(j) = 0
                Case Is = RGB(0, 255, 0)
                    ActiveFlags(j) = 1
            End Select
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
    Loop
    
    Cells(3, XDim + 1).Font.Bold = True
    Cells(3, XDim + 1).Value = "Layer #"
    
    Cells(3, XDim + 2).Font.Bold = True
    Cells(3, XDim + 2).Value = "Active?"
    
    Cells(3, XDim + 3).Font.Bold = True
    Cells(3, XDim + 3).Value = "Top Depth (ft TVD)"
    
    Cells(3, XDim + 4).Font.Bold = True
    Cells(3, XDim + 4).Value = "dZ (ft)"
    
    Cells(3, XDim + 5).Font.Bold = True
    Cells(3, XDim + 5).Value = "Layer Perm (mD)"
    
    Cells(3, XDim + 6).Font.Bold = True
    Cells(3, XDim + 6).Value = "Layer Porosity (Fraction)"
    
    Cells(3, XDim + 7).Font.Bold = True
    Cells(3, XDim + 7).Value = "Layer Porosity (%)"
    
    Cells(3, XDim + 8).Font.Bold = True
    Cells(3, XDim + 8).Value = "Layer Sw (Fraction)"
    
    Cells(3, XDim + 9).Font.Bold = True
    Cells(3, XDim + 9).Value = "Layer Sw (%)"
    
    ThisRow = 4
    j = 1
    
    Do While ThisRow <= LastRow
        If Cells(ThisRow, XDim).Borders(xlEdgeTop).LineStyle = xlContinuous Then
            Cells(ThisRow, XDim + 1).Value = j
            
            If ActiveFlags(j) = 0 Then
                Cells(ThisRow, XDim + 2).Value = "No"
            Else
                Cells(ThisRow, XDim + 2).Value = "Yes"
            End If
            
            Cells(ThisRow, XDim + 3).Value = LayerTops(j)
            Cells(ThisRow, XDim + 5).Value = LayerAvgPerms(j)
            Cells(ThisRow, XDim + 6).Value = LayerAvgPoros(j)
            Cells(ThisRow, XDim + 7).Value = LayerAvgPoros(j) * 100
            Cells(ThisRow, XDim + 8).Value = LayerAvgSws(j)
            Cells(ThisRow, XDim + 9).Value = LayerAvgSws(j) * 100
            Range(Cells(ThisRow, XDim + 1), Cells(ThisRow, XDim + 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
            
            j = j + 1
        End If
    
        Cells(ThisRow, XDim + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 6).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 7).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 8).Borders(xlEdgeRight).LineStyle = xlContinuous
        Cells(ThisRow, XDim + 9).Borders(xlEdgeRight).LineStyle = xlContinuous
        ThisRow = ThisRow + 1
    Loop
    
    ThisRow = 4
    
    Do While ThisRow <= LastRow
        j = ThisRow + 1
        
        If Cells(ThisRow, XDim).Borders(xlEdgeTop).LineStyle = xlContinuous Then
            Do While j <= LastRow
                If Cells(j, XDim + 3).Value <> vbNullString Then
                    Cells(ThisRow, XDim + 4).Value = CDbl(Cells(j, XDim + 3).Value) - CDbl(Cells(ThisRow, XDim + 3).Value)
                    Exit Do
                End If
                
                j = j + 1
            Loop
        End If
    
        ThisRow = ThisRow + 1
    Loop
    
    Range(Cells(LastRow + 1, XDim + 1), Cells(LastRow + 1, XDim + 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    For i = 1 To 9
        Columns(XDim + i).HorizontalAlignment = xlCenter
        Columns(XDim + i).AutoFit
        
        With Cells(3, XDim + i)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders.Weight = xlThick
        End With
    Next i
    
    With Range(Cells(1, "A"), Cells(2, XDim))
        CSRectLeft = .Left
        CSRectTop = .Top
        CSRectWidth = .Width
        CSRectHeight = .Height
    End With
    
    Set CSRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, CSRectLeft, CSRectTop, CSRectWidth, CSRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, CSRectLeft, CSRectTop, CSRectWidth, CSRectHeight).Name = "Grid Cross-Section (Without Frac)"
    
    With ActiveSheet.Shapes("Grid Cross-Section (Without Frac)")
        .TextFrame.Characters.Text = "GRID CROSS-SECTION (WITHOUT FRAC)"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    With Range(Cells(1, XDim + 1), Cells(2, XDim + 9))
        DataRectLeft = .Left
        DataRectTop = .Top
        DataRectWidth = .Width
        DataRectHeight = .Height
    End With
    
    Set DataRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, DataRectLeft, DataRectTop, DataRectWidth, DataRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, DataRectLeft, DataRectTop, DataRectWidth, DataRectHeight).Name = "Grid Layer Data"
    
    With ActiveSheet.Shapes("Grid Layer Data")
        .TextFrame.Characters.Text = "GRID LAYER DATA"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With

End Sub

'================================================================================================================================
' GENERATES THE LAYER PROPERTIES NEXT TO THE VISUAL GRID CROSS-SECTION WITH A FRACTURE
'================================================================================================================================
Private Sub GenerateLayerPropertiesFrac(ByVal XDim As Integer, ByVal LastFrac As Integer)

    Dim ThisRow As Long 'Row index for the active sheet
    Dim LASRow As Long 'Row index for the ".LAS File Data" sheet
    Dim i As Long 'Count variable for layer tops
    Dim j As Long 'Count variable for layer average properties
    Dim LastRow As Long 'Row index for the last row of the ".LAS File Data" sheet
    Dim NumLayers As Integer 'Total number of grid layers
    Dim ActiveFlags(2000) As Integer 'Contains whether or not each layer is active in the simulation
    Dim LayerTops(2000) As Double 'Top depth value for each layer
    Dim dZs(2000) As Double 'dZ value for each layer
    Dim LayerNumReadings As Long 'Number of depth readings in the layer currently being parsed
    Dim LayerAvgPerms(2000) As Double 'Average permeability for each layer
    Dim LayerAvgPoros(2000) As Double 'Average porosity for each layer
    Dim LayerAvgSws(2000) As Double 'Average water saturation for each layer
    Dim CSRect As Shape 'Grid Cross-Section title rectangle
    Dim CSRectLeft As Double 'Grid Cross-Section title rectangle left
    Dim CSRectTop As Double 'Grid Cross-Section title rectangle top
    Dim CSRectWidth As Double 'Grid Cross-Section title rectangle width
    Dim CSRectHeight As Double 'Grid Cross-Section title rectangle height
    Dim DataRect As Shape 'Grid Data title rectangle
    Dim DataRectLeft As Double 'Grid Data title rectangle left
    Dim DataRectTop As Double 'Grid Data title rectangle top
    Dim DataRectWidth As Double 'Grid Data title rectangle width
    Dim DataRectHeight As Double 'Grid Data title rectangle height
    Dim FracDXs() As Double 'Array containing one fracture wing's grid block dX values
    Dim HalfLengths() As Double 'Array of distance along the half-length for each fracture grid block (far endpoint of each block)
    Dim HLMidPoints() As Double 'Array of distance along the half-length for each fracture grid block (midpoint of each block)
    Dim VFCDRect As Shape 'Variable FCD title rectangle
    Dim VFCDRectLeft As Double 'Variable FCD title rectangle left
    Dim VFCDRectTop As Double 'Variable FCD title rectangle top
    Dim VFCDRectWidth As Double 'Variable FCD title rectangle width
    Dim VFCDRectHeight As Double 'Variable FCD title rectangle height
    
    ReDim FracDXs(LastFrac)
    ReDim HalfLengths(LastFrac)
    ReDim HLMidPoints(LastFrac)
    
    LastRow = Sheets(".LAS File Data").Cells(Rows.Count, "B").End(xlUp).Row - 1
    NumLayers = GetNumLayers
    
    ThisRow = 4
    LASRow = 5
    j = 1
    LayerNumReadings = 0
    
    Do While ThisRow <= LastRow
        LayerTops(1) = CDbl(Sheets(".LAS File Data").Cells(5, "C").Value)
        
        If Cells(ThisRow, "A").Borders(xlEdgeTop).LineStyle = xlContinuous And LASRow <> 5 Then
            LayerAvgPerms(j) = LayerAvgPerms(j) / LayerNumReadings
            LayerAvgPoros(j) = LayerAvgPoros(j) / LayerNumReadings
            LayerAvgSws(j) = LayerAvgSws(j) / LayerNumReadings
    
            Select Case Cells(ThisRow - 1, "A").Interior.Color
                Case Is = RGB(255, 0, 0)
                    ActiveFlags(j) = 0
                Case Is = RGB(0, 255, 0)
                    ActiveFlags(j) = 1
            End Select
            
            j = j + 1
            
            If j <= NumLayers Then
                LayerTops(j) = CDbl(Sheets(".LAS File Data").Cells(LASRow, "C").Value)
            End If
            
            LayerNumReadings = 0
        End If
        
        If j <= NumLayers Then
            LayerAvgPerms(j) = LayerAvgPerms(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "E").Value)
            LayerAvgPoros(j) = LayerAvgPoros(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "F").Value)
            LayerAvgSws(j) = LayerAvgSws(j) + CDbl(Sheets(".LAS File Data").Cells(LASRow, "H").Value)
        End If
        
        LayerNumReadings = LayerNumReadings + 1
        
        If ThisRow = LastRow And j <= NumLayers Then
            LayerAvgPerms(j) = LayerAvgPerms(j) / LayerNumReadings
            LayerAvgPoros(j) = LayerAvgPoros(j) / LayerNumReadings
            LayerAvgSws(j) = LayerAvgSws(j) / LayerNumReadings
            
            Select Case Sheets(".LAS File Data").Cells(ThisRow, "A").Interior.Color
                Case Is = RGB(255, 0, 0)
                    ActiveFlags(j) = 0
                Case Is = RGB(0, 255, 0)
                    ActiveFlags(j) = 1
            End Select
        End If
        
        ThisRow = ThisRow + 1
        LASRow = LASRow + 1
    Loop
    
    Cells(3, XDim + 1).Font.Bold = True
    Cells(3, XDim + 1).Value = "Layer #"
    
    Cells(3, XDim + 2).Font.Bold = True
    Cells(3, XDim + 2).Value = "Active?"
    
    Cells(3, XDim + 3).Font.Bold = True
    Cells(3, XDim + 3).Value = "Top Depth (ft TVD)"
    
    Cells(3, XDim + 4).Font.Bold = True
    Cells(3, XDim + 4).Value = "dZ (ft)"
    
    Cells(3, XDim + 5).Font.Bold = True
    Cells(3, XDim + 5).Value = "Layer Perm (mD)"
    
    Cells(3, XDim + 6).Font.Bold = True
    Cells(3, XDim + 6).Value = "Layer Porosity (Fraction)"
    
    Cells(3, XDim + 7).Font.Bold = True
    Cells(3, XDim + 7).Value = "Layer Porosity (%)"
    
    Cells(3, XDim + 8).Font.Bold = True
    Cells(3, XDim + 8).Value = "Layer Sw (Fraction)"
    
    Cells(3, XDim + 9).Font.Bold = True
    Cells(3, XDim + 9).Value = "Layer Sw (%)"
    
    Cells(3, XDim + 10).Font.Bold = True
    Cells(3, XDim + 10).Value = "Frac Perm (mD)"
    
    For i = 1 To LastFrac * 2 + 1
        Dim X As Integer
        
        If i = 1 Then
            X = XDim \ 2 + 1 - LastFrac
        End If
        
        Cells(2, XDim + 11 + i).Font.Bold = True
        Cells(2, XDim + 11 + i).Value = "X = " & X
        
        X = X + 1
    Next i
    
    ThisRow = 14
    j = 1
    
    Do While ThisRow <= LastRow
        If Sheets("Grid Statistics").Cells(ThisRow, "A").Interior.Color = RGB(0, 255, 255) Then
            FracDXs(j) = CDbl(Sheets("Grid Statistics").Cells(ThisRow, "C").Value)
            
            j = j + 1
            
            If Sheets("Grid Statistics").Cells(ThisRow + 1, "A").Interior.Color <> RGB(0, 255, 255) Then
                Exit Do
            End If
        End If
        
        ThisRow = ThisRow + 1
    Loop
    
    j = 1
    
    Dim SumDXs As Double
    
    SumDXs = 0
    
    For i = 1 To LastFrac
        
        For k = j To LastFrac
            SumDXs = SumDXs + FracDXs(k)
        Next k
        
        HalfLengths(i) = SumDXs
        
        SumDXs = 0
        j = j + 1
    Next i
    
    For i = 1 To LastFrac
    
        If i = LastFrac Then
            HLMidPoints(i) = HalfLengths(i) / 2
        Else
            HLMidPoints(i) = HalfLengths(i) - ((HalfLengths(i) - HalfLengths(i + 1)) / 2)
        End If
        
        Cells(3, XDim + 11 + i).Font.Bold = True
        Cells(3, XDim + 11 + i).Value = HLMidPoints(i) + 2.5
    Next i
    
    Cells(3, XDim + 11 + LastFrac + 1).Font.Bold = True
    Cells(3, XDim + 11 + LastFrac + 1).Value = 0
    
    j = XDim + 11 + LastFrac + 2
    k = LastFrac
    
    For i = 1 To LastFrac
        Cells(3, j).Font.Bold = True
        Cells(3, j).Value = HLMidPoints(k) + 2.5
        
        k = k - 1
        j = j + 1
    Next i
    
    ThisRow = 4
    j = 1
    
    Do While ThisRow <= LastRow
        If Cells(ThisRow, XDim).Borders(xlEdgeTop).LineStyle = xlContinuous Then
            Cells(ThisRow, XDim + 1).Value = j
            
            If ActiveFlags(j) = 0 Then
                Cells(ThisRow, XDim + 2).Value = "No"
            Else
                Cells(ThisRow, XDim + 2).Value = "Yes"
            End If
            
            Cells(ThisRow, XDim + 3).Value = LayerTops(j)
            Cells(ThisRow, XDim + 5).Value = LayerAvgPerms(j)
            Cells(ThisRow, XDim + 6).Value = LayerAvgPoros(j)
            Cells(ThisRow, XDim + 7).Value = LayerAvgPoros(j) * 100
            Cells(ThisRow, XDim + 8).Value = LayerAvgSws(j)
            Cells(ThisRow, XDim + 9).Value = LayerAvgSws(j) * 100
            
            If Cells(ThisRow, XDim \ 2).Interior.Color = RGB(0, 255, 255) Then
                Cells(ThisRow, XDim + 10).Value = Sheets("Hydraulic Fracture").Cells(10, "C").Value
                
                For i = XDim + 12 To XDim + 12 + LastFrac * 2
                    Cells(ThisRow, i).Value = Sheets("Hydraulic Fracture").Cells(10, "C").Value
                Next i
                
                Range(Cells(ThisRow, XDim + 12), Cells(ThisRow, XDim + 12 + LastFrac * 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
            Range(Cells(ThisRow, XDim + 1), Cells(ThisRow, XDim + 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
            
            j = j + 1
        End If
    
        For i = XDim + 1 To XDim + 12 + LastFrac * 2
            Cells(ThisRow, i).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next i
        
'        Cells(ThisRow, XDim + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 2).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 3).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 5).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 6).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 7).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 8).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 9).Borders(xlEdgeRight).LineStyle = xlContinuous
'        Cells(ThisRow, XDim + 10).Borders(xlEdgeRight).LineStyle = xlContinuous

        ThisRow = ThisRow + 1
    Loop
    
    For i = XDim + 12 To XDim + 12 + LastFrac * 2
        Range(Cells(2, i), Cells(3, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
        Range(Cells(2, i), Cells(3, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Range(Cells(2, i), Cells(3, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(Cells(2, i), Cells(3, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Range(Cells(2, i), Cells(3, i)).Borders.Weight = xlThick
    Next i
    
    ThisRow = 4
    
    Do While ThisRow <= LastRow
        j = ThisRow + 1
        
        If Cells(ThisRow, XDim).Borders(xlEdgeTop).LineStyle = xlContinuous Then
            Do While j <= LastRow
                If Cells(j, XDim + 2).Value <> vbNullString Then
                    Cells(ThisRow, XDim + 4).Value = CDbl(Cells(j, XDim + 3).Value) - CDbl(Cells(ThisRow, XDim + 3).Value)
                    Exit Do
                End If
                
                j = j + 1
            Loop
        End If
    
        ThisRow = ThisRow + 1
    Loop
    
    Range(Cells(LastRow + 1, XDim + 1), Cells(LastRow + 1, XDim + 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range(Cells(LastRow + 1, XDim + 12), Cells(LastRow + 1, XDim + 12 + LastFrac * 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    For i = 1 To 10
        Columns(XDim + i).HorizontalAlignment = xlCenter
        Columns(XDim + i).AutoFit
        
        With Cells(3, XDim + i)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders.Weight = xlThick
        End With
    Next i
    
    For i = XDim + 12 To XDim + 12 + LastFrac * 2 + 1
        Columns(i).HorizontalAlignment = xlCenter
    Next i
    
    With Range(Cells(1, "A"), Cells(2, XDim))
        CSRectLeft = .Left
        CSRectTop = .Top
        CSRectWidth = .Width
        CSRectHeight = .Height
    End With
    
    Set CSRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, CSRectLeft, CSRectTop, CSRectWidth, CSRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, CSRectLeft, CSRectTop, CSRectWidth, CSRectHeight).Name = "Grid Cross-Section (With Frac)"
    
    With ActiveSheet.Shapes("Grid Cross-Section (With Frac)")
        .TextFrame.Characters.Text = "GRID CROSS-SECTION (WITH FRAC)"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    With Range(Cells(1, XDim + 1), Cells(2, XDim + 10))
        DataRectLeft = .Left
        DataRectTop = .Top
        DataRectWidth = .Width
        DataRectHeight = .Height
    End With
    
    Set DataRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, DataRectLeft, DataRectTop, DataRectWidth, DataRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, DataRectLeft, DataRectTop, DataRectWidth, DataRectHeight).Name = "Grid Layer Data"
    
    With ActiveSheet.Shapes("Grid Layer Data")
        .TextFrame.Characters.Text = "GRID LAYER DATA"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
    With Range(Cells(1, XDim + 12), Cells(1, XDim + 12 + LastFrac * 2))
        VFCDRectLeft = .Left
        VFCDRectTop = .Top
        VFCDRectWidth = .Width
        VFCDRectHeight = .Height
    End With
    
    Set VFCDRect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, VFCDRectLeft, VFCDRectTop, VFCDRectWidth, VFCDRectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, VFCDRectLeft, VFCDRectTop, VFCDRectWidth, VFCDRectHeight).Name = "Variable FCD Fracture Permeabilities"
    
    With ActiveSheet.Shapes("Variable FCD Fracture Permeabilities")
        .TextFrame.Characters.Text = "VARIABLE FCD FRACTURE PERMEABILITIES"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With

End Sub

'================================================================================================================================
' UPDATES THE "GRID STATISTICS" SHEET WITH Z-DIMENSION INFORMATION
'================================================================================================================================
Private Sub UpdateGridStats(ByVal XDim As Integer)

    Dim Row As Integer 'Row index for loops
    Dim NumLayers As Integer 'Total number of layers in the grid without a frac
    Dim NumLayersFrac As Integer 'Total number of layers in the grid with a frac
    Dim ZDistance As Double 'Total (gross) thickness of the grid
    Dim PermTol As Double 'Permeability tolerance (if applicable)
    Dim PoroTol As Double 'Porosity tolerance (if applicable)
    Dim TotalBlocks As Long 'Total number of grid blocks (x*y*z)
    Dim TotalBlocksFrac As Long 'Total number of grid blocks with frac (x*y*z_frac)
    Dim InactiveLayers As Integer 'Number of inactive (non-reservoir) grid layers
    Dim InactiveLayersFrac As Integer 'Number of inactive (non-reservoir) grid layers with frac
    Dim ActiveLayers As Integer 'Number of active (reservoir) grid layers
    Dim ActiveLayersFrac As Integer 'Number of active (reservoir) grid layers with frac
    Dim ActiveBlocks As Long 'Number of active grid blocks (x*y*active_z)
    Dim ActiveBlocksFrac As Long 'Number of active grid blocks with frac (x*y*active_z_frac)
    Dim Rect As Shape 'Grid Statistics title rectangle
    Dim RectLeft As Double 'Grid Statistics title rectangle left
    Dim RectTop As Double 'Grid Statistics title rectangle top
    Dim RectWidth As Double 'Grid Statistics title rectangle width
    Dim RectHeight As Double 'Grid Statistics title rectangle height
    
    Sheets("Cross-Section (Without Frac)").Activate
    
    NumLayers = GetNumLayers

    Sheets("Cross-Section (With Frac)").Activate
    
    NumLayersFrac = GetNumLayers

    ZDistance = CDbl(Sheets("Petrophysical Analysis").Cells(5, "G").Value)
    
    With FormUpScaledGrid
        If .OptPermTol.Value = True Or .OptBothTol.Value = True Then
            PermTol = CDbl(.TxtPermTolValue.Text)
        End If
        
        If .OptPoroTol.Value = True Or .OptBothTol.Value = True Then
            PoroTol = CDbl(.TxtPoroTolValue.Text)
        End If
    End With
    
    With Sheets("Grid Statistics")
        TotalBlocks = CLng(XDim) * CLng(.Cells(8, "C").Value) * NumLayers
        TotalBlocksFrac = CLng(XDim) * CLng(.Cells(8, "C").Value) * NumLayersFrac
    End With
    
    Sheets("Cross-Section (Without Frac)").Activate
    
    Row = 4
    
    Do While Row <= Cells(Rows.Count, XDim + 1).End(xlUp).Row
        If Cells(Row, XDim + 2).Value = "No" Then
            InactiveLayers = InactiveLayers + 1
        End If
        
        Row = Row + 1
    Loop
    
    ActiveLayers = NumLayers - InactiveLayers
    
    Sheets("Cross-Section (With Frac)").Activate
    
    Row = 4
    
    Do While Row <= Cells(Rows.Count, XDim + 1).End(xlUp).Row
        If Cells(Row, XDim + 2).Value = "No" Then
            InactiveLayersFrac = InactiveLayersFrac + 1
        End If
        
        Row = Row + 1
    Loop
    
    ActiveLayersFrac = NumLayersFrac - InactiveLayersFrac
    
    With Sheets("Grid Statistics")
        ActiveBlocks = CLng(XDim) * CLng(.Cells(8, "C").Value) * ActiveLayers
        ActiveBlocksFrac = CLng(XDim) * CLng(.Cells(8, "C").Value) * ActiveLayersFrac
        
        InactiveBlocks = TotalBlocks - ActiveBlocks
        InactiveBlocksFrac = TotalBlocksFrac - ActiveBlocksFrac
        
        .Range("E4:G4").Merge
        .Cells(4, "E").Font.Bold = True
        .Cells(4, "E").Value = "Non-Fractured Grid Layers"
        
        .Range("E5:F5").Merge
        .Cells(5, "E").Font.Underline = True
        .Cells(5, "E").Value = "Z Dimensions (Grid Blocks):"
        .Cells(5, "G").Value = NumLayers
        
        .Range("E6:F6").Merge
        .Cells(6, "E").Font.Underline = True
        .Cells(6, "E").Value = "Total Z Distance (ft):"
        .Cells(6, "G").Value = ZDistance
        
        .Range("E7:F7").Merge
        .Cells(7, "E").Font.Underline = True
        .Cells(7, "E").Value = "Active Layers:"
        .Cells(7, "G").Value = ActiveLayers
        
        .Range("E8:F8").Merge
        .Cells(8, "E").Font.Underline = True
        .Cells(8, "E").Value = "Inactive Layers:"
        .Cells(8, "G").Value = InactiveLayers
        
        .Range("E9:F9").Merge
        .Cells(9, "E").Font.Underline = True
        .Cells(9, "E").Value = "Total Grid Blocks:"
        .Cells(9, "G").Value = TotalBlocks
        
        .Range("E10:F10").Merge
        .Cells(10, "E").Font.Underline = True
        .Cells(10, "E").Value = "Active Grid Blocks:"
        .Cells(10, "G").Value = ActiveBlocks
        
        .Range("E11:F11").Merge
        .Cells(11, "E").Font.Underline = True
        .Cells(11, "E").Value = "Inactive Grid Blocks:"
        .Cells(11, "G").Value = InactiveBlocks
        
        With .Range("E4:G11")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        
        With .Range("E4:G4")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders.Weight = xlThick
        End With
        
        .Range("E13:G13").Merge
        .Cells(13, "E").Font.Bold = True
        .Cells(13, "E").Value = "Fractured Grid Layers"
        
        .Range("E14:F14").Merge
        .Cells(14, "E").Font.Underline = True
        .Cells(14, "E").Value = "Z Dimensions (Grid Blocks):"
        .Cells(14, "G").Value = NumLayersFrac
        
        .Range("E15:F15").Merge
        .Cells(15, "E").Font.Underline = True
        .Cells(15, "E").Value = "Total Z Distance (ft):"
        .Cells(15, "G").Value = ZDistance
        
        .Range("E16:F16").Merge
        .Cells(16, "E").Font.Underline = True
        .Cells(16, "E").Value = "Active Layers:"
        .Cells(16, "G").Value = ActiveLayersFrac
        
        .Range("E17:F17").Merge
        .Cells(17, "E").Font.Underline = True
        .Cells(17, "E").Value = "Inactive Layers:"
        .Cells(17, "G").Value = InactiveLayersFrac
        
        .Range("E18:F18").Merge
        .Cells(18, "E").Font.Underline = True
        .Cells(18, "E").Value = "Total Grid Blocks:"
        .Cells(18, "G").Value = TotalBlocksFrac
        
        .Range("E19:F19").Merge
        .Cells(19, "E").Font.Underline = True
        .Cells(19, "E").Value = "Active Grid Blocks:"
        .Cells(19, "G").Value = ActiveBlocksFrac
        
        .Range("E20:F20").Merge
        .Cells(20, "E").Font.Underline = True
        .Cells(20, "E").Value = "Inactive Grid Blocks:"
        .Cells(20, "G").Value = InactiveBlocksFrac
        
        .Range("E22:G22").Merge
        .Cells(22, "E").Font.Bold = True
        .Cells(22, "E").Value = "S_f/R_we Grid Info"
        
        .Range("E23:F23").Merge
        .Cells(23, "E").Font.Underline = True
        .Cells(23, "E").Value = "Constant dX/dY (ft):"
        .Cells(23, "G").Value = CDbl(.Cells(6, "C").Value) / CDbl(.Cells(5, "C").Value)
        
        With .Range("E22:G23")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        
        With .Range("E22:G22")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders.Weight = xlThick
        End With
        
        With .Range("E13:G20")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        
        With .Range("E13:G13")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders.Weight = xlThick
        End With
        
        Sheets("Grid Statistics").Activate
        
        Columns("E").ColumnWidth = 11.22
        Columns("F").ColumnWidth = 11.22
        Columns("G").ColumnWidth = 11.22
        Range("E4:G23").HorizontalAlignment = xlCenter
        
        With Range(Cells(1, "A"), Cells(2, "G"))
            RectLeft = .Left
            RectTop = .Top
            RectWidth = .Width
            RectHeight = .Height
        End With
        
        Set Rect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, RectLeft, RectTop, RectWidth, RectHeight)
    
        ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, RectLeft, RectTop, RectWidth, RectHeight).Name = "Grid Statistics"
        
        With ActiveSheet.Shapes("Grid Statistics")
            .TextFrame.Characters.Text = "GRID STATISTICS"
            .TextFrame.HorizontalAlignment = xlCenter
            .TextFrame.VerticalAlignment = xlCenter
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = True
            .Fill.Transparency = 1
            .Line.Visible = msoFalse
        End With
    
    End With
    
End Sub

'================================================================================================================================
' GENERATES THE "INCLUDE" FILE FOR THE NON-FRACTURED GRID
'================================================================================================================================
Sub GenerateNonFracInclude()

    Dim XDim As Integer 'X dimensions for both grids
    Dim YDim As Integer 'Y dimensions for both grids
    Dim ZDim As Integer 'Z dimensions (number of layers) for the non-fractured grid
    Dim dXs() As Double 'dX profile for both grids
    Dim dYs() As Double 'dY profile for both grids
    Dim ActiveFlags() As Integer 'Active flag for each layer of the non-fractured grid
    Dim TopDepths() As Double 'Top depth for each layer of the non-fractured grid
    Dim dZs() As Double 'dZ profile for the non-fractured grid
    Dim LayerPerms() As Double 'Permeability for each layer of the non-fractured grid
    Dim LayerPoros() As Double 'Porosity for each layer of the non-fractured grid
    Dim LayerSws() As Double 'Water saturation for each layer of the non-fractured grid
    Dim Row As Integer 'Row index for loops
    Dim i As Integer 'Count variable for loops
    Dim Fso As Object 'For creating the include file
    Dim OutputFile As Object 'Output include file
    Dim PrintLine As String 'The line to be written to the output file
    
    Sheets("Grid Statistics").Activate
    
    XDim = Cells(5, "C").Value
    YDim = Cells(8, "C").Value
    ZDim = Cells(5, "G").Value
    
    ReDim dXs(XDim)
    ReDim dYs(YDim)
    ReDim ActiveFlags(ZDim)
    ReDim TopDepths(ZDim)
    ReDim dZs(ZDim)
    ReDim LayerPerms(ZDim)
    ReDim LayerPoros(ZDim)
    ReDim LayerSws(ZDim)
    
    Row = 14
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dXs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Row = Row + 3
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dYs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Sheets("Cross-Section (Without Frac)").Activate
    
    Row = 4
    i = 1
    
    Do While Row <= Cells(Rows.Count, XDim + 1).End(xlUp).Row
        If Cells(Row, XDim + 1).Value <> vbNullString Then
            If Cells(Row, XDim + 2).Value = "Yes" Then
                ActiveFlags(i) = 1
            Else
                ActiveFlags(i) = 0
            End If
            
            TopDepths(i) = Cells(Row, XDim + 3).Value
            dZs(i) = Cells(Row, XDim + 4).Value
            LayerPerms(i) = Cells(Row, XDim + 5).Value
            LayerPoros(i) = Cells(Row, XDim + 6).Value
            LayerSws(i) = Cells(Row, XDim + 8).Value
            
            i = i + 1
        End If
        
        Row = Row + 1
    Loop
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set OutputFile = Fso.CreateTextFile(FormExpressRun.TxtNonFracturedGridPath.Text & "-GRID.INC", True, True)
    
    OutputFile.WriteLine "--Requests output of an INIT file."
    OutputFile.WriteLine "INIT"
    OutputFile.WriteLine vbNullString
    
    OutputFile.WriteLine "DXV"
            
    i = 1
    PrintLine = vbNullString
        
    Do While i <= XDim
        If i <> XDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dXs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dXs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dXs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
        
    OutputFile.WriteLine "DYV"
        
    i = 1
    PrintLine = vbNullString
        
    Do While i <= YDim
        If i <> YDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dYs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dYs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dYs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
    
    i = 1
        
    Do While i <= ZDim
        OutputFile.WriteLine "BOX"
        OutputFile.WriteLine vbTab & "1" & vbTab & vbTab & XDim & vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & _
            i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine vbNullString
            
        OutputFile.WriteLine "EQUALS"
        OutputFile.WriteLine vbTab & "'DZ'" & vbTab & vbTab & dZs(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'ACTNUM'" & vbTab & vbTab & ActiveFlags(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PORO'" & vbTab & vbTab & LayerPoros(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'TOPS'" & vbTab & vbTab & TopDepths(i) & vbTab & vbTab & "1" & vbTab & vbTab & XDim & _
            vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine "/"
        OutputFile.WriteLine vbNullString
            
        i = i + 1
    Loop
    
    OutputFile.Close

End Sub

'================================================================================================================================
' GENERATES THE "INCLUDE" FILE FOR THE FRACTURED GRID
'================================================================================================================================
Sub GenerateFracInclude(ByVal LastFrac As Integer)

    Dim XDim As Integer 'X dimensions for both grids
    Dim YDim As Integer 'Y dimensions for both grids
    Dim ZDim As Integer 'Z dimensions (number of layers) for the non-fractured grid
    Dim dXs() As Double 'dX profile for both grids
    Dim dYs() As Double 'dY profile for both grids
    Dim ActiveFlags() As Integer 'Active flag for each layer of the non-fractured grid
    Dim TopDepths() As Double 'Top depth for each layer of the non-fractured grid
    Dim dZs() As Double 'dZ profile for the non-fractured grid
    Dim LayerPerms() As Double 'Permeability for each layer of the non-fractured grid
    Dim LayerPoros() As Double 'Porosity for each layer of the non-fractured grid
    Dim LayerSws() As Double 'Water saturation for each layer of the non-fractured grid
    Dim FracTopZ As Integer 'Top layer of the frac
    Dim FracBaseZ As Integer 'Bottom layer of the frac
    Dim Row As Integer 'Row index for loops
    Dim i As Integer 'Count variable for loops
    Dim Fso As Object 'For creating the include file
    Dim OutputFile As Object 'Output include file
    Dim PrintLine As String 'The line to be written to the output file
    
    FracTopZ = -1
    
    Sheets("Grid Statistics").Activate
    
    XDim = Cells(5, "C").Value
    YDim = Cells(8, "C").Value
    ZDim = Cells(5, "G").Value
    
    ReDim dXs(XDim)
    ReDim dYs(YDim)
    ReDim ActiveFlags(ZDim + 2)
    ReDim TopDepths(ZDim + 2)
    ReDim dZs(ZDim + 2)
    ReDim LayerPerms(ZDim + 2)
    ReDim LayerPoros(ZDim + 2)
    ReDim LayerSws(ZDim + 2)
    
    Row = 14
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dXs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Row = Row + 3
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dYs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Sheets("Cross-Section (With Frac)").Activate
    
    Row = 4
    i = 1
    
    Do While Row <= Cells(Rows.Count, XDim + 1).End(xlUp).Row
        If Cells(Row, XDim + 1).Value <> vbNullString Then
            If FracTopZ = -1 And Cells(Row, XDim + 10).Value <> vbNullString Then
                FracTopZ = Cells(Row, XDim + 1).Value
            End If
            
            If Cells(Row, XDim + 2).Value = "Yes" Then
                ActiveFlags(i) = 1
            Else
                ActiveFlags(i) = 0
            End If
            
            TopDepths(i) = Cells(Row, XDim + 3).Value
            dZs(i) = Cells(Row, XDim + 4).Value
            LayerPerms(i) = Cells(Row, XDim + 5).Value
            LayerPoros(i) = Cells(Row, XDim + 6).Value
            LayerSws(i) = Cells(Row, XDim + 8).Value
            
            i = i + 1
        End If
        
        Row = Row + 1
    Loop
    
    FracBaseZ = CDbl(Cells(Cells(Rows.Count, XDim + 10).End(xlUp).Row, XDim + 1).Value)
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set OutputFile = Fso.CreateTextFile(FormExpressRun.TxtFracturedGridPath.Text & "-GRID.INC", True, True)
    
    OutputFile.WriteLine "--Requests output of an INIT file."
    OutputFile.WriteLine "INIT"
    OutputFile.WriteLine vbNullString
    
    OutputFile.WriteLine "DXV"
            
    i = 1
    PrintLine = vbNullString
        
    Do While i <= XDim
        If i <> XDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dXs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dXs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dXs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
        
    OutputFile.WriteLine "DYV"
        
    i = 1
    PrintLine = vbNullString
        
    Do While i <= YDim
        If i <> YDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dYs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dYs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dYs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
    
    i = 1
        
    Do While i <= ZDim
        OutputFile.WriteLine "BOX"
        OutputFile.WriteLine vbTab & "1" & vbTab & vbTab & XDim & vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & _
            i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine vbNullString
            
        OutputFile.WriteLine "EQUALS"
        OutputFile.WriteLine vbTab & "'DZ'" & vbTab & vbTab & dZs(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'ACTNUM'" & vbTab & vbTab & ActiveFlags(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PORO'" & vbTab & vbTab & LayerPoros(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'TOPS'" & vbTab & vbTab & TopDepths(i) & vbTab & vbTab & "1" & vbTab & vbTab & XDim & _
            vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine "/"
        OutputFile.WriteLine vbNullString
            
        i = i + 1
    Loop
    
    OutputFile.WriteLine "--********************Hydraulic Fracture Input********************"
    OutputFile.WriteLine "BOX"
    OutputFile.WriteLine vbTab & XDim \ 2 + 1 - LastFrac & vbTab & vbTab & XDim \ 2 + 1 + LastFrac & vbTab & vbTab & YDim \ 2 + 1 & _
        vbTab & vbTab & YDim \ 2 + 1 & vbTab & vbTab & FracTopZ & vbTab & vbTab & FracBaseZ & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
    
    OutputFile.WriteLine "EQUALS"
    OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & Sheets("Hydraulic Fracture").Cells(10, "C").Value & vbTab & vbTab & "/"
    OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & Sheets("Hydraulic Fracture").Cells(10, "C").Value & vbTab & vbTab & "/"
    OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & Sheets("Hydraulic Fracture").Cells(10, "C").Value & vbTab & vbTab & "/"
    OutputFile.WriteLine "/"
    OutputFile.WriteLine vbNullString
    
    OutputFile.Close

End Sub

'================================================================================================================================
' GENERATES THE "INCLUDE" FILE FOR THE FRACTURED GRID
'================================================================================================================================
Sub GenerateVarFCDInclude(ByVal LastFrac As Integer)

    Dim XDim As Integer 'X dimensions for both grids
    Dim YDim As Integer 'Y dimensions for both grids
    Dim ZDim As Integer 'Z dimensions (number of layers) for the non-fractured grid
    Dim dXs() As Double 'dX profile for both grids
    Dim dYs() As Double 'dY profile for both grids
    Dim ActiveFlags() As Integer 'Active flag for each layer of the non-fractured grid
    Dim TopDepths() As Double 'Top depth for each layer of the non-fractured grid
    Dim dZs() As Double 'dZ profile for the non-fractured grid
    Dim LayerPerms() As Double 'Permeability for each layer of the non-fractured grid
    Dim LayerPoros() As Double 'Porosity for each layer of the non-fractured grid
    Dim LayerSws() As Double 'Water saturation for each layer of the non-fractured grid
    Dim FracTopZ As Integer 'Top layer of the frac
    Dim FracBaseZ As Integer 'Bottom layer of the frac
    Dim Row As Integer 'Row index for loops
    Dim i As Integer 'Count variable for loops
    Dim Fso As Object 'For creating the include file
    Dim OutputFile As Object 'Output include file
    Dim PrintLine As String 'The line to be written to the output file
    Dim FracTopRow As Integer 'Frac top row on cross-section sheet
    Dim FracBottomRow As Integer 'Frac bottom row on cross section-sheet
    Dim Column As Integer 'Column number index
    Dim FirstColumn As Integer 'First column for frac input box
    
    FracTopZ = -1
    
    Sheets("Grid Statistics").Activate
    
    XDim = Cells(5, "C").Value
    YDim = Cells(8, "C").Value
    ZDim = Cells(5, "G").Value
    
    ReDim dXs(XDim)
    ReDim dYs(YDim)
    ReDim ActiveFlags(ZDim + 2)
    ReDim TopDepths(ZDim + 2)
    ReDim dZs(ZDim + 2)
    ReDim LayerPerms(ZDim + 2)
    ReDim LayerPoros(ZDim + 2)
    ReDim LayerSws(ZDim + 2)
    
    Row = 14
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dXs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Row = Row + 3
    i = 1
    
    Do While Cells(Row, "C").Value <> vbNullString
        dYs(i) = Cells(Row, "C").Value
        
        Row = Row + 1
        i = i + 1
    Loop
    
    Sheets("Cross-Section (With Frac)").Activate
    
    Row = 4
    i = 1
    
    Do While Row <= Cells(Rows.Count, XDim + 1).End(xlUp).Row
        If Cells(Row, XDim + 1).Value <> vbNullString Then
            If FracTopZ = -1 And Cells(Row, XDim + 10).Value <> vbNullString Then
                FracTopZ = Cells(Row, XDim + 1).Value
            End If
            
            If Cells(Row, XDim + 2).Value = "Yes" Then
                ActiveFlags(i) = 1
            Else
                ActiveFlags(i) = 0
            End If
            
            TopDepths(i) = Cells(Row, XDim + 3).Value
            dZs(i) = Cells(Row, XDim + 4).Value
            LayerPerms(i) = Cells(Row, XDim + 5).Value
            LayerPoros(i) = Cells(Row, XDim + 6).Value
            LayerSws(i) = Cells(Row, XDim + 8).Value
            
            i = i + 1
        End If
        
        Row = Row + 1
    Loop
    
    FracBaseZ = CDbl(Cells(Cells(Rows.Count, XDim + 10).End(xlUp).Row, XDim + 1).Value)
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set OutputFile = Fso.CreateTextFile(FormVarFCD.TxtVFCDGridPath.Text & "-GRID.INC", True, True)
    
    OutputFile.WriteLine "--Requests output of an INIT file."
    OutputFile.WriteLine "INIT"
    OutputFile.WriteLine vbNullString
    
    OutputFile.WriteLine "DXV"
            
    i = 1
    PrintLine = vbNullString
        
    Do While i <= XDim
        If i <> XDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dXs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dXs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dXs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
        
    OutputFile.WriteLine "DYV"
        
    i = 1
    PrintLine = vbNullString
        
    Do While i <= YDim
        If i <> YDim Then
            If i Mod 5 <> 0 Then
                PrintLine = PrintLine & dYs(i) & vbTab & vbTab
            Else
                PrintLine = PrintLine & dYs(i) & vbCrLf & vbTab
            End If
                
        Else
            PrintLine = PrintLine & dYs(i)
        End If
            
        i = i + 1
    Loop
        
    OutputFile.WriteLine vbTab & PrintLine & vbTab & vbTab & "/"
    OutputFile.WriteLine vbNullString
    
    i = 1
        
    Do While i <= ZDim
        OutputFile.WriteLine "BOX"
        OutputFile.WriteLine vbTab & "1" & vbTab & vbTab & XDim & vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & _
            i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine vbNullString
            
        OutputFile.WriteLine "EQUALS"
        OutputFile.WriteLine vbTab & "'DZ'" & vbTab & vbTab & dZs(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'ACTNUM'" & vbTab & vbTab & ActiveFlags(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & LayerPerms(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'PORO'" & vbTab & vbTab & LayerPoros(i) & vbTab & vbTab & "/"
        OutputFile.WriteLine vbTab & "'TOPS'" & vbTab & vbTab & TopDepths(i) & vbTab & vbTab & "1" & vbTab & vbTab & XDim & _
            vbTab & vbTab & "1" & vbTab & vbTab & YDim & vbTab & vbTab & i & vbTab & vbTab & i & vbTab & vbTab & "/"
        OutputFile.WriteLine "/"
        OutputFile.WriteLine vbNullString
            
        i = i + 1
    Loop
    
    OutputFile.WriteLine "--********************Hydraulic Fracture Input********************"
    
    Sheets("Cross-Section (With Frac)").Activate
    
    Row = 4
    
    Do While Row <= 10000
        If Cells(Row, XDim + 1).Value = FracTopZ Then
            FracTopRow = Row
        End If
        
        If Cells(Row, XDim + 1).Value = FracBaseZ Then
            FracBottomRow = Row
            Exit Do
        End If
        
        Row = Row + 1
    Loop
    
    Row = FracTopRow
    
    Do While Row <= FracBottomRow
        Column = XDim + 12
        
        While Cells(Row, Column).Value <> vbNullString
            FirstColumn = Column
            
            Do While Column <= XDim + 12 + LastFrac * 2 + 1
                
                If Cells(Row, Column + 1).Value <> Cells(Row, Column).Value Then
                    OutputFile.WriteLine "BOX"
                    OutputFile.WriteLine vbTab & Mid(Cells(2, FirstColumn).Value, 5, 2) & vbTab & vbTab & _
                        Mid(Cells(2, Column).Value, 5, 2) & vbTab & vbTab & YDim \ 2 + 1 & vbTab & vbTab & YDim \ 2 + 1 & _
                        vbTab & vbTab & Cells(Row, XDim + 1).Value & vbTab & vbTab & Cells(Row, XDim + 1).Value & vbTab & vbTab & "/"
                    OutputFile.WriteLine vbNullString
                        
                    OutputFile.WriteLine "EQUALS"
                    
                    If FormVarFCD.OptUsedDx = False Then
                        OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & Cells(Row, Column).Value * _
                            (Sheets("Hydraulic Fracture").Cells(6, "C").Value / 12 / 5) & vbTab & vbTab & "/"
                        OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & Cells(Row, Column).Value * _
                            (Sheets("Hydraulic Fracture").Cells(6, "C").Value / 12 / 5) & vbTab & vbTab & "/"
                        OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & Cells(Row, Column).Value * _
                            (Sheets("Hydraulic Fracture").Cells(6, "C").Value / 12 / 5) & vbTab & vbTab & "/"
                    Else
                        OutputFile.WriteLine vbTab & "'PERMX'" & vbTab & vbTab & Cells(Row, Column).Value & vbTab & vbTab & "/"
                        OutputFile.WriteLine vbTab & "'PERMY'" & vbTab & vbTab & Cells(Row, Column).Value & vbTab & vbTab & "/"
                        OutputFile.WriteLine vbTab & "'PERMZ'" & vbTab & vbTab & Cells(Row, Column).Value & vbTab & vbTab & "/"
                    End If
                    
                    OutputFile.WriteLine "/"
                    OutputFile.WriteLine vbNullString
                    
                    FirstColumn = Column + 1
                End If
                
                Column = Column + 1
            Loop
        Wend
        
        Row = Row + 1
    Loop
    
    OutputFile.Close

End Sub

'================================================================================================================================
' CONVERTS THE INPUT RANGE OF NUMBERS TO THE MINIMUM/MAXIMUM COLUMN WIDTH AND ROW HEIGHT PROCESSED BY EXCEL
'================================================================================================================================
Function GetExcelDim(ByVal X As Double, ByVal Minimum As Double, ByVal Maximum As Double) As Double

    Const a As Double = 1.2 'Minimum row height/column width
    Const b As Double = 409.2 'Maximum row height/column width
    
    GetExcelDim = (((b - a) * (X - Minimum)) / (Maximum - Minimum)) + a

End Function
