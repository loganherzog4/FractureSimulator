Attribute VB_Name = "LASParser"

'================================================================================================================================
' OPENS THE FILE EXPLORER TO ALLOW THE USER TO SEARCH FOR AN .LAS FILE TO PARSE
'================================================================================================================================
Sub OpenLASFileExplorer()

    Dim FileExplorer As Office.FileDialog 'File Explorer object
    Dim FileName As String 'LAS file path and name
    Dim FileNum As Integer 'FreeFile number
    Dim DataReached As Boolean 'Determines when the actual data to parse has been reached in the file
    Dim ParsedFile As String 'Stores whatever gets read from the file
    Dim ColumnHeadersString As String 'Column headers all as one string
    Dim ColumnHeadersArray() As String 'All column headers of the .LAS file
    Dim NumColumns As Integer 'Number of columns of data in the .LAS file
    Dim FileTopDepth As String 'The top depth from selected .LAS file
    Dim FileBaseDepth As String 'The base depth from the selected .LAS file
    Dim TopAndBase() As String 'Stores ParsedFile elements as an array
    
    Set FileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    With FileExplorer

      .AllowMultiSelect = False
      .Title = "Please select an .LAS file."
      .Filters.Clear
      .Filters.Add "Log ASCII Standard", "*.las"
      .Filters.Add "Text Documents", "*.txt"

      If .Show = True Then
        FileName = .SelectedItems(1)
        
        FormExpressRun.TxtSelectedLASFile = FileName
        
        FileNum = FreeFile()
        
        Open FileName For Input As #FileNum
        
        Do While Not EOF(FileNum)
            Line Input #FileNum, DataLine
            
            '=============================================================
            ' Only applicable for Nexen-specific files, not general cases.
            '=============================================================
            'If InStr(DataLine, "STRT .FT") Then
                'ParsedFile = ParsedFile + DataLine
            'End If
            
            'If InStr(DataLine, "STOP .FT") Then
                'ParsedFile = ParsedFile + DataLine
                'Exit Do
            'End If
            
            If DataReached Then
                ParsedFile = ParsedFile + DataLine
            End If
            
            If InStr(DataLine, "~A") Then
                DataReached = True
                ColumnHeadersString = DataLine
                ColumnHeadersString = Replace(ColumnHeadersString, "          ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "         ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "        ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "       ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "      ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "     ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "    ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "   ", " ")
                ColumnHeadersString = Replace(ColumnHeadersString, "  ", " ")
                ColumnHeadersString = Trim(ColumnHeadersString)
                
                ColumnHeadersArray() = Split(ColumnHeadersString, " ")
                NumColumns = UBound(ColumnHeadersArray) - LBound(ColumnHeadersArray)
            End If
        Loop
        
        DataReached = False
        
        Close #FileNum
        
        ParsedFile = Replace(ParsedFile, "          ", " ")
        ParsedFile = Replace(ParsedFile, "         ", " ")
        ParsedFile = Replace(ParsedFile, "        ", " ")
        ParsedFile = Replace(ParsedFile, "       ", " ")
        ParsedFile = Replace(ParsedFile, "      ", " ")
        ParsedFile = Replace(ParsedFile, "     ", " ")
        ParsedFile = Replace(ParsedFile, "    ", " ")
        ParsedFile = Replace(ParsedFile, "   ", " ")
        ParsedFile = Replace(ParsedFile, "  ", " ")
        ParsedFile = Replace(ParsedFile, ":", " ")
        ParsedFile = Trim(ParsedFile)
        
        TopAndBase() = Split(ParsedFile, " ")
        
        'For i = 0 To UBound(TopAndBase)
            'If IsNumeric(TopAndBase(i)) Then
                'If FileTopDepth = vbNullString Then
                    'FileTopDepth = TopAndBase(i)
                'Else
                    'FileBaseDepth = TopAndBase(i)
                'End If
            'End If
        'Next i
        
        FileTopDepth = TopAndBase(0)
        FileBaseDepth = TopAndBase((UBound(TopAndBase) - LBound(TopAndBase) + 1) - NumColumns)
        
        With FormExpressRun
            
            .TxtFileTopDepth.Text = FileTopDepth
            .TxtFileBaseDepth.Text = FileBaseDepth
            
            .FrameInputFile.Enabled = False
            .LblSelectedLASFile.Enabled = False
            .LblFileTopDepth.Enabled = False
            .LblFileBaseDepth.Enabled = False
            
            .TxtSelectedLASFile.BackColor = vbButtonFace
            .TxtFileTopDepth.BackColor = vbButtonFace
            .TxtFileBaseDepth.BackColor = vbButtonFace
            .BtnBrowseLASFiles.Enabled = False
            
            .FrameDepthInterval.Enabled = True
            .LblInputTopDepth.Enabled = True
            .TxtInputTopDepth.Enabled = True
            .TxtInputTopDepth.BackColor = vbWindowBackground
            .LblInputBaseDepth.Enabled = True
            .TxtInputBaseDepth.Enabled = True
            .TxtInputBaseDepth.BackColor = vbWindowBackground
            .LblDepthIntervalErrors.Enabled = True
            .LblDepthIntervalError.Enabled = True
            .BtnDepthIntervalContinue.Enabled = True
            
            .FrameColumnHeaders.Enabled = True
            .LblColumnHeaders.Enabled = True
            .BtnColumnHeaders.Enabled = True
            
        End With

      Else
        MsgBox "A file was not selected.", vbExclamation, "Macro Terminated"
        End
      End If
   End With

End Sub

'================================================================================================================================
' CHECKS IF THE USER-INPUTTED DEPTH INTERVAL IS VALID FOR THE SELECTED .LAS FILE
'================================================================================================================================
Function ValidDepthInterval() As Boolean

    Dim FileTopDepth As String 'Top depth from the .LAS file
    Dim FileBaseDepth As String 'Base depth from the .LAS file
    Dim InputTopDepth As String 'Top depth that the user has entered
    Dim InputBaseDepth As String 'Base depth that the user has entered
    
    FileTopDepth = FormExpressRun.TxtFileTopDepth.Text
    FileBaseDepth = FormExpressRun.TxtFileBaseDepth.Text
    InputTopDepth = FormExpressRun.TxtInputTopDepth.Text
    InputBaseDepth = FormExpressRun.TxtInputBaseDepth.Text
    
    If InputTopDepth = vbNullString Then
        FormExpressRun.LblDepthIntervalError.Caption = "Please enter a valid top depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf FileTopDepth = vbNullString Then
        FormExpressRun.LblDepthIntervalError.Caption = "The file cannot be parsed correctly. Please double check for errors."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf Not IsNumeric(InputTopDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "An invalid character was entered in top depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputTopDepth) < 0 Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input top depth value cannot be negative."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputTopDepth) < CDbl(FileTopDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input top depth value cannot be less than the file's top depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputTopDepth) > CDbl(FileBaseDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input top depth value cannot be greater than the file base depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf InputBaseDepth = vbNullString Then
        FormExpressRun.LblDepthIntervalError.Caption = "Please enter a valid base depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf Not IsNumeric(InputBaseDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "An invalid character was entered in base depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputBaseDepth) < 0 Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input base depth value cannot be negative."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputBaseDepth) < CDbl(FileTopDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input base depth value cannot be less than the file's top depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputBaseDepth) > CDbl(FileBaseDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input base depth value cannot be greater than the file's base depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    ElseIf CDbl(InputBaseDepth) < CDbl(InputTopDepth) Then
        FormExpressRun.LblDepthIntervalError.Caption = "Input base depth value cannot be less than the input top depth."
        FormExpressRun.LblDepthIntervalError.Visible = True
        ValidDepthInterval = False
    Else
        ValidDepthInterval = True
    End If

End Function

'================================================================================================================================
' DISABLES ALL CONTROLS IN THE DEPTH INTERVAL FRAME
'================================================================================================================================
Sub DisableDepthIntervalFrame()

    With FormExpressRun
        .FrameDepthInterval.Enabled = False
        .LblInputTopDepth.Enabled = False
        .TxtInputTopDepth.Enabled = False
        .TxtInputTopDepth.BackColor = vbButtonFace
        .LblInputBaseDepth.Enabled = False
        .TxtInputBaseDepth.Enabled = False
        .TxtInputBaseDepth.BackColor = vbButtonFace
        .LblDepthIntervalErrors.Enabled = False
        .LblDepthIntervalError.Visible = False
        .BtnDepthIntervalContinue.Enabled = False
    End With

End Sub

'================================================================================================================================
' ENABLES AND SWITCHES TO THE DIRECTIONAL DATA PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub SwitchToDirectionalDataPage()

    With FormExpressRun.MultiPageExpressRun
        .Pages(1).Enabled = True
        .Value = .Value + 1
    End With

End Sub

'================================================================================================================================
' PARSES THE SPECIFIED .LAS FILE AND CREATES A TABLE OF EXTRACTED DATA IN A NEW SHEET
'================================================================================================================================
Sub ParseLASFile()

    Dim FileName As String 'The .LAS file to be parsed
    Dim TopDepth As String 'Depth to start reading from
    Dim BaseDepth As String 'Depth to stop reading at
    Dim ParsedFile As String 'Stores whatever has been read from the .LAS file
    Dim ValueArray() As String 'An array of solely the words/numbers that have been read from the .LAS file (no white spaces)
    Dim RecordedKB As Boolean 'Indicates whether or not the KB height has been read already
    Dim KBHeight As String 'Elevation measurement reference from the .LAS file (Kelly Bushing height)
    Dim FileNum As Integer 'FreeFile number
    Dim ReachedDepths As Boolean 'Indicates whether or not the file parser has reached the depth readings
    Dim ReachedTop As Boolean 'Indicates whether or not the file parser has reached the top depth specified by the user
    Dim ColPerm As Integer 'Indicates the column number of permeabilities from the .LAS file
    Dim ColPorosity As Integer 'Indicates the column number of porosities from the .LAS file
    Dim ColWaterSat As Integer 'Indicates the column number of water saturations from the .LAS file
    Dim ColPay As Integer 'Indicates the column number of pay flags from the .LAS file
    Dim ColRes As Integer 'Indicates the column number of reservoir flags from the .LAS file
    Dim Pos As Integer 'String position
    Dim Row As Long 'Row index
    Dim LASNull As String 'The first null value used for the .LAS file to determine undefined data points
    Dim LASNull2 As String 'The second null value used for the .LAS file to determine undefined data points
    Dim LastRow As Long 'Row index for the last row of .LAS file data
    Dim RectLeft As Double 'Rectangle shape dimension
    Dim RectTop As Double 'Rectangle shape dimension
    Dim RectWidth As Double 'Rectangle shape dimension
    Dim RectHeight As Double 'Rectangle shape dimension
    Dim Rect As Shape 'Blue rectangle title
    Dim DepthHeader As String 'User-inputted or default depth column header
    Dim PermHeader As String 'User-inputted or default permeability column header
    Dim PoroHeader As String 'User-inputted or default porosity column header
    Dim SwHeader As String 'User-inputted or default water saturation column header
    Dim PayHeader As String 'User-inputted or default pay flag column header
    Dim ResHeader As String 'User-inputted or default reservoir flag column header
    
    Sheets.Add
    ActiveSheet.Name = ".LAS File Data"
    
    Cells(4, "A").Font.Bold = True
    Cells(4, "B").Font.Bold = True
    Cells(4, "C").Font.Bold = True
    Cells(4, "D").Font.Bold = True
    Cells(4, "E").Font.Bold = True
    Cells(4, "F").Font.Bold = True
    Cells(4, "G").Font.Bold = True
    Cells(4, "H").Font.Bold = True
    Cells(4, "I").Font.Bold = True
    Cells(4, "J").Font.Bold = True
    Cells(4, "K").Font.Bold = True
    
    Cells(4, "A").Value = "KB Height (ft)"
    Cells(4, "B").Value = "Measured Depth (ft)"
    Cells(4, "C").Value = "Total Vertical Depth (ft)"
    Cells(4, "D").Value = "Total Vertical Depth Subsea (ft)"
    Cells(4, "E").Value = "Permeability (mD)"
    Cells(4, "F").Value = "Porosity (Fraction)"
    Cells(4, "G").Value = "Porosity (%)"
    Cells(4, "H").Value = "Water Saturation (Fraction)"
    Cells(4, "I").Value = "Water Saturation (%)"
    Cells(4, "J").Value = "Pay Flag"
    Cells(4, "K").Value = "Reservoir Flag"
    
    With FormExpressRun
        
        FileName = .TxtSelectedLASFile.Text
        TopDepth = .TxtInputTopDepth.Text
        BaseDepth = .TxtInputBaseDepth.Text
        
    End With
    
    With FormColumnHeaders
        If .TxtDepthHeader.Text = vbNullString Then
            DepthHeader = "DEPTH"
        Else
            DepthHeader = .TxtDepthHeader.Text
        End If
        
        If .TxtPermHeader.Text = vbNullString Then
            PermHeader = "KAIR_NXN"
        Else
            PermHeader = .TxtPermHeader.Text
        End If
        
        If .TxtPoroHeader.Text = vbNullString Then
            PoroHeader = "PHIE_NXN"
        Else
            PoroHeader = .TxtPoroHeader.Text
        End If
        
        'If .TxtPayHeader.Text = vbNullString Then
            'PayHeader = "PAYFLG_NXN"
        'Else
            PayHeader = .TxtPayHeader.Text
        'End If
        
        'If .TxtResHeader.Text = vbNullString Then
            'ResHeader = "RESFLG_NXN"
        'Else
            ResHeader = .TxtResHeader.Text
        'End If
        
        If .TxtSwHeader.Text = vbNullString Then
            SwHeader = "SWT_NXN"
        Else
            SwHeader = .TxtSwHeader.Text
        End If
    End With
    
    Row = 5
    LASNull = "-999.2500"
    LASNull2 = "-999.250000"
        
    FileNum = FreeFile()
        
    Open FileName For Input As #FileNum
        
    Do While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        
        If RecordedKB = False Then
            If InStr(DataLine, "ELEV_MEAS_REF.FT") Then
                ParsedFile = DataLine
                
                ParsedFile = Replace(ParsedFile, "          ", " ")
                ParsedFile = Replace(ParsedFile, "         ", " ")
                ParsedFile = Replace(ParsedFile, "        ", " ")
                ParsedFile = Replace(ParsedFile, "       ", " ")
                ParsedFile = Replace(ParsedFile, "      ", " ")
                ParsedFile = Replace(ParsedFile, "     ", " ")
                ParsedFile = Replace(ParsedFile, "    ", " ")
                ParsedFile = Replace(ParsedFile, "   ", " ")
                ParsedFile = Replace(ParsedFile, "  ", " ")
                ParsedFile = Replace(ParsedFile, ":", " ")
                ParsedFile = Trim(ParsedFile)
                
                ValueArray = Split(ParsedFile, " ")
                KBHeight = ValueArray(1)
                Cells(5, "A").Value = KBHeight
                Erase ValueArray
                ParsedFile = vbNullString
                RecordedKB = True
            End If
        End If
        
        If InStr(DataLine, "~A") Then
            ParsedFile = DataLine
            
            ParsedFile = Replace(ParsedFile, "          ", " ")
            ParsedFile = Replace(ParsedFile, "         ", " ")
            ParsedFile = Replace(ParsedFile, "        ", " ")
            ParsedFile = Replace(ParsedFile, "       ", " ")
            ParsedFile = Replace(ParsedFile, "      ", " ")
            ParsedFile = Replace(ParsedFile, "     ", " ")
            ParsedFile = Replace(ParsedFile, "    ", " ")
            ParsedFile = Replace(ParsedFile, "   ", " ")
            ParsedFile = Replace(ParsedFile, "  ", " ")
            ParsedFile = Trim(ParsedFile)
            
            ValueArray = Split(ParsedFile, " ")
            
            For i = 1 To UBound(ValueArray)
                Select Case ValueArray(i)
                    Case Is = PermHeader
                        ColPerm = i - 1
                    Case Is = PayHeader
                        ColPay = i - 1
                    Case Is = ResHeader
                        ColRes = i - 1
                    Case Is = PoroHeader
                        ColPorosity = i - 1
                    Case Is = SwHeader
                        ColWaterSat = i - 1
                End Select
            Next i
            
            If ColWaterSat = 0 Then
                For i = 1 To UBound(ValueArray)
                    If ValueArray(i) = "SWT_NXN" Then
                        ColWaterSat = i - 1
                    End If
                Next i
            End If
            
            ReachedDepths = True
            
            ParsedFile = vbNullString
            Erase ValueArray
        End If
        
        If ReachedDepths = True Then
            Pos = InStr(DataLine, TopDepth)
            
            If Pos <> 0 And Pos < 20 Then
                ReachedTop = True
            End If
            
            If ReachedTop = True Then
                ParsedFile = DataLine
                
                ParsedFile = Replace(ParsedFile, "          ", " ")
                ParsedFile = Replace(ParsedFile, "         ", " ")
                ParsedFile = Replace(ParsedFile, "        ", " ")
                ParsedFile = Replace(ParsedFile, "       ", " ")
                ParsedFile = Replace(ParsedFile, "      ", " ")
                ParsedFile = Replace(ParsedFile, "     ", " ")
                ParsedFile = Replace(ParsedFile, "    ", " ")
                ParsedFile = Replace(ParsedFile, "   ", " ")
                ParsedFile = Replace(ParsedFile, "  ", " ")
                ParsedFile = Trim(ParsedFile)
                
                ValueArray = Split(ParsedFile, " ")
                
                Cells(Row, "B").Value = ValueArray(0)
                
                If ValueArray(ColPerm) = LASNull Or ValueArray(ColPerm) = LASNull2 Then
                    Cells(Row, "E").Value = 0
                Else
                    Cells(Row, "E").Value = ValueArray(ColPerm)
                End If
                
                If ValueArray(ColPorosity) = LASNull Or ValueArray(ColPorosity) = LASNull2 Then
                    Cells(Row, "F").Value = 0
                    Cells(Row, "G").Value = 0
                Else
                    If FormColumnHeaders.ChkPoroPercentage Then
                        Cells(Row, "F").Value = ValueArray(ColPorosity) / 100
                        Cells(Row, "G").Value = ValueArray(ColPorosity)
                    Else
                        Cells(Row, "F").Value = ValueArray(ColPorosity)
                        Cells(Row, "G").Value = CDbl(Cells(Row, "F").Value) * 100
                    End If
                End If
                
                If ValueArray(ColWaterSat) = LASNull Or ValueArray(ColWaterSat) = LASNull2 Then
                    Cells(Row, "H").Value = 0
                    Cells(Row, "I").Value = 0
                Else
                    If FormColumnHeaders.ChkSwPercentage Then
                        Cells(Row, "H").Value = ValueArray(ColWaterSat) / 100
                        Cells(Row, "I").Value = ValueArray(ColWaterSat)
                    Else
                        Cells(Row, "H").Value = ValueArray(ColWaterSat)
                        Cells(Row, "I").Value = CDbl(Cells(Row, "H").Value) * 100
                    End If
                End If
                
                If ValueArray(ColPay) = 1 Or PayHeader = vbNullString Then
                    Cells(Row, "J").Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(Row, "J").Interior.Color = RGB(255, 0, 0)
                End If
                
                If ValueArray(ColRes) = 1 Or ResHeader = vbNullString Then
                    Cells(Row, "K").Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(Row, "K").Interior.Color = RGB(255, 0, 0)
                End If
                
                ParsedFile = vbNullString
                Erase ValueArray
                Row = Row + 1
            End If
            
            Pos = InStr(DataLine, BaseDepth)
            
            If Pos <> 0 And Pos < 20 Then
                Exit Do
            End If
        End If
    Loop
    
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    With Columns("A:K")
        .AutoFit
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("A4:A5")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("B4:K" & LastRow)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
    End With
    
    With Range("A4:K4")
        .Borders.Weight = xlThick
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("A1:K2")
        RectLeft = .Left
        RectTop = .Top
        RectHeight = .Height
        RectWidth = .Width
    End With
    
    Set Rect = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, RectLeft, RectTop, RectWidth, RectHeight)
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, RectLeft, RectTop, RectWidth, RectHeight).Name = ".LAS File Data"
    
    With ActiveSheet.Shapes(".LAS File Data")
        .TextFrame.Characters.Text = ".LAS FILE DATA"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = True
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
    End With
    
End Sub

