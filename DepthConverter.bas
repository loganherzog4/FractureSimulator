Attribute VB_Name = "DepthConverter"

'================================================================================================================================
' OPENS THE FILE EXPLORER TO ALLOW THE USER TO SEARCH FOR A SURVEY FILE TO PARSE
'================================================================================================================================
Sub OpenSurveyFileExplorer()

    Dim FileExplorer As Office.FileDialog 'File Explorer object
    Dim FileName As String 'LAS file path and name
    
    Set FileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    With FileExplorer

      .AllowMultiSelect = False
      .Title = "Please select a survey file."
      .Filters.Clear
      .Filters.Add "Text Documents", "*.txt"

      If .Show = True Then
        FileName = .SelectedItems(1)
        FormExpressRun.TxtSelectedSurveyFile = FileName
        DisableInputSurveyFileFrame
        EnableColumnHeadersFrame
      Else
        MsgBox "A file was not selected.", vbExclamation, "Macro Terminated"
        End
      End If
   End With

End Sub

'================================================================================================================================
' DISABLES THE "INPUT SURVEY FILE" FRAME ON THE DIRECTIONAL DATA PAGE OF THE MULTIPAGE FORM
'================================================================================================================================
Private Sub DisableInputSurveyFileFrame()

    With FormExpressRun
        .FrameInputSurveyFile.Enabled = False
        .LblSelectedSurveyFile.Enabled = False
        .TxtSelectedSurveyFile.Enabled = False
        .TxtSelectedSurveyFile.BackColor = vbButtonFace
        .BtnBrowseSurveyFiles.Enabled = False
    End With

End Sub

'================================================================================================================================
' ENABLES THE "COLUMN HEADERS" FRAME FOR THE SURVEY FILE
'================================================================================================================================
Private Sub EnableColumnHeadersFrame()

    With FormExpressRun
        .FrameColumnHeadersSurvey.Enabled = True
        .LblColumnHeadersSurvey.Enabled = True
        .BtnColumnHeadersSurvey.Enabled = True
        .BtnSurveyContinue.Enabled = True
    End With

End Sub

'================================================================================================================================
' ENABLES AND SWITCHES TO THE PETROPHYSICAL ANALYSIS PAGE OF THE MULTIPAGE CONTROL
'================================================================================================================================
Sub SwitchToPetrophysicalAnalysisPage()

    With FormExpressRun.MultiPageExpressRun
        .Pages(2).Enabled = True
        .Value = .Value + 1
    End With

End Sub

'================================================================================================================================
' CONVERTS MEASURED DEPTH TO TVD BY INTERPOLATING USING THE SURVEY FILE, AND TVD TO TVDSS USING THE KB HEIGHT
'================================================================================================================================
Sub ConvertDepths()

    Dim FileName As String 'Survey file path and name
    Dim FileNum As Integer 'FreeFile number
    Dim ParsedFile As String 'Stores the DataLine currently being read from the file being parsed
    Dim MDHeader As String 'User-inputted or default measured depth column header
    Dim TVDHeader As String 'User-inputted or default vertical depth column header
    Dim ColMD As Integer 'Column index for MD column in the survey file
    Dim ColTVD As Integer 'Column index for TVD column in the survey file
    Dim ReachedData As Boolean 'Indicates whether or not the file parser has reached the data values
    Dim Row As Long 'Row index
    Dim i As Integer 'Count variable for loops
    Dim MDs() As Double 'Measured Depth column from the survey file
    Dim TVDs() As Double 'TVD column from the survey file
    Dim LastRow As Long 'Last row index of .LAS file data
    Dim MD As Double 'Current MD value
    Dim LowMD As Double 'The closest MD value in the survey file that is less than the current MD value
    Dim HighMD As Double 'The closest MD value in the survey file that is less than the current MD value
    Dim LowTVD As Double 'The TVD value in the survey file corresponding to LowMD
    Dim HighTVD As Double 'The TVD value in the survey file corresponding to HighMD
    Dim Slope As Double 'Interpolation slope
    Dim ValueArray() As String 'Column headers in directional survey
    
    FileName = FormExpressRun.TxtSelectedSurveyFile.Text
    FileNum = FreeFile()
    
    Row = 5
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    ReDim ValueArray(100)
    ReDim MDs(10000)
    ReDim TVDs(10000)
        
    With FormColumnHeadersSurvey
        If .TxtMDHeader.Text = vbNullString Then
            MDHeader = "Measured Depth"
        Else
            MDHeader = .TxtMDHeader.Text
        End If
        
        If .TxtTVDHeader.Text = vbNullString Then
            TVDHeader = "Vertical Depth"
        Else
            TVDHeader = .TxtTVDHeader.Text
        End If
    End With
    
    Open FileName For Input As #FileNum
        
    Do While Not EOF(FileNum)
        Line Input #FileNum, DataLine
        
        If InStr(DataLine, MDHeader) Then
            ParsedFile = DataLine
            
            ParsedFile = Replace(ParsedFile, "          ", " ")
            ParsedFile = Replace(ParsedFile, "         ", " ")
            ParsedFile = Replace(ParsedFile, "        ", " ")
            ParsedFile = Replace(ParsedFile, "       ", " ")
            ParsedFile = Replace(ParsedFile, "      ", " ")
            ParsedFile = Replace(ParsedFile, "     ", " ")
            ParsedFile = Replace(ParsedFile, "    ", " ")
            ParsedFile = Replace(ParsedFile, "   ", " ")
            ParsedFile = Replace(ParsedFile, vbTab, " ")
            ParsedFile = Trim(ParsedFile)
            
            ValueArray = Split(ParsedFile, " ")
            
            For i = 0 To UBound(ValueArray)
                If ValueArray(i) = "!" & MDHeader Or ValueArray(i) = MDHeader Then
                    ColMD = i
                End If
                
                If ValueArray(i) = "!" & TVDHeader Or ValueArray(i) = TVDHeader Then
                    ColTVD = i
                End If
            Next i
            
            i = 0
        End If
        
        If IsNumeric(Left(DataLine, 1)) Then
            ReachedData = True
        End If
        
        If ReachedData = True Then
            ParsedFile = DataLine
            
            ParsedFile = Replace(ParsedFile, "          ", " ")
            ParsedFile = Replace(ParsedFile, "         ", " ")
            ParsedFile = Replace(ParsedFile, "        ", " ")
            ParsedFile = Replace(ParsedFile, "       ", " ")
            ParsedFile = Replace(ParsedFile, "      ", " ")
            ParsedFile = Replace(ParsedFile, "     ", " ")
            ParsedFile = Replace(ParsedFile, "    ", " ")
            ParsedFile = Replace(ParsedFile, "   ", " ")
            ParsedFile = Replace(ParsedFile, vbTab, " ")
            ParsedFile = Trim(ParsedFile)
            
            ValueArray = Split(ParsedFile, " ")
            
            MDs(i) = CDbl(ValueArray(ColMD))
            TVDs(i) = CDbl(ValueArray(ColTVD))
            
            i = i + 1
        End If
        
    Loop
    
    Do While Row <= LastRow
        MD = Cells(Row, "B").Value
        
        i = 0
        
        Do While i < UBound(MDs)
            If MDs(i) <= MD And MDs(i + 1) > MD Then
                LowMD = MDs(i)
                LowTVD = TVDs(i)
                HighMD = MDs(i + 1)
                HighTVD = TVDs(i + 1)
                Exit Do
            End If
            
            i = i + 1
        Loop
        
        If HighMD - LowMD <> 0 Then
            Slope = (HighTVD - LowTVD) / (HighMD - LowMD)
            Cells(Row, "C").Value = LowTVD + (Slope * (MD - LowMD))
            Cells(Row, "D").Value = Cells(Row, "C").Value - Cells(5, "A").Value
        End If
        
        Row = Row + 1
    Loop

End Sub
