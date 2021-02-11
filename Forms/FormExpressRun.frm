VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExpressRun 
   Caption         =   "Fracture Simulator - Run"
   ClientHeight    =   6840
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6735
   OleObjectBlob   =   "FormExpressRun.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExpressRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================================================================================
' OPENS THE FILE EXPLORER TO ALLOW THE USER TO SEARCH FOR AN .LAS FILE TO PARSE
'================================================================================================================================
Private Sub BtnBrowseLASFiles_Click()

    OpenLASFileExplorer

End Sub

'================================================================================================================================
' OPENS THE .LAS FILE COLUMN HEADERS FORM
'================================================================================================================================
Private Sub BtnColumnHeaders_Click()

    FormColumnHeaders.Show

End Sub

'================================================================================================================================
' ADVANCES THE USER ALONG THE FORM IF THEY HAVE ENTERED A VALID DEPTH INTERVAL; OTHERWISE DISPLAYS AN ERROR MESSAGE
'================================================================================================================================
Private Sub BtnDepthIntervalContinue_Click()
    
    If ValidDepthInterval = True Then
        DisableDepthIntervalFrame
        SwitchToDirectionalDataPage
        ParseLASFile
    End If
    
End Sub

'================================================================================================================================
' OPENS THE FILE EXPLORER TO ALLOW THE USER TO SEARCH FOR A SURVEY FILE TO PARSE
'================================================================================================================================
Private Sub BtnBrowseSurveyFiles_Click()

    OpenSurveyFileExplorer

End Sub

'================================================================================================================================
' OPENS THE SURVEY FILE COLUMN HEADERS FORM
'================================================================================================================================
Private Sub BtnColumnHeadersSurvey_Click()

    FormColumnHeadersSurvey.Show

End Sub

'================================================================================================================================
' SWITCHES TO THE "PETROPHYSICAL ANALYSIS" PAGE
'================================================================================================================================
Private Sub BtnSurveyContinue_Click()

    SwitchToPetrophysicalAnalysisPage
    ConvertDepths
    PetrophysicsAnalysis

End Sub

'================================================================================================================================
' ENABLES OR DISABLES THE RADIUS OF INVESTIGATION LABEL
'================================================================================================================================
Private Sub OptDefaultArea_Click()

    If OptDefaultArea.Value = True Then
        BtnWellTestingEqn.Enabled = False
    End If
    
    If LblRi.Visible = True Then
        LblRi.Enabled = False
        LblRiValue.Enabled = False
        LblRiFt.Enabled = False
    End If

End Sub

'================================================================================================================================
' ENABLES OR DISABLES THE RADIUS OF INVESTIGATION LABEL
'================================================================================================================================
Private Sub OptEqn_Click()
    
    If OptEqn.Value = True Then
        BtnWellTestingEqn.Enabled = True
    End If
    
    If LblRi.Visible = True Then
        LblRi.Enabled = True
        LblRiValue.Enabled = True
        LblRiFt.Enabled = True
    End If
    
End Sub

'================================================================================================================================
' OPENS THE RADIUS OF INVESTIGATION CALCULATON FORM
'================================================================================================================================
Private Sub BtnWellTestingEqn_Click()
    
    FormRadiusOfInvestigation.Show
    
End Sub

'================================================================================================================================
' ENABLES THE PRODUCTIVITY INDEX'S SUB-CONTROLS
'================================================================================================================================
Private Sub ChkProductivityIndex_Click()
    If ChkProductivityIndex.Value = True Then
        BtnPIDarcy.Enabled = True
    Else
        BtnPIDarcy.Enabled = False
    End If

End Sub

'================================================================================================================================
' DISPLAYS THE FORM TO CALCULATE PI USING DARCY FORMULATION
'================================================================================================================================
Private Sub BtnPIDarcy_Click()

    FormPIDarcy.Show

End Sub

'================================================================================================================================
' SWITCHES TO "ENGINEERING" ANALYSIS PAGE IF INPUT CRITERIA IS MET
'================================================================================================================================
Private Sub BtnEngineeringAnalysisContinue_Click()
    
    If OptDefaultArea.Value = False And OptEqn.Value = False Then
        LblEngineeringAnalysisError.Caption = "Please choose an option for radius of investigation."
        LblEngineeringAnalysisError.Visible = True
    ElseIf OptEqn.Value = True And (FormRadiusOfInvestigation.TxtTime.Text = vbNullString Or _
        FormRadiusOfInvestigation.TxtViscosity = vbNullString Or FormRadiusOfInvestigation.TxtComp = vbNullString) Then
        
        LblEngineeringAnalysisError.Caption = "Invalid radius of investigation input parameters."
        LblEngineeringAnalysisError.Visible = True
    ElseIf ChkProductivityIndex.Value = True And (FormPIDarcy.TxtViscosity.Text = vbNullString Or FormPIDarcy.TxtFVF.Text = vbNullString Or _
        FormPIDarcy.TxtRe.Text = vbNullString Or FormPIDarcy.TxtRw.Text = vbNullString) Then
            
        LblEngineeringAnalysisError.Caption = "Invalid PI (Darcy) input parameters."
        LblEngineeringAnalysisError.Visible = True
    Else
        SwitchToHydraulicFracturePage
        DisableAllEngineeringControls
        EngineeringsAnalysis
    End If
    
End Sub

'================================================================================================================================
' SWITCHES TO "RESERVOIR SIMULATION" PAGE IF INPUT CRITERIA IS MET
'================================================================================================================================
Private Sub BtnFractureContinue_Click()

    'Fracture half-length from user in ft
    Dim FracHL As String
    
    'Average fracture width from user in inches
    Dim FracWidth As String
    
    'Fracture height from user in ft
    Dim FracHeight As String
    
    'Fracture top depth from user in ft
    Dim FracTop As String
    
    'Dimensionless fracture conductivity from user
    Dim FCD As String
    
    'Top depth in ft TVD from the .LAS file
    Dim TopTVD As Double
    
    'Base depth in ft TVD from the .LAS file
    Dim BaseTVD As Double
    
    FracHL = TxtHL.Text
    FracWidth = TxtFracWidth.Text
    FracHeight = TxtFracHeight.Text
    FracTop = TxtFracTop.Text
    FCD = TxtFcd.Text
    TopTVD = Sheets(".LAS File Data").Cells(5, "C").Value
    BaseTVD = Sheets(".LAS File Data").Cells(Sheets(".LAS File Data").Cells(Rows.Count, "C").End(xlUp).Row, "C").Value
    
    If FracHL = vbNullString Then
        LblFractureError.Caption = "Please enter a fracture half-length."
        LblFractureError.Visible = True
    ElseIf Not IsNumeric(FracHL) Then
        LblFractureError.Caption = "An invalid character was entered in fracture half-length."
        LblFractureError.Visible = True
    ElseIf CDbl(FracHL) = 0 Then
        LblFractureError.Caption = "Fracture half-length cannot equal zero."
        LblFractureError.Visible = True
    ElseIf CDbl(FracHL) < 0 Then
        LblFractureError.Caption = "Fracture half-length cannot be negative."
        LblFractureError.Visible = True
    ElseIf FracWidth = vbNullString Then
        LblFractureError.Caption = "Please enter an average fracture width."
        LblFractureError.Visible = True
    ElseIf Not IsNumeric(FracWidth) Then
        LblFractureError.Caption = "An invalid character was entered in average fracture width."
        LblFractureError.Visible = True
    ElseIf CDbl(FracWidth) = 0 Then
        LblFractureError.Caption = "Average fracture width cannot equal zero."
        LblFractureError.Visible = True
    ElseIf CDbl(FracWidth) < 0 Then
        LblFractureError.Caption = "Average fracture width cannot be negative."
        LblFractureError.Visible = True
    ElseIf FracHeight = vbNullString Then
        LblFractureError.Caption = "Please enter a fracture height."
        LblFractureError.Visible = True
    ElseIf Not IsNumeric(FracHeight) Then
        LblFractureError.Caption = "An invalid character was entered in fracture height."
        LblFractureError.Visible = True
    ElseIf CDbl(FracHeight) = 0 Then
        LblFractureError.Caption = "Fracture height cannot equal zero."
        LblFractureError.Visible = True
    ElseIf CDbl(FracHeight) < 0 Then
        LblFractureError.Caption = "Fracture height cannot be negative."
        LblFractureError.Visible = True
    ElseIf FracTop = vbNullString Then
        LblFractureError.Caption = "Please enter a fracture top depth."
        LblFractureError.Visible = True
    ElseIf Not IsNumeric(FracTop) Then
        LblFractureError.Caption = "An invalid character was entered in fracture top depth."
        LblFractureError.Visible = True
    ElseIf CDbl(FracTop) = 0 Then
        LblFractureError.Caption = "Fracture top depth cannot equal zero."
        LblFractureError.Visible = True
    ElseIf CDbl(FracTop) < 0 Then
        LblFractureError.Caption = "Fracture top depth cannot be negative."
        LblFractureError.Visible = True
    ElseIf CDbl(FracTop) < TopTVD Then
        LblFractureError.Caption = "Fracture top depth cannot be less than the .LAS file top depth."
        LblFractureError.Visible = True
    ElseIf CDbl(FracTop) > BaseTVD Then
        LblFractureError.Caption = "Fracture top depth cannot be greater than the .LAS file base depth."
        LblFractureError.Visible = True
    ElseIf CDbl(FracTop) + CDbl(FracHeight) > BaseTVD Then
        LblFractureError.Caption = "Fracture top depth plus fracture height cannot be greater than the .LAS file base depth."
        LblFractureError.Visible = True
    ElseIf FCD = vbNullString Then
        LblFractureError.Caption = "Please enter a dimensionless fracture conductivity."
        LblFractureError.Visible = True
    ElseIf Not IsNumeric(FCD) Then
        LblFractureError.Caption = "An invalid character was entered in dimensionless fracture conductivity."
        LblFractureError.Visible = True
    ElseIf CDbl(FCD) = 0 Then
        LblFractureError.Caption = "Dimensionless fracture conductivity cannot equal zero."
        LblFractureError.Visible = True
    ElseIf CDbl(FCD) < 0 Then
        LblFractureError.Caption = "Dimensionless fracture conductivity cannot be negative."
        LblFractureError.Visible = True
    Else
        HydraulicFractures
        DisableFracturePage
        SwitchToSimulationPage
    End If

End Sub

'================================================================================================================================
' ENABLES OR DISABLES THE ROUGH SCALE GRID BUTTON
'================================================================================================================================
Private Sub OptRoughScaleGrid_Click()

    If BtnUpScaledGrid.Enabled = True Then
        BtnUpScaledGrid.Enabled = False
    End If

End Sub

'================================================================================================================================
' ENABLES OR DISABLES THE FINE SCALE GRID BUTTON
'================================================================================================================================
Private Sub OptFineScaleGrid_Click()

    If BtnUpScaledGrid.Enabled = True Then
        BtnUpScaledGrid.Enabled = False
    End If

End Sub

'================================================================================================================================
' ENABLES OR DISABLES THE UP-SCALED GRID BUTTON
'================================================================================================================================
Private Sub OptUpScaledGrid_Click()

    If OptUpScaledGrid.Value = True Then
        BtnUpScaledGrid.Enabled = True
    Else
        BtnUpScaledGrid.Enabled = False
    End If

End Sub

'================================================================================================================================
' DISPLAYS THE UP-SCALED GRID FORM
'================================================================================================================================
Private Sub BtnUpScaledGrid_Click()

    FormUpScaledGrid.Show

End Sub

'================================================================================================================================
' OPENS THE FILE EXPLORER TO SEARCH FOR A PATH FOR THE NON-FRACTURED GRID
'================================================================================================================================
Private Sub BtnBrowseNonFracturedGridPath_Click()

    Dim Path As String
    
    Path = GetFilePath("C:\", True, False)
    
    If Path <> vbNullString Then
        TxtNonFracturedGridPath.Text = Path
    End If
    
End Sub

'================================================================================================================================
' OPENS THE FILE EXPLORER TO SEARCH FOR A PATH FOR THE FRACTURED GRID
'================================================================================================================================
Private Sub BtnBrowseFracturedGridPath_Click()

    Dim Path As String
    
    Path = GetFilePath("C:\", False, True)
    
    If Path <> vbNullString Then
        TxtFracturedGridPath.Text = Path
    End If

End Sub

'================================================================================================================================
' GENERATES THE VISUAL GRID VIEWS AND INCLUDE FILES IF THE INPUT CRITERIA IS MET
'================================================================================================================================
Private Sub BtnGridContinue_Click()
    
    'Non-fractured grid file path and name from user
    Dim NonFracturedGridPath As String
    
    'Fractured grid file path and name from user
    Dim FracturedGridPath As String
    
    'Indicates whether or not the user has entered valid parameters if they have chosen the "up-scaled" grid option
    Dim ValidUpScaledParams As Boolean
    
    'Wellbore radius in ft from user
    Dim Rw As String
    
    NonFracturedGridPath = TxtNonFracturedGridPath.Text
    FracturedGridPath = TxtFracturedGridPath.Text
    Rw = TxtRw.Text
    
    With FormUpScaledGrid
    
        If .OptPermTol.Value = False And .OptPoroTol.Value = False And .OptBothTol.Value = False Then
            ValidUpScaledParams = False
        ElseIf (.OptPermTol.Value = True Or .OptBothTol.Value = True) And .OptPermNumeric.Value = False And .OptPermPercentage.Value = False Then
            ValidUpScaledParams = False
        ElseIf .TxtPermTolValue.Enabled = True And .TxtPermTolValue.Text = vbNullString Then
            ValidUpScaledParams = False
        ElseIf (.OptPoroTol.Value = True Or .OptBothTol.Value = True) And .OptPoroNumeric.Value = False And .OptPoroPercentage.Value = False Then
            ValidUpScaledParams = False
        ElseIf .TxtPoroTolValue.Enabled = True And .TxtPoroTolValue.Text = vbNullString Then
            ValidUpScaledParams = False
        Else
            ValidUpScaledParams = True
        End If
    
    End With
    
    If OptRoughScaleGrid.Value = False And OptFineScaleGrid.Value = False And OptUpScaledGrid.Value = False Then
        LblGridError.Caption = "Please choose a grid quality."
        LblGridError.Visible = True
    ElseIf OptUpScaledGrid.Value = True And ValidUpScaledParams = False Then
        LblGridError.Caption = "Invalid up-scaled grid parameters were entered."
        LblGridError.Visible = True
    ElseIf NonFracturedGridPath = vbNullString Then
        LblGridError.Caption = "Please choose a non-fractured grid file path."
        LblGridError.Visible = True
    ElseIf FracturedGridPath = vbNullString Then
        LblGridError.Caption = "Please choose a fractured grid file path."
        LblGridError.Visible = True
    ElseIf Rw = vbNullString Then
        LblGridError.Caption = "Please enter a wellbore radius."
        LblGridError.Visible = True
    ElseIf Not IsNumeric(Rw) Then
        LblGridError.Caption = "An invalid character was entered in wellbore radius."
        LblGridError.Visible = True
    ElseIf CDbl(Rw) = 0 Then
        LblGridError.Caption = "Wellbore radius cannot equal zero."
        LblGridError.Visible = True
    ElseIf CDbl(Rw) < 0 Then
        LblGridError.Caption = "Wellbore radius cannot be negative."
        LblGridError.Visible = True
    Else
        FormExpressRun.Hide
        GenerateCrossSections
        Sheets("Fracture Simulator").Activate
        MsgBox ("Your INCLUDE files have been generated. Enjoy!")
    End If

End Sub
