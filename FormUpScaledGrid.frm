VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormUpScaledGrid 
   Caption         =   "Reservoir Simulation - Up-Scaled Grid"
   ClientHeight    =   10560
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4005
   OleObjectBlob   =   "FormUpScaledGrid.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormUpScaledGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================================================================================
' ENABLES THE PERMEABILITY TOLERANCE CONTROLS AND DISABLES THE POROSITY TOLERANCE CONTROLS
'================================================================================================================================
Private Sub OptPermTol_Click()

    If OptPermTol.Value = True Then
    
        FramePermTol.Enabled = True
        LblPermTol.Enabled = True
        OptPermNumeric.Enabled = True
        OptPermPercentage.Enabled = True
        LblPermTolValue.Enabled = True
        TxtPermTolValue.Enabled = True
        
        FramePoroTol.Enabled = False
        LblPoroTol.Enabled = False
        LblPoroTolNote.Enabled = False
        OptPoroNumeric.Enabled = False
        OptPoroPercentage.Enabled = False
        LblPoroTolValue.Enabled = False
        TxtPoroTolValue.Enabled = False
        
    End If

End Sub

'================================================================================================================================
' ENABLES THE POROSITY TOLERANCE CONTROLS AND DISABLES THE PERMEABILITY TOLERANCE CONTROLS
'================================================================================================================================
Private Sub OptPoroTol_Click()

    If OptPoroTol.Value = True Then
    
        FramePoroTol.Enabled = True
        LblPoroTol.Enabled = True
        LblPoroTolNote.Enabled = True
        OptPoroNumeric.Enabled = True
        OptPoroPercentage.Enabled = True
        LblPoroTolValue.Enabled = True
        TxtPoroTolValue.Enabled = True
        
        FramePermTol.Enabled = False
        LblPermTol.Enabled = False
        OptPermNumeric.Enabled = False
        OptPermPercentage.Enabled = False
        LblPermTolValue.Enabled = False
        TxtPermTolValue.Enabled = False
        
    End If

End Sub

'================================================================================================================================
' ENABLES BOTH THE PERMEABILITY AND POROSITY TOLERANCE CONTROLS
'================================================================================================================================
Private Sub OptBothTol_Click()

    If OptBothTol.Value = True Then
    
        FramePermTol.Enabled = True
        LblPermTol.Enabled = True
        OptPermNumeric.Enabled = True
        OptPermPercentage.Enabled = True
        LblPermTolValue.Enabled = True
        TxtPermTolValue.Enabled = True
        
        FramePoroTol.Enabled = True
        LblPoroTol.Enabled = True
        LblPoroTolNote.Enabled = True
        OptPoroNumeric.Enabled = True
        OptPoroPercentage.Enabled = True
        LblPoroTolValue.Enabled = True
        TxtPoroTolValue.Enabled = True
        
    End If

End Sub

'================================================================================================================================
' CHECKS FOR VALID INPUT PARAMETERS
'================================================================================================================================
Private Sub BtnOK_Click()

    'Numeric or percentage permeability tolerance from user
    Dim PermTol As String
    
    'Numeric or percentage porosity tolerance from user
    Dim PoroTol As String
    
    PermTol = TxtPermTolValue.Text
    PoroTol = TxtPoroTolValue.Text
    
    If OptPermTol.Value = False And OptPoroTol.Value = False And OptBothTol.Value = False Then
        LblUpScaledError.Caption = "Please make an up-scaling decision."
        LblUpScaledError.Visible = True
    ElseIf (OptPermTol.Value = True Or OptBothTol.Value = True) And OptPermNumeric.Value = False And OptPermPercentage.Value = False Then
        LblUpScaledError.Caption = "Please choose a permeability tolerance type."
        LblUpScaledError.Visible = True
    ElseIf TxtPermTolValue.Enabled = True And PermTol = vbNullString Then
        LblUpScaledError.Caption = "Please enter a permeability tolerance."
        LblUpScaledError.Visible = True
    ElseIf TxtPermTolValue.Enabled = True And Not IsNumeric(PermTol) Then
        LblUpScaledError.Caption = "An invalid character was entered in permeability tolerance."
        LblUpScaledError.Visible = True
    ElseIf TxtPermTolValue.Enabled = True And Val(PermTol) < 0 Then
        LblUpScaledError.Caption = "Permeability tolerance cannot be negative."
        LblUpScaledError.Visible = True
    ElseIf (OptPoroTol.Value = True Or OptBothTol.Value = True) And OptPoroNumeric.Value = False And OptPoroPercentage.Value = False Then
        LblUpScaledError.Caption = "Please choose a porosity tolerance type."
        LblUpScaledError.Visible = True
    ElseIf TxtPoroTolValue.Enabled = True And PoroTol = vbNullString Then
        LblUpScaledError.Caption = "Please enter a porosity tolerance."
        LblUpScaledError.Visible = True
    ElseIf TxtPoroTolValue.Enabled = True And Not IsNumeric(PoroTol) Then
        LblUpScaledError.Caption = "An invalid character was entered in porosity tolerance."
        LblUpScaledError.Visible = True
    ElseIf TxtPoroTolValue.Enabled = True And Val(PoroTol) < 0 Then
        LblUpScaledError.Caption = "Porosity tolerance cannot be negative."
        LblUpScaledError.Visible = True
    Else
        LblUpScaledError.Visible = False
        FormUpScaledGrid.Hide
    End If

End Sub

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES OR THE ERROR LABEL
'================================================================================================================================
Private Sub BtnClear_Click()

    If TxtPermTolValue.Text <> vbNullString Or TxtPoroTolValue.Text <> vbNullString Then
        TxtPermTolValue.Text = vbNullString
        TxtPoroTolValue.Text = vbNullString
        LblUpScaledError.Caption = "The variable text boxes have been cleared."
        LblUpScaledError.Visible = True
    Else
        LblUpScaledError.Visible = False
    End If

End Sub

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES AND CLOSES THE FORM
'================================================================================================================================
Private Sub BtnCancel_Click()

    OptPermTol.Value = False
    OptPoroTol.Value = False
    OptBothTol.Value = False
    OptPermNumeric.Value = False
    OptPermPercentage.Value = False
    TxtPermTolValue.Text = vbNullString
    OptPoroNumeric.Value = False
    OptPoroPercentage.Value = False
    TxtPoroTolValue.Text = vbNullString
    LblUpScaledError.Visible = False
    
    FormUpScaledGrid.Hide

End Sub





