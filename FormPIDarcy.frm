VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPIDarcy 
   Caption         =   "Productivity Index - Darcy Law Form"
   ClientHeight    =   8310
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4005
   OleObjectBlob   =   "FormPIDarcy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPIDarcy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================================================================================
' CHECKS FOR VALID INPUT PARAMETERS
'================================================================================================================================
Private Sub BtnOK_Click()

    'Reservoir fluid viscosity from user in cP
    Dim Viscosity As String
    
    'Formation volume factor from user in Res BBL/STB
    Dim FVF As String
    
    'Drainage radius from user in ft
    Dim Re As String
    
    'Wellbore radius from user in ft
    Dim Rw As String
    
    'Skin factor from user
    Dim Sf As String
    
    Viscosity = TxtViscosity.Text
    FVF = TxtFVF.Text
    Re = TxtRe.Text
    Rw = TxtRw.Text
    Sf = TxtSf.Text
    
    If Viscosity = vbNullString Then
        LblError.Caption = "Please enter a reservoir fluid viscosity."
        LblError.Visible = True
    ElseIf Not IsNumeric(Viscosity) Then
        LblError.Caption = "An invalid character was entered in reservoir fluid viscosity."
        LblError.Visible = True
    ElseIf CDbl(Viscosity) = 0 Then
        LblError.Caption = "Reservoir fluid viscosity cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(Viscosity) < 0 Then
        LblError.Caption = "Reservoir fluid viscosity cannot be negative."
        LblError.Visible = True
    ElseIf FVF = vbNullString Then
        LblError.Caption = "Please enter a formation volume factor."
        LblError.Visible = True
    ElseIf Not IsNumeric(FVF) Then
        LblError.Caption = "An invalid character was entered in formation volume factor."
        LblError.Visible = True
    ElseIf CDbl(FVF) = 0 Then
        LblError.Caption = "Formation volume factor cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(FVF) < 0 Then
        LblError.Caption = "Formation volume factor cannot be negative."
        LblError.Visible = True
    ElseIf Re = vbNullString Then
        LblError.Caption = "Please enter a drainage radius."
        LblError.Visible = True
    ElseIf Not IsNumeric(Re) Then
        LblError.Caption = "An invalid character was entered in drainage radius."
        LblError.Visible = True
    ElseIf CDbl(Re) = 0 Then
        LblError.Caption = "Drainage radius cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(Re) < 0 Then
        LblError.Caption = "Drainage radius cannot be negative."
        LblError.Visible = True
    ElseIf Rw = vbNullString Then
        LblError.Caption = "Please enter a wellbore radius."
        LblError.Visible = True
    ElseIf Not IsNumeric(Rw) Then
        LblError.Caption = "An invalid character was entered in wellbore radius."
        LblError.Visible = True
    ElseIf CDbl(Rw) = 0 Then
        LblError.Caption = "Wellbore radius cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(Rw) < 0 Then
        LblError.Caption = "Wellbore radius cannot be negative."
        LblError.Visible = True
    ElseIf Sf = vbNullString Then
        LblError.Caption = "Please enter a skin factor."
        LblError.Visible = True
    ElseIf Not IsNumeric(Sf) Then
        LblError.Caption = "An invalid character was entered in skin factor."
        LblError.Visible = True
    Else
        LblError.Visible = False
        FormPIDarcy.Hide
    End If

End Sub

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES OR THE ERROR LABEL
'================================================================================================================================
Private Sub BtnClear_Click()

    If TxtViscosity.Text <> vbNullString Or TxtFVF.Text <> vbNullString Or TxtRe.Text <> vbNullString Or TxtRw.Text <> _
        vbNullString Then
        
        TxtViscosity.Text = vbNullString
        TxtFVF.Text = vbNullString
        TxtRe.Text = vbNullString
        TxtRw.Text = vbNullString
        
        LblError.Caption = "The variable text boxes have been cleared."
        LblError.Visible = True
    Else
        LblError.Visible = False
    End If

End Sub

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES AND CLOSES THE FORM
'================================================================================================================================
Private Sub BtnCancel_Click()

    TxtViscosity.Text = vbNullString
    TxtFVF.Text = vbNullString
    TxtRe.Text = vbNullString
    TxtRw.Text = vbNullString
    LblError.Visible = False
    FormPIDarcy.Hide

End Sub

Private Sub LblEqnDenominator_Click()

End Sub
