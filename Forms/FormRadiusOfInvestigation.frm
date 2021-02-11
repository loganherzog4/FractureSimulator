VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRadiusOfInvestigation 
   Caption         =   "Radius of Investigation - Well Testing"
   ClientHeight    =   6765
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4005
   OleObjectBlob   =   "FormRadiusOfInvestigation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormRadiusOfInvestigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================================================================================
' CHECKS FOR VALID INPUT PARAMETERS
'================================================================================================================================
Private Sub BtnOK_Click()

    'Production time input from user in days
    Dim ProductionTime As String
    
    'Reservoir fluid viscosity input from user in cP
    Dim Viscosity As String
    
    'Total compressibility input from user in microsips
    Dim Comp As String
    
    'Average reservoir permeability
    Dim ResAvgPerm
    
    'Average reservoir porosity
    Dim ResAvgPorosity
    
    'Radius of investigation calculated and displayed
    Dim Ri As Double
    
    ProductionTime = TxtTime.Text
    Viscosity = TxtViscosity.Text
    Comp = TxtComp.Text
    
    If ProductionTime = vbNullString Then
        LblError.Caption = "Please enter a production time."
        LblError.Visible = True
    ElseIf Not IsNumeric(ProductionTime) Then
        LblError.Caption = "An invalid character was entered in production time."
        LblError.Visible = True
    ElseIf CDbl(ProductionTime) = 0 Then
        LblError.Caption = "Production time cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(ProductionTime) < 0 Then
        LblError.Caption = "Production time cannot be negative."
        LblError.Visible = True
    ElseIf Viscosity = vbNullString Then
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
    ElseIf Comp = vbNullString Then
        LblError.Caption = "Please enter a total compressibility."
        LblError.Visible = True
    ElseIf Not IsNumeric(Comp) Then
        LblError.Caption = "An invalid character was entered in total compressibility."
        LblError.Visible = True
    ElseIf CDbl(Comp) = 0 Then
        LblError.Caption = "Total compressibility cannot equal zero."
        LblError.Visible = True
    ElseIf CDbl(Comp) < 0 Then
        LblError.Caption = "Total compressibility cannot be negative."
        LblError.Visible = True
    Else
        LblError.Visible = False
        FormRadiusOfInvestigation.Hide
        
        ResAvgPerm = Cells(7, "C").Value
        ResAvgPorosity = Cells(17, "C").Value
        
        Ri = Sqr((ResAvgPerm * CDbl(ProductionTime) * 24) / (948 * ResAvgPorosity * CDbl(Viscosity) * CDbl(Comp) * 10 ^ (-6)))
        
        With FormExpressRun
        
            .LblRiValue.Caption = Format(Ri, "#.00")
            .LblRi.Visible = True
            .LblRiValue.Visible = True
            .LblRiFt.Visible = True
        
        End With
        
    End If

End Sub

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES OR THE ERROR LABEL
'================================================================================================================================
Private Sub BtnClear_Click()

    If TxtTime.Text <> vbNullString Or TxtViscosity.Text <> vbNullString Or TxtComp.Text <> vbNullString Then
        TxtTime.Text = vbNullString
        TxtViscosity.Text = vbNullString
        TxtComp.Text = vbNullString
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

    TxtTime.Text = vbNullString
    TxtViscosity.Text = vbNullString
    TxtComp.Text = vbNullString
    LblError.Caption = vbNullString
    LblError.Visible = False
    
    FormRadiusOfInvestigation.Hide
    
    With FormExpressRun
        If .LblRi.Visible = True Then
            .LblRi.Visible = False
            .LblRiValue.Visible = False
            .LblRiFt.Visible = False
        End If
    End With

End Sub


