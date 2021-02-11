VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVFCDQuickCalcs 
   Caption         =   "Fracture Simulator - Quick Calculations"
   ClientHeight    =   8295
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6735
   OleObjectBlob   =   "FormVFCDQuickCalcs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVFCDQuickCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnFCDCalc_Click()

    Dim FracPerm As Double
    Dim AvgFracWidth As Double
    Dim AvgResPerm As Double
    Dim FracHalfLength As Double
    Dim FracConductivity As Double
    Dim FCD As Double
    
    FracPerm = TxtFCDFracPerm.Text
    AvgFracWidth = TxtFCDFracWidth.Text
    AvgResPerm = TxtFCDAvgResPerm.Text
    FracHalfLength = TxtFCDHalfLength.Text
    
    AvgFracWidth = AvgFracWidth / 12
    
    FracConductivity = FracPerm * AvgFracWidth
    
    FCD = FracConductivity / (AvgResPerm * FracHalfLength)
    
    LblFCDFractureConductivityOutput.Visible = True
    LblFCDFractureConductivityOutput.Caption = Format(FracConductivity, "#.00")
    
    LblFCDOutput.Visible = True
    LblFCDOutput.Caption = Format(FCD, "0.00")

End Sub

Private Sub BtnFracPermCalc_Click()

    Dim FracConductivity As Double
    Const FracWidth As Double = 5
    Dim FracPerm As Double
    
    FracConductivity = TxtFracPermFracConductivity.Text
    
    FracPerm = FracConductivity / FracWidth
    
    LblFracPermOutput.Visible = True
    LblFracPermOutput.Caption = Format(FracPerm, "#.00")

End Sub
