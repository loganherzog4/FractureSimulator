VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormQuickCalcs 
   Caption         =   "Fracture Simulator - Quick Calculations"
   ClientHeight    =   9585
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6735
   OleObjectBlob   =   "FormQuickCalcs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormQuickCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnRiCalc_Click()

    Dim AvgResPerm As Double
    Dim ProductionTime As Double
    Dim AvgResPoro As Double
    Dim Viscosity As Double
    Dim Comp As Double
    Dim Ri As Double
    Dim GridAreaFt As Double
    Dim GridAreaAcres As Double
    
    AvgResPerm = TxtRiAvgResPerm.Text
    ProductionTime = TxtRiProductionTime.Text
    AvgResPoro = TxtRiAvgResPoro.Text
    Viscosity = TxtRiViscosity.Text
    Comp = TxtRiComp.Text
    
    ProductionTime = ProductionTime * 24
    Comp = Comp * 10 ^ -6
    
    Ri = Sqr((AvgResPerm * ProductionTime) / (948 * AvgResPoro * Viscosity * Comp))
    GridAreaFt = Ri ^ 2
    GridAreaAcres = GridAreaFt * 2.2957 * 10 ^ -5
    
    LblRiOutput.Visible = True
    LblRiOutput.Caption = Format(Ri, "#.00")
    
    LblGridAreaFtOutput.Visible = True
    LblGridAreaFtOutput.Caption = Format(GridAreaFt, "#.00")
    
    LblGridAreaAcresOutput.Visible = True
    LblGridAreaAcresOutput.Caption = Format(GridAreaAcres, "#.00")

End Sub
