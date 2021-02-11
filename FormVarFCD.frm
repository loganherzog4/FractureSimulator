VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVarFCD 
   Caption         =   "Fracture Simulator - Variable FCD"
   ClientHeight    =   3975
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6735
   OleObjectBlob   =   "FormVarFCD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVarFCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBrowseVFCDGridPath_Click()

    Dim Path As String
    
    Path = GetFilePath("C:\", False, False)
    
    If Path <> vbNullString Then
        TxtVFCDGridPath.Text = Path
        TxtVFCDGridPath.BackColor = vbButtonFace
    End If

End Sub

Private Sub BtnVFCDContinue_Click()

    Dim Row As Integer
    Dim LastFracIndex As Integer
    
    Row = 14
    
    Do While Row <= 1000
        If Sheets("Grid Statistics").Cells(Row, "A").Interior.Color = RGB(0, 255, 255) Then
            If Sheets("Grid Statistics").Cells(Row + 1, "A").Interior.Color = RGB(0, 255, 255) Then
                LastFracIndex = LastFracIndex + 1
            Else
                LastFracIndex = LastFracIndex + 1
                Exit Do
            End If
        End If
    
        Row = Row + 1
    Loop
    
    FormVarFCD.Hide
    GenerateVarFCDInclude LastFracIndex

End Sub
