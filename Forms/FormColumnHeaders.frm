VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormColumnHeaders 
   Caption         =   ".LAS File Column Headers"
   ClientHeight    =   8595
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4005
   OleObjectBlob   =   "FormColumnHeaders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormColumnHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================================================================================
' CLEARS THE VARIABLE TEXT BOXES
'================================================================================================================================
Private Sub BtnClear_Click()

    TxtDepthHeader.Text = vbNullString
    TxtPermHeader.Text = vbNullString
    TxtPoroHeader.Text = vbNullString
    TxtSwHeader.Text = vbNullString
    TxtPayHeader.Text = vbNullString
    TxtResHeader.Text = vbNullString

End Sub

'================================================================================================================================
' CLOSES THE FORM
'================================================================================================================================
Private Sub BtnCancel_Click()

    TxtDepthHeader.Text = vbNullString
    TxtPermHeader.Text = vbNullString
    TxtPoroHeader.Text = vbNullString
    TxtSwHeader.Text = vbNullString
    TxtPayHeader.Text = vbNullString
    TxtResHeader.Text = vbNullString
    
    FormColumnHeaders.Hide

End Sub

'================================================================================================================================
' INITIALIZES COLUMN HEADER VARIABLES IN "LASPARSER" MODULE
'================================================================================================================================
Private Sub BtnOK_Click()
    
    FormColumnHeaders.Hide
    
End Sub

Private Sub OptPoroFraction_Click()

    If OptPoroPercentage.Value = True Then
        OptPoroPercentage.Value = False
        OptPoroFraction.Value = True
    End If

End Sub

Private Sub OptPoroPercentage_Click()

    If OptPoroFraction.Value = True Then
        OptPoroFraction.Value = False
        OptPoroPercentage.Value = True
    End If

End Sub

Private Sub OptSwFraction_Click()

    If OptSwPercentage.Value = True Then
        OptSwPercentage.Value = False
        OptSwFraction.Value = True
    End If
    
End Sub

Private Sub OptSwPercentage_Click()

    If OptSwFraction.Value = True Then
        OptSwFraction.Value = False
        OptSwPercentage.Value = True
    End If
    
End Sub
