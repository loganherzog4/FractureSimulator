VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormColumnHeadersSurvey 
   Caption         =   "Survey File Column Headers"
   ClientHeight    =   3390
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4005
   OleObjectBlob   =   "FormColumnHeadersSurvey.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormColumnHeadersSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClear_Click()

    TxtMDHeader.Text = vbNullString
    TxtTVDHeader.Text = vbNullString

End Sub


Private Sub BtnCancel_Click()

    TxtMDHeader.Text = vbNullString
    TxtTVDHeader.Text = vbNullString
    
    FormColumnHeadersSurvey.Hide

End Sub

Private Sub BtnOK_Click()

    FormColumnHeadersSurvey.Hide

End Sub
