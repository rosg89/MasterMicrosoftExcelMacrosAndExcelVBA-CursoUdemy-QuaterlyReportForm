VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFinalReport 
   Caption         =   "Final Report"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "FrmFinalReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmFinalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSheet_Change()
    Worksheets(Me.cboSheet.Value).Select
End Sub

Private Sub cmdAddSheet_Click()

    Worksheets.Add before:=Worksheets(1)
    ActiveSheet.Name = InputBox("Please enter a name for the new worksheet")
    
    
         
    
End Sub

Private Sub cmdRunReport_Click()

    FinalReport

End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    i = 1
    
    Do While i <= Worksheets.Count
    Me.cboSheet.AddItem Worksheets(i).Name
    i = i + 1
    Loop
    
End Sub
