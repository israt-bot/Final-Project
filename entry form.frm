VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   OleObjectBlob   =   "entry form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubmit_Click()
    
    Dim ws As Worksheet
    Dim newRow As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    
    ws.Cells(newRow, 1).Value = txtname.Value
    ws.Cells(newRow, 2).Value = txtorigin.Value
    ws.Cells(newRow, 3).Value = txthardiness.Value
    ws.Cells(newRow, 4).Value = txtflower.Value
    ws.Cells(newRow, 5).Value = txtmethod.Value
    ws.Cells(newRow, 6).Value = txtdiseases.Value
    ws.Cells(newRow, 7).Value = txtagent.Value
    
    ' Confirmation Message
    MsgBox "Data Submitted Successfully!", vbInformation
    
    
    ClearForm
End Sub


Private Sub ClearForm()
    txtname.Value = ""
    txtorigin.Value = ""
    txthardiness.Value = ""
    txtflower.Value = ""
    txtmethod.Value = ""
    txtdiseases.Value = ""
    txtagent.Value = ""
End Sub

End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

