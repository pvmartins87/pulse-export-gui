VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_interface 
   Caption         =   "Export to Excel"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   OleObjectBlob   =   "uf_interface.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public export As Boolean

Private Sub bt_export_Click()
    export = True
End Sub

Private Sub bt_selectall_Click()
    For i_ = 0 To uf_interface.lb_records.ListCount - 1
        uf_interface.lb_records.Selected(i_) = True
    Next i_
End Sub

Private Sub bt_clearselection_Click()
    For i_ = 0 To uf_interface.lb_records.ListCount - 1
        uf_interface.lb_records.Selected(i_) = False
    Next i_
End Sub

Private Sub cbb_displaygroup_Change()
    dgName = uf_interface.cbb_displaygroup.Text
    
    If dgName <> "" Then
        uf_interface.lb_displays.Clear
        For Each ds In Project.DisplayOrganiser.DisplayGroups(dgName).Displays
            uf_interface.lb_displays.AddItem ds.Name
        Next ds
    End If
End Sub

Private Sub lb_displays_Click()

End Sub

Private Sub lb_records_Click()

End Sub

Private Sub UserForm_Click()

End Sub
