Attribute VB_Name = "exportGUI"
Sub Interface()
    
    On Error GoTo finishLoop
    uf_interface.lb_records.Clear
    m_ = 0
    previousName = ""
    While True
        measurementName = Project.MeasurementOrganiser.Templates(1).Measurement(m_).Name
        If measurementName <> previousName Then
            uf_interface.lb_records.AddItem measurementName
        End If
        previousName = measurementName
        m_ = m_ + 1
    Wend

finishLoop:

    On Error GoTo 0
    
    uf_interface.cbb_displaygroup.Clear
    For Each dg In Project.DisplayOrganiser.DisplayGroups
        uf_interface.cbb_displaygroup.AddItem dg.Name
    Next dg
    uf_interface.cbb_displaygroup.ListIndex = 0

    uf_interface.export = False
    uf_interface.Show vbModeless
    While uf_interface.Visible And Not uf_interface.export
        DoEvents
    Wend
    uf_interface.Hide
    
    For m_ = 0 To uf_interface.lb_records.ListCount - 1
        If uf_interface.lb_records.Selected(m_) Then
            Debug.Print Project.MeasurementOrganiser.Templates(1).Measurement(m_).Name
            For d_ = 0 To uf_interface.lb_displays.ListCount - 1
                If uf_interface.lb_displays.Selected(d_) Then
                    Debug.Print vbTab & Project.DisplayOrganiser.DisplayGroups(uf_interface.cbb_displaygroup.Text).Displays.Item(d_ + 1).Name
                End If
            Next d_
        End If
    Next m_
End Sub
