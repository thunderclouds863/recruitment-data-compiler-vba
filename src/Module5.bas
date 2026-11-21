Attribute VB_Name = "Module5"
Public Sub OpenEditSourcingFromDetail(ByVal rowNum As Long)
    frmEditSourcing.LoadRecord rowNum
    frmEditSourcing.Show vbModal
End Sub
