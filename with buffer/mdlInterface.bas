Attribute VB_Name = "mdlInterface"
Public Sub LockControls(vForm As Form, vChoice As Boolean)
  With vForm
    Screen.MousePointer = IIf(vChoice = True, vbHourglass, vbDefault)
    .cmdReDraw.Enabled = Not vChoice
    .cmdZoomIn.Enabled = Not vChoice
    .cmdZoomOut.Enabled = Not vChoice
    .cmdLoadBuffer.Enabled = Not vChoice
    .cmdSavePicture.Enabled = Not vChoice
    .pgbProgress.Visible = vChoice
  End With
End Sub



