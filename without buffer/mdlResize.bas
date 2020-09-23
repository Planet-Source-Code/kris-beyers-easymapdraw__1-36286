Attribute VB_Name = "mdlResize"
Public Sub FormResize(vForm As Form)
  If Not vForm.WindowState = vbMinimized Then
    '
    With vForm
      .EasyMapDraw.Width = .Width - 860
      .EasyMapDraw.Height = .Height - 2400
      .cmdReDraw.Top = .Height - 1350
      .cmdReDraw.Left = .Width - 2000
      .cmdSavePicture.Top = .Height - 1350
      .cmdSavePicture.Left = .Width - 3400
      .lblF1.Top = 700
      '
      ' position navigate buttons
      .cmdnavigate(0).Left = .EasyMapDraw.Left + .EasyMapDraw.Width / 2 - .cmdnavigate(4).Width / 2
      .cmdnavigate(0).Top = .EasyMapDraw.Top - 330
      .cmdnavigate(1).Left = .EasyMapDraw.Left + .EasyMapDraw.Width + 25
      .cmdnavigate(1).Top = .EasyMapDraw.Top - 355
      .cmdnavigate(2).Left = .EasyMapDraw.Left + .EasyMapDraw.Width + 55
      .cmdnavigate(2).Top = .EasyMapDraw.Top + .EasyMapDraw.Height / 2 - .cmdnavigate(2).Height / 2
      .cmdnavigate(3).Left = .EasyMapDraw.Left + .EasyMapDraw.Width + 25
      .cmdnavigate(3).Top = .EasyMapDraw.Height + .EasyMapDraw.Top + 35
      .cmdnavigate(4).Left = .EasyMapDraw.Left + .EasyMapDraw.Width / 2 - .cmdnavigate(4).Width / 2
      .cmdnavigate(4).Top = .EasyMapDraw.Height + .EasyMapDraw.Top + 35
      .cmdnavigate(5).Left = .EasyMapDraw.Left - 370
      .cmdnavigate(5).Top = .EasyMapDraw.Height + .EasyMapDraw.Top + 30
      .cmdnavigate(6).Left = .EasyMapDraw.Left - 350
      .cmdnavigate(6).Top = .EasyMapDraw.Top + .EasyMapDraw.Height / 2 - .cmdnavigate(6).Height / 2
      .cmdnavigate(7).Left = .EasyMapDraw.Left - 370
      .cmdnavigate(7).Top = .EasyMapDraw.Top - 355
      '
      .lblStep.Top = .Height - 1350
      .lblStep.Left = .cmdnavigate(4).Left - 1500 '5490
      .txtStep.Top = .Height - 1350
      .txtStep.Left = .cmdnavigate(4).Left - 900
    End With
  End If
End Sub

