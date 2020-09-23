Attribute VB_Name = "mdlPropertyBag"
Public Sub WriteProperties(vEasyMapDraw As EasyMapDraw, vPropBag As PropertyBag)
  With vPropBag
    .WriteProperty "LimitO", vEasyMapDraw.LimitO, 180
    .WriteProperty "LimitW", vEasyMapDraw.LimitW, -180
    .WriteProperty "LimitN", vEasyMapDraw.LimitN, 90
    .WriteProperty "LimitZ", vEasyMapDraw.LimitZ, -90
    '
    .WriteProperty "Redraw", CBool(vEasyMapDraw.Redraw), False
    .WriteProperty "DrawLines", CBool(vEasyMapDraw.DrawLines), False
    '
    .WriteProperty "MapBackColor", vEasyMapDraw.MapBackColor, vbWhite
    .WriteProperty "DrawColor", vEasyMapDraw.DrawColor, vbBlack
    .WriteProperty "MouseFunction", vEasyMapDraw.MouseFunction, MouseZoom
    '
    .WriteProperty "NagivateStep", vEasyMapDraw.NagivateStep, 20
    .WriteProperty "KeyNagivate", vEasyMapDraw.KeyNagivate, False
    '
    .WriteProperty "SelectionRatio", vEasyMapDraw.SelectionRatio, 0
    .WriteProperty "SelectionRange", vEasyMapDraw.SelectionRange, 0
  End With
End Sub

Public Sub InitProperties(vEasyMapDraw As EasyMapDraw, ByRef vMapBackColor As OLE_COLOR)
  'Only called when the OCX is selected from the Toolbar
  '
  With vEasyMapDraw
    .LimitW = -180
    .LimitO = 180
    .LimitZ = -90
    .LimitN = 90
    .NagivateStep = 20
    '
    .MapBackColor = &H80000000
    .DrawColor = vbBlack
    .MouseFunction = MouseZoom
    '
    vMapBackColor = .MapBackColor
    '
    .SelectionRatio = 0
  End With
End Sub

Public Sub ReadProperties(vEasyMapDraw As EasyMapDraw, vPropBag As PropertyBag)
  With vPropBag
    vEasyMapDraw.LimitO = CDbl(.ReadProperty("LimitO", 180))
    vEasyMapDraw.LimitW = CDbl(.ReadProperty("LimitW", -180))
    vEasyMapDraw.LimitN = CDbl(.ReadProperty("LimitN", 90))
    vEasyMapDraw.LimitZ = CDbl(.ReadProperty("LimitZ", -90))
    '
    vEasyMapDraw.Redraw = CBool(.ReadProperty("Redraw", False))
    vEasyMapDraw.DrawLines = CBool(.ReadProperty("DrawLines", False))
    '
    vEasyMapDraw.MapBackColor = .ReadProperty("MapBackColor", vbWhite)
    vEasyMapDraw.DrawColor = .ReadProperty("DrawColor", vbBlack)
    vEasyMapDraw.MouseFunction = .ReadProperty("MouseFunction", MouseZoom)
    '
    vEasyMapDraw.NagivateStep = .ReadProperty("NagivateStep", 20)
    vEasyMapDraw.KeyNagivate = .ReadProperty("KeyNagivate", False)
    vEasyMapDraw.SelectionRatio = .ReadProperty("SelectionRatio", 0)
    vEasyMapDraw.SelectionRange = .ReadProperty("SelectionRange", 0)
  End With
End Sub
