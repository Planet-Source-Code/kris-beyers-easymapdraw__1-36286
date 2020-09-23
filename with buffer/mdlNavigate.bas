Attribute VB_Name = "mdlNavigate"
Public Function Navigate(ByVal IsDrawing As Boolean, ByRef mCurrentLimit As typZoom, ByRef ZoomHistory() As typZoom, vDirection As enmDirection, Optional vStep As Byte = 25) As typZoom
  Dim hulpLimitW As Double, hulpLimitO As Double, hulpLimitN As Double, hulpLimitZ As Double, hulpStepX As Double, hulpStepY As Double
  Dim hulpCurrentLimit As typZoom
  '
  If vStep < 1 Or vStep > 100 Then
    Err.Raise 1000, "CMap", "Step has to be between 1 and 100"
  ElseIf Not IsDrawing Then
    hulpLimitW = mCurrentLimit.W
    hulpLimitO = mCurrentLimit.O
    hulpLimitZ = mCurrentLimit.Z
    hulpLimitN = mCurrentLimit.N
    hulpStepY = (mCurrentLimit.N - mCurrentLimit.Z) * vStep / 100
    hulpStepX = (mCurrentLimit.O - mCurrentLimit.W) * vStep / 100
    If vDirection = DirectionNW Or vDirection = DirectionN Or vDirection = DirectionNO Then
      hulpLimitZ = hulpLimitZ + IIf(hulpStepY > Abs(ZoomHistory(1).N) - Abs(mCurrentLimit.N), _
      Abs(ZoomHistory(1).N) - Abs(mCurrentLimit.N), hulpStepY)
      hulpLimitN = hulpLimitN + IIf(hulpStepY > Abs(ZoomHistory(1).N) - Abs(mCurrentLimit.N), _
      Abs(ZoomHistory(1).N) - Abs(mCurrentLimit.N), hulpStepY)
    End If
    If vDirection = DirectionNO Or vDirection = DirectionO Or vDirection = DirectionZO Then
      hulpLimitW = hulpLimitW + IIf(hulpStepX > Abs(ZoomHistory(1).O) - Abs(mCurrentLimit.O), _
      Abs(ZoomHistory(1).O) - Abs(mCurrentLimit.O), hulpStepX)
      hulpLimitO = hulpLimitO + IIf(hulpStepX > Abs(ZoomHistory(1).O) - Abs(mCurrentLimit.O), _
      Abs(ZoomHistory(1).O) - Abs(mCurrentLimit.O), hulpStepX)
    End If
    If vDirection = DirectionZW Or vDirection = DirectionZ Or vDirection = DirectionZO Then
      hulpLimitZ = hulpLimitZ - IIf(hulpStepY > Abs(ZoomHistory(1).Z) - Abs(mCurrentLimit.Z), _
      Abs(ZoomHistory(1).Z) - Abs(mCurrentLimit.Z), hulpStepY)
      hulpLimitN = hulpLimitN - IIf(hulpStepY > Abs(ZoomHistory(1).Z) - Abs(mCurrentLimit.Z), _
      Abs(ZoomHistory(1).Z) - Abs(mCurrentLimit.Z), hulpStepY)
    End If
    If vDirection = DirectionNW Or vDirection = DirectionW Or vDirection = DirectionZW Then
      hulpLimitW = hulpLimitW - IIf(hulpStepX > Abs(ZoomHistory(1).W) - Abs(mCurrentLimit.W), _
      Abs(ZoomHistory(1).W) - Abs(mCurrentLimit.W), hulpStepX)
      hulpLimitO = hulpLimitO - IIf(hulpStepX > Abs(ZoomHistory(1).W) - Abs(mCurrentLimit.W), _
      Abs(ZoomHistory(1).W) - Abs(mCurrentLimit.W), hulpStepX)
    End If
    hulpCurrentLimit.W = hulpLimitW
    hulpCurrentLimit.O = hulpLimitO
    hulpCurrentLimit.Z = hulpLimitZ
    hulpCurrentLimit.N = hulpLimitN
    Navigate = hulpCurrentLimit
  End If
End Function

Public Sub Swap(ByRef value1 As Double, ByRef value2 As Double)
  Dim hulp As Double
  hulp = value1
  value1 = value2
  value2 = hulp
End Sub
