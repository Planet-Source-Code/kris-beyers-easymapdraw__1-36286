VERSION 5.00
Begin VB.UserControl EasyMapDraw 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   MousePointer    =   2  'Cross
   ScaleHeight     =   2610
   ScaleWidth      =   4500
   ToolboxBitmap   =   "EasyMapDraw.ctx":0000
   Begin VB.PictureBox fPicture 
      Height          =   615
      Left            =   2400
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TmrDraw 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3840
      Top             =   1920
   End
   Begin VB.Line shpDistance 
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      Visible         =   0   'False
      X1              =   240
      X2              =   1920
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Shape shpSelection 
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      Height          =   375
      Left            =   1080
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "EasyMapDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-------------------------------------------------------------------
' Copyright © by Kris Beyers 26-06-2002 Belgium
'
' version 1.4
'
' nothing of this program may be redistributed without
' prior permission from the author.
' you may use this code free for non-commercial purposes only
' and author can not be held responsible for loss of data or any damage
' by using this application.
'-------------------------------------------------------------------
'
' What's new in EasyMapDraw 1.4 ?
' -------------------------------
' * Some calculations are simplified (mouseposition)
' * SaveImage(vPictureFileName As String) function added
' * SelectionCompleted event is splitted into two different kind of events: selection and zoomSelection
' * BUG solved: MousePosition has a correct range now: from -180° to 180° and -90° to 90°
' * BUG solved: Navation function, navigates with correct steps always now
'
'locals
'------
'properties
Private mKeyNagivate As Boolean
Private mRedraw As Boolean
Private mDrawLines As Boolean
Private mSelectionRatio As Byte
Private mNagivateStep As Byte
Private mSelectionRange As Double
Private mMouseFunction As enmMouseFunction
Private mDrawColor As OLE_COLOR
Private mMapBackColor As OLE_COLOR

'status
Private mIsDrawing As Boolean

'zoom
Private Type typZoom
  N As Double
  O As Double
  Z As Double
  W As Double
End Type
Private ZoomHistory() As typZoom
Private mZoomHistoryTel As Byte
Private mCurrentLimit As typZoom

'for zoom with rectangle drawn with mouse
Private mPreviousX As Double
Private mPreviousY As Double
Private mFirstX As Double
Private mFirstY As Double

'for line
Private mPreviousXp As Double
Private mPreviousYp As Double
Private mBreakLine As Boolean

'for SelectionRectangle
Private mPreviousMouseX As Integer 'no degrees like mPreviousX
Private mPreviousMouseY As Integer

'public enums
'------------
Public Enum enmZoomError
  West = 1
  East
  South
  North
  East_West_Equal
  North_South_Equal
  East_SmallerThan_West
  North_SmallerThan_South
  ToMuchSteps
  SelectionIgnored
End Enum

Public Enum enmMouseFunction
  MouseZoom = 0
  MouseSelection
  MouseDistance
End Enum

Public Enum enmZoomDirection
  ZoomInDirection = 1
  ZoomOutDirection
  ZoomRefresh
End Enum

Public Enum enmDrawEquatorMeridian
  DrawEquator = 1
  DrawMeridian
  DrawEquatorAndMeridian
End Enum

Public Enum enmMouseEvent
  CurrentPosition = 1
  SelectionStartPosition
  DistanceStartPosition
  ClickedPosition
End Enum

Public Enum enmDirection
  DirectionN = 1
  DirectionNO
  DirectionO
  DirectionZO
  DirectionZ
  DirectionZW
  DirectionW
  DirectionNW
End Enum

Public Enum enmSelectionEvent
  Selection = 1
  ZoomSelection
End Enum

'events
'------
Public Event MousePostion(vMouseEvent As enmMouseEvent, vLongitude As Double, vLatitude As Double)
Public Event ZoomError(vZoomError As enmZoomError)
Public Event DistanceCalculated(vDistance As Double, vLatitude1 As Double, vLongitude1 As Double, vLatitude2 As Double, vLongitude2 As Double)
Public Event SelectionCompleted(vSelectionEvent As enmSelectionEvent, vLimitN As Double, vLimitO As Double, vLimitZ As Double, vLimitW As Double)
Public Event ZoomConfigured(vZoomDirection As enmZoomDirection)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' just forwarding the basic MouseDown event

'getters
'-------
Public Property Get IsMaxZoomedOut() As Boolean
  IsMaxZoomedOut = IIf(mZoomHistoryTel > 1, False, True)
End Property

Public Property Get LimitW() As Double
  LimitW = mCurrentLimit.W
End Property

Public Property Get LimitO() As Double
  LimitO = mCurrentLimit.O
End Property

Public Property Get LimitZ() As Double
  LimitZ = mCurrentLimit.Z
End Property

Public Property Get LimitN() As Double
  LimitN = mCurrentLimit.N
End Property

Public Property Get ZoomHistoryTel() As Byte
  ZoomHistoryTel = mZoomHistoryTel - 1
End Property

Public Property Get Redraw() As Boolean
  Redraw = mRedraw
End Property

Public Property Get IsDrawing() As Boolean
  IsDrawing = mIsDrawing
End Property

Public Property Get SelectionRange() As Double
  SelectionRange = mSelectionRange
End Property

'extra
Public Property Get DrawLines() As Boolean
  DrawLines = mDrawLines
End Property

Public Property Get DrawColor() As OLE_COLOR
  DrawColor = mDrawColor
End Property

Public Property Get ScaleHeight() As Single
  ScaleHeight = UserControl.ScaleHeight - 15 ' -15 because mouse differs a little in heightcalculation
End Property

Public Property Get ScaleWidth() As Single
  ScaleWidth = UserControl.ScaleWidth - 15  ' -15 because mouse differs a little in widthcalculation
End Property

Public Property Get MapBackColor() As OLE_COLOR
  MapBackColor = mMapBackColor
End Property

Public Property Get MouseFunction() As enmMouseFunction
  MouseFunction = mMouseFunction
End Property

Public Property Get KeyNagivate() As Boolean
  KeyNagivate = mKeyNagivate
End Property

Public Property Get NagivateStep() As Byte
  NagivateStep = mNagivateStep
End Property

Public Property Get SelectionRatio() As Byte
  SelectionRatio = mSelectionRatio
End Property

'setters
'-------
Public Property Let DrawLines(vStatus As Boolean)
  mDrawLines = vStatus
End Property

Public Property Let MouseFunction(vMouseFunction As enmMouseFunction)
  mMouseFunction = vMouseFunction
  PropertyChanged ("MouseFunction")
End Property

Public Property Let DrawColor(vDrawColor As OLE_COLOR)
  mDrawColor = vDrawColor
  PropertyChanged ("DrawColor")
End Property

Public Property Let Redraw(vRedraw As Boolean)
  mRedraw = vRedraw
  UserControl.AutoRedraw = mRedraw
  PropertyChanged ("Redraw")
End Property

Public Property Let SelectionRange(vSelectionRange As Double)
  If vSelectionRange > 0 Then
    mSelectionRange = vSelectionRange
    PropertyChanged ("SelectionRange")
  End If
End Property

'extra
Public Property Let SelectionRatio(vSelectionRatio As Byte)
  If vSelectionRatio <= 100 Then
    mSelectionRatio = vSelectionRatio
    PropertyChanged ("SelectionRatio")
  End If
End Property

Public Property Let KeyNagivate(vKeyNagivate As Boolean)
  mKeyNagivate = vKeyNagivate
  PropertyChanged ("KeyNagivate")
End Property

Public Property Let NagivateStep(vNagivateStep As Byte)
  If vNagivateStep >= 1 And vNagivateStep <= 100 Then
    mNagivateStep = vNagivateStep
    PropertyChanged ("NagivateStep")
  End If
End Property

Public Property Let LimitW(vLimitW As Double)
  If vLimitW >= -180 And vLimitW <= 180 And vLimitW < mCurrentLimit.O Then
    mCurrentLimit.W = vLimitW
    PropertyChanged ("LimitW")
  End If
End Property

Public Property Let LimitO(vLimitO As Double)
  If vLimitO >= -180 And vLimitO <= 180 And vLimitO > mCurrentLimit.W Then
    mCurrentLimit.O = vLimitO
    PropertyChanged ("LimitO")
  End If
End Property

Public Property Let LimitZ(vLimitZ As Double)
  If vLimitZ >= -90 And vLimitZ <= 90 And vLimitZ < mCurrentLimit.N Then
    mCurrentLimit.Z = vLimitZ
    PropertyChanged ("LimitZ")
  End If
End Property

Public Property Let LimitN(vLimitN As Double)
  If vLimitN >= -90 And vLimitN <= 90 And vLimitN > mCurrentLimit.Z Then
    mCurrentLimit.N = vLimitN
    PropertyChanged ("LimitN")
  End If
End Property

Public Property Let MapBackColor(vMapBackColor As OLE_COLOR)
  mMapBackColor = vMapBackColor
  UserControl.BackColor = mMapBackColor
  PropertyChanged ("MapBackColor")
End Property

'public functions/subs
'---------------------
Public Sub SaveImage(vPictureFileName As String)
  Call SavePicture(UserControl.Image, vPictureFileName)
End Sub

Public Function ZoomOut(Optional vStep As Byte = 1)
  Dim hulp As Integer
  'old values
  hulp = CInt(mZoomHistoryTel) - CInt(vStep)
  If hulp > 0 Then
    '
    'zorg er voor dat het First punt niet
    'verbonden wordt met de rechterbovenhoek
    mBreakLine = True
    '
    If mZoomHistoryTel > 1 Then ' do not delete the first zoomed out values
      mZoomHistoryTel = hulp
      mCurrentLimit.W = ZoomHistory(mZoomHistoryTel).W
      mCurrentLimit.O = ZoomHistory(mZoomHistoryTel).O
      mCurrentLimit.Z = ZoomHistory(mZoomHistoryTel).Z
      mCurrentLimit.N = ZoomHistory(mZoomHistoryTel).N
      ReDim Preserve ZoomHistory(1 To mZoomHistoryTel)  'delete one or more zoom(s)
      Call PrepareZoomConfiguredEvent(ZoomOutDirection)
    End If
  Else
    RaiseEvent ZoomError(ToMuchSteps)
  End If
End Function

Public Function ZoomIn(vLimitN As Double, vLimitO As Double, _
vLimitZ As Double, vLimitW As Double)
  Dim ZoomRet As enmZoomError
  '
  If Not IsDrawing Then
    ZoomRet = SystemZoomIn(vLimitN, vLimitO, vLimitZ, vLimitW)
    If ZoomRet > 0 Then RaiseEvent ZoomError(ZoomRet)
  End If
End Function

Public Sub DrawPoint(ByVal vLongitude As Double, ByVal vLatitude As Double)
  Dim Xp As Double, Yp As Double
  '
  mIsDrawing = True
  Xp = Calculate_X(vLongitude)
  If Xp <> -1 Then
    Yp = Calculate_Y(vLatitude)
    If Yp <> -1 Then
      If mDrawLines = True Then
        If mBreakLine = False Then
          UserControl.Line (mPreviousXp, mPreviousYp)-(Xp, Yp), mDrawColor
        Else
          mBreakLine = False
        End If
        mPreviousXp = Xp
        mPreviousYp = Yp
      Else
        UserControl.PSet (Xp, Yp), mDrawColor
      End If
    End If
  End If
End Sub

Public Sub BreakLine()
  mIsDrawing = True
  mBreakLine = True
End Sub

Public Sub MapReload()
  Call EndDraw
  UserControl.Cls
  TmrDraw.Enabled = False
  TmrDraw.Enabled = True
End Sub

Public Sub EndDraw()
  mIsDrawing = False
End Sub

Public Sub DrawPicture(vFileName As String, vWidth As Integer, vHeight As Integer, vLongitude As Double, vLatitude As Double, vNormalZoomWidth As Double, vNormalZoomHeight As Double)
  Dim Xp As Double, Yp As Double, test As Double
  '
  If mCurrentLimit.N - Abs(mCurrentLimit.Z) <= vNormalZoomHeight And mCurrentLimit.O - Abs(mCurrentLimit.W) <= vNormalZoomWidth And _
  mCurrentLimit.N - Abs(mCurrentLimit.Z) > 0 And mCurrentLimit.O - Abs(mCurrentLimit.W) > 0 Then
    '
    vHeight = (((vNormalZoomHeight - (mCurrentLimit.N - Abs(mCurrentLimit.Z))) / vNormalZoomHeight / 1.5) + 1) * vHeight
    vWidth = (((vNormalZoomWidth - (mCurrentLimit.O - Abs(mCurrentLimit.W))) / vNormalZoomWidth * 2) + 1) * vWidth
    '
    If vWidth > 10 And vHeight > 10 Then
      Xp = Calculate_X(vLongitude)
      If Xp <> -1 Then
        Yp = Calculate_Y(vLatitude)
        If Yp <> -1 Then
          fPicture.Picture = LoadPicture(vFileName)
          UserControl.PaintPicture fPicture, CSng(Xp - vWidth / 2), CSng(Yp - vHeight / 2), vWidth, vHeight
        End If
      End If
    End If
  End If
End Sub

Public Function Carc_Distance_Tussen(vLatitude1 As Double, vLongitude1 As Double, vLatitude2 As Double, vLongitude2 As Double) As Double
  'this function calculates the CARC distance on a world surface
  '
  Const dr As Double = 1.74532777777778E-02 ' pi / 180 constant to convert degrees into radians
  Dim Latitude1dr As Double, Latitude2dr As Double, CA As Double, GCa As Double
  '
  Latitude1dr = vLatitude1 * dr
  Latitude2dr = vLatitude2 * dr
  CA = Math.Cos(Latitude1dr) * Math.Cos(Latitude2dr) * _
  Math.Cos((vLongitude2 - vLongitude1) * dr) + _
  Math.Sin(Latitude1dr) * Math.Sin(Latitude2dr)
  GCa = Math.Atn(Math.Sqr(1 - CA * CA) / CA)
  Carc_Distance_Tussen = IIf(GCa <= 0, GCa + 3.14159, GCa) * 6372
End Function

Public Sub Navigate(vDirection As enmDirection, Optional vStep As Byte = 25)
  Dim hulpLimitW As Double, hulpLimitO As Double, hulpLimitN As Double, hulpLimitZ As Double, hulpStepX As Double, hulpStepY As Double
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
    Call SystemZoomIn(hulpLimitN, hulpLimitO, hulpLimitZ, hulpLimitW, True)
  End If
End Sub

Public Sub DrawEquatorMeridian(vDrawEquatorMeridian As enmDrawEquatorMeridian)
  Dim Equator As Double, Meridian As Double
  '
  If vDrawEquatorMeridian = DrawEquator Or _
  vDrawEquatorMeridian = DrawEquatorAndMeridian Then
    Equator = CalculateEquator()
    If Not Equator = 0 Then _
    UserControl.Line (0, Abs(Equator))-(Me.ScaleWidth, Abs(Equator))
  End If
  If vDrawEquatorMeridian = DrawMeridian Or _
  vDrawEquatorMeridian = DrawEquatorAndMeridian Then
    Meridian = CalculateMeridian()
    If Not Meridian = 0 Then _
    UserControl.Line (Meridian, 0)-(Meridian, Me.ScaleHeight)  ' UserControl.Height)
  End If
End Sub

'private functions/subs
'----------------------
Private Sub PrepareZoomConfiguredEvent(vZoomDirection As enmZoomDirection)
  If mRedraw = False Then UserControl.AutoRedraw = True
  UserControl.Cls
  If mRedraw = False Then UserControl.AutoRedraw = False
  ''mIsDrawing = True
  RaiseEvent ZoomConfigured(vZoomDirection)
End Sub

Private Sub TmrDraw_Timer() 'drawing need sometimes a delay to avoid reading values from properties before the objects are drawn
  Call PrepareZoomConfiguredEvent(ZoomRefresh)
  TmrDraw.Enabled = False
End Sub

Private Function CalculateEquator() As Double
  If mCurrentLimit.Z < 0 And mCurrentLimit.N > 0 Then _
  CalculateEquator = Abs((mCurrentLimit.N / (Abs(mCurrentLimit.Z) + mCurrentLimit.N)) _
  * Me.ScaleHeight)
End Function

Private Function CalculateMeridian() As Double
  If mCurrentLimit.W < 0 And mCurrentLimit.O > 0 Then _
  CalculateMeridian = Abs((mCurrentLimit.W / (Abs(mCurrentLimit.W) + mCurrentLimit.O)) _
  * Me.ScaleWidth)
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
  Dim Xg2 As Double, Yg2 As Double
  Dim Equator As Double, Meridian As Double, hulpTop As Double
  '
  'draw line or rectangle
  If Redraw And Button = 1 Then
    If mMouseFunction = MouseZoom Or mMouseFunction = MouseSelection Then
      shpSelection.Visible = True
    Else
      shpDistance.Visible = True
    End If
    If mMouseFunction = MouseZoom Or mMouseFunction = MouseSelection Then
      '
      With shpSelection
        .Left = IIf(mPreviousMouseX < X, mPreviousMouseX, X)
        hulpTop = IIf(mSelectionRatio = 0, _
        IIf(mPreviousMouseY < Y, mPreviousMouseY, Y), Y - shpSelection.Height)
        .Top = hulpTop
        .Width = IIf(X - mPreviousMouseX < mPreviousMouseX - X, mPreviousMouseX - X, _
        X - mPreviousMouseX)
        .Height = IIf(mSelectionRatio = 0, _
        IIf(Y - mPreviousMouseY < mPreviousMouseY - Y, mPreviousMouseY - Y, _
        Y - mPreviousMouseY), .Width * mSelectionRatio / 100)
        If mSelectionRatio > 0 Then Y = shpSelection.Top + shpSelection.Height
      End With
      '
    Else
      '
      With shpDistance
        .X1 = mPreviousMouseX
        .Y1 = mPreviousMouseY
        .X2 = X
        .Y2 = Y
      End With
      '
    End If
  End If
  '
  ' Calculate YDegree from cursor
  If mCurrentLimit.Z >= 0 And mCurrentLimit.N > 0 Then
    Yg2 = ((Me.ScaleHeight - Y) / Me.ScaleHeight) * _
    (mCurrentLimit.N - mCurrentLimit.Z) + mCurrentLimit.Z
  ElseIf mCurrentLimit.Z <= 0 And mCurrentLimit.N <= 0 Then
    Yg2 = -(Y / Me.ScaleHeight) * _
    (mCurrentLimit.N - mCurrentLimit.Z) + mCurrentLimit.N
  Else
    'we look first where the Equator is positioned on y
    Equator = CalculateEquator()
    If Y >= Equator Then 'can not with IIF
      Yg2 = ((Y - Equator) / (Me.ScaleHeight - Equator)) * mCurrentLimit.Z
    Else
      Yg2 = ((Equator - Y) / Equator) * mCurrentLimit.N
    End If
  End If
  '
  ' Calculate XDegree from cursor
  If mCurrentLimit.W >= 0 And mCurrentLimit.O > 0 Then
    Xg2 = (X / Me.ScaleWidth) * _
    (mCurrentLimit.O - mCurrentLimit.W) + mCurrentLimit.W
  ElseIf mCurrentLimit.W <= 0 And mCurrentLimit.O <= 0 Then
    Xg2 = -((Me.ScaleWidth - X) / Me.ScaleWidth) * _
    (mCurrentLimit.O - mCurrentLimit.W) + mCurrentLimit.O
  Else
    'we look first where the Equator is positioned on X
    Meridian = CalculateMeridian
    If X > Meridian Then 'kan niet met IIF
      Xg2 = ((X - Meridian) / (Me.ScaleWidth - Meridian)) * mCurrentLimit.O
    Else
      Xg2 = ((Meridian - X) / Meridian) * mCurrentLimit.W
    End If
  End If
  '
  'keep for calculating with mouse
  mPreviousX = Xg2
  mPreviousY = Yg2
  RaiseEvent MousePostion(CurrentPosition, Xg2, Yg2)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
  If Button = 1 Then
    mFirstX = mPreviousX
    mFirstY = mPreviousY
    mPreviousMouseX = X
    mPreviousMouseY = Y
    RaiseEvent MousePostion(IIf(mMouseFunction = MouseZoom Or mMouseFunction = MouseSelection, SelectionStartPosition, DistanceStartPosition), mFirstX, mFirstY)
  End If
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim hulp As Double
  '
  If Button = 1 Then
    If mRedraw Then
      If mMouseFunction = MouseZoom Or mMouseFunction = MouseSelection Then
        shpSelection.Visible = False
      Else
        shpDistance.Visible = False
      End If
    End If
    '
    If mFirstX = mPreviousX Or mFirstY = mPreviousY Then
      'Single Mouseklik
      RaiseEvent MousePostion(ClickedPosition, mFirstX, mFirstY)
    ElseIf IIf(mPreviousMouseX < X, X - mPreviousMouseX, _
    mPreviousMouseX - X) < mSelectionRange Or _
      IIf(mPreviousMouseY < Y, Y - mPreviousMouseY, mPreviousMouseY - Y) < mSelectionRange Then
      'the user meant to click on one position in stead of zooming in
      RaiseEvent ZoomError(SelectionIgnored)
    ElseIf mMouseFunction = MouseZoom Or mMouseFunction = MouseSelection Then
      'if the user draws with the mouse from right to left we have to switch the values
      If mFirstX > mPreviousX Then Swap mFirstX, mPreviousX
      If mFirstY < mPreviousY Then Swap mFirstY, mPreviousY
      If mMouseFunction = MouseSelection Then
        RaiseEvent SelectionCompleted(Selection, mFirstY, mPreviousX, mPreviousY, mFirstX)
      Else
        RaiseEvent SelectionCompleted(ZoomSelection, mFirstY, mPreviousX, mPreviousY, mFirstX)
      End If
    ElseIf mMouseFunction = MouseDistance Then
      RaiseEvent DistanceCalculated(Carc_Distance_Tussen(mFirstX, mFirstY, mPreviousX, mPreviousY), mFirstX, mFirstY, mPreviousX, mPreviousY)
    End If
  End If
End Sub

Private Sub Swap(ByRef value1 As Double, ByRef value2 As Double)
  Dim hulp As Double
  hulp = value1
  value1 = value2
  value2 = hulp
End Sub

' for drawing the map
Private Function Calculate_X(vLongitude As Double) As Double
  If vLongitude > mCurrentLimit.W And vLongitude < mCurrentLimit.O Then
    vLongitude = vLongitude - mCurrentLimit.W
    Calculate_X = Me.ScaleWidth * _
    (Abs(vLongitude) / (mCurrentLimit.O - mCurrentLimit.W))
  Else
    Calculate_X = -1 ' vLongitude out of zoomarea
  End If
End Function

Private Function Calculate_Y(vAltitude As Double) As Double
  If vAltitude > mCurrentLimit.Z And vAltitude < mCurrentLimit.N Then
    vAltitude = vAltitude - mCurrentLimit.N
    Calculate_Y = Me.ScaleHeight * _
    (Abs(vAltitude) / (mCurrentLimit.N - mCurrentLimit.Z))
  Else
    Calculate_Y = -1 ' valtitude out of zoomarea
  End If
End Function

Private Function SystemZoomIn(vLimitN As Double, vLimitO As Double, vLimitZ As Double, vLimitW As Double, Optional IsMove As Boolean = False, Optional IsInit As Boolean) As enmZoomError
  'because we receive values through a function
  'we can not use the checkfunctions from the properties
  If IsInit = True Or Not (mCurrentLimit.W = vLimitW And mCurrentLimit.O = vLimitO And mCurrentLimit.Z = vLimitZ And mCurrentLimit.N = vLimitN) Then
    If vLimitW > ZoomHistory(1).O Or vLimitW < ZoomHistory(1).W Then
      SystemZoomIn = West
    ElseIf vLimitO > ZoomHistory(1).O Or vLimitO < ZoomHistory(1).W Then
      SystemZoomIn = East
    ElseIf vLimitZ > ZoomHistory(1).N Or vLimitZ < ZoomHistory(1).Z Then
      SystemZoomIn = South
    ElseIf vLimitN > ZoomHistory(1).N Or vLimitN < ZoomHistory(1).Z Then
      SystemZoomIn = North
    ElseIf vLimitO = vLimitW Then
      SystemZoomIn = East_West_Equal
    ElseIf vLimitN = vLimitZ Then
      SystemZoomIn = North_South_Equal
    ElseIf vLimitW > vLimitO Then
      SystemZoomIn = East_SmallerThan_West
    ElseIf vLimitZ > vLimitN Then
      SystemZoomIn = North_SmallerThan_South
    Else
      '
      'all values are OK
      mCurrentLimit.W = vLimitW
      mCurrentLimit.O = vLimitO
      mCurrentLimit.Z = vLimitZ
      mCurrentLimit.N = vLimitN
      '
      'the first point may not be connected
      'to the upper-left corner
      mBreakLine = True
      '
      '
      'if isInit=true then the Limit values are allready been checked
      '(zie UserControl_Readproperties())
      If IsMove = False And Not IsInit Then _
      Call VoegZoomHistoryToTableToe(mCurrentLimit.N, mCurrentLimit.O, mCurrentLimit.Z, mCurrentLimit.W) ' sla geen moves op door move()
      '
      'do not draw if IsInit because it would be drawn twice
      If Not IsInit Then Call PrepareZoomConfiguredEvent(ZoomInDirection)
      '
    End If
  End If
End Function

Private Sub VoegZoomHistoryToTableToe(vLimitN As Double, vLimitO As Double, vLimitZ As Double, vLimitW As Double)
  mZoomHistoryTel = mZoomHistoryTel + 1
  ReDim Preserve ZoomHistory(1 To mZoomHistoryTel)
  ZoomHistory(mZoomHistoryTel).N = vLimitN
  ZoomHistory(mZoomHistoryTel).O = vLimitO
  ZoomHistory(mZoomHistoryTel).Z = vLimitZ
  ZoomHistory(mZoomHistoryTel).W = vLimitW
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyNagivate Then
    Select Case KeyCode
      Case 37 ' Left
        Navigate DirectionW, mNagivateStep
      Case 38 ' Up
        Navigate DirectionN, mNagivateStep
      Case 39 ' Right
        Navigate DirectionO, mNagivateStep
      Case 40 ' Down
        Navigate DirectionZ, mNagivateStep
    End Select
  End If
End Sub

Private Sub UserControl_Resize()
  Call MapReload
End Sub

'property bag
'------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '
  With PropBag
    Me.LimitO = CDbl(.ReadProperty("LimitO", 180))
    Me.LimitW = CDbl(.ReadProperty("LimitW", -180))
    Me.LimitN = CDbl(.ReadProperty("LimitN", 90))
    Me.LimitZ = CDbl(.ReadProperty("LimitZ", -90))
    '
    Me.Redraw = CBool(.ReadProperty("Redraw", False))
    Me.DrawLines = CBool(.ReadProperty("DrawLines", False))
    '
    Me.MapBackColor = .ReadProperty("MapBackColor", vbWhite)
    Me.DrawColor = .ReadProperty("DrawColor", vbBlack)
    Me.MouseFunction = .ReadProperty("MouseFunction", MouseZoom)
    '
    Me.NagivateStep = .ReadProperty("NagivateStep", 20)
    Me.KeyNagivate = .ReadProperty("KeyNagivate", False)
    Me.SelectionRatio = .ReadProperty("SelectionRatio", 0)
    Me.SelectionRange = .ReadProperty("SelectionRange", 0)
  End With
  '
  'The First Limits are allready checked during design-modus,
  'and can be safely be placed at the first position in the zoomtable
  'These Limits at Position 1 are never deleted
  Call VoegZoomHistoryToTableToe(Me.LimitN, Me.LimitO, Me.LimitZ, Me.LimitW)
  Call SystemZoomIn(ZoomHistory(1).N, ZoomHistory(1).O, ZoomHistory(1).Z, ZoomHistory(1).W, False, True)
End Sub

Private Sub UserControl_Show()
  UserControl.BackColor = MapBackColor
  UserControl.BackColor = MapBackColor
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "LimitO", Me.LimitO, 180
    .WriteProperty "LimitW", Me.LimitW, -180
    .WriteProperty "LimitN", Me.LimitN, 90
    .WriteProperty "LimitZ", Me.LimitZ, -90
    '
    .WriteProperty "Redraw", CBool(Me.Redraw), False
    .WriteProperty "DrawLines", CBool(Me.DrawLines), False
    '
    .WriteProperty "MapBackColor", Me.MapBackColor, vbWhite
    .WriteProperty "DrawColor", Me.DrawColor, vbBlack
    .WriteProperty "MouseFunction", Me.MouseFunction, MouseZoom
    '
    .WriteProperty "NagivateStep", Me.NagivateStep, 20
    .WriteProperty "KeyNagivate", Me.KeyNagivate, False
    '
    .WriteProperty "SelectionRatio", Me.SelectionRatio, 0
    .WriteProperty "SelectionRange", Me.SelectionRange, 0
  End With
End Sub

Private Sub UserControl_InitProperties()
  'Only called when the OCX is selected from the Toolbar
  '
  mCurrentLimit.W = -180
  mCurrentLimit.O = 180
  mCurrentLimit.Z = -90
  mCurrentLimit.N = 90
  mNagivateStep = 20
  '
  mMapBackColor = &H80000000
  mDrawColor = vbBlack
  mMouseFunction = MouseZoom
  '
  UserControl.BackColor = MapBackColor
  UserControl.BackColor = MapBackColor
  '
  SelectionRatio = 0
End Sub
