VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\AprjEasyMapDraw.vbp"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   Caption         =   "Easy map draw"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8415
   Icon            =   "Map.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSavePicture 
      Caption         =   "&Save Image"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   4920
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   3540
      TabIndex        =   19
      Top             =   5475
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "20"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   540
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   540
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   315
   End
   Begin VB.CommandButton cmdnavigate 
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   315
   End
   Begin MSComctlLib.StatusBar strStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   5445
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9667
            MinWidth        =   1941
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "long: 0.00°"
            TextSave        =   "long: 0.00°"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "lat: 0.00°"
            TextSave        =   "lat: 0.00°"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "Zoom &out"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   165
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoomIn 
      Caption         =   "Zoom &in"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   165
      Width           =   1215
   End
   Begin VB.CommandButton cmdReDraw 
      Caption         =   "&Redraw"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame frmZoom 
      Caption         =   "Zoom"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtLimit 
         Height          =   285
         Index           =   3
         Left            =   4080
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtLimit 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtLimit 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtLimit 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblNegBreedte 
         Caption         =   "W:"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblNegLengte 
         Caption         =   " Z:"
         Height          =   255
         Left            =   2580
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblPosBreedte 
         Caption         =   "O:"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblPosLengte 
         Caption         =   "N:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin prjEasyMapDraw.EasyMapDraw EasyMapDraw 
      Height          =   3855
      Left            =   360
      TabIndex        =   25
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
      Redraw          =   -1  'True
      DrawLines       =   -1  'True
      MapBackColor    =   -2147483648
      KeyNagivate     =   -1  'True
      SelectionRange  =   100
   End
   Begin VB.Label lblF1 
      Caption         =   "Use F1 and right mousebutton for more options"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblStep 
      Caption         =   "Step %"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   4920
      Width           =   735
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCalculate 
         Caption         =   "&Zoom with mouse"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuCalculate 
         Caption         =   "&Select with mouse"
         Index           =   1
      End
      Begin VB.Menu mnuCalculate 
         Caption         =   "&Calculate distance with mouse"
         Index           =   2
      End
      Begin VB.Menu mnuLijn1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawLines 
         Caption         =   "&Draw with &lines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRedrawEnabled 
         Caption         =   "&Redraw enabled (slower)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuKeepAspectRatio 
         Caption         =   "&Keep aspect ration"
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this sample is based on the maps of "The World Digitized" which you can find here
'http://uiarchive.uiuc.edu/mirrors/ftp/ftphost.simtel.net/pub/simtelnet/msdos/worldmap/
'put the MP1 files in the maps folder of this project
'or by searching on "The World Digitized" and Allison and ftp
'The World Digitized is copyrighted by John B. Allison 1986
'
'Navigation buttons uses fonttype Wingdings 3
'
' What's new in EasyMapDraw 1.4 example ?
' ---------------------------------------
' * new SaveImage button, to save the image in a BMP file
' * disturbing lines are removed when drawing with lines, thanks for the tip Maarten :-)
'
Option Explicit

'Map events
'----------
Private Sub EasyMapDraw_MousePostion(vMouseEvent As enmMouseEvent, vLongitude As Double, vLatitude As Double)
  With strStatus
    Select Case vMouseEvent
      Case CurrentPosition
        .Panels(2).Text = "long: " & Format(vLongitude, "0.00°")
        .Panels(3).Text = "lat: " & Format(vLatitude, "0.00°")
      Case SelectionStartPosition
        .Panels(1).Text = "Selection from: " & _
        Format(vLongitude, "0.00°") & ", " & Format(vLatitude, "0.00°") & " too ->  "
      Case DistanceStartPosition
        .Panels(1).Text = "Distance Calculation from: " & _
        Format(vLongitude, "0.00°") & " " & Format(vLatitude, "0.00°") & " too ->  "
      Case ClickedPosition
        .Panels(1).Text = "Position: " & _
        Format(vLongitude, "0.00°") & ", " & Format(vLatitude, "0.00°") & " has been clicked"
    End Select
  End With
End Sub

Private Sub EasyMapDraw_DistanceCalculated(vDistance As Double, vLatitude1 As Double, vLongitude1 As Double, vLatitude2 As Double, vLongitude2 As Double)
  strStatus.Panels(1).Text = "The distance from " & Format(vLatitude1, "0.00°") & ", " & Format(vLongitude1, "0.00°") & " too " & Format(vLatitude2, "0.00°") & ", " & Format(vLongitude2, "0.00°") & " is: " & Format(vDistance, "#0.00") & " km"
End Sub

Private Sub EasyMapDraw_SelectionCompleted(vSelectionEvent As prjEasyMapDraw.enmSelectionEvent, vLimitN As Double, vLimitO As Double, vLimitZ As Double, vLimitW As Double)
  If vSelectionEvent = Selection Then
    MsgBox "you selected:" & vbCrLf & vbCrLf & vbTab & _
    " N: " & Format(vLimitN, "0.00°") & vbCrLf & _
    " W: " & Format(vLimitW, "0.00°") & vbTab & _
    " O: " & Format(vLimitO, "0.00°") & vbCrLf & vbTab & _
    " Z: " & Format(vLimitZ, "0.00°")
  ElseIf vSelectionEvent = ZoomSelection Then
    EasyMapDraw.ZoomIn vLimitN, vLimitO, vLimitZ, vLimitW
  End If
End Sub

Private Sub EasyMapDraw_ZoomError(vZoomError As enmZoomError)
  Select Case vZoomError
    Case North_South_Equal
      MsgBox "N and Z may not be equal"
      txtLimit(1).SetFocus
    Case East_West_Equal
      MsgBox "O and W may not be equal"
      txtLimit(0).SetFocus
    Case East_SmallerThan_West
      MsgBox "value W must be smaller than O"
      txtLimit(1).SetFocus
    Case North_SmallerThan_South
      MsgBox "value Z must be smaller than N"
      txtLimit(0).SetFocus
    Case ToMuchSteps ' The user tried to go back more steps in the zoomhistory than there are available
      MsgBox "Value must be between 1 and " & EasyMapDraw.ZoomHistoryTel
      cmdZoomOut_Click
    Case SelectionIgnored ' The selection with the mouse was smaller than SelectionRange
      strStatus.Panels(1).Text = "the selection was too small"
  End Select
End Sub

Private Sub EasyMapDraw_ZoomConfigured(vZoomdirection As enmZoomDirection)
  'Add here your code to query a database of coördinates
  Dim latDeg As Double, lngDeg As Double, color As ColorConstants
  Dim FileTeller As Byte, tmpFile As String, F As Integer
  Dim A As String, I As Long
  '
  pgbProgress.Top = IIf(EasyMapDraw.Redraw = True, _
  EasyMapDraw.ScaleHeight / 2 - pgbProgress.Height / 2 + EasyMapDraw.Top, frmMap.Height - 778)
  pgbProgress.Left = IIf(EasyMapDraw.Redraw = True, _
  EasyMapDraw.ScaleWidth / 2 - pgbProgress.Width / 2 + EasyMapDraw.Left, frmMap.Width - 4990)
  '
  Select Case vZoomdirection
    Case ZoomInDirection
      frmMap.Caption = "Easy map draw - zoomed in"
    Case ZoomOutDirection
      frmMap.Caption = "Easy map draw - zoomed out"
    Case ZoomRefresh
      frmMap.Caption = "Easy map draw - refreshed"
  End Select
  '
  Screen.MousePointer = vbHourglass
  cmdReDraw.Enabled = False
  cmdZoomIn.Enabled = False
  cmdZoomOut.Enabled = False
  pgbProgress.Visible = True
  EasyMapDraw.DrawEquatorMeridian DrawEquatorAndMeridian
  '
  tmpFile = Dir(App.Path & "\maps\*.MP1")
  If tmpFile <> "" Then
    '
    ' The Drawloop
    '-------------
    Do
      '
      strStatus.Panels(1).Text = "Loading card-fragment: " & tmpFile & " ..."
      '
      EasyMapDraw.BreakLine  ' interrupt drawing of a line between two points
      '
      ' change the color a couple of times
      If tmpFile Like "AN*" Then
        color = vbCyan
      ElseIf tmpFile Like "AU*" Then
        color = vbBlue
      ElseIf tmpFile Like "GR*" Or _
      tmpFile Like "USA*" Or tmpFile Like "NA*" Or tmpFile Like "SA*" Then
        color = vbGreen
      ElseIf tmpFile Like "E*" Then
        color = vbBlack
      ElseIf tmpFile Like "AS*" Then
        color = vbYellow
      ElseIf tmpFile Like "PA*" Then
        color = vbMagenta
      ElseIf tmpFile Like "AF*" Then
        color = vbRed
      End If
      EasyMapDraw.DrawColor = color
      '
      F = FreeFile()
      Open App.Path & "\maps\" & tmpFile For Input As #F
      If Not EOF(F) Then
        Do
          '
          'get the point coördinates
          Line Input #F, A
          '
          ' this line has been added recently,
          ' to lift the "pen" when moving to another piece of land
          If A = vbNullString Then EasyMapDraw.BreakLine
          '
          I = InStr(1, A, " ")
          If I > 0 Then
             latDeg = Val(Left(A, I - 1))
             If latDeg <> 0 Then
                lngDeg = Val(Mid(A, I + 1))
                EasyMapDraw.DrawPoint lngDeg, latDeg
             End If
          End If
        Loop While Not EOF(F) And EasyMapDraw.IsDrawing
      End If
      '
      Close #F
      tmpFile = Dir
      '
      FileTeller = FileTeller + 1
      pgbProgress.Value = CByte(FileTeller / 30 * 100) ' totaal 29 bestanden
      '
      DoEvents
    Loop While tmpFile <> "" And EasyMapDraw.IsDrawing
  Else
    'no MP1 files found in maps folder
    MsgBox "download first the 30 MP1 files from http://uiarchive.uiuc.edu/mirrors/ftp/ftphost.simtel.net/pub/simtelnet/msdos/worldmap/" & vbCrLf & "and copy them in a folder named Maps, this folder is placed in the project map", vbExclamation
    End
  End If
  '
  cmdReDraw.Enabled = True
  cmdZoomIn.Enabled = True
  cmdZoomOut.Enabled = Not EasyMapDraw.IsMaxZoomedOut
  pgbProgress.Visible = False
  Screen.MousePointer = vbDefault
  '
  strStatus.Panels(1).Text = "Drawing complete N: " & _
  Format(EasyMapDraw.LimitN, "#0.00°") & "  O: " & _
  Format(EasyMapDraw.LimitO, "#0.00°") & "  Z: " & _
  Format(EasyMapDraw.LimitZ, "#0.00°") & "  W: " & _
  Format(EasyMapDraw.LimitW, "#0.00°")
  '
  With EasyMapDraw
    txtLimit(1).Text = Format(.LimitO, "0.00")
    txtLimit(3).Text = Format(.LimitW, "0.00")
    txtLimit(0).Text = Format(.LimitN, "0.00")
    txtLimit(2).Text = Format(.LimitZ, "0.00")
  End With
  '
  ' ////\\\\////\\\\////\\\\////
  EasyMapDraw.EndDraw  ' Important
  ' ////\\\\////\\\\////\\\\////
  '
  ' if you zoom in too Belgium (Europe, near to Meridian Greenwitch line)
  ' you can see where I live ;-)
  EasyMapDraw.DrawPicture App.Path & "\bullets\gl_rd.gif", 100, 100, 4, 51, 12, 19
  '
End Sub

Private Sub Form_Resize()
  mdlResize.FormResize Me
End Sub

Private Sub mnuCalculate_Click(Index As Integer)
  mnuCalculate(2).Checked = IIf(Index = 2, True, False)
  mnuCalculate(1).Checked = IIf(Index = 1, True, False)
  mnuCalculate(0).Checked = IIf(Index = 0, True, False)
  EasyMapDraw.MouseFunction = Index
End Sub

Private Sub mnuDrawlines_Click()
  mnuDrawLines.Checked = IIf(mnuDrawLines.Checked = False, True, False)
  EasyMapDraw.DrawLines = mnuDrawLines.Checked
End Sub

Private Sub EasyMapDraw_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
  If Button = 2 Then Me.PopupMenu mnuPopup
End Sub

Private Sub mnuKeepAspectRatio_Click()
  mnuKeepAspectRatio.Checked = IIf(mnuKeepAspectRatio.Checked = False, True, False)
  EasyMapDraw.SelectionRatio = IIf(mnuKeepAspectRatio.Checked, 50, 0)
End Sub

Private Sub mnuRedrawEnabled_Click()
  mnuRedrawEnabled.Checked = IIf(mnuRedrawEnabled.Checked = False, True, False)
  EasyMapDraw.Redraw = mnuRedrawEnabled.Checked
End Sub

Private Sub txtStep_LostFocus()
  If CInt(txtStep.Text) < 10 Then txtStep.Text = 10
  EasyMapDraw.NagivateStep = CByte(txtStep.Text)
End Sub

'form events
'-----------
Private Sub Form_Load()
  Me.KeyPreview = True ' otherwise the buttons are not being intercepted
  pgbProgress.Value = 0
  
  'load coördinates in memory
  '--------------------------
  
  
End Sub

Private Sub txtLimit_GotFocus(Index As Integer)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtLimit_KeyPress(Index As Integer, KeyAscii As Integer) 'laat alleen [backspace] - . en getallen toe
  Dim hulpValue As Integer
  If KeyAscii <> 8 And KeyAscii <> 45 And KeyAscii <> 46 And _
  KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
  ElseIf Not (KeyAscii < 48 Or KeyAscii > 57) Then
    hulpValue = IIf(Index = 0 Or Index = 2, 90, 180)
    If CDbl(txtLimit(Index).Text & Chr(KeyAscii)) > hulpValue Or _
    CDbl(txtLimit(Index).Text & Chr(KeyAscii)) < -hulpValue Then KeyAscii = 0
  End If
End Sub

Private Sub cmdZoomIn_Click()
  EasyMapDraw.ZoomIn CDbl(txtLimit(0).Text), CDbl(txtLimit(1).Text), _
  CDbl(txtLimit(2).Text), CDbl(txtLimit(3).Text)
End Sub

Private Sub cmdZoomOut_Click()
  Dim antw As String, ErrorVlag As Boolean
  '
  If EasyMapDraw.ZoomHistoryTel > 1 Then
    Do
      ErrorVlag = False
      antw = InputBox("Give the number of steps you which to go back in the zoomhistory" & _
      vbCrLf & "give a value from 1 to " & EasyMapDraw.ZoomHistoryTel, "Zoom uit", 1)
      If antw <> vbNullString Then
        If IsNumeric(antw) = True And Len(antw) < 4 Then
          If CInt(antw) < 1 Or CInt(antw) > 256 Then
            ErrorVlag = True
          Else
            Call EasyMapDraw.ZoomOut(CByte(antw))
          End If
        Else
          ErrorVlag = True
        End If
        If ErrorVlag = True Then MsgBox "Only numbers between 1 and 255 are allowed"
      End If
    Loop While ErrorVlag = True
  Else
    Call EasyMapDraw.ZoomOut(1)
  End If
End Sub

Private Sub txtLimit_LostFocus(Index As Integer)
  If Not IsNumeric(txtLimit(Index).Text) Then
    Select Case Index
      Case 0
        txtLimit(Index).Text = 90
      Case 1
        txtLimit(Index).Text = 180
      Case 2
        txtLimit(Index).Text = -90
      Case 3
        txtLimit(Index).Text = -180
    End Select
  End If
End Sub

Private Sub txtStep_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
  ElseIf Not (KeyAscii < 48 Or KeyAscii > 57) Then
    If CInt(txtStep.Text & Chr(KeyAscii)) > 100 Then KeyAscii = 0
  End If
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
  EasyMapDraw.Navigate Index + 1, CDbl(txtStep.Text)
End Sub

Private Sub cmdSavePicture_Click()
  Call EasyMapDraw.SaveImage(App.Path & "\outputImage.bmp")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim newStep As Byte
  '
  With EasyMapDraw
    Select Case KeyCode
      Case 27 ' ESC
        .EndDraw
      Case 112 ' F1
        MsgBox vbCrLf & _
        "ESC: cancel loading of the map" & vbCrLf & _
        "F1: show this helpmessage" & vbCrLf & _
        "F2: switch between drawing of lines or points" & vbCrLf & _
        "F3, F4: increasing or decreasing the step size of the navigation buttons" & vbCrLf & _
        "F5: redraw the map" & vbCrLf & _
        "F6: switch between zooming in or distance calculation with the mouse" & vbCrLf & _
        "F7, ENTER: zooming in with the choosen N, O, Z, W values" & vbCrLf & _
        "F8: zooming out too the previous N, O, Z, W values" & vbCrLf & _
        "F9, F10, F11, F12: selecting the inputbox N, O, Z or W" & vbCrLf & _
        "ARROWS: navigate too N, O, Z, W with the configured step size"
      Case 113 ' F2
        mnuDrawlines_Click
      Case 114 ' F3
        newStep = IIf(CByte(txtStep.Text) - 10 < 10, 10, CByte(txtStep.Text) - 10)
        txtStep.Text = newStep
        .NagivateStep = newStep
      Case 115 ' F4
        newStep = IIf(CByte(txtStep.Text) + 10 > 100, 100, CByte(txtStep.Text) + 10)
        txtStep.Text = newStep
        .NagivateStep = newStep
      Case 116 ' F5
        .MapReload
      Case 117 ' F6
        If mnuCalculate(0).Checked Then
          Call mnuCalculate_Click(1)
        Else
          Call mnuCalculate_Click(2)
        End If
      Case 118 ' F7
        cmdZoomIn_Click
      Case 119 ' F8
        If Not .IsMaxZoomedOut Then cmdZoomOut_Click
      Case 120 ' F9
        txtLimit(0).SetFocus
      Case 121 ' F10
        txtLimit(1).SetFocus
      Case 122 ' F11
        txtLimit(2).SetFocus
      Case 123 ' F12
        txtLimit(3).SetFocus
    End Select
  End With
End Sub

Private Sub cmdReDraw_Click()
  EasyMapDraw.MapReload
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If EasyMapDraw.IsDrawing Then
    EasyMapDraw.EndDraw
    Cancel = True
  End If
End Sub
