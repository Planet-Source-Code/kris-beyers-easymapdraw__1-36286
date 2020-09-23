VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'local variable(s) to hold property value(s)
Private mvarX As Double 'local copy
Private mvarY As Double 'local copy
Private mvarColor As ColorConstants 'local copy
Private mvarBreakLine As Boolean 'local copy
Public Property Let BreakLine(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BreakLine = 5
    mvarBreakLine = vData
End Property


Public Property Get BreakLine() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BreakLine
    BreakLine = mvarBreakLine
End Property



Public Property Let Color(ByVal vData As ColorConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Color = Form1
    mvarColor = vData
End Property


Public Property Get Color() As ColorConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    Color = mvarColor
End Property



Public Property Let Y(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get Y() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property



Public Property Let X(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get X() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    X = mvarX
End Property



