Attribute VB_Name = "mdlCalculateDistance"
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

