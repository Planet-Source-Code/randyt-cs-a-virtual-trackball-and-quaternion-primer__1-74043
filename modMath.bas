Attribute VB_Name = "modMath"
'
'//////////////////////[ Math Module ]/////////////////////////
'// Coded for single-precision floating point accuracy.
'// Numbers with ! suffix are single-precision floating point.
'//////////////////////////////////////////////////////////////

Option Explicit

'Public Const PI As Double = 3.14159265358979
Public Const PI As Single = 3.141593!

'Public Const INFINITY As Long = 2147483647
'Public Const INFINITY As Double = 2147483647#
Public Const INFINITY As Single = 2.147484E+09!

'Return Inverse Sine, y = Arcsin(x)
Public Function ArcSin(ByVal X1 As Single) As Single
'domain: -1 <= x <= 1, range: -pi/2 <= y <= pi/2
    Select Case X1
        Case 1!
            ArcSin = PI / 2!
        Case -1!
            ArcSin = -PI / 2!
        Case Else
            If Abs(X1) < 1! Then
                ArcSin = Atn(X1 / (-X1 * X1 + 1!) ^ 0.5!)
            Else
                MsgBox "ArcSin(X) domain error [-1 to +1]: X = " & X1
            End If
    End Select
End Function

'Return Inverse Cosine, y = ArcCos(x)
Public Function ArcCos(X1 As Single) As Single
'domain: -1 <= x <= 1, range: 0 <= y <= pi
    Select Case X1
        Case 1!
            ArcCos = 0!
        Case -1!
            ArcCos = PI
        Case Else
            If Abs(X1) < 1! Then
                ArcCos = Atn(-X1 / (-X1 * X1 + 1!) ^ 0.5!) + PI / 2!
            Else
                MsgBox "ArcCos(X) domain error [-1 to +1]: X = " & X1
            End If
    End Select
End Function

'Degrees to Radians
Public Function d2r(deg As Single) As Single
    d2r = deg * PI / 180!
End Function

'Radians to Degrees
Public Function r2d(rad As Single) As Single
    r2d = rad * 180! / PI
End Function
