Attribute VB_Name = "modVector"
'
'/////////////////////[ Vector Module ]///////////////////////////
'// Coded for single-precision floating point accuracy...
'// Numbers with ! suffix are single-precision floating point.
'/////////////////////////////////////////////////////////////////
'// This module does not use user-defined data types.
'// This module is designed for quick conversion to 'C'.
'/////////////////////////////////////////////////////////////////
'//
'// Module for standard 3-D vector operations upon the first
'// three components, (0,1,2) of:
'// (1) Standard 3-D vectors:  <x,y,z>.
'// (2) Homogeneous points:    <x,y,z,s>.
'// (3) Quaternions:           <b,c,d,a>.
'//
'/////////////////////[ Implementation ]//////////////////////////
'//
'// Procedure arguments are transfered by reference (default).
'//
'// For subroutine procedures that modify argument values, this
'// module positions all of the arguments in the same order as
'// standard computer programming assignment statements:
'// vsum = va + vb <-> v3add(vsum, va, vb).
'//
'// Commutative arguments are suffixed:     a, b
'// Noncommutative arguments are suffixed:  1, 2
'//
'////////////////////////////////////////////////////////////////

Option Explicit

'Set vector:
'v() = <x,y,z>
Public Sub v3set(v() As Single, x As Single, Y As Single, Z As Single)
    v(0) = x
    v(1) = Y
    v(2) = Z
End Sub

'Set vector copy:
'vdst() = vsrc()
Public Sub v3copy(vdst() As Single, vsrc() As Single)
    vdst(0) = vsrc(0)
    vdst(1) = vsrc(1)
    vdst(2) = vsrc(2)
End Sub

'Set vector sum:
'vsum() = va() + vb()
Public Sub v3add(vsum() As Single, va() As Single, vb() As Single)
    vsum(0) = va(0) + vb(0)
    vsum(1) = va(1) + vb(1)
    vsum(2) = va(2) + vb(2)
End Sub

'Set vector difference:
'vdif() = va() - vb()
Public Sub v3dif(vdif() As Single, va() As Single, vb() As Single)
    vdif(0) = va(0) - vb(0)
    vdif(1) = va(1) - vb(1)
    vdif(2) = va(2) - vb(2)
End Sub

'Return scalar dot product:
'dot = va() dot vb()
Public Function v3getDot(va() As Single, vb() As Single) As Single
    v3getDot = va(0) * vb(0) + va(1) * vb(1) + va(2) * vb(2)
End Function

'Set vector cross product:
'vcross() = v1() cross v2()
'NONCOMMUTATIVE! [v1 cross v2] <> [v2 cross v1]
Public Sub v3cross(vcross() As Single, v1() As Single, v2() As Single)
'v1 and/or v2 may be the same vector as vcross.
Dim va(0 To 2) As Single
Dim vb(0 To 2) As Single

    'Copy v1 and v2 (factor) arguments:
    Call v3copy(va, v1)
    Call v3copy(vb, v2)
    
    'We cross the copies of v1 and v2:
    vcross(0) = (va(1) * vb(2)) - (va(2) * vb(1))
    vcross(1) = (va(2) * vb(0)) - (va(0) * vb(2))
    vcross(2) = (va(0) * vb(1)) - (va(1) * vb(0))
End Sub

'Return scalar magnitude (length).
Public Function v3getMag(v() As Single) As Single
    v3getMag = (v(0) * v(0) + v(1) * v(1) + v(2) * v(2)) ^ 0.5!
End Function

'Set vector scalar magnitude = 1 (normalize).
Public Sub v3normalize(v() As Single)
Dim mag As Single

    mag = v3getMag(v)
    v(0) = v(0) / mag
    v(1) = v(1) / mag
    v(2) = v(2) / mag
End Sub

'Set vector (scalar multiply).
Public Sub v3scale(v() As Single, sf As Single)
    v(0) = v(0) * sf
    v(1) = v(1) * sf
    v(2) = v(2) * sf
End Sub

'Print vector <x,y,z> to the Immediate Window.
Public Sub v3Print(v() As Single)
    Debug.Print "x,y,z: <"; v(0); ", "; v(1); ", "; v(2); ">"
End Sub
