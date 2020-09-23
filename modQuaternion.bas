Attribute VB_Name = "modQuaternion"
'
'/////////////////////[ Quaternion Module ]///////////////////////
'// Coded for single-precision floating point accuracy...
'// Numbers with ! suffix are single-precision floating point.
'/////////////////////////////////////////////////////////////////
'// This module does not use user-defined data types.
'// This module is designed for quick conversion to 'C'.
'/////////////////////////////////////////////////////////////////
'//
'// By Randy Manning, Jul-2011.
'//
'/////////////////////[ About Quaternions ]///////////////////////
'//
'// A quaternion is an extension of the complex numbers having
'// a real part and three imaginary parts: q = a + bi + cj + dk.
'//
'// Quaternion multiplication rules:
'// i^2 = j^2 = k^2 = -1
'// ij = k, ji = -k
'// jk = i, kj = -i
'// ki = j, ik = -j
'//
'// Many people are intimidated by complex and imaginary numbers.
'// There's nothing really complex or imaginary about quaternions.
'// All that these rules really mean is that when you multiply
'// two quaternions and you get an i*i in one of your terms, you
'// just replace it with a -1, that's all. It's just a way of
'// making sure that you get the right signs for the terms in
'// your answer.
'//
'/////////////////////[ Implementation ]//////////////////////////
'//
'// Procedure arguments are transfered by reference (default).
'//
'// For subroutine procedures that modify argument values, this
'// module positions all of the arguments in the same order as
'// standard computer programming assignment statements:
'// qp = q1 * q2 <-> q4multiply(qp, q1, q2).
'//
'// Commutative arguments are suffixed:     a, b
'// Noncommutative arguments are suffixed:  1, 2
'//
'/////////////////////////////////////////////////////////////////
'//
'// This module uses the following quaternion format:
'// q = a + bi + cj + dk <-> q(b,c,d,a),
'// where q(3)=(a) and q(0,1,2)=(b,c,d)
'//
'/////////////////////////////////////////////////////////////////

Option Explicit

'Set quaternion:
'q() = <b,c,d,a>
Public Sub q4set(q() As Single, b As Single, c As Single, d As Single, a As Single)
    q(0) = b
    q(1) = c
    q(2) = d
    q(3) = a
End Sub

'Set quaternion copy:
'qdst() = qsrc()
Public Sub q4copy(qdst() As Single, qsrc() As Single)
    qdst(0) = qsrc(0)
    qdst(1) = qsrc(1)
    qdst(2) = qsrc(2)
    qdst(3) = qsrc(3)
End Sub

'Set quaternion product:
'qp() = q1() * q2()
'qp is normalized on return, (qp magnitude = 1)
'NONCOMMUTATIVE! [q1 * q2] <> [q2 * q1]
'q1(0,-6,3,2) * q2(3,-2,2,1) = qp(0, -1, 25, -16)    <- qp not normalized
'q2(3,-2,2,1) * q1(0,-6,3,2) = qp(12, -19, -11, -16) <- qp not normalized
'24 x, 15 +, 1 sqr:
Public Sub q4multiply(qp() As Single, q1() As Single, q2() As Single)
'NOTE: q1 and/or q2 may be the same quaternion as qp.
Dim qa(0 To 3) As Single 'copy of q1
Dim qb(0 To 3) As Single 'copy of q2

    'Copy q1 and q2 (factor) arguments:
    Call q4copy(qa, q1)
    Call q4copy(qb, q2)
    
    'q1 = a + bi + cj + dk,  [a=q1(3), b=q1(0), c=q1(1), d=q1(2)]
    'q2 = e + fi + gj + hk,  [e=q2(3), f=q2(0), g=q2(1), h=q2(2)]
    
    'To verify the multiplication product by hand, use
    'the following sign convention when combining the
    'product terms:
    'i^2 = j^2 = k^2 = -1
    'ij = k, ji = -k
    'jk = i, kj = -i
    'ki = j, ik = -j
    
    'I used Mathematica to symbolically verify the following
    'quaternion multiplication formula (commonly found using
    'Google internet searches). The formula is indeed correct.
    'The symbolic product yields the following formula:
    'qp(3) = (ae - bf - cg - dh)    <- the real part
    'qp(0) = (af + be + ch - dg)    <- the i part
    'qp(1) = (ag - bh + ce + df)    <- the j part
    'qp(2) = (ah + bg - cf + de)    <- the k part
    
    'We multiply the copies of q1 and q2:
    '16 x, 12 +:
    qp(3) = (qa(3) * qb(3) - qa(0) * qb(0) - qa(1) * qb(1) - qa(2) * qb(2))
    qp(0) = (qa(3) * qb(0) + qa(0) * qb(3) + qa(1) * qb(2) - qa(2) * qb(1))
    qp(1) = (qa(3) * qb(1) - qa(0) * qb(2) + qa(1) * qb(3) + qa(2) * qb(0))
    qp(2) = (qa(3) * qb(2) + qa(0) * qb(1) - qa(1) * qb(0) + qa(2) * qb(3))
    
    'We normalize the product on return:
    '8 x, 3 +, 1 sqr:
    q4normalize qp
End Sub

'Return quaternion magnitude.
'4 x, 3 +, 1 sqr.
Public Function q4getMag(q() As Single) As Single
'Just as vectors, quaternions have a magnitude too.
    q4getMag = (q(0) * q(0) + q(1) * q(1) + q(2) * q(2) + q(3) * q(3)) ^ 0.5!
End Function

'Set quaternion magnitude = 1, (normalize it).
'8 x, 3 +, 1 sqr.
Public Sub q4normalize(q() As Single)
'Just as vectors, quaternions can be normalized too.
'
'Unit rotation quaternions converted into rotation
'matrices always produce exact rotation matrices:
'(1) R_transpose * R = identity matrix
'(2) The determinant of R = 1
'
'Accumulated error in a should be 'unit' rotation
'quaternion (such as the m_Current_Rotation_Quaternion,
'which in this program, is normalized at every
'multiplication/combination) will, over time, cause
'changes to both the scale and rotation axes/angle of
'your objects.
'
'You can do it all without quaternions - using only
'accumulated combinations of rotation matrices created
'by routines such as m3vRotate(). But it's much easier
'to normalize an accumulated rotation quaternion than it
'is to normalize an accumulated rotation matrix - see
'the m3normalize() routine in the modMatrix module.
Dim mag As Single

    mag = q4getMag(q)
    q(0) = q(0) / mag
    q(1) = q(1) / mag
    q(2) = q(2) / mag
    q(3) = q(3) / mag
End Sub

'Return a unit quaternion from unit axis vector and angle:
Public Sub q4fromAxis(q() As Single, vaxis() As Single, theta As Single)
'q0 = v0 * sin(t/2)
'q1 = v1 * sin(t/2)
'q2 = v2 * sin(t/2)
'q3 = cos(t/2)
    
    'q(0 to 2) defines the rotation axis.
    Call v3normalize(vaxis) 'just to make sure.
    Call v3copy(q, vaxis)
    Call v3scale(q, Sin(theta / 2!))
    'q(3) defines the real part; the amount of rotation (theta).
    q(3) = Cos(theta / 2!)
    'q4print q
End Sub

'Return a unit axis vector and angle from unit quaternion.
Public Sub q4toAxis(vaxis() As Single, theta As Single, q() As Single)
Dim S1 As Single
Dim S2 As Single
't = 2 * arc_cos(q3)
'v0 = q0 / sin(t/2)
'v1 = q1 / sin(t/2)
'v2 = q2 / sin(t/2)

    theta = 2! * ArcCos(q(3)) 'return theta
    S1 = Sin(theta / 2!)
    If (Abs(S1) < 0.0005!) Then S1 = 1!
    S2 = 1! / S1
    Call v3copy(vaxis, q)
    Call v3scale(vaxis, S2) 'return vaxis()
End Sub

'Convert a unit rotation quaternion (magnitude = 1) into an
'orthogonal unit rotation matrix:
'
' //////////////////////////////////////////////////////
' // !!! WARNING -> q() must be a unit quaternion !!! //
' //////////////////////////////////////////////////////
'27 x, 16 +
Public Sub q4toMatrix(M() As Single, q() As Single)
'The orthogonal matrix M(row,col) corresponding to a rotation by the
'unit quaternion q = a + bi + cj + dk, (with |q| = 1) is given by:
'
'   | a^2+b^2-c^2-d^2      2(bc-ad)         2(bd+ac)    |
'   |    2(bc+ad)      a^2-b^2+c^2-d^2      2(cd-ab)    |
'   |    2(bd-ac)          2(cd+ab)     a^2-b^2-c^2+d^2 |
'
'Where (in this module): a=q(3), b=q(0), c=q(1) and d=q(2)
'
'By unit quaternion definition: a^2+b^2+c^2+d^2 = 1
'If we add -2(c^2+d^2) to both sides of the equation above, we get:
'a^2+b^2-c^2-d^2 = 1-2(c^2+d^2), which is equivalent to M(0,0) above.
'
'We use the 1-2(c^2+d^2) form to gain a computational speed advantage
'when calculating the M(0,0), M(1,1) and M(2,2) matrix values.

    M(0, 0) = 1! - 2! * (q(1) * q(1) + q(2) * q(2))
    M(0, 1) = 2! * (q(0) * q(1) - q(2) * q(3))
    M(0, 2) = 2! * (q(2) * q(0) + q(1) * q(3))
    M(0, 3) = 0!

    M(1, 0) = 2! * (q(0) * q(1) + q(2) * q(3))
    M(1, 1) = 1! - 2! * (q(2) * q(2) + q(0) * q(0))
    M(1, 2) = 2! * (q(1) * q(2) - q(0) * q(3))
    M(1, 3) = 0!

    M(2, 0) = 2! * (q(2) * q(0) - q(1) * q(3))
    M(2, 1) = 2! * (q(1) * q(2) + q(0) * q(3))
    M(2, 2) = 1! - 2! * (q(1) * q(1) + q(0) * q(0))
    M(2, 3) = 0!

    M(3, 0) = 0!
    M(3, 1) = 0!
    M(3, 2) = 0!
    M(3, 3) = 1!
End Sub

'Same as the matrix above but optimized for speed.
'12 x, 12 +
Public Sub q4toMatrixF(M() As Single, q() As Single)
Dim X2 As Single, Y2 As Single, Z2 As Single
Dim AX As Single, AY As Single, AZ As Single
Dim XX As Single, YY As Single, ZZ As Single
Dim XY As Single, XZ As Single, YZ As Single
    
    X2 = q(0) * 2!: Y2 = q(1) * 2!: Z2 = q(2) * 2!
    XX = q(0) * X2: XY = q(0) * Y2: XZ = q(0) * Z2
    YY = q(1) * Y2: YZ = q(1) * Z2: ZZ = q(2) * Z2
    AX = q(3) * X2: AY = q(3) * Y2: AZ = q(3) * Z2
    
    M(0, 0) = 1! - (YY + ZZ)
    M(0, 1) = XY - AZ
    M(0, 2) = XZ + AY
    M(0, 3) = 0!

    M(1, 0) = XY + AZ
    M(1, 1) = 1! - (XX + ZZ)
    M(1, 2) = YZ - AX
    M(1, 3) = 0!

    M(2, 0) = XZ - AY
    M(2, 1) = YZ + AX
    M(2, 2) = 1! - (XX + YY)
    M(2, 3) = 0!

    M(3, 0) = 0!
    M(3, 1) = 0!
    M(3, 2) = 0!
    M(3, 3) = 1!
End Sub

'This matrix not only rotates the object but also scales
'the object about the object's axes by scale vector sv().
'Same as multiplying a rotation matrix and a scale matrix.
'You could do this same thing with a rotation matrix also:
'row_0 * vx
'row_1 * vy
'row_2 * vz
'(optimized for speed)
'21 x, 12 +
Public Sub q4toMatrixSO(M() As Single, q() As Single, sv() As Single)
Dim X2 As Single, Y2 As Single, Z2 As Single
Dim AX As Single, AY As Single, AZ As Single
Dim XX As Single, YY As Single, ZZ As Single
Dim XY As Single, XZ As Single, YZ As Single
Dim SX As Single, SY As Single, SZ As Single
    
    X2 = q(0) * 2!: Y2 = q(1) * 2!: Z2 = q(2) * 2!
    XX = q(0) * X2: XY = q(0) * Y2: XZ = q(0) * Z2
    YY = q(1) * Y2: YZ = q(1) * Z2: ZZ = q(2) * Z2
    AX = q(3) * X2: AY = q(3) * Y2: AZ = q(3) * Z2
    
    SX = sv(0) 'Object X-Axis scale factor
    SY = sv(1) 'Object Y-Axis scale factor
    SZ = sv(2) 'Object Z-Axis scale factor
    
    M(0, 0) = (1! - (YY + ZZ)) * SX
    M(0, 1) = (XY - AZ) * SX
    M(0, 2) = (XZ + AY) * SX
    M(0, 3) = 0!

    M(1, 0) = (XY + AZ) * SY
    M(1, 1) = (1! - (XX + ZZ)) * SY
    M(1, 2) = (YZ - AX) * SY
    M(1, 3) = 0!

    M(2, 0) = (XZ - AY) * SZ
    M(2, 1) = (YZ + AX) * SZ
    M(2, 2) = (1! - (XX + YY)) * SZ
    M(2, 3) = 0!

    M(3, 0) = 0!
    M(3, 1) = 0!
    M(3, 2) = 0!
    M(3, 3) = 1!
End Sub

'This matrix not only rotates the object but also scales
'the object about the screen's axes by scale vector sv().
'Same as multiplying a rotation matrix and a scale matrix.
'You could do this same thing with a rotation matrix also:
'col_0 * vx
'col_1 * vy
'col_2 * vz
'(optimized for speed)
'21 x, 12 +
Public Sub q4toMatrixSS(M() As Single, q() As Single, sv() As Single)
Dim X2 As Single, Y2 As Single, Z2 As Single
Dim AX As Single, AY As Single, AZ As Single
Dim XX As Single, YY As Single, ZZ As Single
Dim XY As Single, XZ As Single, YZ As Single
Dim SX As Single, SY As Single, SZ As Single
    
    X2 = q(0) * 2!: Y2 = q(1) * 2!: Z2 = q(2) * 2!
    XX = q(0) * X2: XY = q(0) * Y2: XZ = q(0) * Z2
    YY = q(1) * Y2: YZ = q(1) * Z2: ZZ = q(2) * Z2
    AX = q(3) * X2: AY = q(3) * Y2: AZ = q(3) * Z2
    
    SX = sv(0) 'Screen X-Axis scale factor
    SY = sv(1) 'Screen Y-Axis scale factor
    SZ = sv(2) 'Screen Z-Axis scale factor
    
    M(0, 0) = (1! - (YY + ZZ)) * SX
    M(0, 1) = (XY - AZ) * SY
    M(0, 2) = (XZ + AY) * SZ
    M(0, 3) = 0!

    M(1, 0) = (XY + AZ) * SX
    M(1, 1) = (1! - (XX + ZZ)) * SY
    M(1, 2) = (YZ - AX) * SZ
    M(1, 3) = 0!

    M(2, 0) = (XZ - AY) * SX
    M(2, 1) = (YZ + AX) * SY
    M(2, 2) = (1! - (XX + YY)) * SZ
    M(2, 3) = 0!

    M(3, 0) = 0!
    M(3, 1) = 0!
    M(3, 2) = 0!
    M(3, 3) = 1!
End Sub

'Given a normalized, unit, exact rotation matrix,
'compute a unit rotation quaternion:
Public Sub q4fromMatrix(q() As Single, M() As Single)
Dim T As Single 'T is theta?
Dim S As Single 'S is sin(theta) ?

    T = 1! + M(0, 0) + M(1, 1) + M(2, 2)
    If (T > 1E-08!) Then
        S = T ^ 0.5! * 2!
        q(0) = (M(2, 1) - M(1, 2)) / S
        q(1) = (M(0, 2) - M(2, 0)) / S
        q(2) = (M(1, 0) - M(0, 1)) / S
        q(3) = 0.25! * S
    'Else, to avoid distortion, which is major diagonal?
    ElseIf (M(0, 0) > M(1, 1) And M(0, 0) > M(2, 2)) Then
        S = (1! + M(0, 0) - M(1, 1) - M(2, 2)) ^ 0.5! * 2!
        q(0) = 0.25! * S
        q(1) = (M(1, 0) + M(0, 1)) / S
        q(2) = (M(0, 2) + M(2, 0)) / S
        q(3) = (M(2, 1) - M(1, 2)) / S
     ElseIf (M(1, 1) > M(2, 2)) Then
        S = (1! + M(1, 1) - M(0, 0) - M(2, 2)) ^ 0.5! * 2!
        q(0) = (M(1, 0) + M(0, 1)) / S
        q(1) = 0.25! * S
        q(2) = (M(2, 1) + M(1, 2)) / S
        q(3) = (M(0, 2) - M(2, 0)) / S
     Else
        S = (1! + M(2, 2) - M(0, 0) - M(1, 1)) ^ 0.5! * 2!
        q(0) = (M(0, 2) + M(2, 0)) / S
        q(1) = (M(2, 1) + M(1, 2)) / S
        q(2) = 0.25! * S
        q(3) = (M(1, 0) - M(0, 1)) / S
    End If
End Sub

'Print quaternion <b,c,d,a> to the Immediate Window.
Public Sub q4print(q() As Single)
    Debug.Print "b,c,d,a: <"; q(0); ", "; q(1); ", "; q(2); ", "; q(3); ">"
End Sub
