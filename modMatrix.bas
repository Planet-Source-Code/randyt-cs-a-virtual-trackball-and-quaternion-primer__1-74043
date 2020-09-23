Attribute VB_Name = "modMatrix"
'
'/////////////////////[ Matrix Module ]///////////////////////////
'// Coded for single-precision floating point accuracy...
'// Numbers with ! suffix are single-precision floating point.
'/////////////////////////////////////////////////////////////////
'// This module does not use user-defined data types.
'// This module is designed for quick conversion to 'C'.
'/////////////////////////////////////////////////////////////////
'//
'// By Randy Manning, Aug-2011.
'//
'///////////////[ About Rotation Matricies ]//////////////////////
'//
'//         For any exact 3x3 rotation matrix R()
'//
'// (1) R_Transpose * R = Identity matrix.
'// (2) The determinant of R = 1.
'// (3) R is normalized: The squares of the elements in any row
'//     or column sum to 1.
'//     ...Any of the rows or columns represent unit vectors.
'// (4) R is orthogonal: The dot product of any pair of rows or
'//     any pair of columns is 0.
'// (5) The rows of R represent the coordinates in the original
'//     space of unit vectors along the coordinate axes of the
'//     rotated space.
'//     ...In this example, the ROWS of R are the unit vectors
'//     which, if multiplied by scaling factors, will produce
'//     proportional scalings of the represented object along
'//     the OBJECT'S reference axes.
'// (6) The columns of R represent the coordinates in the
'//     rotated space of unit vectors along the axes of the
'//     original space.
'//     ...In this example, the COLUMNS of R are the unit
'//     vectors which, if multiplied by scaling factors, will
'//     produce proportional scalings of the represented object
'//     along the SCREEN'S reference axes.
'//
'/////////////////////[ Implementation ]//////////////////////////
'//
'// Procedure arguments are transfered by reference (default).
'//
'// For subroutine procedures that modify argument values, this
'// module positions all of the arguments in the same order as
'// standard computer programming assignment statements:
'// mp = m1 * m2 <-> m3multiply(mp, m1, m2).
'//
'// Commutative arguments are suffixed:     a, b
'// Noncommutative arguments are suffixed:  1, 2
'//
'/////////////////////////////////////////////////////////////////

Option Explicit

Public Const M4_PARALLEL As Integer = 0     'Parallel projection
Public Const M4_PERSPECTIVE As Integer = 1  'Perspective projection

'Set identity matrix:
'4x4 identity matrix.
Public Sub m4identity(mi() As Single)
Dim i As Integer
Dim j As Integer

    For i = 0 To 3
        For j = 0 To 3
            If i = j Then
                mi(i, j) = 1!
            Else
                mi(i, j) = 0!
            End If
        Next j
    Next i
End Sub

'Normalize the X, Y and Z coordinates of a 3-D homogeneous
'point.
'
'This is where we convert from a point's 3-D coordinates
'(X,Y,Z) in memory to the point's 2-D coordinates (X,Y)
'needed for drawing on the computer monitor screen.
'
'It's also where the perspective projection information gets
'converted into a visual perspective effect just before
'drawing the point's X,Y coordinates directly to the screen.

'Note, the projection transform matrix changed the s
'(scalefactor) component of all the homogeneous points.
'This is WHY we use homogeneous points (with 4 coordinates).
'
'Center of projection, Z-clipping, is not supported here...
'See p4normalizeXY() for Z-cliping support.
'
'WARNING! any point with a zero scalefactor will cause /0!
Public Sub p4normalizeXYZ(p() As Single)
Dim scaleFactor As Single

    scaleFactor = p(3)
    p(0) = p(0) / scaleFactor
    p(1) = p(1) / scaleFactor
    p(2) = p(2) / scaleFactor
    p(3) = 1!
End Sub

'Set z-perspective matrix:
'For projection along the Z axis onto the X-Y
'plane with focus at the origin and the
'center of projection at distance (0, 0, CoP).
Public Sub m4zPerspective(M() As Single, ByVal CoP As Single)
    m4identity M
    If CoP <> 0! Then M(2, 3) = -1! / CoP
End Sub

'Set scale matrix:
'scale by factors Sx, Sy, and Sz.
Public Sub m3scale(ms() As Single, ByVal SX As Single, ByVal SY As Single, ByVal SZ As Single)
    m4identity ms
    ms(0, 0) = SX
    ms(1, 1) = SY
    ms(2, 2) = SZ
End Sub

'Set translation matrix:
'translation by Tx, Ty, and Tz.
Public Sub m3translate(M() As Single, ByVal Tx As Single, ByVal Ty As Single, ByVal Tz As Single)
    m4identity M
    M(3, 0) = Tx
    M(3, 1) = Ty
    M(3, 2) = Tz
End Sub

'Set rotate x-axis matrix:
'theta measured in radians.
Public Sub m3xRotate(M() As Single, ByVal theta As Single)
    m4identity M
    M(1, 1) = Cos(theta)
    M(2, 2) = M(1, 1)
    M(1, 2) = Sin(theta)
    M(2, 1) = -M(1, 2)
End Sub

'Set rotate y-axis matrix:
'theta measured in radians.
Public Sub m3yRotate(M() As Single, ByVal theta As Single)
    m4identity M
    M(0, 0) = Cos(theta)
    M(2, 2) = M(0, 0)
    M(2, 0) = Sin(theta)
    M(0, 2) = -M(2, 0)
End Sub

'Set rotate z-axis matrix:
'theta measured in radians.
Public Sub m3zRotate(M() As Single, ByVal theta As Single)
    m4identity M
    M(0, 0) = Cos(theta)
    M(1, 1) = M(0, 0)
    M(0, 1) = Sin(theta)
    M(1, 0) = -M(0, 1)
End Sub

'Set rotate vector-axis matrix:
'theta measured in radians.
'Google: "Generalized Rotation Matrix"
'generates the same matrix as: q4fromAxis() and q4toMatrix(),
Public Sub m3vRotate(M() As Single, vaxis() As Single, theta As Single)
'Return the generalized 4x4 3-D rotation matrix: M(). Where theta
'represents the angle of object rotation about a line passing
'through the origin that is parallel to a unit-vaxis.
'Right-hand rule: Wrap the fingers of your right hand around the
'unit-vaxis arrow shaft and point your thumb in the direction of
'the unit-vaxis arrow head. Your right hand fingers will then be
'pointing in the direction of positive rotation angle measurement.

Dim unit_vaxis(0 To 2) As Single
Dim x As Single, Y As Single, Z As Single
Dim XX As Single, XY As Single, XZ As Single
Dim YY As Single, YZ As Single, ZZ As Single
Dim CT As Single    'CT = [1 - Cos(theta)]
Dim ST As Single    'ST = Sin(theta)
    
    'Do not alter vaxis():
    Call v3copy(unit_vaxis, vaxis)
    'Set vaxis copy to be a unit-vector:
    Call v3normalize(unit_vaxis)
        
    x = unit_vaxis(0)
    Y = unit_vaxis(1)
    Z = unit_vaxis(2)
    XX = x * x: XY = x * Y: XZ = x * Z
    YY = Y * Y: YZ = Y * Z: ZZ = Z * Z
    
    CT = 1! - Cos(theta)
    ST = Sin(theta)
    
    m4identity M
    M(0, 0) = 1! + CT * (XX - 1!): M(0, 1) = -Z * ST + CT * XY: M(0, 2) = Y * ST + CT * XZ
    M(1, 0) = Z * ST + CT * XY: M(1, 1) = 1! + CT * (YY - 1!): M(1, 2) = -x * ST + CT * YZ
    M(2, 0) = -Y * ST + CT * XZ: M(2, 1) = x * ST + CT * YZ: M(2, 2) = 1! + CT * (ZZ - 1!)
    
    'Transposed version of the above matrix.
    'm4Identity M
    'M(0, 0) = 1! + CT * (xx - 1!): M(0, 1) = Z * ST + CT * xy: M(0, 2) = -Y * ST + CT * xz
    'M(1, 0) = -Z * ST + CT * xy: M(1, 1) = 1! + CT * (yy - 1!): M(1, 2) = X * ST + CT * yz
    'M(2, 0) = Y * ST + CT * xz: M(2, 1) = -X * ST + CT * yz: M(2, 2) = 1! + CT * (zz - 1!)
End Sub

'Set matrix normalize:
'Restore an accumulated rotation matrix back to an exact
'rotation matrix. Such that:
'(1) R_transpose * R = identity matrix
'(2) The determinant of R = 1
'...To remove accumulated distortions.
Public Sub m3normalize(M() As Single)
Dim col0_in(0 To 2) As Single
Dim col1_in(0 To 2) As Single
Dim col2_in(0 To 2) As Single
Dim cross0(0 To 3) As Single
Dim cross1(0 To 3) As Single
'Dim cross2(0 To 3) As Single
Dim col0_out(0 To 2) As Single
Dim col1_out(0 To 2) As Single
Dim col2_out(0 To 2) As Single

    'You can do the same with rows if you want.
    'One column/row may be selected to be distorted the least.
    
    'Column0=Normalized(CrossProduct(Column1,Column2));
    'Column1=Normalized(CrossProduct(Column2,Column0));
    'don't really need to Column2=Normalized(CrossProduct(Column0,Column1));
    'Column2=Normalized(Column2);
    
    'M(row,col)
    'Set matrix column vectors:
    col0_in(0) = M(0, 0): col0_in(1) = M(1, 0): col0_in(2) = M(2, 0)
    col1_in(0) = M(0, 1): col1_in(1) = M(1, 1): col1_in(2) = M(2, 1)
    col2_in(0) = M(0, 2): col2_in(1) = M(1, 2): col2_in(2) = M(2, 2)
    
    'Set the cross products:
    Call v3cross(cross0, col1_in, col2_in)
    Call v3cross(cross1, col2_in, col0_in)
    'Call v3cross(cross2, col0_in, col1_in)
    
    'Normalize the cross products:
    Call v3normalize(cross0)
    Call v3normalize(cross1)
    'Call v3normalize(cross2)
    
    'Set normalized cross products to their
    'associated out vectors:
    Call v3copy(col0_out, cross0)
    Call v3copy(col1_out, cross1)
    
    'Normalize Column 2 out vector:
    'Column2=Normalized(Column2);
    Call v3normalize(col2_in)
    Call v3copy(col2_out, col2_in)
    
    'Transfer calculated out vectors back into
    'the rotation matrix:
    'M(row,col)
    'Set matrix column vectors:
    M(0, 0) = col0_out(0): M(1, 0) = col0_out(1): M(2, 0) = col0_out(2)
    M(0, 1) = col1_out(0): M(1, 1) = col1_out(1): M(2, 1) = col1_out(2)
    M(0, 2) = col2_out(0): M(1, 2) = col2_out(1): M(2, 2) = col2_out(2)
    
End Sub

'Set matrix copy:
'copy <- orig.
Public Sub m4copy(copy() As Single, orig() As Single)
Dim i As Integer
Dim j As Integer

    For i = 0 To 3
        For j = 0 To 3
            copy(i, j) = orig(i, j)
        Next j
    Next i
End Sub

'Apply a transformation matrix to a point.
'Use this transform when the matrix M() does
'not contain 0, 0, 0, 1 in its last column.
'i.e., when M() is a perspective matrix type
'M4_PERSPECTIVE.
Public Sub p4transform(p_scr() As Single, M() As Single, p_cur() As Single)
'p_scr() and p_cur() may be the same point.
Dim p_cur_copy(0 To 3) As Single
Dim i As Integer
Dim j As Integer
Dim value As Single

    'Use copy of p_cur():
    Call p4copy(p_cur_copy, p_cur)
    
    For i = 0 To 3 'column
        value = 0!
        For j = 0 To 3 'row
            value = value + p_cur_copy(j) * M(j, i)
        Next j
        p_scr(i) = value
    Next i
End Sub

'Set point copy:
'pdst() = psrc()
Public Sub p4copy(pdst() As Single, psrc() As Single)
    pdst(0) = psrc(0)
    pdst(1) = psrc(1)
    pdst(2) = psrc(2)
    pdst(3) = psrc(3)
End Sub

'Set matrix transpose:
'4x4 matrix T() <- M().
Public Sub m4transpose(T() As Single, M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 0 To 3
        For j = 0 To 3
            T(j, i) = M(i, j)
        Next j
    Next i
End Sub

'Return 3x3 determinant of matrix M()
Public Function m3getDet(M() As Single) As Single
'Formula derived via Mathematica.
    m3getDet = _
       -M(0, 2) * M(1, 1) * M(2, 0) + _
        M(0, 1) * M(1, 2) * M(2, 0) + _
        M(0, 2) * M(1, 0) * M(2, 1) - _
        M(0, 0) * M(1, 2) * M(2, 1) - _
        M(0, 1) * M(1, 0) * M(2, 2) + _
        M(0, 0) * M(1, 1) * M(2, 2)
End Function

'Normalize only the X and Y coordinates of a 3-D
'homogeneous point. Preserve the Z coordinate for Z-clipping.
'
'If you're not going to Z-clip, use p4normalizeXYZ().
'
'This is where we convert from a point's 3-D coordinates
'(X,Y,Z) in memory to the point's 2-D coordinates (X,Y) needed
'for drawing on the computer monitor screen.
'
'It's also where the perspective projection information gets
'converted into a visual perspective effect just before
'drawing the point's X,Y coordinates directly to the screen.
'
'Note, the projection transform matrix changed the s
'(scalefactor) component of all the homogeneous points.
'This is WHY we use homogeneous points (with 4 coordinates).
Public Sub p4normalizeXY(p() As Single)
Dim scaleFactor As Single

    'Normalize only the X,Y coordinates of the
    '3-D homogeneous point p<x,y,z,s>, where s
    'is the scale factor.
    scaleFactor = p(3)
    If scaleFactor <> 0! Then
        p(0) = p(0) / scaleFactor
        p(1) = p(1) / scaleFactor
        'Do not normalize the Z coordinate.
        'Preserve the Z coordinate for Z-clipping.
    Else
        'Force Z-clipping of any point with a zero
        'scaleFactor... We make the point's Z-coordinate
        'value greater than that of the center of
        'projection so that the point will be clipped.
        p(2) = INFINITY
    End If
    p(3) = 1!
End Sub

'Apply a transformation matrix to a point.
'Use this transform when the matrix M()
'contains 0, 0, 0, 1 in its last column.
'Note: faster than p4transform().
Public Sub p3transform(p_cur() As Single, M() As Single, p_bas() As Single)
'p_cur() and p_bas() may be the same point.
Dim p_bas_copy(0 To 3) As Single

    'Use copy of p_bas():
    Call p4copy(p_bas_copy, p_bas)
    
    p_cur(0) = p_bas_copy(0) * M(0, 0) + _
                p_bas_copy(1) * M(1, 0) + _
                p_bas_copy(2) * M(2, 0) + M(3, 0)
    p_cur(1) = p_bas_copy(0) * M(0, 1) + _
                p_bas_copy(1) * M(1, 1) + _
                p_bas_copy(2) * M(2, 1) + M(3, 1)
    p_cur(2) = p_bas_copy(0) * M(0, 2) + _
                p_bas_copy(1) * M(1, 2) + _
                p_bas_copy(2) * M(2, 2) + M(3, 2)
    p_cur(3) = 1!
End Sub

'Set matrix product:
'mp() = m1() * m2()
'NONCOMMUTATIVE! [m1 * m2] <> [m2 * m1]
'Use this multiplication when one or both
'of the matrices m1() and m2() do not contain
'0, 0, 0, 1 in their last columns.
'Complete 4x4 matrix multiply.
'64 x, 64 +
Public Sub m4multiply(mp() As Single, m1() As Single, m2() As Single)
'NOTE: m1 and/or m2 may be the same matrix as mp.
Dim ma(0 To 3, 0 To 3) As Single 'copy of m1
Dim mb(0 To 3, 0 To 3) As Single 'copy of m2
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim value As Single
    
    'Copy m1 and m2 (factor) arguments:
    Call m4copy(ma, m1)
    Call m4copy(mb, m2)
    
    For i = 0 To 3
        For j = 0 To 3
            value = 0!
            For k = 0 To 3
                value = value + ma(i, k) * mb(k, j)
            Next k
            mp(i, j) = value
        Next j
    Next i
End Sub

'Set matrix product:
'mp() = m1() * m2()
'NONCOMMUTATIVE! [m1 * m2] <> [m2 * m1]
'Use this multiplication when the matrices m1()
'and m2() both contain 0, 0, 0, 1 in their last
'columns.
'2x faster than m4Multiply().
'36 x, 27 +
Public Sub m3multiply(mp() As Single, m1() As Single, m2() As Single)
'NOTE: m1 and/or m2 may be the same matrix as mp.
Dim ma(0 To 3, 0 To 3) As Single 'copy of m1
Dim mb(0 To 3, 0 To 3) As Single 'copy of m2

    'Copy m1 and m2 (factor) arguments:
    Call m4copy(ma, m1)
    Call m4copy(mb, m2)
    
    mp(0, 0) = ma(0, 0) * mb(0, 0) + ma(0, 1) * mb(1, 0) + ma(0, 2) * mb(2, 0)
    mp(0, 1) = ma(0, 0) * mb(0, 1) + ma(0, 1) * mb(1, 1) + ma(0, 2) * mb(2, 1)
    mp(0, 2) = ma(0, 0) * mb(0, 2) + ma(0, 1) * mb(1, 2) + ma(0, 2) * mb(2, 2)
    mp(0, 3) = 0!
    mp(1, 0) = ma(1, 0) * mb(0, 0) + ma(1, 1) * mb(1, 0) + ma(1, 2) * mb(2, 0)
    mp(1, 1) = ma(1, 0) * mb(0, 1) + ma(1, 1) * mb(1, 1) + ma(1, 2) * mb(2, 1)
    mp(1, 2) = ma(1, 0) * mb(0, 2) + ma(1, 1) * mb(1, 2) + ma(1, 2) * mb(2, 2)
    mp(1, 3) = 0!
    mp(2, 0) = ma(2, 0) * mb(0, 0) + ma(2, 1) * mb(1, 0) + ma(2, 2) * mb(2, 0)
    mp(2, 1) = ma(2, 0) * mb(0, 1) + ma(2, 1) * mb(1, 1) + ma(2, 2) * mb(2, 1)
    mp(2, 2) = ma(2, 0) * mb(0, 2) + ma(2, 1) * mb(1, 2) + ma(2, 2) * mb(2, 2)
    mp(2, 3) = 0!
    mp(3, 0) = ma(3, 0) * mb(0, 0) + ma(3, 1) * mb(1, 0) + ma(3, 2) * mb(2, 0) + mb(3, 0)
    mp(3, 1) = ma(3, 0) * mb(0, 1) + ma(3, 1) * mb(1, 1) + ma(3, 2) * mb(2, 1) + mb(3, 1)
    mp(3, 2) = ma(3, 0) * mb(0, 2) + ma(3, 1) * mb(1, 2) + ma(3, 2) * mb(2, 2) + mb(3, 2)
    mp(3, 3) = 1!
End Sub

'Print matrix M(row,col) to the Immediate Window.
Public Sub m4print(M() As Single)
Dim i As Integer
Dim j As Integer
    For i = 0 To 3 'row
        For j = 0 To 3 'col
            Debug.Print M(i, j);
        Next j
        Debug.Print
    Next i
End Sub

'Print homogeneous point <x,y,z,s> to the Immediate Window.
Public Sub p4print(p() As Single)
    Debug.Print "x,y,z,s: <"; p(0); ", "; p(1); ", "; p(2); ", "; p(3); ">"
End Sub
