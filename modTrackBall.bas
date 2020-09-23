Attribute VB_Name = "modTrackball"
'
'//////////////////[ Virtual TrackBall Module ]///////////////////
'// Coded for single-precision floating point accuracy.
'// Numbers with ! suffix are single-precision floating point.
'/////////////////////////////////////////////////////////////////
'//
'// By Randy Manning, Jul-2011.
'//
'/////////////////////////////////////////////////////////////////
'//
'//        Virtual TrackBall Basics: Via Quaternions
'//
'// If we have two arbitrary points A and B (position vectors)
'// on the surface of a unit sphere, we can represent the
'// rotation along the shortest path (geodesic) from A to B
'// as a single rotation quaternion, where the angle (amount)
'// of the rotation is given by ArcCos(A dot B) and the rotation
'// vector (axis of rotation) is given by (A cross B).
'//
'// The normalized quaternion expressing the rotation from
'// A to C on a unit sphere is numerically equal to the
'// normalized product of the two quaternions expressing the
'// rotations from A to B and B to C for any arbitrary point B
'// on the unit sphere.
'//
'/////////////////////////////////////////////////////////////////

Option Explicit

'Public variables:
'Public TrackBall (mouse position change) Rotation Quaternion:
Public p_TrackBall_Rotation_Quaternion(0 To 3) As Single

'//////////////////[ !!! WARNING !!! ]////////////////////
'// Set the p_TrackBall_Radius_Pixels variable to some  //
'// reasonable value (say 200 pixels) before using this //
'// module!!!                                           //
'/////////////////////////////////////////////////////////
Public p_TrackBall_Radius_Pixels As Single

'Module-Level Variables:
Private m_PBoxLeftButtonDown As Boolean
Private m_PBoxRightButtonDown As Boolean

'Old mouse point (x,y) pixels:
Private m_MouAX_Pixels As Single
Private m_MouAY_Pixels As Single
'New mouse point (x,y) pixels:
Private m_MouBX_Pixels As Single
Private m_MouBY_Pixels As Single

'/////////////////[ Public Section - Input ]///////////////////

'//////////[ Mapping Functions: GDI <-> Cartesian ]////////////
'//                                                          //
'//                     Mouse-GDI [Input]:                   //
'//          Cartesian.X = GDI.X  - PBox_HalfWidth           //
'//          Cartesian.Y = PBox_HalfHeight - GDI.Y           //
'//                                                          //
'//                     Drawing [Output]:                    //
'//          GDI.X = PBox_HalfWidth  + Cartesian.X           //
'//          GDI.Y = PBox_HalfHeight - Cartesian.Y           //
'//                                                          //
'//////////////////////////////////////////////////////////////

'Echo from: Form.PBox_MouseDown()
'Assumes that X,Y are mapped to Cartesian coordinates.
'i.e., (0,0) = Center of PBox:
Public Sub TrackBall_PBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Button
        Case Is = vbLeftButton
            m_PBoxLeftButtonDown = True
            'Set mouse-down (Old-Point) here:
            m_MouAX_Pixels = x
            m_MouAY_Pixels = Y
        Case Is = vbRightButton
            m_PBoxRightButtonDown = True
    End Select
End Sub

'Echo from: Form.PBox_MouseMove()
'Assumes that X,Y are mapped to Cartesian coordinates.
'i.e., (0,0) = Center of PBox:
Public Sub TrackBall_PBox_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim p1x As Single
Dim p1y As Single
Dim p2x As Single
Dim p2y As Single

    If m_PBoxLeftButtonDown Then
        'Set mouse-move (New-Point) here:
        m_MouBX_Pixels = x
        m_MouBY_Pixels = Y
        
        'Scale (normalize) mouse-move-points (in pixel units)
        'relative to trackball-radius (in pixel units) for
        'projection onto a unit sphere:
        p1x = m_MouAX_Pixels / p_TrackBall_Radius_Pixels
        p1y = m_MouAY_Pixels / p_TrackBall_Radius_Pixels
        p2x = m_MouBX_Pixels / p_TrackBall_Radius_Pixels
        p2y = m_MouBY_Pixels / p_TrackBall_Radius_Pixels
        'Debug.Print Sqr(p1x ^ 2 + p1y ^ 2) '<- make it obvious
        
        'Simulate a track-ball.
        'Set Public quaternion variable: p_TrackBall_Rotation_Quaternion.
        Call TrackBall(p_TrackBall_Rotation_Quaternion, p1x, p1y, p2x, p2y)
                
        'Set old mouse-location to new mouse-location:
        m_MouAX_Pixels = m_MouBX_Pixels
        m_MouAY_Pixels = m_MouBY_Pixels
    End If
End Sub

'Echo from: Form.PBox_MouseUp()
Public Sub TrackBall_PBox_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Button
        Case Is = vbLeftButton
            m_PBoxLeftButtonDown = False
        Case Is = vbRightButton
            m_PBoxRightButtonDown = False
    End Select
End Sub

'////////////////////[ Private Section ]///////////////////////

'///////////////////[ Simulate a TrackBall ]///////////////////////
'// (1) Project the normalized x,y point (arguments) as position
'//     vectors onto the surface of a unit sphere.
'// (2) Determine the rotation axis vector and the angle of
'//     rotation.
'// (3) Calculate and return an equivalent rotation quaternion
'//     Rotation_Quaternion().
Private Sub TrackBall(Rotation_Quaternion() As Single, p1x As Single, p1y As Single, p2x As Single, p2y As Single)
Dim p1(0 To 2) As Single    'Position vector A (start point)
Dim p2(0 To 2) As Single    'Position vector B (end point)
Dim p1z As Single           'z component of p1
Dim p2z As Single           'z component of p2
Dim dot As Single           'Dot product of p1 and p2
Dim axisV(0 To 2) As Single 'Rotation axis vector
Dim theta As Single         'Positive rotation angle about axisV()

    If (p1x = p2x And p1y = p2y) Then
        'zero rotation.
        'return unit quaternion with no rotation.
        Call q4set(Rotation_Quaternion, 0!, 0!, 0!, 1!)
        Exit Sub
    End If

    'Get the z-components of p1 and p2 such that p1 and p2
    '(position vectors) project to the surface of a unit sphere:
    p1z = Project_pz_ToUnitSphere(p1x, p1y)
    p2z = Project_pz_ToUnitSphere(p2x, p2y)
    
    'Set p1 and p2 as position vectors on the
    'surface of a unit sphere:
    Call v3set(p1, p1x, p1y, p1z)
    Call v3set(p2, p2x, p2y, p2z)
    'Debug.Print v3getMag(p1), v3getMag(p2)
    
    'Get rotation axis vector: axisV().
    Call v3cross(axisV, p2, p1)
    'Debug.Print v3getMag(p2), v3getMag(p1)
    
    'Get rotation angle about axisV(): theta.
    v3normalize p1
    v3normalize p2
    dot = v3getDot(p1, p2)
    'Debug.Print dot
    'Avoid problems with out-of-range dot values:
    If (dot > 1!) Then dot = 1!
    If (dot < -1!) Then dot = -1!
    'Angle of rotation: theta.
    theta = ArcCos(dot)
    'Setting, theta = -1 * theta, will reverse trackball
    'responce. This is handy for when you've zoomed to
    'the inside of an object.
    
    'Calculate and return the rotation quaternion:
    'Rotation_Quaternion().
    Call q4fromAxis(Rotation_Quaternion, axisV, theta)
    'Or, if you accumulate rotation matrices, you could return a
    'rotation matrix here: Call m3vRotate(tbrMatrix, axisV, theta)
End Sub

'Project the z component of an x,y pair onto a unit sphere or
'onto a hyperbolic sheet if we are away from the center of the
'unit sphere.
'Note, this is a deformed unit sphere. Spherical from the center
'up to an x-y radius of 0.7071, but deformed into a hyperbolic
'sheet of rotation away from an x-y radius of 0.7071.
'Return the z component of the x,y pair:
Private Function Project_pz_ToUnitSphere(px As Single, py As Single) As Single
Dim xy_radius As Single
Dim pz As Single

    xy_radius = (px * px + py * py) ^ 0.5!
    
    If (xy_radius < 0.7071068!) Then
        'Project pz to unit sphere:
        pz = (1! - xy_radius * xy_radius) ^ 0.5!
    Else
        'Project pz to hyperbolic sheet:
        'This hyperbolic sheet function is exactly tangent to
        'the unit sphere surface at a 45 degree angle measured
        'from the sphere origin (0,0,0) relative to the xy
        'viewing plane that contains the sphere origin (0,0,0).
        pz = 0.5! / xy_radius
    End If
    
    'Return the z-component of the position vector
    '<px,py,pz> on the unit sphere surface:
    Project_pz_ToUnitSphere = pz
End Function
