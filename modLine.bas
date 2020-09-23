Attribute VB_Name = "modLine"
'
'///////////////////[ Line Segment Module ]////////////////////
'// Coded for single-precision floating point accuracy.
'// Numbers with ! suffix are single-precision floating point.
'//////////////////////////////////////////////////////////////

Option Explicit

Public Type Line_Segment
    'Line-segment base (reference-position) points.
    from_base_point(0 To 3) As Single
    to_base_point(0 To 3) As Single
    'Line-segment working (current-position) points.
    from_current_point(0 To 3) As Single
    to_current_point(0 To 3) As Single
    'Line-segment screen (screen-position) points.
    from_screen_point(0 To 3) As Single
    to_screen_point(0 To 3) As Single
    'Line-segment color
    color As Long
End Type

'Public variables:

'Array to hold line-segment points:
Public p_Lines() As Line_Segment
'Number of line-segments contained in p_Lines().
Public p_NumLines As Integer

'Create (add) a new line-segment to the p_Lines() array.
Public Sub AddLine( _
    ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single, _
    ByVal X2 As Single, ByVal Y2 As Single, ByVal Z2 As Single, _
    ByVal color As Long)
    
    p_NumLines = p_NumLines + 1
    ReDim Preserve p_Lines(1 To p_NumLines)
    p_Lines(p_NumLines).color = color
    p_Lines(p_NumLines).from_base_point(0) = X1
    p_Lines(p_NumLines).from_base_point(1) = Y1
    p_Lines(p_NumLines).from_base_point(2) = Z1
    p_Lines(p_NumLines).from_base_point(3) = 1!
    p_Lines(p_NumLines).to_base_point(0) = X2
    p_Lines(p_NumLines).to_base_point(1) = Y2
    p_Lines(p_NumLines).to_base_point(2) = Z2
    p_Lines(p_NumLines).to_base_point(3) = 1!
End Sub

' Check that all of the lines in this object
' have the same length. Return true if the
' lines all have the same length.
Public Function SameSideLengths(ByVal ndx1 As Integer, ByVal ndx2 As Integer) As Boolean
Dim a As Single
Dim b As Single
Dim c As Single
Dim S As Single
Dim i As Integer

    ' (S) <- Get first lines's length.
    a = p_Lines(ndx1).from_base_point(0) - p_Lines(ndx1).to_base_point(0)
    b = p_Lines(ndx1).from_base_point(1) - p_Lines(ndx1).to_base_point(1)
    c = p_Lines(ndx1).from_base_point(2) - p_Lines(ndx1).to_base_point(2)
    S = Sqr(a * a + b * b + c * c)
    
    ' Compare all other line lengths to first line length.
    SameSideLengths = False
    For i = ndx1 + 1 To ndx2
        a = p_Lines(i).from_base_point(0) - p_Lines(i).to_base_point(0)
        b = p_Lines(i).from_base_point(1) - p_Lines(i).to_base_point(1)
        c = p_Lines(i).from_base_point(2) - p_Lines(i).to_base_point(2)
        If Abs(S - Sqr(a * a + b * b + c * c)) > 0.001 Then Exit Function
    Next i
    
    SameSideLengths = True
End Function

'Apply the transformation matrix M() to all of the line
'segment end points.
Public Sub TransformAllLineBasePointsToCurrentPoints(M() As Single)
    TransformSomeLineBasePointsToCurrentPoints M, 1, p_NumLines
End Sub

'Apply the transformation matrix M() to the indicated
'line segment end points.
Public Sub TransformSomeLineBasePointsToCurrentPoints(M() As Single, ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
    
    For i = ndx1 To ndx2
        p3transform p_Lines(i).from_current_point, M, p_Lines(i).from_base_point
        p3transform p_Lines(i).to_current_point, M, p_Lines(i).to_base_point
    Next i
End Sub

'Overwrite the line-segment base-points with the current-points:
'i.e., permanently move the points.
Public Sub CopyAllLineCurrentPointsToBasePoints()
    Call CopySomeLineCurrentPointsToBasePoints(1, p_NumLines)
End Sub
'Overwrite the line-segment base-points with the current-points:
'i.e., permanently move the points.
Public Sub CopySomeLineCurrentPointsToBasePoints(ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
Dim j As Integer

    For i = ndx1 To ndx2
        For j = 0 To 3
            p_Lines(i).from_base_point(j) = p_Lines(i).from_current_point(j)
            p_Lines(i).to_base_point(j) = p_Lines(i).to_current_point(j)
        Next j
    Next i
End Sub

'Apply a 4x4 projection matrix to all the
'line segments and normalize the points for
'drawing to the screen.
Public Sub ProjectAllLineCurrentPointsToScreenPoints(ProjMatrix() As Single)
    ProjectSomeLineCurrentPointsToScreenPoints ProjMatrix, 1, p_NumLines
End Sub
'Apply a 4x4 projection matrix to the indicated
'line segments and normalize the points for
'drawing to the screen.
Public Sub ProjectSomeLineCurrentPointsToScreenPoints(ProjMatrix() As Single, ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
    
    For i = ndx1 To ndx2
        'Apply 4x4 transform from current to screen points:
        p4transform p_Lines(i).from_screen_point, ProjMatrix, p_Lines(i).from_current_point
        p4transform p_Lines(i).to_screen_point, ProjMatrix, p_Lines(i).to_current_point
        'Normalize screen points (for drawing to screen):
        p4normalizeXY p_Lines(i).from_screen_point
        p4normalizeXY p_Lines(i).to_screen_point
    Next i
End Sub

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
'Draw all transformed line segments
Public Sub DrawAllLines_Clip(pic As PictureBox, pic_HalfWidth As Single, pic_HalfHeight As Single, eye_Z As Single)
    Call DrawSomeLines_Clip(pic, pic_HalfWidth, pic_HalfHeight, eye_Z, 1, p_NumLines)
End Sub
'Draw the indicated transformed line segments
Public Sub DrawSomeLines_Clip(pic As PictureBox, pic_HalfWidth As Single, pic_HalfHeight As Single, eye_Z As Single, ndx1 As Integer, ndx2 As Integer)
Dim oldFC As Long
Dim i As Integer
Dim X1 As Single
Dim Y1 As Single
Dim Z1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim Z2 As Single
Dim draw_line As Boolean

    oldFC = pic.ForeColor '<- save ForeColor
    'Draw the lines, using On Error to avoid
    'overflows when drawing lines far out of bounds.
    On Error Resume Next
    draw_line = True
    For i = ndx1 To ndx2
        'Point clipping:
        Z1 = p_Lines(i).from_screen_point(2)
        Z2 = p_Lines(i).to_screen_point(2)
        'Don't draw if either point is farther
        'from the focus point than the center of
        'projection (which is distance eye_Z away).
        'Debug.Print eye_Z
        draw_line = (Z1 < eye_Z And Z2 < eye_Z)
        If draw_line Then
            'Map X,Y from Cartesian to GDI coordinates:
            X1 = pic_HalfWidth + p_Lines(i).from_screen_point(0)
            Y1 = pic_HalfHeight - p_Lines(i).from_screen_point(1)
            X2 = pic_HalfWidth + p_Lines(i).to_screen_point(0)
            Y2 = pic_HalfHeight - p_Lines(i).to_screen_point(1)
            'Draw the line segment:
            pic.ForeColor = p_Lines(i).color
            pic.Line (X1, Y1)-(X2, Y2)
        End If
    Next i
    pic.ForeColor = oldFC '<- restore ForeColor
End Sub

