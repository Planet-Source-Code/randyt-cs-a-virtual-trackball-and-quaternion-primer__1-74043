Attribute VB_Name = "modTextPoint"
'
'////////////////////[ Text Point Module ]//////////////////////////
'// Coded for single-precision floating point accuracy.
'// Numbers with ! suffix are single-precision floating point.
'///////////////////////////////////////////////////////////////////
'//
'// A module very similar to this one could be constructed for the
'// manipulation and display of individual 3-D points.
'//
'///////////////////////////////////////////////////////////////////

Option Explicit

'API declarations:

'For drawing Text:
'[GDI Cooridnates: in vbPixels Units]
Private Declare Function TextOut Lib "gdi32" _
    Alias "TextOutA" (ByVal hdc As Long, _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long) As Long
    
Public Type Text_Point
    'Text (a single character)
    txt As String * 1
    'Text-point color
    color As Long
    'Text-point base (reference-position) points.
    base_point(0 To 3) As Single
    'Text-point working (current-position) points.
    current_point(0 To 3) As Single
    'Text-point screen (screen-position) points.
    screen_point(0 To 3) As Single
End Type

'Public variables:

'Array to hold text-points:
Public p_TextPoints() As Text_Point
'Number of text-points contained in p_TextPoints():
Public p_NumTextPoints As Integer

'Create (add) a new text-point.
Public Sub AddText(txt As String, ByVal color As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single)
    p_NumTextPoints = p_NumTextPoints + 1
    ReDim Preserve p_TextPoints(1 To p_NumTextPoints)
    p_TextPoints(p_NumTextPoints).txt = txt
    p_TextPoints(p_NumTextPoints).color = color
    p_TextPoints(p_NumTextPoints).base_point(0) = X1
    p_TextPoints(p_NumTextPoints).base_point(1) = Y1
    p_TextPoints(p_NumTextPoints).base_point(2) = Z1
    p_TextPoints(p_NumTextPoints).base_point(3) = 1!
End Sub

'Apply the transformation matrix M() to all of the
'text-point points.
Public Sub TransformAllTextBasePointsToCurrentPoints(M() As Single)
    TransformSomeTextBasePointsToCurrentPoints M, 1, p_NumTextPoints
End Sub
'Apply the transformation matrix M() to the indicated
'text-point points.
Public Sub TransformSomeTextBasePointsToCurrentPoints(M() As Single, ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
    
    For i = ndx1 To ndx2
        Call p3transform(p_TextPoints(i).current_point, M, p_TextPoints(i).base_point)
    Next i
End Sub

'Overwrite the text-point base-points with the current-points:
'i.e., permanently move the points.
Public Sub CopyAllTextCurrentPointsToBasePoints()
    Call CopySomeTextCurrentPointsToBasePoints(1, p_NumTextPoints)
End Sub
'Overwrite the text-point base-points with the current-points:
'i.e., permanently move the points.
Public Sub CopySomeTextCurrentPointsToBasePoints(ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
Dim j As Integer

    For i = ndx1 To ndx2
        For j = 0 To 3
            p_TextPoints(i).base_point(j) = p_TextPoints(i).current_point(j)
        Next j
    Next i
End Sub

'Apply a 4x4 projection matrix to all the
'text-points and normalize the points for
'drawing to the screen.
Public Sub ProjectAllTextCurrentPointsToScreenPoints(ProjMatrix() As Single)
    ProjectSomeTextCurrentPointsToScreenPoints ProjMatrix, 1, p_NumTextPoints
End Sub
'Apply a 4x4 projection matrix to the indicated
'text-points and normalize the points for
'drawing to the screen.
Public Sub ProjectSomeTextCurrentPointsToScreenPoints(ProjMatrix() As Single, ByVal ndx1 As Integer, ByVal ndx2 As Integer)
Dim i As Integer
    
    For i = ndx1 To ndx2
        'Apply 4x4 transform from current to screen points:
        p4transform p_TextPoints(i).screen_point, ProjMatrix, p_TextPoints(i).current_point
        'p4print p_TextPoints(i).current_point
        'p4print p_TextPoints(i).screen_point
        
        'Normalize screen points (for drawing to screen):
        p4normalizeXY p_TextPoints(i).screen_point
        'p4print p_TextPoints(i).screen_point
    Next i
End Sub

'//////////[ Mapping Functions: GDI <-> Cartesian ]////////////
'//                                                          //
'//                    Mouse-GDI [Input]:                    //
'//          Cartesian.X = GDI.X  - PBox_HalfWidth           //
'//          Cartesian.Y = PBox_HalfHeight - GDI.Y           //
'//                                                          //
'//                     Drawing [Output]:                    //
'//          GDI.X = PBox_HalfWidth  + Cartesian.X           //
'//          GDI.Y = PBox_HalfHeight - Cartesian.Y           //
'//                                                          //
'//////////////////////////////////////////////////////////////
'Draw all transformed text-points
Public Sub DrawAllText_Clip(pic As PictureBox, pic_HalfWidth As Single, pic_HalfHeight As Single, eye_Z As Single)
    Call DrawSomeText_Clip(pic, pic_HalfWidth, pic_HalfHeight, eye_Z, 1, p_NumTextPoints)
End Sub
'Draw the indicated transformed text-points
Public Sub DrawSomeText_Clip(pic As PictureBox, pic_HalfWidth As Single, pic_HalfHeight As Single, eye_Z As Single, ndx1 As Integer, ndx2 As Integer)
Dim oldFC As Long
Dim i As Integer
Dim X1 As Single
Dim Y1 As Single
Dim Z1 As Single
Dim draw_txt As Boolean

    oldFC = pic.ForeColor '<- save ForeColor
    'Draw the text-points, using On Error to avoid
    'overflows when drawing text far out of bounds.
    On Error Resume Next
    draw_txt = True
    For i = ndx1 To ndx2
        'Point clipping:
        Z1 = p_TextPoints(i).screen_point(2)
        'Don't draw if the point is farther
        'from the focus point than the center of
        'projection (which is distance eye_Z away).
        'Debug.Print eye_Z
        draw_txt = Z1 < eye_Z
        If draw_txt Then
            'Map X,Y from Cartesian to GDI coordinates:
            X1 = pic_HalfWidth + p_TextPoints(i).screen_point(0)
            Y1 = pic_HalfHeight - p_TextPoints(i).screen_point(1)
            'Draw the text:
            pic.ForeColor = p_TextPoints(i).color
            Call TextOut(pic.hdc, X1, Y1, p_TextPoints(i).txt, 1)
        End If
    Next i
    pic.ForeColor = oldFC '<- restore ForeColor
End Sub

