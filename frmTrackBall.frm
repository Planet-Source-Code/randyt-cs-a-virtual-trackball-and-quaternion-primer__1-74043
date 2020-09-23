VERSION 5.00
Begin VB.Form frmTrackBall 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "A Virtual Trackball and Quaternion Primer for 3-D Graphics"
   ClientHeight    =   5910
   ClientLeft      =   1395
   ClientTop       =   1140
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmTrackBall.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   Begin VB.PictureBox picControlBox 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   5640
      ScaleHeight     =   367
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      Begin VB.TextBox txtMilliSeconds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   31
         Text            =   "milliseconds"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdHome 
         Caption         =   "Rotate Home"
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   4800
         Width           =   1455
      End
      Begin VB.PictureBox picShapeSelection 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   2280
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   13
         Top             =   120
         Width           =   2295
         Begin VB.CheckBox Choice 
            Caption         =   "8 Cubes"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   20
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Tetrahedron"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Octahedron"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Cube"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Icosahedron"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Dodecahedron"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Choice 
            Caption         =   "Axes"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
      End
      Begin VB.VScrollBar VScroll_SF 
         Height          =   3135
         LargeChange     =   5
         Left            =   1740
         Max             =   100
         Min             =   -100
         TabIndex        =   4
         Top             =   2160
         Width           =   255
      End
      Begin VB.VScrollBar VScroll_TbR 
         Height          =   3135
         LargeChange     =   24
         Left            =   1020
         Max             =   20
         Min             =   500
         TabIndex        =   3
         Top             =   2160
         Value           =   161
         Width           =   255
      End
      Begin VB.VScrollBar VScroll_CoP 
         Height          =   3135
         LargeChange     =   5
         Left            =   300
         Max             =   100
         Min             =   -100
         TabIndex        =   2
         Top             =   2160
         Value           =   -15
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Hardware Timer:"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "CoP"
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblScaleFactor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Scale Factor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTrackballDiameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trackball Diameter"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblCenterOfProjection 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CenterOfProjection"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "TbD     Zoom"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Scale Factor:"
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Trackball Diameter:"
         Height          =   255
         Left            =   300
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Center of Projection:"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   0
      ScaleHeight     =   358
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   366
      TabIndex        =   0
      Top             =   360
      Width           =   5490
   End
   Begin VB.PictureBox picKeyFunctions 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   691
      TabIndex        =   21
      Top             =   0
      Width           =   10365
      Begin VB.PictureBox picProjectionControls 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6120
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   26
         Top             =   0
         Width           =   3780
         Begin VB.OptionButton optParallel 
            Caption         =   "Parallel"
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   45
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optPerspective 
            Caption         =   "Perspective"
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   45
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Projection:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   45
            Width           =   975
         End
      End
      Begin VB.PictureBox picAxesSet 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   375
         TabIndex        =   22
         Top             =   0
         Width           =   5655
         Begin VB.OptionButton optScreen 
            Caption         =   "Screen   Axes"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   45
            Width           =   1575
         End
         Begin VB.OptionButton optObject 
            Caption         =   "Object"
            Height          =   255
            Left            =   3000
            TabIndex        =   23
            Top             =   45
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "X,Y,Z  KeyDowns Rotate About:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   60
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmTrackBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'/////////////////////////////////////////////////////////////////
'//                                                             //
'// A Virtual Trackball and Quaternion Primer for 3-D Graphics  //
'//                                                             //
'/////////////////////////////////////////////////////////////////
'//                                                             //
'// By Randy Manning, Aug-2011.   mrandycs@swbell.net           //
'// Language: Microsoft Visual Basic 6.0 (SP4)                  //
'//                                                             //
'/////////////////////////////////////////////////////////////////
'//                                                             //
'// This example program contains concise solutions to the      //
'// three most difficult aspects, for me, of 3-D graphics       //
'// programming:                                                //
'// (1) Perspective Projection.                                 //
'// (2) Object Rotation Combinations.                           //
'// (3) Virtual Trackball Interface.                            //
'//                                                             //
'// And most importantly: "How to put it all together."         //
'//                                                             //
'/////////////////////////////////////////////////////////////////
'//                                                             //
'// This program intentionally includes only 3-D line-segment   //
'// drawing capability - To keep the program structure as       //
'// simple as possible.                                         //
'//                                                             //
'/////////////////////////////////////////////////////////////////
'
'/////////[ About Free-Space Rotations and Quaternions ]//////////
'//
'//  This program uses Quaternions for 3-D Rotation Combinations.
'//
'// A quaternion is an extension of the complex numbers having
'// a real part and three imaginary parts: q = a + bi + cj + dk,
'// where i^2 = j^2 = k^2 = -1.
'// An object rotation in free-space can be represented as a single
'// quaternion, where:
'// 1) q(a) the real part, represents an angle (amount) of rotation.
'// 2) q(b,c,d) the three imaginary parts, represent a vector (in
'//    ijk-space) that defines an axis of rotation.
'//
'//                    The Unit Sphere
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
'//            Basic 3-D Graphic Quaternion Principals:
'//
'// (1) A quaternion holds a rotation.
'// (2) Quaternion multiplication combines two rotations.
'// (3) The order of the quaternion multiplication factors determine
'//     the set of reference-axes about which the object rotation
'//     will occur, either:
'//     (a) The SCREEN reference-axes:  current = current * new
'//         or
'//     (b) The OBJECT reference-axes:  current = new * current
'//
'//     ...See the picCanvas_KeyDown() routine for an example.
'//
'//
'//                      Rotation Procedure
'//
'// This system is very simple. Begin by defining a current rotation
'// quaternion and set it to home position: q(0,0,0,1). Call
'// q4toMatrix() to convert the current rotation quaternion into a
'// current rotation matrix. Use the current rotation matrix to draw
'// your initial graphics.
'//
'// From then on...
'// To rotate your object from its current orientation; Create a new
'// quaternion by calling q4fromAxis() with the desired rotation
'// vector and angle. Then multiply your current rotation quaternion
'// with the new quaternion and stuff the resulting product back into
'// your current rotation quaternion variable. Then call q4toMatrix()
'// to convert the current rotation quaternion into a current rotation
'// matrix. Use the current rotation matrix to draw your graphics.
'//
'/////////////////////////////////////////////////////////////////
'//
'// For more information:
'// Google: "Parameterizing the space of rotations"
'//
'/////////////////////////////////////////////////////////////////
'
'////////////////[ Basic 3-D Graphics Programming ]/////////////////////////
'//                                                                       //
'// 1) Create a current quaternion.                                       //
'// 2) Calculate a new rotation quaternion.                               //
'// 3) Multiply current quaternion by new rotation quaternion.            //
'// 4) Set current quaternion equal to the multiplication product.        //
'// '  Call RenderScene(picCanvas), which performs the following:         //
'// 5) Convert current quaternion into a rotation matrix.                 //
'// 6) Combine rotation matrix with scale matrix -> transform matrix      //
'// 7) Transform base-points to current-points with transform matrix.     //
'// 8) Transform current-points to screen-points with projection matrix.  //
'// 9) Draw screen-points directly to the screen.                         //
'// 10) Go back to step 2.                                                //
'//                                                                       //
'///////////////////////////////////////////////////////////////////////////

Option Explicit

'///////////////////////////[ My Rules ]////////////////////////////////////
'//
'// Rule 1 - I always use pixel-units for everything, because
'//          that's what the Windows API functions use. Keeps me
'//          from having to compare apples to oranges in my code
'//          and making unit conversion errors.
'// Rule 2 - I always convert from/to GDI and Cartesian coordinate
'//          systems at mouse or GDI function (input) and drawing
'//          (output) interfaces. This way I can always do all my
'//          stuff in Cartesian coordinates and Windows can always
'//          do all it's stuff in GDI coordinates.
'//
''/////////////////////////////////////////////////////////////
'//                                                         '//
'// The following mapping-functions assume that your        '//
'// 3-D drawing canvas (PictureBox) is set to:              '//
'// ScaleMode = 3 - Pixel.                                  '//
'//                                                         '//
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
Private m_PBox_HalfWidth_Pixels As Single                   '//
Private m_PBox_HalfHeight_Pixels As Single                  '//
'//////////////////////////////////////////////////////////////

'Center of projection:
'Location of viewing eye (along Z-axis).
Private m_EyeZ_Pixels As Single '<- in pixel units.

'Current Rotation Quaternion and Rotation Matrix.
Private m_Current_Rotation_Quaternion(0 To 3) As Single
Private m_Current_Rotation_Matrix(0 To 3, 0 To 3) As Single

'Rotation-Home Quaternion.
Private m_Rotation_Home_Quaternion(0 To 3) As Single

'Current Translation Matrix.
Private m_Current_Translate_matrix(0 To 3, 0 To 3) As Single

'Current Scale matrix:
Private m_Current_Scale_Matrix(0 To 3, 0 To 3) As Single

'Current Transform Matrix.
Private m_Current_Transform_Matrix(0 To 3, 0 To 3) As Single

'Current projection matrix type:
'M4_PARALLEL or M4_PERSPECTIVE
Private m_Current_Projection_Matrix_Type As Integer

'Current Projection Matrix.
Private m_Current_Projection_Matrix(0 To 3, 0 To 3) As Single

'Draw trackball boolean.
Private m_Draw_Trackball As Boolean

'picCanvas Right-Button mouse-down and mouse-move
'value for trackball diameter adjustment.
Private m_picCanvasRB_LastY As Single

'Hardware counter variables.
'See modHardwareCounter and RenderScene()
Private hardwareCount As Currency
Private elapsedTime_mS As Double

'Platonic shape index arrays, into p_Lines()
'array, located in module modLine:
Private AxisLines(0 To 1) As Integer
Private Tetrahedron(0 To 1) As Integer
Private OneCube(0 To 1) As Integer
Private Octahedron(0 To 1) As Integer
Private Dodecahedron(0 To 1) As Integer
Private Icosahedron(0 To 1) As Integer
Private EightCubes(0 To 1) As Integer

Private Sub Form_Load()
   
    'Set the home quaternion:
    'qHome = <0,0,0,1>
    Call q4set(m_Rotation_Home_Quaternion, 0!, 0!, 0!, 1!)
    
    'Create 3-D graphic data:
    CreateGraphicData
    
    'Form start-up position:
    'Set to upper-left of view-screen.
    Me.Move (Screen.Width - Me.Width) _
        / 3.5, (Screen.Height - Me.Height) / 3
End Sub

'Do any initialization stuff that can only be done
'after all controls on the form have been created:
Private Sub Form_Activate()
Static oneShot As Boolean

    If Not oneShot Then
    'Perform the following code statements only
    'once; On the first Form_Activate() event.
        
        'Set the drawing canvas halfwidth and halfheight
        'variables. For GDI<->Cartesian conversions:
        m_PBox_HalfWidth_Pixels = picCanvas.ScaleWidth / 2!
        m_PBox_HalfHeight_Pixels = picCanvas.ScaleHeight / 2!
        
        'Set the m_EyeZ_Pixels variable and
        'Update the Center of Projection label:
        Call VScroll_CoP_Handler(VScroll_CoP.value, False)
        
        'Set the p_TrackBall_Radius_Pixels variable and
        'update the Trackball Diameter label:
        Call VScroll_TbR_Handler(VScroll_TbR.value, False)
        
        'Set the current scale matrix
        'and update the Scale Factor label:
        'First call is dummy start up, won't start with 0.
        Call VScroll_SF_Handler(1, False)
        Call VScroll_SF_Handler(VScroll_SF.value, False)
        
        'Set current projection type:
        'm_Current_Projection_Matrix_Type = M4_PERSPECTIVE
        m_Current_Projection_Matrix_Type = M4_PARALLEL
        
        'Set the ProjectionMatrix:
        Call SetProjectionMatrix
        
        'Start up with no rotation:
        'Set the current rotation quaternion
        'equal to the home quaternion: <0,0,0,1>
        'Note: 2 * arc_cos(1) = 0 radians rotation.
        Call q4copy( _
            m_Current_Rotation_Quaternion, _
            m_Rotation_Home_Quaternion)
        'q4print m_Current_Rotation_Quaternion
                
        'Draw the graphics:
        Call RenderScene(picCanvas)
        
        picCanvas.SetFocus
        
        'Block any further entry attempts.
        oneShot = True
    End If
End Sub

'Create the graphic data.
Private Sub CreateGraphicData()
Dim Sc(0 To 3, 0 To 3) As Single     'Scale Matrix
Dim Tr(0 To 3, 0 To 3) As Single     'Translate Matrix
Dim ScTr(0 To 3, 0 To 3) As Single   'Scale and Translate combo Matrix
Dim theta1 As Single
Dim theta2 As Single
Dim S1 As Single
Dim S2 As Single
Dim c1 As Single
Dim c2 As Single
Dim S As Single
Dim R As Single
Dim H As Single
Dim a As Single
Dim b As Single
Dim c As Single
Dim d As Single
Dim x As Single
Dim Y As Single
Dim Y2 As Single
Dim M As Single
Dim N As Single

    'Zero the number of lines contained in the
    'p_Lines() array - located in modLine.
    p_NumLines = 0
    
    'Adds the line segments to the p_Lines() array
    'defined in module modLine.
    AxisLines(0) = p_NumLines + 1
    AddLine 0, 0, 0, 0.5, 0, 0, vbWhite ' X axis.
    AddLine 0, 0, 0, 0, 0.5, 0, vbWhite ' Y axis.
    AddLine 0, 0, 0, 0, 0, 0.5, vbWhite ' Z axis.
    AxisLines(1) = p_NumLines
           
    'Tetrahedron.
    Tetrahedron(0) = p_NumLines + 1
    S = Sqr(6)
    a = S / Sqr(3)
    b = -a / 2
    c = a * Sqr(2) - 1
    d = S / 2
    AddLine 0, c, 0, a, -1, 0, vbRed
    AddLine 0, c, 0, b, -1, d, vbRed
    AddLine 0, c, 0, b, -1, -d, vbRed
    AddLine b, -1, -d, b, -1, d, vbRed
    AddLine b, -1, d, a, -1, 0, vbRed
    AddLine a, -1, 0, b, -1, -d, vbRed
    Tetrahedron(1) = p_NumLines
    
    'Cube.
    OneCube(0) = p_NumLines + 1
    AddLine -1, -1, -1, -1, 1, -1, RGB(0, 128, 0)
    AddLine -1, 1, -1, 1, 1, -1, RGB(0, 128, 0)
    AddLine 1, 1, -1, 1, -1, -1, RGB(0, 128, 0)
    AddLine 1, -1, -1, -1, -1, -1, RGB(0, 128, 0)
    
    AddLine -1, -1, 1, -1, 1, 1, RGB(0, 128, 0)
    AddLine -1, 1, 1, 1, 1, 1, RGB(0, 128, 0)
    AddLine 1, 1, 1, 1, -1, 1, RGB(0, 128, 0)
    AddLine 1, -1, 1, -1, -1, 1, RGB(0, 128, 0)
    
    AddLine -1, -1, -1, -1, -1, 1, RGB(0, 128, 0)
    AddLine -1, 1, -1, -1, 1, 1, RGB(0, 128, 0)
    AddLine 1, 1, -1, 1, 1, 1, RGB(0, 128, 0)
    AddLine 1, -1, -1, 1, -1, 1, RGB(0, 128, 0)
    OneCube(1) = p_NumLines
    
    'Octahedron.
    Octahedron(0) = p_NumLines + 1
    AddLine 0, 1, 0, 1, 0, 0, vbBlue
    AddLine 0, 1, 0, -1, 0, 0, vbBlue
    AddLine 0, 1, 0, 0, 0, 1, vbBlue
    AddLine 0, 1, 0, 0, 0, -1, vbBlue
    
    AddLine 0, -1, 0, 1, 0, 0, vbBlue
    AddLine 0, -1, 0, -1, 0, 0, vbBlue
    AddLine 0, -1, 0, 0, 0, 1, vbBlue
    AddLine 0, -1, 0, 0, 0, -1, vbBlue
    
    AddLine 0, 0, 1, 1, 0, 0, vbBlue
    AddLine 0, 0, 1, -1, 0, 0, vbBlue
    AddLine 0, 0, -1, 1, 0, 0, vbBlue
    AddLine 0, 0, -1, -1, 0, 0, vbBlue
    Octahedron(1) = p_NumLines
    
    'Dodecahedron.
    Dodecahedron(0) = p_NumLines + 1
    theta1 = PI * 0.4
    theta2 = PI * 0.8
    S1 = Sin(theta1)
    c1 = Cos(theta1)
    S2 = Sin(theta2)
    c2 = Cos(theta2)
    
    M = 1 - (2 - 2 * c1 - 4 * S1 * S1) / (2 * c1 - 2)
    N = Sqr((2 - 2 * c1) - M * M) * (1 + (1 - c2) / (c1 - c2))
    R = 2 / N
    S = R * Sqr(2 - 2 * c1)
    a = R * S1
    b = R * S2
    c = R * c1
    d = R * c2
    H = R * (c1 - S1)
    
    x = (R * R * (2 - 2 * c1) - 4 * a * a) / (2 * c - 2 * R)
    Y = Sqr(S * S - (R - x) * (R - x))
    Y2 = Y * (1 - c2) / (c1 - c2)
    
    AddLine R, 1, 0, c, 1, a, RGB(0, 192, 192) ' Top
    AddLine c, 1, a, d, 1, b, RGB(0, 192, 192)
    AddLine d, 1, b, d, 1, -b, RGB(0, 192, 192)
    AddLine d, 1, -b, c, 1, -a, RGB(0, 192, 192)
    AddLine c, 1, -a, R, 1, 0, RGB(0, 192, 192)
    
    AddLine R, 1, 0, x, 1 - Y, 0, RGB(0, 192, 192) ' Top downward edges.
    AddLine c, 1, a, x * c1, 1 - Y, x * S1, RGB(0, 192, 192)
    AddLine c, 1, -a, x * c1, 1 - Y, -x * S1, RGB(0, 192, 192)
    AddLine d, 1, b, x * c2, 1 - Y, x * S2, RGB(0, 192, 192)
    AddLine d, 1, -b, x * c2, 1 - Y, -x * S2, RGB(0, 192, 192)
    
    AddLine x, 1 - Y, 0, -x * c2, 1 - Y2, -x * S2, RGB(0, 192, 192) ' Middle.
    AddLine x, 1 - Y, 0, -x * c2, 1 - Y2, x * S2, RGB(0, 192, 192)
    AddLine x * c1, 1 - Y, x * S1, -x * c2, 1 - Y2, x * S2, RGB(0, 192, 192)
    AddLine x * c1, 1 - Y, x * S1, -x * c1, 1 - Y2, x * S1, RGB(0, 192, 192)
    AddLine x * c2, 1 - Y, x * S2, -x * c1, 1 - Y2, x * S1, RGB(0, 192, 192)
    AddLine x * c2, 1 - Y, x * S2, -x, 1 - Y2, 0, RGB(0, 192, 192)
    AddLine x * c2, 1 - Y, -x * S2, -x, 1 - Y2, 0, RGB(0, 192, 192)
    AddLine x * c2, 1 - Y, -x * S2, -x * c1, 1 - Y2, -x * S1, RGB(0, 192, 192)
    AddLine x * c1, 1 - Y, -x * S1, -x * c1, 1 - Y2, -x * S1, RGB(0, 192, 192)
    AddLine x * c1, 1 - Y, -x * S1, -x * c2, 1 - Y2, -x * S2, RGB(0, 192, 192)
        
    AddLine -R, -1, 0, -x, 1 - Y2, 0, RGB(0, 192, 192) ' Bottom upward edges.
    AddLine -c, -1, a, -x * c1, 1 - Y2, x * S1, RGB(0, 192, 192) ' Bottom upward edges.
    AddLine -d, -1, b, -x * c2, 1 - Y2, x * S2, RGB(0, 192, 192)
    AddLine -d, -1, -b, -x * c2, 1 - Y2, -x * S2, RGB(0, 192, 192)
    AddLine -c, -1, -a, -x * c1, 1 - Y2, -x * S1, RGB(0, 192, 192)
    
    AddLine -R, -1, 0, -c, -1, a, RGB(0, 192, 192) ' Bottom
    AddLine -c, -1, a, -d, -1, b, RGB(0, 192, 192)
    AddLine -d, -1, b, -d, -1, -b, RGB(0, 192, 192)
    AddLine -d, -1, -b, -c, -1, -a, RGB(0, 192, 192)
    AddLine -c, -1, -a, -R, -1, 0, RGB(0, 192, 192)
    Dodecahedron(1) = p_NumLines
    
    'Icosahedron.
    Icosahedron(0) = p_NumLines + 1
    R = 2 / (2 * Sqr(1 - 2 * c1) + Sqr(3 / 4 * (2 - 2 * c1) - 2 * c2 - c2 * c2 - 1))
    S = R * Sqr(2 - 2 * c1)
    H = 1 - Sqr(S * S - R * R)
    a = R * S1
    b = R * S2
    c = R * c1
    d = R * c2
    AddLine R, H, 0, c, H, a, RGB(192, 0, 192)   ' Top
    AddLine c, H, a, d, H, b, RGB(192, 0, 192)
    AddLine d, H, b, d, H, -b, RGB(192, 0, 192)
    AddLine d, H, -b, c, H, -a, RGB(192, 0, 192)
    AddLine c, H, -a, R, H, 0, RGB(192, 0, 192)
    AddLine R, H, 0, 0, 1, 0, RGB(192, 0, 192)      ' Point
    AddLine c, H, a, 0, 1, 0, RGB(192, 0, 192)
    AddLine d, H, b, 0, 1, 0, RGB(192, 0, 192)
    AddLine d, H, -b, 0, 1, 0, RGB(192, 0, 192)
    AddLine c, H, -a, 0, 1, 0, RGB(192, 0, 192)
    
    AddLine -R, -H, 0, -c, -H, a, RGB(192, 0, 192)  ' Bottom
    AddLine -c, -H, a, -d, -H, b, RGB(192, 0, 192)
    AddLine -d, -H, b, -d, -H, -b, RGB(192, 0, 192)
    AddLine -d, -H, -b, -c, -H, -a, RGB(192, 0, 192)
    AddLine -c, -H, -a, -R, -H, 0, RGB(192, 0, 192)
    AddLine -R, -H, 0, 0, -1, 0, RGB(192, 0, 192)   ' Point
    AddLine -c, -H, a, 0, -1, 0, RGB(192, 0, 192)
    AddLine -d, -H, b, 0, -1, 0, RGB(192, 0, 192)
    AddLine -d, -H, -b, 0, -1, 0, RGB(192, 0, 192)
    AddLine -c, -H, -a, 0, -1, 0, RGB(192, 0, 192)

    AddLine R, H, 0, -d, -H, b, RGB(192, 0, 192)    ' Middle
    AddLine R, H, 0, -d, -H, -b, RGB(192, 0, 192)
    AddLine c, H, a, -d, -H, b, RGB(192, 0, 192)
    AddLine c, H, a, -c, -H, a, RGB(192, 0, 192)
    AddLine d, H, b, -c, -H, a, RGB(192, 0, 192)
    AddLine d, H, b, -R, -H, 0, RGB(192, 0, 192)
    AddLine d, H, -b, -R, -H, 0, RGB(192, 0, 192)
    AddLine d, H, -b, -c, -H, -a, RGB(192, 0, 192)
    AddLine c, H, -a, -c, -H, -a, RGB(192, 0, 192)
    AddLine c, H, -a, -d, -H, -b, RGB(192, 0, 192)
    Icosahedron(1) = p_NumLines

    'Eight Cubes:
    EightCubes(0) = p_NumLines + 1
    Call CreateEightCubes
    EightCubes(1) = p_NumLines
    
    'Ensure that the side-lenghts of any given
    'Platonic solid are equal.
    'Easy check - If not, it's not Platonic.
    If Not SameSideLengths(Tetrahedron(0), Tetrahedron(1)) Then MsgBox "Error in tetrahedron."
    If Not SameSideLengths(OneCube(0), OneCube(1)) Then MsgBox "Error in cube."
    If Not SameSideLengths(Octahedron(0), Octahedron(1)) Then MsgBox "Error in octahedron."
    If Not SameSideLengths(Dodecahedron(0), Dodecahedron(1)) Then MsgBox "Error in dodecahedron."
    If Not SameSideLengths(Icosahedron(0), Icosahedron(1)) Then MsgBox "Error in icosahedron."
    
    'Zero the number of text points contained in the
    'p_TextPoints() array - located in modTextPoint.
    p_NumTextPoints = 0
    
    'Create 3 axis labels (at distance sl from origin):
    'Adds the text points to the p_TextPoints() array
    'defined in module modTextPoint.
    AddText "X", vbWhite, 0.5, 0, 0 'X-axis label
    AddText "Y", vbWhite, 0, 0.5, 0 'Y-axis label
    AddText "Z", vbWhite, 0, 0, 0.5 'Z-axis label
    
    'Scale and translate the initial data points:
    'The view port (picCanvas) size is initially 350x350 pixels.
    'So setting a side-length of 70 pixels for each axis-line should
    'make all objects fill up most of the viewing area.
    'The side-length of each axis-line is currently 0.5 pixels, so we
    'we scale up (magnify) the size of our data points by 140
    'times larger:
    m3scale Sc, 70, 70, 70
    'No initial translation:
    m3translate Tr, 0, 0, 0
    'The following statement will translate
    'everything 50 pixels to the right:
    'm3translate T, 50, 0, 0
    
    'Combine the Scale and Translate transform matrices:
    m3multiply ScTr, Sc, Tr
    'Apply the combination transform to the current
    'base-point data:
    TransformAllLineBasePointsToCurrentPoints ScTr
    TransformAllTextBasePointsToCurrentPoints ScTr
    'Copy the transformed current-points back to the
    'base-points (permanantly resets all initial base-
    'point position values):
    CopyAllLineCurrentPointsToBasePoints
    CopyAllTextCurrentPointsToBasePoints
End Sub

'Create 8 cubes:
Private Sub CreateEightCubes()
Dim sl As Single        'cube side length (in pixel units)
Dim sl2 As Single       'half cube side length
Dim x As Single         'X-coordinate increment variable
Dim Y As Single         'Y-coordinate increment variable
Dim Z As Single         'Z-coordinate increment variable
Dim cnt As Integer      'color counter variable
Dim color As Long       'cube color
    
    'Initialize.
    sl = 1 '<- side length of each cube (in pixel units)
    sl2 = sl / 2! '<- half side length (in pixel units)
    cnt = 0 '<- cube color counter
    
    'Create 8 cubes in 3-D space.
    For x = -sl To sl Step sl * 2
        For Y = -sl To sl Step sl * 2
            For Z = -sl To sl Step sl * 2
                cnt = cnt + 1
                If cnt = 1 Then color = &HC0C0FF    'LIGHT_PINK
                If cnt = 2 Then color = &HFF&       'RED
                If cnt = 3 Then color = &H80FF&     'ORANGE
                If cnt = 4 Then color = &HFFFF&     'YELLOW
                If cnt = 5 Then color = &HFF00&     'GREEN
                If cnt = 6 Then color = &HFFFF00    'CYAN
                If cnt = 7 Then color = &HFF0000    'BLUE
                If cnt = 8 Then color = &HFF00FF    'MAGENTA
                'Adds the line segments to the p_Lines() array
                'defined in module modLine:
                AddLine x - sl2, Y - sl2, Z - sl2, x - sl2, Y - sl2, Z + sl2, color
                AddLine x - sl2, Y - sl2, Z + sl2, x - sl2, Y + sl2, Z + sl2, color
                AddLine x - sl2, Y + sl2, Z + sl2, x - sl2, Y + sl2, Z - sl2, color
                AddLine x - sl2, Y + sl2, Z - sl2, x - sl2, Y - sl2, Z - sl2, color
                AddLine x + sl2, Y - sl2, Z - sl2, x + sl2, Y - sl2, Z + sl2, color
                AddLine x + sl2, Y - sl2, Z + sl2, x + sl2, Y + sl2, Z + sl2, color
                AddLine x + sl2, Y + sl2, Z + sl2, x + sl2, Y + sl2, Z - sl2, color
                AddLine x + sl2, Y + sl2, Z - sl2, x + sl2, Y - sl2, Z - sl2, color
                AddLine x - sl2, Y - sl2, Z - sl2, x + sl2, Y - sl2, Z - sl2, color
                AddLine x - sl2, Y - sl2, Z + sl2, x + sl2, Y - sl2, Z + sl2, color
                AddLine x - sl2, Y + sl2, Z + sl2, x + sl2, Y + sl2, Z + sl2, color
                AddLine x - sl2, Y + sl2, Z - sl2, x + sl2, Y + sl2, Z - sl2, color
            Next Z
        Next Y
    Next x
   
End Sub

Private Sub Form_Resize()
    'The Form ScaleMode should be vbPixels:
    If Me.WindowState <> vbMinimized Then
        If Me.ScaleWidth - picControlBox.Width - 5 > 5 And Me.ScaleHeight > 24 Then
            picCanvas.Move Me.ScaleLeft, 24, Me.ScaleWidth - picControlBox.Width, Me.ScaleHeight - 24
            picControlBox.Move picCanvas.ScaleWidth, 24, picControlBox.Width, Me.ScaleHeight - 24
        End If
    End If
End Sub

Private Sub optObject_Click()
    picCanvas.SetFocus
End Sub

Private Sub optScreen_Click()
    picCanvas.SetFocus
End Sub

Private Sub optParallel_Click()
    'Set current projection type:
    m_Current_Projection_Matrix_Type = M4_PARALLEL
    
    picCanvas.SetFocus
    
    'Draw the graphics:
    Call RenderScene(picCanvas)
End Sub

Private Sub optPerspective_Click()
    'Set current projection type:
    m_Current_Projection_Matrix_Type = M4_PERSPECTIVE
    
    picCanvas.SetFocus
    
    'Draw the graphics:
    Call RenderScene(picCanvas)
End Sub

Private Sub Choice_Click(Index As Integer)
    picCanvas.SetFocus
    'Draw the graphics:
    Call RenderScene(picCanvas)
End Sub

Private Sub cmdHome_Click()

    'Rotate to home position: q = <0,0,0,1>
    'Note: 2 * arc_cos(1) = 0 radians rotation.
    Call q4copy( _
        m_Current_Rotation_Quaternion, _
        m_Rotation_Home_Quaternion)
    'q4print m_Current_Rotation_Quaternion
    
    picCanvas.SetFocus
    
    'Draw the graphics:
    Call RenderScene(picCanvas)
End Sub

'This is an elegant and very important example of what you can do
'with rotation quaternions.
'
'   Sequence of events:
'
'Calls to picCanvas_KeyDown() perform the following:
'1) Calculate a +4 degree rotation quaternion.
'2) Multiply +4d quaternion with current quaternion (or vice-versa).
'3) Set current quaternion equal to the multiplication product.
'   Call RenderScene(picCanvas), which performs the following:
'4) Convert current quaternion into a rotation matrix.
'5) Combine rotation matrix with scale matrix -> transform matrix
'6) Transform base-points to current-points with transform matrix.
'7) Transform current-points to screen-points with projection matrix.
'8) Draw screen-points directly to the screen.
'
Private Sub picCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim d4_rotateQuaternion(0 To 3) As Single
Dim rotateVect(0 To 2) As Single
Dim theta As Single
    
    'Hold 'X' button down to spin object about the X-Axis.
    If KeyCode = vbKeyX Then
        'Create a rotation quaternion for a
        '+4 degree rotation about an X-Axis:
        Call v3set(rotateVect, 1!, 0!, 0!)  '<- (1,0,0) the X-Axis
        'Call v3set(rotateVect, 1!, 1!, 0!)  '<- (1,1,0) instead
        theta = d2r(4!) 'degrees to radians
        Call q4fromAxis(d4_rotateQuaternion, rotateVect, theta)
        
        'About which X-Axis shall we spin?
        If optScreen Then
            'Spin about the Screen X-Axis: cq = cq*d4q
            Call q4multiply(m_Current_Rotation_Quaternion, _
                m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion)
        End If
        If optObject Then
            'Spin about the Object X-Axis: cq=d4q*cq
            Call q4multiply(m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion, _
                m_Current_Rotation_Quaternion)
        End If
        
        'Draw the graphics:
        Call RenderScene(picCanvas)
    End If
    
    'Hold 'Y' button down to spin object about the Y-Axis.
    If KeyCode = vbKeyY Then
        'Create a rotation quaternion for a
        '+4 degree rotation about a Y-Axis:
        Call v3set(rotateVect, 0!, 1!, 0!)  '<- (0,1,0) the Y-Axis
        theta = d2r(4!) 'degrees to radians
        Call q4fromAxis(d4_rotateQuaternion, rotateVect, theta)
        
        'About which Y-Axis shall we spin?
        If optScreen Then
            'Spin about the Screen Y-Axis: cq=cq*d4q
            Call q4multiply(m_Current_Rotation_Quaternion, _
                m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion)
        End If
        If optObject Then
            'Spin about the Object Y-Axis: cq=d4q*cq
            Call q4multiply(m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion, _
                m_Current_Rotation_Quaternion)
        End If
        
        'Draw the graphics:
        Call RenderScene(picCanvas)
    End If
    
    'Hold 'Z' button down to spin object about the Z-Axis.
    If KeyCode = vbKeyZ Then
        'Create a rotation quaternion for a
        '+4 degree rotation about a Z-Axis:
        Call v3set(rotateVect, 0!, 0!, 1!)  '<- (0,0,1) the Z-Axis
        theta = d2r(4!) 'degrees to radians
        Call q4fromAxis(d4_rotateQuaternion, rotateVect, theta)
        
        'About which Z-Axis shall we spin?
        If optScreen Then
            'Spin about the Screen Z-Axis: cq=cq*d4q
            Call q4multiply(m_Current_Rotation_Quaternion, _
                m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion)
        End If
        If optObject Then
            'Spin about the Object Z-Axis: cq=d4q*cq
            Call q4multiply(m_Current_Rotation_Quaternion, _
                d4_rotateQuaternion, _
                m_Current_Rotation_Quaternion)
        End If
        
        'Draw the graphics:
        Call RenderScene(picCanvas)
    End If
    
    'The important lesson to be learned here is that swapping the
    'ORDER of quaterneion multiplication factors determines the SET
    'of reference-axes about which the object rotation will occur,
    'either:
    '(a) The SCREEN reference-axes.
    '    or
    '(b) The OBJECT reference-axes.
    '
    'I used the x,y and z axes here to provide clear visual-feedback of
    'how quaternion multiplication order determines the rotation axes
    'set about which object rotation will occur.
    'You are not confined to rotating only about the x,y or z axes...
    'Under the 'If KeyCode = vbKeyX Then' statement above, change the
    'vector creation statement to: 'Call v3set(rotateVect, 1!, 1!, 0!)'
    'instead. Run the program and press the 'X' key.
    '
    'Are you starting to get the general idea now?
    '1) A quaternion holds a rotation.
    '2) Quaternion multiplication combines two rotations.
    '3) The order of the quaternion multiplication factors determines
    '   the set of reference-axes about which the object rotation will
    '   occur.
End Sub

Private Sub picCanvas_Resize()
    'Calculate new halfwidth and halfheight variables
    'and redraw the scene:
    m_PBox_HalfWidth_Pixels = picCanvas.ScaleWidth / 2!
    m_PBox_HalfHeight_Pixels = picCanvas.ScaleHeight / 2!
    'Draw the graphics:
    Call RenderScene(picCanvas)
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Echo the MouseDown event to the MouseDown handler in the
    'Trackball module. Map X,Y to Cartesian coordinates first:
    Call TrackBall_PBox_MouseDown(Button, Shift, x - m_PBox_HalfWidth_Pixels, m_PBox_HalfHeight_Pixels - Y)
    
    'A Right-Button mouse-down sets the
    'trackball diameter adjustment m_picCanvasRB_LastY value.
    If Button = vbRightButton Then
        m_picCanvasRB_LastY = Y
    End If
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim scrollVal As Single

    Select Case Button
        Case Is = vbLeftButton
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
            'Echo the MouseMove event to the MouseMove handler in the
            'Trackball module. Map X,Y to Cartesian coordinates first:
            'This action creates a new p_TrackBall_Rotation_Quaternion.
            Call TrackBall_PBox_MouseMove( _
                Button, _
                Shift, _
                x - m_PBox_HalfWidth_Pixels, _
                m_PBox_HalfHeight_Pixels - Y)
                        
            'Multiply the current rotation quaternion in this
            'module with the new mouse-move rotation quaternion,
            'just calculated in the Trackball module.
            'Note, we arange the multiplication factor order to make
            'the new rotation occur about the SCREEN reference-axes.
            Call q4multiply(m_Current_Rotation_Quaternion, _
                m_Current_Rotation_Quaternion, _
                p_TrackBall_Rotation_Quaternion)
            
            '   Sequence of events:
            '
            '1) Get the new TrackBall_Rotation_Quaternion increment quaternion.
            '2) Multiply the current quaternion with increment quaternion.
            '3) Set current quaternion equal to the multiplication product.
            '   Call RenderScene(picCanvas), which performs the following:
            '4) Convert current quaternion into a rotation matrix.
            '5) Combine rotation matrix with scale matrix => transform matrix
            '6) Transform base-points to current-points with transform matrix.
            '7) Transform current-points to screen-points with projection matrix.
            '8) Draw screen-points directly to the screen.

            'Draw the graphics:
            Call RenderScene(picCanvas)
        
        Case Is = vbRightButton
            'A Right-Button mouse-move visually
            'adjusts the trackball diameter.
            scrollVal = VScroll_TbR.value + (m_picCanvasRB_LastY - Y)
            'Note, this scrollbar is reverse-acting:
            If scrollVal < VScroll_TbR.Max Then scrollVal = VScroll_TbR.Max
            If scrollVal > VScroll_TbR.Min Then scrollVal = VScroll_TbR.Min
            VScroll_TbR.value = scrollVal
            'Debug.Print scrollVal
            m_picCanvasRB_LastY = Y
            
    End Select
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Echo the MouseUp event to the MouseUp handler in the
    'Trackball module:
    Call TrackBall_PBox_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub SetProjectionMatrix()
    
    Select Case m_Current_Projection_Matrix_Type
        Case Is = M4_PERSPECTIVE
            'Set ProjectionMatrix = a PERSPECTIVE projection:
            Call m4zPerspective(m_Current_Projection_Matrix, m_EyeZ_Pixels)
        Case Is = M4_PARALLEL
            'Set ProjectionMatrix = a PARALLEL projection:
            Call m4identity(m_Current_Projection_Matrix)
        Case Else
            'Set ProjectionMatrix = a PARALLEL projection:
            Call m4identity(m_Current_Projection_Matrix)
    End Select
End Sub

'Draw the current scene:
Private Sub RenderScene(pBox As PictureBox)
Dim ndx1 As Integer
Dim ndx2 As Integer
Dim cnt As Integer
   
    'Start hardware timing:
    hardwareCount = CounterStart
    
    'Note, if your drawing updates get slow and jerky,
    'esp. on large canvases, it's likely you've either
    'commented or deleted the following DoEvents
    'statement:
    DoEvents
    'My guess is that this DoEvents statement may, in
    'some way, let the system either abort or finish
    'any unfinished drawing commands being drawn into
    'the canvas's AutoRedrawn back bitmap.
    'But, in any case, there is definately some kind
    'of Microsoft 'black-box' magic going on here.
        
    'Convert the current rotation quaternion
    'into the m_Current_Rotation_Matrix:
    Call q4toMatrixF( _
        m_Current_Rotation_Matrix, _
        m_Current_Rotation_Quaternion)
    'q4print m_Current_Rotation_Quaternion
    'm4print m_Current_Rotation_Matrix
    
    'Do scaling. Combine current rotation matrix with
    'current scale matrix into current transform matrix:
    'Swap order of multiplication to swap reference axes
    'of scaling:
    'Debug.Print "before scale:"
    'm4print m_Current_Rotation_Matrix
    'Scale about Object_Axes:
    Call m3multiply( _
        m_Current_Transform_Matrix, _
        m_Current_Scale_Matrix, _
        m_Current_Rotation_Matrix)
    'Scale about Screen_Axes:
    'Call m3multiply( _
    '    m_Current_Transform_Matrix, _
    '    m_Current_Rotation_Matrix, _
    '    m_Current_Scale_Matrix)
    'Debug.Print "after scale:"
    'm4print m_Current_Transform_Matrix
    'Debug.Print
    
    'Do translation:
    'Swap order of multiplication to swap
    'reference axes of translation:
    'Debug.Print "before translate:"
    'm4print m_Current_Transform_Matrix
    Call m3translate(m_Current_Translate_matrix, 0, 0, 0)
    'Translate about Object_Axes:
    'Call m3multiply( _
    '    m_Current_Transform_Matrix, _
    '    m_Current_Translate_matrix, _
    '    m_Current_Transform_Matrix)
    'Translate about Screen_Axes:
    Call m3multiply( _
        m_Current_Transform_Matrix, _
        m_Current_Transform_Matrix, _
        m_Current_Translate_matrix)
    'Debug.Print "after translate:"
    'm4print m_Current_Transform_Matrix
    'Debug.Print
    
    'Generate the ProjectionMatrix:
    Call SetProjectionMatrix
    
    'Clear the drawing canvas:
    pBox.Cls
    
    'Draw the selected objects:
    For cnt = 0 To Choice.Count - 1
        If Choice(cnt).value = vbChecked Then
            'Draw the line segments within the given index range.
            If cnt = 0 Then ndx1 = AxisLines(0): ndx2 = AxisLines(1)
            If cnt = 1 Then ndx1 = Tetrahedron(0): ndx2 = Tetrahedron(1)
            If cnt = 2 Then ndx1 = OneCube(0): ndx2 = OneCube(1)
            If cnt = 3 Then ndx1 = Octahedron(0): ndx2 = Octahedron(1)
            If cnt = 4 Then ndx1 = Dodecahedron(0): ndx2 = Dodecahedron(1)
            If cnt = 5 Then ndx1 = Icosahedron(0): ndx2 = Icosahedron(1)
            If cnt = 6 Then ndx1 = EightCubes(0): ndx2 = EightCubes(1)

            'Update the current line-points array:
            Call TransformSomeLineBasePointsToCurrentPoints( _
                m_Current_Transform_Matrix, ndx1, ndx2)
            'Project the current line-points to the screen line-points:
            Call ProjectSomeLineCurrentPointsToScreenPoints( _
                m_Current_Projection_Matrix, ndx1, ndx2)
            'Draw the (screen) line segments:
            Call DrawSomeLines_Clip( _
                pBox, _
                m_PBox_HalfWidth_Pixels, _
                m_PBox_HalfHeight_Pixels, _
                m_EyeZ_Pixels, _
                ndx1, _
                ndx2)
        End If
    Next cnt
    
    'Draw the axes labels:
    If Choice(0).value = vbChecked Then
        'Update the current text-points array:
        Call TransformAllTextBasePointsToCurrentPoints( _
            m_Current_Transform_Matrix)
        'Project the current text-points to the screen text-points:
        Call ProjectAllTextCurrentPointsToScreenPoints( _
            m_Current_Projection_Matrix)
        'Draw the axes labels at the (screen) text-points:
        Call DrawAllText_Clip( _
            pBox, _
            m_PBox_HalfWidth_Pixels, _
            m_PBox_HalfHeight_Pixels, _
            m_EyeZ_Pixels)
    End If
    
    'Draw the trackball circles:
    If m_Draw_Trackball Then
        'This circle shows the trackball's actual size and location
        'on the screen:
        pBox.Circle (m_PBox_HalfWidth_Pixels, m_PBox_HalfHeight_Pixels), _
            p_TrackBall_Radius_Pixels, _
            &H646464 'LIGHT GRAY - RGB(100, 100, 100)
            
        'This circle shows where the screen mouse-move points switch
        'from being projected onto the trackball's spherical surface to
        'being projected onto the hyperbolic sheet of rotation:
        pBox.Circle (m_PBox_HalfWidth_Pixels, m_PBox_HalfHeight_Pixels), _
            p_TrackBall_Radius_Pixels * 0.7071!, _
            vbWhite
    End If
    m_Draw_Trackball = False
    
    'Update the hardware counter display.
    elapsedTime_mS = CounterStop(hardwareCount, fmtMillisecs) 'Stop timing.
    txtMilliSeconds.Text = SigFigs(elapsedTime_mS, 7) & " mS" 'Seven figures.

End Sub

Private Sub picControlBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picCanvas.SetFocus
End Sub

Private Sub picShapeSelection_Click()
    picCanvas.SetFocus
End Sub

'Scale size (Scale Factor, Zoom) of graphics:
Private Sub VScroll_SF_Change()
    VScroll_SF_Handler VScroll_SF.value, True
End Sub
Private Sub VScroll_SF_Scroll()
    VScroll_SF_Handler VScroll_SF.value, True
End Sub
Private Sub VScroll_SF_Handler(scroll_SF_Value As Integer, bRender As Boolean)
'Here, we implement an exponential responce to the scroll
'value changes to provide a more intutitive visual feedback
'of +/- zooming from scale = 1:
'scroll_SF_Value must range from -100 to +100 on input.
Static LastValue As Integer
Dim scaleFactor As Single
    
    If scroll_SF_Value <> LastValue Then
    'y = a^x, where (a) is minium scale value and (x) varies linearly
    'from -1 to +1.
    'We use scaleFactor = 0.1^x, => varies scale from 0.1 to 10:
    'scaleFactor=0.1^(-1)=10, scaleFactor=0.1^(0)=1, scaleFactor=0.1^(1)=0.1.
        
        'To vary scale from 0.2 to 5, use:
        'ScaleFactor = 0.2! ^ (scroll_SF_Value / 100!)
        
        '+/- qubic responce: -10 < ScaleFactor < +10
        'scaleFactor = 10! * (scroll_SF_Value / 100!) ^ 3
        
        '+/- quadratic responce: -10 < ScaleFactor < +10
        'scaleFactor = 10! * Sgn(scroll_SF_Value / 100!) * (scroll_SF_Value / 100!) ^ 2
        
        'Vary scale exponentally from 0.1 to 10, where
        '(scroll_SF_Value / 100!) varies linearly from -1 to +1:
        scaleFactor = 0.1! ^ (scroll_SF_Value / 100!)
        'Debug.Print ScaleFactor
        
        'Update the Scale Factor label:
        lblScaleFactor.Caption = scaleFactor
        
        'Set the current scale matrix
        Call m3scale( _
            m_Current_Scale_Matrix, _
            scaleFactor, _
            scaleFactor, _
            scaleFactor)
       
        'Draw the graphics:
        If bRender Then Call RenderScene(picCanvas)
    End If
    LastValue = scroll_SF_Value
End Sub

'Set size of Trackball radius:
Private Sub VScroll_TbR_Change()
    VScroll_TbR_Handler VScroll_TbR.value, True
End Sub
Private Sub VScroll_TbR_Scroll()
    VScroll_TbR_Handler VScroll_TbR.value, True
End Sub
Private Sub VScroll_TbR_Handler(scroll_TbR_Value As Integer, bRender As Boolean)
Static LastValue As Integer

    If scroll_TbR_Value <> LastValue Then
        
        'Debug.Print scroll_TbR_Value
        
        'Set the trackball radius variable:
        p_TrackBall_Radius_Pixels = CSng(scroll_TbR_Value)
        
        'Update the Trackball Diameter label:
        lblTrackballDiameter.Caption = _
            CStr(2! * p_TrackBall_Radius_Pixels) & " Pixels"
        
        'Draw the graphics?
        If bRender Then
            'draw the trackball.
            m_Draw_Trackball = True
            'draw the graphics
            Call RenderScene(picCanvas)
        End If
    End If
    LastValue = scroll_TbR_Value
End Sub

'Set the center of projection distance.
Private Sub VScroll_CoP_Change()
    VScroll_CoP_Handler VScroll_CoP.value, True
End Sub
Private Sub VScroll_CoP_Scroll()
    VScroll_CoP_Handler VScroll_CoP.value, True
End Sub
Private Sub VScroll_CoP_Handler(scroll_CoP_Value As Integer, bRender As Boolean)
'Here, we implement an exponential responce to the scroll
'value changes to provide a more intutitive visual feedback
'to the user:
'scroll_CoP_Value must range from -100 to +100 on input.
Static LastValue As Integer
Dim CoP_ScaleFactor As Single

    If scroll_CoP_Value <> LastValue Then
    'We use y=a^x, where (a) is minium scale value
    'and (x) varies linearly from -1 to +1. => y(0)=1.
    
        'Debug.Print scroll_CoP_Value
        
        'Vary CoP_ScaleFactor from 0.1 to 10:
        CoP_ScaleFactor = 0.1! ^ (scroll_CoP_Value / 100!)
       
        'Multiply CoP_ScaleFactor by 350 and add an offset of 65
        'to get a range from 100 to 3565 pixels:
        'Sets the Center of Projection (along the screen z-axis):
        m_EyeZ_Pixels = CSng(CInt(CoP_ScaleFactor * 350!) + 65!)
        
        'Update the Center of Projection label:
        lblCenterOfProjection.Caption = _
            CStr(m_EyeZ_Pixels) & " Pixels"
        
        'Draw the graphics:
        If bRender Then Call RenderScene(picCanvas)
    End If
    LastValue = scroll_CoP_Value
End Sub


