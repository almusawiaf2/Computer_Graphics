VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7670
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7670
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   10920
      Top             =   1560
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Box"
      Height          =   490
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   970
   End
   Begin VB.PictureBox pictPerspective 
      Height          =   3730
      Left            =   5400
      ScaleHeight     =   3690
      ScaleWidth      =   5010
      TabIndex        =   3
      Top             =   3840
      Width           =   5050
   End
   Begin VB.PictureBox pictTopView 
      Height          =   3610
      Left            =   5400
      ScaleHeight     =   3570
      ScaleWidth      =   5010
      TabIndex        =   2
      Top             =   120
      Width           =   5050
   End
   Begin VB.PictureBox pictSideView 
      Height          =   3730
      Left            =   120
      ScaleHeight     =   3690
      ScaleWidth      =   5130
      TabIndex        =   1
      Top             =   3840
      Width           =   5170
   End
   Begin VB.PictureBox pictFrontView 
      Height          =   3610
      Left            =   120
      ScaleHeight     =   3570
      ScaleWidth      =   5130
      TabIndex        =   0
      Top             =   120
      Width           =   5170
   End
   Begin VB.Menu mnuShapes 
      Caption         =   "Shapes"
      Begin VB.Menu mnuContinous_lines 
         Caption         =   "Continous Lines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparatedLines 
         Caption         =   "Separated Lines"
      End
      Begin VB.Menu mnuContinousTrianlges 
         Caption         =   "Continous Triangle"
      End
      Begin VB.Menu mnuSeparatedTriangles 
         Caption         =   "Separated Triangle"
      End
      Begin VB.Menu mnuSeparated4Polygons 
         Caption         =   "Separated 4 Polygon"
      End
      Begin VB.Menu mnuResetShape 
         Caption         =   "Reset Shape"
      End
   End
   Begin VB.Menu mnuTransformations 
      Caption         =   "Transformation"
      Begin VB.Menu mnu2DWorld 
         Caption         =   "2D World"
         Begin VB.Menu mnuReflection 
            Caption         =   "Reflection"
            Begin VB.Menu mnu2DRefX 
               Caption         =   "On X Axis"
            End
            Begin VB.Menu mnu2DRefY 
               Caption         =   "On Y-Axis"
            End
            Begin VB.Menu mnu2DRefXY 
               Caption         =   "On XY"
            End
         End
         Begin VB.Menu mnu2DScale 
            Caption         =   "Scale"
         End
         Begin VB.Menu mnu2DRotate 
            Caption         =   "Rotate"
         End
         Begin VB.Menu mnu2DShear 
            Caption         =   "Shear"
         End
         Begin VB.Menu mnu2DMove 
            Caption         =   "Move"
            Begin VB.Menu mnu2DMoveUp 
               Caption         =   "Up"
            End
            Begin VB.Menu mnu2DMoveDown 
               Caption         =   "Down"
            End
            Begin VB.Menu mnu2DMoveLeft 
               Caption         =   "Left"
            End
            Begin VB.Menu mnu2DMoveRight 
               Caption         =   "Right"
            End
         End
      End
      Begin VB.Menu mnu3D 
         Caption         =   "3D World"
         Begin VB.Menu mnu3DRef 
            Caption         =   "Reflection"
         End
         Begin VB.Menu mnu3DScale 
            Caption         =   "Scale"
         End
         Begin VB.Menu mnu3DRotate 
            Caption         =   "Rotate"
            Begin VB.Menu mnu3dRotateX 
               Caption         =   "Around X"
            End
            Begin VB.Menu mnu3DRotateY 
               Caption         =   "Around Y"
            End
            Begin VB.Menu mnu3DRotateZ 
               Caption         =   "Around Z"
            End
         End
         Begin VB.Menu mnu3DShear 
            Caption         =   "Shear"
         End
         Begin VB.Menu mnu3DMove 
            Caption         =   "Move"
            Begin VB.Menu mnu3DMoveUP 
               Caption         =   "Up"
            End
            Begin VB.Menu mnu3DMoveDown 
               Caption         =   "Down"
            End
            Begin VB.Menu mnu3DMoveLeft 
               Caption         =   "Left"
            End
            Begin VB.Menu mnu3DMoveRight 
               Caption         =   "Right"
            End
            Begin VB.Menu mnu3DMoveIn 
               Caption         =   "Closer"
            End
            Begin VB.Menu mnu3DMoveOut 
               Caption         =   "Far"
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Ahmad F. Al Musawi.
'@ 2021.
' University of Thi Qar.
' College of Computer Science and Mathematics
' Computer Graphics : Third Stage : Computer Science Department.
Dim Box(9, 4) As Integer
Dim ResBox(9, 4) As Integer
Dim T(4, 4) As Double

Dim Points(100, 4) As Double ' +-XY
Dim OLD_Points(100, 4) As Double '+XY
Dim iPoint As Integer
Const Pi As Double = 22 / 7
Dim XCen, YCen As Integer
Dim thm As Integer
Private Sub cmdDown_Click()
    Call MoveDown2D
End Sub
Private Sub cmdLeft_Click()
    Call MoveLeft2D
End Sub
Private Sub cmdRight_Click()
    Call MoveRight2D
End Sub
Private Sub cmdUP_Click()
    Call MoveUp2D
End Sub
Private Sub MoveDown2D()
    Call Moving(0, -50)
    Call Multiply2
End Sub
Private Sub MoveUp2D()
    Call Moving(0, 50)
    Call Multiply2
End Sub
Private Sub MoveLeft2D()
    Call Moving(-50, 0)
    Call Multiply2
End Sub
Private Sub MoveRight2D()
    Call Moving(50, 0)
    Call Multiply2
End Sub


Private Sub Moving(ByVal Mx As Double, ByVal My As Double)
    T(1, 1) = 1: T(1, 2) = 0: T(1, 3) = 0
    T(2, 1) = 0: T(2, 2) = 1: T(2, 3) = 0
    T(3, 1) = Mx: T(3, 2) = My: T(3, 3) = 1
End Sub
Private Sub Scaling(ByVal Sx As Double, ByVal Sy As Double)
    T(1, 1) = Sx: T(1, 2) = 0: T(1, 3) = 0
    T(2, 1) = 0: T(2, 2) = Sy: T(2, 3) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1
End Sub
Private Sub shear(ByVal sh_x As Integer, ByVal sh_y As Integer)
    T(1, 1) = 1: T(1, 2) = sh_y: T(1, 3) = 0
    T(2, 1) = sh_x: T(2, 2) = 1: T(2, 3) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1
End Sub
Private Sub Reflection(ByVal Rx As Integer, ByVal Ry As Integer)
    T(1, 1) = Rx: T(1, 2) = 0: T(1, 3) = 0
    T(2, 1) = 0: T(2, 2) = Ry: T(2, 3) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1
End Sub
Private Sub rotate(ByVal th As Integer)
    T(1, 1) = Cos(th * Pi / 180): T(1, 2) = Sin(th * Pi / 180): T(1, 3) = 0
    T(2, 1) = -Sin(th * Pi / 180): T(2, 2) = Cos(th * Pi / 180): T(2, 3) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1
End Sub

Private Sub Moving3D(ByVal Mx As Double, ByVal My As Double, ByVal Mz As Double)
    T(1, 1) = 1: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = 1: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1: T(3, 4) = 0
    T(4, 1) = Mx: T(4, 2) = My: T(4, 3) = Mz: T(4, 4) = 1
End Sub
Private Sub Scaling3D(ByVal Sx As Double, ByVal Sy As Double, ByVal Sz As Double)
    T(1, 1) = Sx: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = Sy: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = Sz: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub Reflection3D(ByVal Rx As Integer, ByVal Ry As Integer, ByVal Rz As Integer)
    T(1, 1) = Rx: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = Ry: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = Rz: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub rotate3D_Z(ByVal th As Integer)
    T(1, 1) = Cos(th * Pi / 180): T(1, 2) = Sin(th * Pi / 180): T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = -Sin(th * Pi / 180): T(2, 2) = Cos(th * Pi / 180): T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub rotate3D_X(ByVal th As Integer)
    T(1, 1) = 1: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = Cos(th * Pi / 180): T(2, 3) = Sin(th * Pi / 180): T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = -Sin(th * Pi / 180): T(3, 3) = Cos(th * Pi / 180): T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub rotate3D_Y(ByVal th As Integer)
    T(1, 1) = Cos(th * Pi / 180): T(1, 2) = 0: T(1, 3) = Sin(th * Pi / 180): T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = 1: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = -Sin(th * Pi / 180): T(3, 2) = 0: T(3, 3) = Cos(th * Pi / 180): T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub


Private Sub Multiply2()
    Dim c(100, 3) As Double
    For i = 1 To iPoint
        For j = 1 To 3
            c(i, j) = 0
            For k = 1 To 3
                c(i, j) = c(i, j) + Points(i, k) * T(k, j)
            Next
        Next
    Next
    For i = 1 To iPoint
        Points(i, 1) = c(i, 1)
        Points(i, 2) = c(i, 2)
        Points(i, 3) = c(i, 3)
    Next
    Call DrawPoint
End Sub
Private Sub Multiply3D_Save()
    For i = 1 To 9
        For j = 1 To 4
            ResBox(i, j) = 0
            For k = 1 To 4
                ResBox(i, j) = ResBox(i, j) + Box(i, k) * T(k, j)
            Next
        Next
    Next
    For i = 1 To 9
        Box(i, 1) = ResBox(i, 1)
        Box(i, 2) = ResBox(i, 2)
        Box(i, 3) = ResBox(i, 3)
    Next
    Call Draw3D
End Sub
Private Sub Multiply3D()
    For i = 1 To 9
        For j = 1 To 4
            ResBox(i, j) = 0
            For k = 1 To 4
                ResBox(i, j) = ResBox(i, j) + Box(i, k) * T(k, j)
            Next
        Next
    Next
End Sub

Private Sub DrawPoint()
    Call convert_New_to_Old
    pictFrontView.Cls
    For i = 1 To iPoint
        X = OLD_Points(i, 1)
        Y = OLD_Points(i, 2)
        pictFrontView.Circle (X, Y), 50
    Next
End Sub

Private Sub convert_New_to_Old()
    For i = 1 To iPoint
        newx = Points(i, 1)
        newy = Points(i, 2)
        Xd = newx + XCen
        Yd = YCen - newy
        OLD_Points(i, 1) = Xd
        OLD_Points(i, 2) = Yd
    Next
End Sub



Private Sub mnu2dMoveDown_Click()
    Call MoveDown2D
End Sub

Private Sub mnu2DMoveLeft_Click()
    Call MoveLeft2D
End Sub

Private Sub mnu2DMoveRight_Click()
    Call MoveRight2D
End Sub

Private Sub mnu2DMoveUp_Click()
    Call MoveUp2D
End Sub

Private Sub mnu2DRefX_Click()
    Call Reflection(1, -1)
    Call Multiply2
End Sub

Private Sub mnu2DRefXY_Click()
    Call Reflection(-1, -1)
    Call Multiply2
End Sub

Private Sub mnu2DRefY_Click()
    Call Reflection(-1, 1)
    Call Multiply2
End Sub

Private Sub mnu2dRotate_Click()
    th = InputBox("Enter angle of rotation")
    Call rotate(th)
    Call Multiply2
End Sub

Private Sub mnu2DScale_Click()
    Sx = InputBox("Enter Scale on X-Axis : Sx")
    Sy = InputBox("Enter Scale on Y-Axis : Sy")
    Call Scaling(Sx, Sy)
    Call Multiply2
End Sub

Private Sub mnu2DShear_Click()
    sh_x = InputBox("Enter shear value on X-Axis : SHx")
    sh_y = InputBox("Enter Shear value on Y-Axis : SHy")
    Call shear(sh_x, sh_y)
    Call Multiply2
End Sub

Private Sub mnu3DMoveDown_Click()
    Call Moving3D(0, -50, 0)
    Call Multiply3D_Save
End Sub
Private Sub mnu3DMoveIn_Click()
    Call Moving3D(0, 0, -50)
    Call Multiply3D_Save
End Sub
Private Sub mnu3DMoveLeft_Click()
    Call Moving3D(-50, 0, 0)
    Call Multiply3D_Save
End Sub
Private Sub mnu3DMoveOut_Click()
    Call Moving3D(0, 0, 50)
    Call Multiply3D_Save
End Sub
Private Sub mnu3DMoveRight_Click()
    Call Moving3D(50, 0, 0)
    Call Multiply3D_Save
End Sub
Private Sub mnu3DMoveUp_Click()
    Call Moving3D(0, 50, 0)
    Call Multiply3D_Save
End Sub

Private Sub mnu3dRotateX_Click()
    th = InputBox("Enter angle of rotation around X")
    Call rotate3D_X(th)
    Call Multiply3D_Save
End Sub

Private Sub mnu3DRotateY_Click()
    th = InputBox("Enter angle of rotation around Y")
    Call rotate3D_Y(th)
    Call Multiply3D_Save
End Sub

Private Sub mnu3DRotateZ_Click()
    th = InputBox("Enter angle of rotation around Z")
    Call rotate3D_Z(th)
    Call Multiply3D_Save
End Sub

Private Sub mnuContinous_lines_Click()
    Call continousLines
End Sub

Private Sub mnuContinousTrianlges_Click()
    Call ContinousTriangles
End Sub

Private Sub mnuResetShape_Click()
    Call ResetShape
End Sub

Private Sub mnuSeparated4Polygons_Click()
    Call SeparatedPolygons
End Sub

Private Sub mnuSeparatedLines_Click()
    Call SeparatedLines
End Sub

Private Sub mnuSeparatedTriangles_Click()
    Call SeparatedTriangles
End Sub

Private Sub pictFrontView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'convert device to Decarzian System
    newx = X - XCen
    newy = YCen - Y
    'add newX, NewY to list1
    iPoint = iPoint + 1
    'Save newX,newY to Points
    Points(iPoint, 1) = newx
    Points(iPoint, 2) = newy
    Points(iPoint, 3) = 1
    DrawPoint
End Sub
Private Sub Form_Load()
    XCen = pictFrontView.ScaleWidth / 2
    YCen = pictFrontView.ScaleHeight / 2
    Call SetBox
End Sub
Private Sub pictFrontView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newx = X - XCen
    newy = YCen - Y
End Sub

Private Sub ResetShape()
    pictFrontView.Cls
    'Reset Shape : Remove everything
    For i = 1 To iPoint
        Points(i, 1) = 0
        Points(i, 2) = 0
    Next
    iPoint = 0
End Sub

Private Sub continousLines()
    pictFrontView.Cls
    DrawPoint
    'line by line
    For i = 1 To iPoint - 1
        X1 = OLD_Points(i, 1)
        Y1 = OLD_Points(i, 2)
        X2 = OLD_Points(i + 1, 1)
        Y2 = OLD_Points(i + 1, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
End Sub
Private Sub SeparatedLines()
    pictFrontView.Cls
    DrawPoint
    For i = 1 To iPoint - 1 Step 2
        X1 = OLD_Points(i, 1)
        Y1 = OLD_Points(i, 2)
        X2 = OLD_Points(i + 1, 1)
        Y2 = OLD_Points(i + 1, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
End Sub
Private Sub ContinousTriangles()
    pictFrontView.Cls
    DrawPoint
    For i = 1 To iPoint - 2
        X1 = OLD_Points(i, 1)
        Y1 = OLD_Points(i, 2)
        X2 = OLD_Points(i + 1, 1)
        Y2 = OLD_Points(i + 1, 2)
        x3 = OLD_Points(i + 2, 1)
        y3 = OLD_Points(i + 2, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
        pictFrontView.Line -(x3, y3)
        pictFrontView.Line -(X1, Y1)
    Next
End Sub

Private Sub SeparatedTriangles()
    'each three points seperated
    pictFrontView.Cls
    DrawPoint
    For i = 1 To iPoint - 2 Step 3
        X1 = OLD_Points(i, 1)
        Y1 = OLD_Points(i, 2)
        X2 = OLD_Points(i + 1, 1)
        Y2 = OLD_Points(i + 1, 2)
        x3 = OLD_Points(i + 2, 1)
        y3 = OLD_Points(i + 2, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
        pictFrontView.Line -(x3, y3)
        pictFrontView.Line -(X1, Y1)
    Next
End Sub
Private Sub SeparatedPolygons()
    pictFrontView.Cls
    DrawPoint
    For i = 1 To iPoint - 3 Step 4
        X1 = OLD_Points(i, 1)
        Y1 = OLD_Points(i, 2)
        X2 = OLD_Points(i + 1, 1)
        Y2 = OLD_Points(i + 1, 2)
        x3 = OLD_Points(i + 2, 1)
        y3 = OLD_Points(i + 2, 2)
        x4 = OLD_Points(i + 3, 1)
        y4 = OLD_Points(i + 3, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
        pictFrontView.Line -(x3, y3)
        pictFrontView.Line -(x4, y4)
        pictFrontView.Line -(X1, Y1)
    Next
End Sub

Private Sub Multiply()
    For i = 1 To iPoint
        X = Points(i, 1)
        Y = Points(i, 2)
        newx = X * T(1, 1) + Y * T(2, 1)
        newy = X * T(1, 2) + Y * T(2, 2)
        Points(i, 1) = newx
        Points(i, 2) = newy
        Me.List2.AddItem (X & vbTab & Y)
    Next
    Call DrawPoint
End Sub
Private Sub SetBox()
    Box(1, 1) = 0: Box(1, 2) = 0: Box(1, 3) = 0: Box(1, 4) = 1
    Box(2, 1) = 1000: Box(2, 2) = 0: Box(2, 3) = 0: Box(2, 4) = 1
    Box(3, 1) = 1000: Box(3, 2) = 1000: Box(3, 3) = 0: Box(3, 4) = 1
    Box(4, 1) = 0: Box(4, 2) = 1000: Box(4, 3) = 0: Box(4, 4) = 1

    Box(5, 1) = 0: Box(5, 2) = 0: Box(5, 3) = 1000: Box(5, 4) = 1
    Box(6, 1) = 1000: Box(6, 2) = 0: Box(6, 3) = 1000: Box(6, 4) = 1
    Box(7, 1) = 1000: Box(7, 2) = 1000: Box(7, 3) = 1000: Box(7, 4) = 1
    Box(8, 1) = 0: Box(8, 2) = 1000: Box(8, 3) = 1000: Box(8, 4) = 1
    
    Box(9, 1) = 500: Box(9, 2) = 500: Box(9, 3) = 1500: Box(9, 4) = 1
End Sub

Private Sub ShowFrontView()
    Call SetFrontViewT
    Call Multiply3D
    '1,2,3,4
    For i = 1 To 3
        j = i + 1
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
    X1 = ResBox(1, 1): Y1 = ResBox(1, 2)
    pictFrontView.Line -(X1, Y1)
    '5,6,7,8
    For i = 5 To 7
        j = i + 1
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
    X1 = ResBox(5, 1): Y1 = ResBox(5, 2)
    pictFrontView.Line -(X1, Y1)
    '1-5, 2-6, 3-7, 4-8
    For i = 1 To 4
        j = i + 4
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
    For i = 5 To 8
        X1 = ResBox(9, 1): Y1 = ResBox(9, 2)
        X2 = ResBox(i, 1): Y2 = ResBox(i, 2)
        pictFrontView.Line (X1, Y1)-(X2, Y2)
    Next
        
End Sub
Private Sub ShowSideView()
    Call SetSideViewT
    Call Multiply3D
    '1,2,3,4
    For i = 1 To 3
        j = i + 1
        Z1 = ResBox(i, 3): Y1 = ResBox(i, 2)
        Z2 = ResBox(j, 3): Y2 = ResBox(j, 2)
        pictSideView.Line (Z1, Y1)-(Z2, Y2)
    Next
    Z1 = ResBox(1, 3): Y1 = ResBox(1, 2)
    pictSideView.Line -(Z1, Y1)
    '5,6,7,8
    For i = 5 To 7
        j = i + 1
        Z1 = ResBox(i, 3): Y1 = ResBox(i, 2)
        Z2 = ResBox(j, 3): Y2 = ResBox(j, 2)
        pictSideView.Line (Z1, Y1)-(Z2, Y2)
    Next
    Z1 = ResBox(5, 3): Y1 = ResBox(5, 2)
    pictSideView.Line -(Z1, Y1)
    '1-5, 2-6, 3-7, 4-8
    For i = 1 To 4
        j = i + 4
        Z1 = ResBox(i, 3): Y1 = ResBox(i, 2)
        Z2 = ResBox(j, 3): Y2 = ResBox(j, 2)
        pictSideView.Line (Z1, Y1)-(Z2, Y2)
    Next
    For i = 5 To 8
        Z1 = ResBox(9, 3): Y1 = ResBox(9, 2)
        Z2 = ResBox(i, 3): Y2 = ResBox(i, 2)
        pictSideView.Line (Z1, Y1)-(Z2, Y2)
    Next

End Sub
Private Sub ShowTopView()
    Call SetTopViewT
    Call Multiply3D
    '1,2,3,4
    For i = 1 To 3
        j = i + 1
        X1 = ResBox(i, 1): Z1 = ResBox(i, 3)
        X2 = ResBox(j, 1): Z2 = ResBox(j, 3)
        pictTopView.Line (X1, Z1)-(X2, Z2)
    Next
    X1 = ResBox(1, 1): Z1 = ResBox(1, 3)
    pictTopView.Line -(X1, Z1)
    '5,6,7,8
    For i = 5 To 7
        j = i + 1
        X1 = ResBox(i, 1): Z1 = ResBox(i, 3)
        X2 = ResBox(j, 1): Z2 = ResBox(j, 3)
        pictTopView.Line (X1, Z1)-(X2, Z2)
    Next
    X1 = ResBox(5, 1): Z1 = ResBox(5, 3)
    pictTopView.Line -(X1, Z1)
    '1-5, 2-6, 3-7, 4-8
    For i = 1 To 4
        j = i + 4
        X1 = ResBox(i, 1): Z1 = ResBox(i, 3)
        X2 = ResBox(j, 1): Z2 = ResBox(j, 3)
        pictTopView.Line (X1, Z1)-(X2, Z2)
    Next
    For i = 5 To 8
        X1 = ResBox(9, 1): Z1 = ResBox(9, 3)
        X2 = ResBox(i, 1): Z2 = ResBox(i, 3)
        pictTopView.Line (X1, Z1)-(X2, Z2)
    Next

End Sub
Private Sub SetFrontViewT()
    T(1, 1) = 1: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = 1: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 0: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub SetSideViewT()
    T(1, 1) = 0: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = 1: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub
Private Sub SetTopViewT()
    T(1, 1) = 1: T(1, 2) = 0: T(1, 3) = 0: T(1, 4) = 0
    T(2, 1) = 0: T(2, 2) = 0: T(2, 3) = 0: T(2, 4) = 0
    T(3, 1) = 0: T(3, 2) = 0: T(3, 3) = 1: T(3, 4) = 0
    T(4, 1) = 0: T(4, 2) = 0: T(4, 3) = 0: T(4, 4) = 1
End Sub


Private Sub ShowPerspective()
    Call SetPerspective
    '1,2,3,4
    For i = 1 To 3
        j = i + 1
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictPerspective.Line (X1, Y1)-(X2, Y2)
    Next
    X1 = ResBox(1, 1): Y1 = ResBox(1, 2)
    pictPerspective.Line -(X1, Y1)
    '5,6,7,8
    For i = 5 To 7
        j = i + 1
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictPerspective.Line (X1, Y1)-(X2, Y2)
    Next
    X1 = ResBox(5, 1): Y1 = ResBox(5, 2)
    pictPerspective.Line -(X1, Y1)
    '1-5, 2-6, 3-7, 4-8
    For i = 1 To 4
        j = i + 4
        X1 = ResBox(i, 1): Y1 = ResBox(i, 2)
        X2 = ResBox(j, 1): Y2 = ResBox(j, 2)
        pictPerspective.Line (X1, Y1)-(X2, Y2)
    Next
    For i = 5 To 8
        X1 = ResBox(9, 1): Y1 = ResBox(9, 2)
        X2 = ResBox(i, 1): Y2 = ResBox(i, 2)
        pictPerspective.Line (X1, Y1)-(X2, Y2)
    Next
        
End Sub

Private Sub SetPerspective()
    xc = pictPerspective.ScaleWidth / 2
    yc = pictPerspective.ScaleHeight / 2
    zp = 2000 ' vanishing point
    r = 1 / zp
    For i = 1 To 9
        X = Box(i, 1)
        Y = Box(i, 2)
        z = Box(i, 3)
        newx = X / ((r * z) + 1)
        newy = Y / ((r * z) + 1)
        ResBox(i, 1) = newx + xc
        ResBox(i, 2) = yc - newy
    Next
End Sub
Private Sub Draw3D()
    Call ClearViews
    Call ShowFrontView
    Call ShowSideView
    Call ShowTopView
    Call ShowPerspective
End Sub
Private Sub ClearViews()
    Me.pictFrontView.Cls
    Me.pictPerspective.Cls
    Me.pictSideView.Cls
    Me.pictTopView.Cls
End Sub
Private Sub cmdDraw_Click()
    Call Draw3D
End Sub

Private Sub Timer1_Timer()
    Call rotate3D_Y(2)
    Call Multiply3D_Save
End Sub
