VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   1980
   ClientTop       =   1710
   ClientWidth     =   5520
   ForeColor       =   &H00000000&
   Height          =   5685
   Left            =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   Top             =   1365
   Width           =   5640
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic3D 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer Timer2 
         Interval        =   250
         Left            =   2640
         Top             =   1560
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim LightSpeed As Integer
Dim Pi As Single


















Public Sub Move_In()
For i = 0 To nrlines
pts3d(i).Z = pts3d(i).Z - LightSpeed    'Move closer to the camera
    If pts3d(i).Z <= 5 Then  'If too close to the camera
        pts3d(i).Z = 500
        pts3d(i).X = Rnd * pic3D.Width
        pts3d(i).Y = Rnd * pic3D.Height
    End If      'Make a new star at a distance of 500
Next i

Redraw
End Sub











Private Sub Form_Activate()
Do
Move_In
DoEvents
Loop
End Sub

Private Sub Form_Load()
Pi = 3.1427
pic3D.Height = Screen.Height \ Screen.TwipsPerPixelY
pic3D.Width = Screen.Width \ Screen.TwipsPerPixelX
New_Star_Field
LightSpeed = 0
End Sub





Public Sub Redraw()
Dim RGBDist As Integer  '(For dimming effect)
Dim ConstDist As Integer
ConstDist = 300  'Basically lower number = lower viewing angle
pic3D.Cls   'Clear the box
For i = 0 To nrlines 'Draw each point
    pts(1).X = pic3D.Width / 2 - (ConstDist * (pic3D.Width / 2 - pts3d(i).X)) / pts3d(i).Z  'X value of the first point
    pts(2).X = pic3D.Width / 2 - (ConstDist * (pic3D.Width / 2 - pts3d(i).X)) / (pts3d(i).Z + LightSpeed / 2) 'X value of secone point (will be half the speed we are going at behind the first)
    pts(1).Y = pic3D.Height / 2 - (ConstDist * (pic3D.Height / 2 - pts3d(i).Y) + Angle) / pts3d(i).Z 'Y for point 1
    pts(2).Y = pic3D.Height / 2 - (ConstDist * (pic3D.Height / 2 - pts3d(i).Y) + Angle) / (pts3d(i).Z + LightSpeed / 2) 'Y for point 2
    RGBDist = 250 - (pts3d(i).Z \ 2) 'RGB value depends on distance from camera
    pic3D.Line (pts(1).X, pts(1).Y)-(pts(2).X, pts(2).Y), RGB(RGBDist, RGBDist, RGBDist)   'Draws line from first point to second point
Next i
pic3D.Refresh
End Sub

Public Sub New_Star_Field()
nrlines = 199 'Change theis as much as you like (less that 2499 though)
Randomize
    For i = 0 To nrlines 'Each of the stars is only one point
        pts3d(i).X = Rnd * pic3D.Width  'X Value
        pts3d(i).Y = Rnd * pic3D.Height 'Y Value
        pts3d(i).Z = Rnd * 450 + 50     'Z Value (Don't make too close to the user)
    Next i
End Sub

Private Sub pic3D_Click()
Unload Me
End
End Sub




Private Sub Timer2_Timer()
LightSpeed = LightSpeed + 1
If LightSpeed > 15 Then Timer2.Enabled = False
LightSpeed = LightSpeed + 1
End Sub


