Attribute VB_Name = "Module1"
Public pts(1 To 2) As POINTAPI
Public pts3d(0 To 2499) As POINTAPI3D
Public nrlines As Integer

Public Type POINTAPI
    X As Long
    Y As Long
End Type


Public Type POINTAPI3D
    X As Long
    Y As Long
    Z As Long
End Type

