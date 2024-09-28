VERSION 5.00
Begin VB.Form GG 
   Caption         =   "GoGreen PC Art Designer Output"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "GG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnDrawFlag As Boolean
Private Sub Form_Load()
blnDrawFlag = False
Me.FillColor = &HFF
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
CurrentX = x
CurrentY = Y
blnDrawFlag = True
'Debug.Print "md: " & X & ", " & Y & " " & blnDrawFlag
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If blnDrawFlag Then Line -(x, Y), &HFF
'Debug.Print "md: " & X & ", " & Y & " " & blnDrawFlag
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
blnDrawFlag = False
End Sub

Private Sub Form_Resize()

x1 = GG.Left
y1 = GG.Top
x2 = GG.Width
y2 = GG.Height

'HH.o.AddItem "(x1, y1 : x2, y2) " & x1 & "," & y1 & " : " & x2 & ", " & y2

End Sub

