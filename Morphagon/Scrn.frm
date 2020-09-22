VERSION 5.00
Begin VB.Form Scrn 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Scrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private pX(0 To 12), pY(0 To 12), pXmov(0 To 12), pYmov(0 To 12), Mx, My, xmov, ymov As Integer
Private MouseMoveInd As Single

Private Sub Form_Click()
    ShowCursor (1)
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ShowCursor (1)
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMoveInd = MouseMoveInd + 1
    If MouseMoveInd > 4 Then
        ShowCursor (1)
        End
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then End
    Randomize
    ShowCursor (0)
    Mx = Int(Rnd * Screen.Width + 1)
    My = Int(Rnd * Screen.Height + 1)
    Dim num As Integer
    For num = 0 To Sides - 1
        pX(num) = Mx + Int(Rnd * 200)
        pY(num) = My + Int(Rnd * 200)
    Next num
    Timer1.Interval = 1
    Me.BackColor = Bckrndcolor
End Sub

Private Sub Timer1_Timer()
    Dim num, num2 As Integer
    Me.Cls
    Me.BackColor = Bckrndcolor
    For num = 0 To Sides
        If pX(num) >= Screen.Width Then xmov = -50
        If pX(num) <= 0 Then xmov = 50
        If pY(num) >= Screen.Height Then ymov = -50
        If pY(num) <= 0 Then ymov = 50
        num2 = Int(Rnd * 28) + 1
        If num2 = 1 Then
            pXmov(num) = 0: pYmov(num) = 0
        ElseIf num2 = 2 Then
            pXmov(num) = 50: pYmov(num) = 50
        ElseIf num2 = 3 Then
            pXmov(num) = 0: pYmov(num) = 50
        ElseIf num2 = 4 Then
            pXmov(num) = 50: pYmov(num) = 0
        ElseIf num2 = 5 Then
            pXmov(num) = -50: pYmov(num) = -50
        ElseIf num2 = 6 Then
            pXmov(num) = 0: pYmov(num) = -50
        ElseIf num2 = 7 Then
            pXmov(num) = -50: pYmov(num) = 0
        ElseIf num2 = 8 Then
            pXmov(num) = 0: pYmov(num) = 0
        ElseIf num2 = 9 Then
            pXmov(num) = 70: pYmov(num) = 70
        ElseIf num2 = 10 Then
            pXmov(num) = 0: pYmov(num) = 70
        ElseIf num2 = 11 Then
            pXmov(num) = 70: pYmov(num) = 0
        ElseIf num2 = 12 Then
            pXmov(num) = -70: pYmov(num) = -70
        ElseIf num2 = 13 Then
            pXmov(num) = 0: pYmov(num) = -70
        ElseIf num2 = 14 Then
            pXmov(num) = -70: pYmov(num) = 0
        End If
        Mx = Mx + xmov
        My = My + ymov
        pX(num) = (pX(num) + pXmov(num)) + xmov
        pY(num) = (pY(num) + pYmov(num)) + ymov
        If num < Sides Then
            Line (pX(num), pY(num))-(pX(num + 1), pY(num + 1)), PColor
        ElseIf num = Sides Then
            Line (pX(num), pY(num))-(pX(0), pY(0)), PColor
        End If
        If TD = 1 Then Call D3
    Next num
End Sub

Sub D3()
    Dim num3, n As Integer
    For num3 = 0 To Sides
        For n = 0 To 12
            If RandColor = 0 Then
                If n <= Sides Then Line (pX(num3), pY(num3))-(pX(n), pY(n)), PColor
            ElseIf RandColor = 1 Then
                If n <= Sides Then Line (pX(num3), pY(num3))-(pX(n), pY(n)), RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            End If
        Next n
    Next num3
End Sub
