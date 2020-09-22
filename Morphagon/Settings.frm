VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Morphagon Settings"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      Caption         =   "&Reset Defaults"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "-"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "-"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Random Color"
      Height          =   255
      Left            =   1440
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "3-D"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Number of sides: (3-13)"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3840
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "B"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "G"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "R"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Morphagon color:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3840
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "B"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "G"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "R"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back ground color:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private r, r1, g, g1, b, b1 As Integer

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    PColor = RGB(r, g, b)
    Bckrndcolor = RGB(r1, g1, b1)
    If Text1.Text = "" Then
        Sides = 2
    Else: Sides = Text1.Text - 1
    End If
    If Sides >= 13 Then Sides = 12
    If Sides <= 3 Then Sides = 2
    Dim Free As Byte
    Free = FreeFile
    Open App.Path & "/" & SetName For Output As Free
        Print #Free, Sides
        Print #Free, PColor
        Print #Free, Bckrndcolor
        Print #Free, TD
        Print #Free, RandColor
    Close #Free
End Sub

Private Sub Command4_Click()
    r1 = r1 + 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command5_Click()
    g1 = g1 + 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command6_Click()
    b1 = b1 + 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command7_Click()
    If r1 <= 2 Then r1 = 2
    r1 = r1 - 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command8_Click()
    If g1 <= 2 Then g1 = 2
    g1 = g1 - 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command9_Click()
    If b1 <= 2 Then b1 = 2
    b1 = b1 - 2
    Shape1.FillColor = RGB(r1, g1, b1)
End Sub

Private Sub Command10_Click()
    r = r + 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command11_Click()
    g = g + 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command12_Click()
    b = b + 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command13_Click()
    If r <= 2 Then r = 2
    r = r - 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command14_Click()
    If g <= 2 Then g = 2
    g = g - 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command15_Click()
    If b <= 2 Then b = 2
    b = b - 2
    Shape2.FillColor = RGB(r, g, b)
End Sub

Private Sub Command16_Click()
    Sides = 2
    Bckrndcolor = RGB(0, 0, 0)
    PColor = RGB(245, 250, 224)
    TD = False
    RandColor = 0
End Sub

Private Sub Form_Load()
    r = 0
    r1 = 0
    g = 0
    g1 = 0
    b = 0
    b1 = 0
    RandColor = 0
    TD = 0
End Sub

Private Sub Label10_Click()
    If TD = 1 Then TD = 0
    If TD = 0 Then TD = 1
    If TD = 1 Then Label10.BorderStyle = 1
    If TD = 0 Then Label10.BorderStyle = 0
End Sub

Private Sub Label11_Click()
    If RandColor = 1 Then RandColor = 0
    If RandColor = 0 Then RandColor = 1
    If RandColor = 1 Then Label11.BorderStyle = 1
    If RandColor = 0 Then Label11.BorderStyle = 0
End Sub
