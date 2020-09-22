Attribute VB_Name = "Module1"
Option Explicit
Public Sides As Single
Public PColor, Bckrndcolor As Long
Public TD, RandColor As Integer
Public Const SetName = "Morphagon.bri"

Sub Main()
    Call Parameters
    Call ReadSettings
End Sub

Public Sub ReadSettings()
'Reads settings from file
    Dim Free As Byte
    Dim ChkFile As String
    Dim Temp As String
    ChkFile = App.Path & "/" & SetName
    If FileExists(ChkFile) And SetName <> "" Then
        Free = FreeFile
        Open ChkFile For Input As Free
            Input #Free, Sides
            Input #Free, PColor
            Input #Free, Bckrndcolor
            Input #Free, TD
            Input #Free, RandColor
        Close #Free
    Else
        Call SetDefaults
    End If
End Sub

Sub SetDefaults()
'Gives default settings if settings file not found
    Sides = 2
    PColor = RGB(245, 224, 250)
    Bckrndcolor = RGB(0, 0, 0)
    TD = 1
    RandColor = 0
End Sub

Public Function FileExists(FileName As String) As Boolean
'Checks if a file exists.
    If FileName <> "" Then
        Dim CheckThis As String
        On Error Resume Next
        CheckThis = Dir(FileName)
        If CheckThis = "" Then
            FileExists = False
        Else
            FileExists = True
        End If
    End If
End Function

Sub Parameters()
    Dim Param As String
    Param = Left(Command, 2)
    If Param <> "" Then
        Select Case Param
            Case "/p"
                End
            Case "/c"
                Settings.Show
            Case "/s"
                Scrn.Show
        End Select
    Else
        Scrn.Show
    End If
End Sub
