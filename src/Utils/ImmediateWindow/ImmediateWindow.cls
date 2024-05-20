VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediateWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Utils.ImmediateWindow"
Option Explicit

Private Type TImmediateWindow
    Func As String
End Type

Private this As TImmediateWindow

Public Sub WriteLine(ParamArray Args() As Variant)
    Dim Text As String: Text = Strings.Join(Args, " ")

    If Strings.InStr(1, Text, "\\t", vbTextCompare) = 0 Then
        Debug.Print Text
        Exit Sub
    End If

    Dim Lines As Variant: Lines = Strings.Split(Text, vbNewLine)
    Dim Index As Long
    For Index = 0 To UBound(Lines)
        Dim Line As String: Line = Lines(Index)
        Dim i As Long: i = 1
        Do While i <= Strings.Len(Line)
            If Strings.Mid(Line, i, 1) = "\" And _
               Strings.Mid(Line, i + 1, 1) = "\" And _
               Strings.Mid(Line, i + 2, 1) = "t" Then
                Line = Strings.Replace(Line, "\\t", Strings.Space(4 - ((i - 1) Mod 4)), Count:=1)
            End If

            Lines(Index) = Line
            i = i + 1
            If i > Strings.Len(Line) Then Exit Do
        Loop
    Next

    Debug.Print Strings.Join(Lines, vbNewLine)
End Sub

Public Sub ReadLine(ByVal Text As String, Optional ByVal NextFunction As String)
    If Strings.Len(NextFunction) > 0 Then NextFunction = NextFunction & " _"
    Debug.Print Text & vbNewLine & Space(100) & NextFunction
End Sub
