Attribute VB_Name = "ArgsParser"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Function Parse(ByVal StringArgs As String) As Variant
    Dim i As Long: i = 1
    Dim Char As String
    Char = Strings.Mid(StringArgs, i, 1)

    Dim Args As Collection: Set Args = New Collection

    Do While i <= Strings.Len(StringArgs)
        Do While IsWhiteSpace(Char)
            i = i + 1
            If i > Strings.Len(StringArgs) Then Exit Do
            Char = Strings.Mid(StringArgs, i, 1)
        Loop

        Dim Arg As String: Arg = ""
        Do While Not IsWhiteSpace(Char)
            Arg = Arg & Char

            i = i + 1
            If i > Strings.Len(StringArgs) Then Exit Do
            Char = Strings.Mid(StringArgs, i, 1)
        Loop
        Args.Add Arg
    Loop

    Parse = ToArray(Args)
End Function
