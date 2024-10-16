VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Immediate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PearPMProject.src.Utils"
Option Explicit

Const TAB_SIZE As Integer = 4

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
                Line = Strings.Replace(Line, "\\t", Strings.Space(TAB_SIZE - ((i - 1) Mod TAB_SIZE)), Count:=1)
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

Public Sub WriteTable(ByRef Data As Variant, ByRef Header As Variant, Optional ByVal MaxWidth As Long = -1)
    Const COL_DELIMITER As String = " | "
    Dim Widths As Collection: Set Widths = GetWidths(Data, Header, MaxWidth)

    Dim i As Long
    For i = LBound(Header) To UBound(Header)
        Dim Width As Long: Width = Widths(Conversion.CStr(i))
        Dim SpaceCount As Long: SpaceCount = Width - IIf(MaxWidth > Strings.Len(Header(i)), Strings.Len(Header(i)), MaxWidth)
        Dim HeaderText As String
        If MaxWidth = -1 Or MaxWidth >= Strings.Len(Header(i)) Then
            HeaderText = Header(i)
        Else
            HeaderText = Strings.Left(Header(i), MaxWidth - 3) & "..."
        End If
        Dim Text As String: Text = Text & HeaderText & Strings.Space(SpaceCount) & COL_DELIMITER
    Next

    Debug.Print Text
    Text = Empty

    Dim j As Long
    For j = LBound(Data, 1) To UBound(Data, 1)
        For i = LBound(Header) To UBound(Header)
            Width = Widths(Conversion.CStr(i))
            SpaceCount = Width - IIf(MaxWidth > Strings.Len(Data(j, i)), Strings.Len(Data(j, i)), MaxWidth)
            Dim DataText As String
            If MaxWidth = -1 Or MaxWidth >= Strings.Len(Data(j, i)) Then
                DataText = Data(j, i)
            Else
                DataText = Strings.Left(Data(j, i), MaxWidth - 3) & "..."
            End If
            Text = Text & DataText & Strings.Space(SpaceCount) & COL_DELIMITER
        Next

        Debug.Print Text
        Text = Empty
    Next
End Sub

Private Function GetWidths(ByRef Data As Variant, ByRef Header As Variant, ByVal MaxWidth As Long) As Collection
    Dim Widths As Collection: Set Widths = New Collection

    Dim i As Long
    For i = LBound(Header) To UBound(Header)
        Dim MaxLen As Long: MaxLen = Strings.Len(Header(i))
        Dim j As Long
        For j = LBound(Data, 1) To UBound(Data, 1)
            If Strings.Len(Data(j, i)) > MaxLen Then MaxLen = Strings.Len(Data(j, i))
        Next
        If MaxLen > MaxWidth And MaxWidth <> -1 Then MaxLen = MaxWidth
        Widths.Add Key:=Conversion.CStr(i), Item:=MaxLen
    Next

    Set GetWidths = Widths
End Function
