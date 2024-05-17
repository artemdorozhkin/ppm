VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.ConfigIO"
Option Explicit

Private Enum ConfigIOMode
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Private Type TConfigReader
    ConfigPath As String
End Type

Private this As TConfigReader

Public Property Let ConfigPath(ByVal RHS As String)
    this.ConfigPath = RHS
End Property

Public Function ReadSection(ByVal Name As String) As Object
    Dim Lines As Variant: Lines = ReadLines()
    If Information.IsEmpty(Lines) Then Lines = Array()
    Dim Section As Object: Set Section = NewDictionary(vbTextCompare)

    Dim i As Long
    For i = 0 To UBound(Lines)
        Dim Line As Variant: Line = Lines(i)
        If Not IsEqual(Line, FString("[{0}]", Name)) Then GoTo Continue
        i = i + 1
        Line = Lines(i)
        Do While Not StartsWith(Line, "[")
            Dim KeyValuePair As Variant
            KeyValuePair = Strings.Split(Line, "=")
            Section(Strings.Trim(KeyValuePair(0))) = Strings.Trim(KeyValuePair(1))
            i = i + 1
            If i >= UBound(Lines) Then Exit Do
        Loop
        Exit For
Continue:
    Next

    Set ReadSection = Section
End Function

Public Sub WriteData(ByVal SectionName As String, ByVal Key As String, ByVal Value As String)
    Dim Lines As Variant: Lines = ReadLines()
    If Information.IsEmpty(Lines) Then Lines = Array()

    Dim i As Long
    For i = 0 To UBound(Lines)
        Dim Line As Variant: Line = Lines(i)
        If Not IsEqual(Line, FString("[{0}]", SectionName)) Then GoTo Continue
        i = i + 1
        Line = Lines(i)
        Do While Not StartsWith(Line, "[")
            If Strings.Len(Strings.Replace(Line, " ", "")) > 0 Then
                Dim KeyValuePair As Variant
                KeyValuePair = Strings.Split(Line, "=")
                If IsEqual(KeyValuePair(0), Key) Then
                    Lines(i) = FString("{0}={1}", Key, Value)
                    WriteLines Lines
                    Exit Sub
                End If
            End If
            i = i + 1
            If i >= UBound(Lines) Then Exit Do
            Line = Lines(i)
        Loop
        If i > UBound(Lines) Then ReDim Preserve Lines(UBound(Lines) + 1)
        Lines(i) = FString("{0}={1}", Key, Value)
        WriteLines Lines
        Exit Sub
Continue:
    Next

    ReDim Preserve Lines(i + 2)
    If i > 0 Then i = i + 1
    Lines(i) = FString("[{0}]", SectionName)
    Lines(i + 1) = FString("{0}={1}", Key, Value)
    WriteLines Lines
End Sub
    
Private Function ReadLines() As Variant
    If Strings.Len(this.ConfigPath) = 0 Then Err.Raise 76, TypeName(Me)

    With NewFileSystemObject()
        If Not .FileExists(this.ConfigPath) Then .CreateTextFile this.ConfigPath
        With .GetFile(this.ConfigPath).OpenAsTextStream(ConfigIOMode.ForReading)
            If Not .AtEndOfLine Then ReadLines = Strings.Split(.ReadAll(), vbNewLine)
            .Close
        End With
    End With
End Function

Private Sub WriteLines(ByRef Lines As Variant)
    If Strings.Len(this.ConfigPath) = 0 Then Err.Raise 76, TypeName(Me)

    With NewFileSystemObject()
        If Not .FileExists(this.ConfigPath) Then .CreateTextFile this.ConfigPath
        With .GetFile(this.ConfigPath).OpenAsTextStream(ConfigIOMode.ForWriting)
            .Write Strings.Join(Lines, vbNewLine)
            .Close
        End With
    End With
End Sub
