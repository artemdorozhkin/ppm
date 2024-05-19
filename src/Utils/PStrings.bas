Attribute VB_Name = "PStrings"
'@Folder("PearPMProject.src.Utils")
Option Explicit

Public Function Contains( _
    ByVal Text As String, _
    ByVal Value As String, _
    Optional ByVal Compare As VbCompareMethod = vbTextCompare _
) As Boolean
    Contains = Strings.InStr(1, Text, Value, Compare) > 0
End Function

Public Function EndsWith( _
    ByVal Text As String, _
    ByVal Value As String, _
    Optional ByVal Compare As VbCompareMethod = vbTextCompare _
) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    EndsWith = Strings.StrComp(Strings.Right(Text, ValueLen), Value, Compare) = 0
End Function

Public Function StartsWith( _
    ByVal Text As String, _
    ByVal Value As String, _
    Optional ByVal Compare As VbCompareMethod = vbTextCompare _
) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    StartsWith = Strings.StrComp(Strings.Left(Text, ValueLen), Value, Compare) = 0
End Function

Public Function IsEqual( _
    ByVal Str1 As String, _
    ByVal Str2 As String, _
    Optional ByVal Compare As VbCompareMethod = vbTextCompare _
) As Boolean
    IsEqual = Strings.StrComp(Str1, Str2, Compare) = 0
End Function

Public Function FString(ByVal Text As String, ParamArray Variables() As Variant) As String
    Dim FormatedText As String: FormatedText = Text

    Dim i As Integer
    For i = LBound(Variables) To UBound(Variables)
        Dim Plug As String: Plug = "{" & i & "}"
        FormatedText = Strings.Replace(FormatedText, Plug, Variables(i))
    Next

    FormatedText = Strings.Replace(FormatedText, "\\n", vbNewLine)
    FormatedText = Strings.Replace(FormatedText, "\\r", vbNewLine)
    FormatedText = Strings.Replace(FormatedText, "\\t", vbTab)

    FString = FormatedText
End Function

Public Function IsWhiteSpace(ByVal Char As String) As Boolean
    IsWhiteSpace = Char = " "
End Function
