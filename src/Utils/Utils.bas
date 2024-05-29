Attribute VB_Name = "Utils"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Function UncommentString(ByVal Value As String) As String
    Dim Lines As Variant: Lines = Strings.Split(Value, vbNewLine)
    Dim i As Long
    For i = 0 To UBound(Lines)
        Dim j As Long: j = 1
        Dim Char As String: Char = Strings.Mid(Lines(i), j, 1)
        Do While PStrings.IsWhiteSpace(Char)
            j = j + 1
            If j > Strings.Len(Lines(i)) Then Exit Do
            Char = Strings.Mid(Lines(i), j, 1)
        Loop
        If j > Strings.Len(Lines(i)) Then
            Lines(i) = Lines(i)
        Else
            Lines(i) = Strings.Right(Lines(i), Strings.Len(Lines(i)) - j)
        End If
    Next
    UncommentString = Strings.Join(Lines, vbNewLine)
End Function

Public Function CommentString(ByVal Value As String) As String
    Dim Lines As Variant: Lines = Strings.Split(Value, vbNewLine)
    Dim i As Long
    For i = 0 To UBound(Lines)
        Lines(i) = "'" & Lines(i)
    Next
    CommentString = Strings.Join(Lines, vbNewLine)
End Function

Public Function IsDictionary(ByVal Value As Object) As Boolean
    IsDictionary = Information.TypeName(Value) = "Dictionary"
End Function

Public Function IsCollection(ByVal Value As Object) As Boolean
    IsCollection = Information.TypeName(Value) = "Collection"
End Function

Public Function ConvertToType(ByVal Value As Variant, ByVal DataType As VbVarType) As Variant
    If Information.VarType(Value) = DataType Then
        ConvertToType = Value
    Else
        Select Case DataType
            Case VbVarType.vbString: ConvertToType = Conversion.CStr(Value)
            Case VbVarType.vbBoolean: ConvertToType = Conversion.CBool(Value)
            Case Else: Information.Err.Raise 9, "ConvertToType", "Type not defined: " & DataType
        End Select
    End If
End Function

Public Sub CreateCstr(ByVal Name As String)
    Dim Project As Project: Set Project = NewProject(Application.VBE.ActiveVBProject)
    If Project.IsModuleExists(Name) Then
        Name = Project.GetModule(Name).Name
    End If

    Dim Folder As String
    Folder = PStrings.FString("'@Folder ""{0}.{1}""", Project.Name, Name)

    Dim CstrCode As String
    CstrCode = PStrings.FString( _
        "Public Function New{0}() As {0}\\n" & _
        "\\tSet New{0} = New {0}\\n" & _
        "End Function", _
        Name _
    )
    With Project.AddModule(PStrings.FString("{0}Cstr", Name)).CodeModule
        .InsertLines 1, Folder
        .AddFromString CstrCode
    End With
End Sub

Public Function ConvertTime(ByVal Value As Double) As String
    Dim s As Double: s = 1000
    Dim m As Double: m = s * 60
    Dim h As Double: h = m * 60
    Dim d As Double: d = h * 24
    Dim w As Double: w = d * 7
    Dim y As Double: y = d * 365.25

    Value = Math.Round(Math.Abs(Value * s))
    If Value >= d Then
        ConvertTime = PStrings.FString("{0}d", Math.Round(Value / d))
    ElseIf Value >= h Then
        ConvertTime = PStrings.FString("{0}h", Math.Round(Value / h))
    ElseIf Value >= m Then
        ConvertTime = PStrings.FString("{0}m", Math.Round(Value / m))
    ElseIf Value >= s Then
        ConvertTime = PStrings.FString("{0}s", Math.Round(Value / s))
    Else
        ConvertTime = PStrings.FString("{0}ms", Value)
    End If
End Function

Public Function IsTrue(ByVal Value As Variant) As Boolean
    IsTrue = Not IsFalse(Value)
End Function

Public Function IsFalse(ByVal Value As Variant) As Boolean
    Dim ValueType As VbVarType: ValueType = Information.VarType(Value)

    If (ValueType And vbArray) = vbArray Then
        If Information.IsObject(Value) Then
            IsFalse = Value Is Nothing
        ElseIf Information.IsArray(Value) Then
            On Error Resume Next
            Dim IsEmptyArray As Boolean
            IsEmptyArray = UBound(Value) = -1
            IsFalse = IsEmptyArray Or Information.Err.Number > 0
        End If
    ElseIf (ValueType And vbObject) = vbObject Then
        IsFalse = Value Is Nothing
    ElseIf (ValueType And vbString) = vbString Then
        IsFalse = Strings.Len(Value) = 0
    ElseIf (ValueType And vbBoolean) = vbBoolean Then
        IsFalse = Value
    ElseIf (ValueType And vbVariant) = vbVariant Then
        IsFalse = Information.IsEmpty(Value)
    ElseIf (ValueType And vbByte) = vbByte Or _
           (ValueType And vbCurrency) = vbCurrency Or _
           (ValueType And vbDecimal) = vbDecimal Or _
           (ValueType And vbInteger) = vbInteger Or _
           (ValueType And vbLong) = vbLong Or _
           (ValueType And vbLongLong) = vbLongLong Then
        IsFalse = Value = 0
    ElseIf (ValueType And vbNull) = vbNull Then
        IsFalse = True
    Else
        Information.Err.Raise 13, "IsFalse", "Cannot detect the type of Value"
    End If
End Function


