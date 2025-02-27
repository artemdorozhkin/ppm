Attribute VB_Name = "Utils"
'@Folder "PearPMProject.src.Utils"
Option Explicit

#If DEV Then
  Public Function ChangeDictionaryCompareMode( _
      ByRef Dictionary As Dictionary, _
      ByVal CompareMethod As VbCompareMethod _
  ) As Dictionary
#Else
  Public Function ChangeDictionaryCompareMode( _
      ByRef Dictionary As Object, _
      ByVal CompareMethod As VbCompareMethod _
  ) As Object
#End If
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary()
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary()
  #End If
    Buffer.CompareMode = CompareMethod

    Dim Key As Variant
    For Each Key In Dictionary
        Buffer.Add Key, Dictionary(Key)
    Next

    Set ChangeDictionaryCompareMode = Buffer
End Function

' R3uK: https://stackoverflow.com/a/39912842/21597893
Public Function SizeInStr(ByVal Size_Bytes As Double) As String
    Dim TS()
    ReDim TS(4)
    TS(0) = "b"
    TS(1) = "kb"
    TS(2) = "Mb"
    TS(3) = "Gb"
    TS(4) = "Tb"

    Dim Size_Counter As Integer
    Size_Counter = 0

    If Size_Bytes <= 1 Then
        Size_Counter = 1
    Else
        While Size_Bytes > 1
            Size_Bytes = Size_Bytes / 1000
            Size_Counter = Size_Counter + 1
        Wend
    End If

    SizeInStr = Strings.Format(Size_Bytes * 1000, "##0.0#") & " " & TS(Size_Counter - 1)
End Function

Public Function CalculateFileCheckSum(ByVal Path As String) As String
    Dim Converter As BinaryConverter: Set Converter = New BinaryConverter
    CalculateFileCheckSum = CalculateBytesCheckSum(Converter.FileToBytes(Path))
End Function

Public Function CalculateBytesCheckSum(ByRef Bytes() As Byte) As String
    With CreateObject("MSXML2.DOMDocument")
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        Dim SHA256 As Object
        Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
        .DocumentElement.nodeTypedValue = SHA256.ComputeHash_2((Bytes))
        CalculateBytesCheckSum = Strings.Replace(.DocumentElement.Text, vbLf, "")
    End With
End Function

Public Function GetNewestVersion(ByVal Current As String, ByVal Other As String) As String
    Const Major As Integer = 0
    Const Minor As Integer = 1
    Const Patch As Integer = 2

    If Strings.Len(Other) = 0 Then
        GetNewestVersion = Current
        Exit Function
    End If

    If Strings.Len(Current) = 0 Then
        GetNewestVersion = Other
        Exit Function
    End If

    If PStrings.IsEqual(Current, "latest") Then
        GetNewestVersion = Current
        Exit Function
    End If

    Dim CurrentParts As Variant: CurrentParts = Strings.Split(Current, ".")
    Dim OtherParts As Variant: OtherParts = Strings.Split(Other, ".")
    CurrentParts(Major) = Conversion.CInt(CurrentParts(Major))
    CurrentParts(Minor) = Conversion.CInt(CurrentParts(Minor))
    CurrentParts(Patch) = Conversion.CInt(CurrentParts(Patch))
    OtherParts(Major) = Conversion.CInt(OtherParts(Major))
    OtherParts(Minor) = Conversion.CInt(OtherParts(Minor))
    OtherParts(Patch) = Conversion.CInt(OtherParts(Patch))

    If CurrentParts(Major) > OtherParts(Major) Then
        GetNewestVersion = Current
    ElseIf CurrentParts(Major) < OtherParts(Major) Then
        GetNewestVersion = Other
    ElseIf CurrentParts(Minor) > OtherParts(Minor) Then
        GetNewestVersion = Current
    ElseIf CurrentParts(Minor) < OtherParts(Minor) Then
        GetNewestVersion = Other
    ElseIf CurrentParts(Patch) > OtherParts(Patch) Then
        GetNewestVersion = Current
    ElseIf CurrentParts(Patch) < OtherParts(Patch) Then
        GetNewestVersion = Other
    Else
        GetNewestVersion = Current
    End If
End Function

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

Public Function IsVBComponent(ByVal Value As Object) As Boolean
    IsVBComponent = Information.TypeName(Value) = "VBComponent"
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
            Case VbVarType.vbBoolean: ConvertToType = IsTrue(Value)
            Case Else: Information.Err.Raise 9, "ConvertToType", "Type not defined: " & DataType
        End Select
    End If
End Function

Public Function ConvertTime(ByVal Value As Double) As String
    Dim s As Double: s = 1000
    Dim m As Double: m = s * 60
    Dim h As Double: h = m * 60
    Dim d As Double: d = h * 24

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
    ElseIf ValueType = vbObject Then
        IsFalse = Value Is Nothing
    ElseIf ValueType = vbString Then
        IsFalse = Strings.Len(Value) = 0
    ElseIf ValueType = vbBoolean Then
        IsFalse = Not Value
    ElseIf ValueType = vbVariant Then
        IsFalse = Information.IsEmpty(Value)
    ElseIf ValueType = vbByte Or _
           ValueType = vbCurrency Or _
           ValueType = vbDecimal Or _
           ValueType = vbInteger Or _
           ValueType = vbLong Or _
           ValueType = vbLongLong Then
        IsFalse = Value = 0
    ElseIf ValueType = vbNull Then
        IsFalse = True
    ElseIf ValueType = vbEmpty Then
        IsFalse = True
    Else
        Information.Err.Raise 13, "IsFalse", "Cannot detect the type of Value"
    End If
End Function

Public Function GetFirstTrue(ParamArray Values() As Variant) As Variant
    Dim Value As Variant
    For Each Value In Values
        If IsTrue(Value) Then
            If Information.IsObject(Value) Then
                Set GetFirstTrue = Value
            Else
                GetFirstTrue = Value
            End If
            Exit Function
        End If
    Next

    GetFirstTrue = Null
End Function

Public Function GetFirstValueFrom(ByVal DefName As String, ByRef Tokens As Tokens, ByRef Config As Config) As Variant
    If Not Definitions.Items().Exists(DefName) Then
        Err.Raise 9, "Utils", "Can't find definition: " & DefName
    End If

    Dim Def As Definition
    Set Def = Definitions(DefName)

    If Tokens.IncludeDefinition(Def) Then
        Dim Token As SyntaxToken
        Set Token = Tokens.GetTokenByDefinition(Def)
        If IsFalse(Token) Then
            GetFirstValueFrom = True
        Else
            GetFirstValueFrom = ConvertToType(Token.Text, Def.DataType)
        End If
    Else
        Dim Value As Variant
        Value = GetFirstTrue( _
            Config.GetValue(DefName), _
            Def.Default _
        )
        If Not Information.IsNull(Value) Then
            GetFirstValueFrom = ConvertToType(Value, Def.DataType)
        Else
            GetFirstValueFrom = ConvertToType(Def.Default, Def.DataType)
        End If
    End If
End Function
