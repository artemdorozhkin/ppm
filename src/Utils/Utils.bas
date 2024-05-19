Attribute VB_Name = "Utils"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Function ConvertTime(ByVal Value As Double) As String
    Dim s As Double: s = 1000
    Dim m As Double: m = s * 60
    Dim h As Double: h = m * 60
    Dim d As Double: d = h * 24
    Dim w As Double: w = d * 7
    Dim y As Double: y = d * 365.25

    Value = Math.Round(Math.Abs(Value * s))
    If Value >= d Then
        ConvertTime = FString("{0}d", Math.Round(Value / d))
    ElseIf Value >= h Then
        ConvertTime = FString("{0}h", Math.Round(Value / h))
    ElseIf Value >= m Then
        ConvertTime = FString("{0}m", Math.Round(Value / m))
    ElseIf Value >= s Then
        ConvertTime = FString("{0}s", Math.Round(Value / s))
    Else
        ConvertTime = FString("{0}ms", Value)
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
            IsFalse = IsEmptyArray Or Err.Number > 0
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
        Err.Raise 13, "IsFalse", "Cannot detect the type of Value"
    End If
End Function
