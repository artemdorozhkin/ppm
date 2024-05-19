Attribute VB_Name = "PArray"
'@Folder("PearPMProject.src.Utils")
Option Explicit

Public Function IncludesAny(ByRef Data As Variant, ParamArray OneOfValue() As Variant) As Boolean
    Dim Item As Variant
    For Each Item In Data
        Dim Value As Variant
        For Each Value In OneOfValue
            If Item = Value Then
                IncludesAny = True
                Exit Function
            End If
        Next
    Next
End Function
