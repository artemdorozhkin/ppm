Attribute VB_Name = "PCollection"
'@Folder("PearPMProject.src.Utils")
Option Explicit

Public Function ToArray(ByRef Collection As Collection) As Variant
    If Collection.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    Dim Arr() As Variant: ReDim Arr(0 To Collection.Count - 1)

    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If IsObject(Collection.Item(i + 1)) Then
            Set Arr(i) = Collection.Item(i + 1)
        Else
            Arr(i) = Collection.Item(i + 1)
        End If
    Next

    ToArray = Arr
End Function
