Attribute VB_Name = "PCollection"
'@Folder "PearPMProject.src.Utils"
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

Public Function KeyExists(ByRef Collection As Collection, ByVal Key As Variant) As Boolean
    On Error Resume Next
    Dim Dummy As Boolean
    Dummy = Information.IsObject(Collection(Key))

    KeyExists = Information.Err.Number <> 0
End Function

Public Function ItemExists(ByRef Collection As Collection, ByVal Item As Variant) As Boolean
    Dim Check As Variant
    For Each Check In Collection
        Dim IsSame As Boolean
        If Information.IsObject(Item) Then
            IsSame = Item Is Check
        Else
            IsSame = Item = Check
        End If

        If IsSame Then
            ItemExists = True
            Exit Function
        End If
    Next
End Function
