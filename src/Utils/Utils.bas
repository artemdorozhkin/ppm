Attribute VB_Name = "Utils"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Sub CreateFoldersRecoursive(ByVal Path As String)
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim Parts As Variant: Parts = Strings.Split(Path, Application.PathSeparator)
    Dim Part As Variant
    For Each Part In Parts
        Dim Current As String
        If Strings.Len(Current) = 0 Then
            Current = Part
        Else
            Current = FString("{0}{1}{2}", Current, Application.PathSeparator, Part)
        End If

        If Not FSO.FolderExists(Current) Then
            FSO.CreateFolder Current
        End If
    Next
End Sub

Public Function GetProjectPath() As String
    Dim PPMProjectsPath As String: PPMProjectsPath = GetPPMProjectsPath()
    With NewFileSystemObject()
        Dim ProjectName As String
        ProjectName = .GetFileName(SelectedProject.Path)

        Dim ProjectTimeStamp As String
        ProjectTimeStamp = Strings.Format(.GetFile(SelectedProject.Path).DateCreated, "ddmmyyyy_hhnnss")

        Dim FolderName As String
        FolderName = FString("{0}_{1}", ProjectName, ProjectTimeStamp)

        Dim ThisProjectPath As String
        ThisProjectPath = .BuildPath(PPMProjectsPath, FolderName)
        If Not .FolderExists(ThisProjectPath) Then FileSystem.MkDir ThisProjectPath
    End With

    GetProjectPath = ThisProjectPath
End Function

Public Function GetPPMProjectsPath() As String
    With NewFileSystemObject()
        Dim PPMPath As String
        PPMPath = .BuildPath(Interaction.Environ("LOCALAPPDATA"), "ppm")
        Dim PPMProjectsPath As String
        PPMProjectsPath = .BuildPath(PPMPath, "projects")
        If Not .FolderExists(PPMProjectsPath) Then CreateFoldersRecoursive PPMProjectsPath
        GetPPMProjectsPath = PPMProjectsPath
    End With
End Function

Public Sub AddArrayToCollection(ByRef SourceArray As Variant, ByRef Output As Collection)
    If Not Information.IsArray(SourceArray) Then Exit Sub
    If IsFalse(SourceArray) Then Exit Sub

    Dim Item As Variant
    For Each Item In SourceArray
        Output.Add Item
    Next
End Sub

Public Function Contains(ByVal Text As String, ByVal Value As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Contains = Strings.InStr(1, Text, Value, Compare) > 0
End Function

Public Function EndsWith(ByVal Text As String, ByVal Value As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    EndsWith = Strings.StrComp(Strings.Right(Text, ValueLen), Value, Compare) = 0
End Function

Public Function StartsWith(ByVal Text As String, ByVal Value As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    StartsWith = Strings.StrComp(Strings.Left(Text, ValueLen), Value, Compare) = 0
End Function

Public Function EndsWithLike(ByVal Text As String, ByVal Value As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    EndsWithLike = Strings.Right(Text, ValueLen) Like Value
End Function

Public Function StartsWithLike(ByVal Text As String, ByVal Value As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim ValueLen As Long: ValueLen = Strings.Len(Value)
    StartsWithLike = Strings.Left(Text, ValueLen) Like Value
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

Public Function IsEqual(ByVal Str1 As String, ByVal Str2 As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    IsEqual = Strings.StrComp(Str1, Str2, Compare) = 0
End Function

Public Function ArrayIncludes(ByRef Data As Variant, ParamArray OneOfValue() As Variant) As Boolean
    Dim Item As Variant
    For Each Item In Data
        Dim Value As Variant
        For Each Value In OneOfValue
            If Item = Value Then
                ArrayIncludes = True
                Exit Function
            End If
        Next
    Next
End Function

Public Function ArgsToOptions(ByRef Args As Variant) As Variant
    Dim Options As Object: Set Options = NewDictionary(VbCompareMethod.vbTextCompare)

    Dim i As Long
    For i = 1 To UBound(Args)
        Dim OptionName As String
        Dim Value As Variant: Value = Null
        If Strings.Left(Args(i), 2) = "--" Then
            OptionName = Strings.Mid(Args(i), 3, Strings.Len(Args(i)) - 2)
            Options(OptionName) = Value
        ElseIf Strings.Left(Args(i), 1) = "-" Then
            OptionName = Strings.Mid(Args(i), 2, Strings.Len(Args(i)) - 1)
            If Strings.Len(OptionName) > 1 Then
                Value = Strings.Mid(OptionName, 2, Strings.Len(OptionName) - 1)
                OptionName = Strings.Left(OptionName, 1)
            End If
            Options(OptionName) = Value
        End If
    Next

    Set ArgsToOptions = Options
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

    If InStr(1, FormatedText, "\\t", vbTextCompare) = 0 Then
        FString = FormatedText
        Exit Function
    End If

    Dim Lines As Variant: Lines = Strings.Split(FormatedText, vbNewLine)
    Dim Index As Long
    For Index = 0 To UBound(Lines)
        Dim Line As String: Line = Lines(Index)
        i = 1
        Do While i <= Strings.Len(Line)
            If Strings.Mid(Line, i, 1) = "\" And _
               Strings.Mid(Line, i + 1, 1) = "\" And _
               Strings.Mid(Line, i + 2, 1) = "t" Then
                Line = Strings.Replace(Line, "\\t", Strings.Space(4 - ((i - 1) Mod 4)), Count:=1)
            End If

            Lines(Index) = Line
            i = i + 1
            If i > Strings.Len(Line) Then Exit Do
        Loop
    Next

    FString = Strings.Join(Lines, vbNewLine)
End Function

Public Function IsWhiteSpace(ByVal Char As String) As Boolean
    IsWhiteSpace = Char = " "
End Function

Public Function ToArray(ByRef Collectable As Collection) As Variant
    If Collectable.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    Dim Arr() As Variant: ReDim Arr(0 To Collectable.Count - 1)

    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If IsObject(Collectable.Item(i + 1)) Then
            Set Arr(i) = Collectable.Item(i + 1)
        Else
            Arr(i) = Collectable.Item(i + 1)
        End If
    Next

    ToArray = Arr
End Function

Public Function GetFileNameWithoutExt(ByVal Path As String) As String
    Dim FileName As String: FileName = GetFileName(Path)
    Dim Ext As String: Ext = GetFileExt(Path)
    GetFileNameWithoutExt = Strings.Left(FileName, Strings.Len(FileName) - Strings.Len(Ext))
End Function

Public Function GetFileName(ByVal Path As String) As String
    GetFileName = Strings.Mid(Path, Strings.InStrRev(Path, Application.PathSeparator) + 1)
End Function

Public Function GetFileExt(ByVal Path As String) As String
    Dim FileName As String: FileName = GetFileName(Path)
    Dim DotPosition As Long: DotPosition = Strings.InStrRev(FileName, ".")
    If DotPosition = 0 Then Exit Function
    GetFileExt = Strings.Mid(FileName, DotPosition)
End Function

Public Function NewDictionary(Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = Compare
End Function

Public Function NewFileSystemObject() As FileSystemObject
    Set NewFileSystemObject = CreateObject("Scripting.FileSystemObject")
End Function

Public Function NewFolder(ByVal Path As String) As Folder
    With NewFileSystemObject()
        If (FileSystem.GetAttr(Path) And vbDirectory) = vbDirectory Then
            Set NewFolder = .GetFolder(Path)
        ElseIf .FileExists(Path) Then
            Set NewFolder = .GetFile(Path).ParentFolder
        End If
    End With
End Function
