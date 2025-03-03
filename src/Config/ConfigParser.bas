Attribute VB_Name = "ConfigParser"
'@Folder("PearPMProject.src.Config")
Option Explicit

Public Function ReadDefinitions(Optional ByRef Output As Object) As Object
    Dim Config As Object
    Set Config = NewDictionary()
    If Not Output Is Nothing Then
        Set Config = Output
    End If

    Dim ConfigKeys As Variant
    ConfigKeys = Array( _
        "author-name", _
        "author-url", _
        "create-constructor", _
        "encoding", _
        "language", _
        "local", _
        "name", _
        "no-clear", _
        "registry", _
        "save-struct", _
        "version" _
    )

    Debug.Assert (UBound(Definitions.Items.Items()) = 15) = False

    Dim Definition As Variant
    For Each Definition In Definitions.Items.Items()
        If PArray.IncludesAny(ConfigKeys, Definition.Key) Then
            Dim Value As String
            If Information.IsMissing(Definition.Default) Then
                Value = ""
            Else
                Value = Definition.Default
            End If
            Config(Definition.Key) = Value
        End If
    Next

    If Not Output Is Nothing Then
        Set Output = Config
    End If

    Set ReadDefinitions = Config
End Function

Public Function Read(ByVal Path As String, Optional ByRef Output As Object) As Object
    Dim Config As Object
    Set Config = NewDictionary()
    If Not Output Is Nothing Then
        Set Config = Output
    End If

    Dim Data As Object
    Set Data = ParseData(Path)

    Dim Key As Variant
    For Each Key In Data
        Config(Key) = Conversion.CStr(Data(Key))
    Next

    If Not Output Is Nothing Then
        Set Output = Config
    End If

    Set Read = Config
End Function

Public Sub Save(ByVal Path As String, ByRef Data As Object)
    Dim Content() As String
    ReDim Content(Data.Count)

    Dim Key As Variant
    Dim i As Long
    For Each Key In Data
        Dim Pair As String
        Pair = PStrings.FString("{0}={1}", Key, Data(Key))
        Content(i) = Pair
        i = i + 1
    Next

    PFileSystem.SaveToFile Path, Strings.Join(Content, Constants.vbNewLine)
End Sub

Private Function ParseData(ByVal Path As String) As Object
    Dim Lines As Variant
    Lines = ReadLines(Path)

    If Information.IsEmpty(Lines) Then Lines = Array()
    Dim Data As Object
    Set Data = NewDictionary()

    Dim Line As Variant
    For Each Line In Lines
        If PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), ";") Or _
           PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), "#") Or _
           PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), "") Then
            GoTo Continue
        End If

        Dim KeyValuePair As Variant
        KeyValuePair = Strings.Split(Line, "=")
        Dim Value As String
        If UBound(KeyValuePair) = 0 Then
            Value = ""
        ElseIf KeyValuePair(1) = "" Then
            Value = ""
        Else
            Value = Strings.Split(KeyValuePair(1), "#")(0)
            Value = Strings.Split(Value, ";")(0)
        End If
        Data(Strings.Trim(KeyValuePair(0))) = Strings.Trim(Value)
Continue:
    Next

    Set ParseData = Data
End Function

Private Function ReadLines(ByVal Path As String) As Variant
    If Strings.Len(Path) = 0 Then Information.Err.Raise 76, "ConfigParser.ReadLines"

    With NewFileSystemObject()
        If Not .FileExists(Path) Then
            .CreateTextFile Path
        End If
    End With

    With NewStream()
        .Charset = "utf-8"
        .Open
        .LoadFromFile Path
        ReadLines = Strings.Split(.ReadText(), Constants.vbNewLine)
    End With
End Function
