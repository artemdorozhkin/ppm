Attribute VB_Name = "Config"
'@Folder("PearPMProject.src.Config")
Option Explicit

Private Type TConfig
    Data As Object
End Type

Private this As TConfig

Public Function GetText() As String
    Dim Content() As String
    ReDim Content(this.Data.Count - 1)

    Dim i As Long
    Dim Key As Variant
    For Each Key In this.Data
        Dim Pair As String
        Pair = PStrings.FString("{0}={1}", Key, this.Data(Key))
        Content(i) = Pair
        i = i + 1
    Next

    GetText = Strings.Join(Content, Constants.vbNewLine)
End Function

Public Sub ReadFileByScope(ByVal Scope As ConfigScopes)
    If Scope = ConfigScopes.DefaultScope Then Exit Sub

    If Scope = ConfigScopes.GlobalScode Then
        Set this.Data = ConfigParser.Read(ConfigPaths.GetGlobalConfigPath())
    ElseIf Scope = ConfigScopes.UserScope Then
        Set this.Data = ConfigParser.Read(ConfigPaths.GetUserConfigPath())
    ElseIf Scope = ConfigScopes.ProjectScope Then
        Dim FilePath As String
        FilePath = ConfigPaths.GetProjectConfigPath()
        If Strings.Len(FilePath) = 0 Then Exit Sub
        Set this.Data = ConfigParser.Read(ConfigPaths.GetProjectConfigPath())
    End If
End Sub

Public Sub ReadScope(Optional ByVal Scope As ConfigScopes = ConfigScopes.ProjectScope)
    Set this.Data = ConfigParser.ReadDefinitions()
    If Scope = ConfigScopes.DefaultScope Then Exit Sub

    If Scope >= ConfigScopes.GlobalScode Then
        ConfigParser.Read ConfigPaths.GetGlobalConfigPath(), this.Data
    End If

    If Scope >= ConfigScopes.UserScope Then
        ConfigParser.Read ConfigPaths.GetUserConfigPath(), this.Data
    End If

    If Scope >= ConfigScopes.ProjectScope Then
        Dim FilePath As String
        FilePath = ConfigPaths.GetProjectConfigPath()
        If Strings.Len(FilePath) = 0 Then Exit Sub
        ConfigParser.Read ConfigPaths.GetProjectConfigPath(), this.Data
    End If
End Sub

Public Function GetValue(ByVal Key As String, Optional ByVal CastTo As VbVarType = VbVarType.vbString) As Variant
    GetValue = Utils.ConvertToType(this.Data(Key), CastTo)
End Function

Public Sub DeleteKey(ByVal Key As String)
    this.Data.Remove Key
End Sub

Public Sub SetValue(ByVal Key As String, ByVal Value As String)
    this.Data(Key) = Value
End Sub

Public Sub Save(ByVal Scope As ConfigScopes, Optional ByRef Data As Object)
    Dim FilePath As String
    If Scope = ConfigScopes.DefaultScope Then
        FilePath = ConfigPaths.GetGlobalConfigPath()
    ElseIf Scope = ConfigScopes.GlobalScode Then
        FilePath = ConfigPaths.GetGlobalConfigPath()
    ElseIf Scope = ConfigScopes.UserScope Then
        FilePath = ConfigPaths.GetUserConfigPath()
    ElseIf Scope = ConfigScopes.ProjectScope Then
        FilePath = ConfigPaths.GetProjectConfigPath()
    Else
        Information.Err.Raise 9, "Config.Save", "Unknown scope: " & Scope
    End If

    If Data Is Nothing Then
        ConfigParser.Save FilePath, this.Data
    Else
        ConfigParser.Save FilePath, Data
    End If
End Sub

Public Function IsMissing(ByVal ForScope As ConfigScopes) As Boolean
    Dim FSO As Object
    Set FSO = GetFileSystemObject()

    Select Case ForScope
        Case ConfigScopes.GlobalScode
            IsMissing = Not FSO.FileExists(ConfigPaths.GetGlobalConfigPath())
        Case ConfigScopes.UserScope
            IsMissing = Not FSO.FileExists(ConfigPaths.GetUserConfigPath())
        Case ConfigScopes.ProjectScope
            IsMissing = Not FSO.FileExists(ConfigPaths.GetProjectConfigPath())
        Case Else
            IsMissing = ForScope <> ConfigScopes.DefaultScope
    End Select
End Function

Public Sub GenerateDefault(ByVal ForScope As ConfigScopes)
    Debug.Assert ForScope <> ConfigScopes.DefaultScope ' Config cannot be generated for DefaultScope
    Config.ReadScope ForScope - 1

    Dim ConfigData As Object
    Set ConfigData = NewDictionary()

    Dim SkipableKeys As Variant
    If ForScope = ConfigScopes.ProjectScope Then
        SkipableKeys = Array( _
            "author-name", _
            "author-url", _
            "save-dev", _
            "help", _
            "location", _
            "name", _
            "path", _
            "version", _
            "yes" _
        )
    Else
        SkipableKeys = Array( _
            "save-dev", _
            "help", _
            "location", _
            "path", _
            "yes" _
        )
    End If

    Dim Definition As Variant
    For Each Definition In Definitions.Items.Items
        If PArray.IncludesAny(SkipableKeys, Definition.Key) Then GoTo Continue
        If PStrings.StartsWith(Definition.Key, "_") Then GoTo Continue

        ConfigData(Definition.Key) = Config.GetValue(Definition.Key, Definition.DataType)
Continue:
    Next

    Config.Save ForScope, ConfigData
End Sub
