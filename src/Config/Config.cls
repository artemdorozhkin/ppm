VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Config"
Option Explicit

Private Type TConfig
  #If DEV Then
    Definitions As Dictionary
  #Else
    Definitions As Object
  #End If
    Tokens As Tokens

    GlobalConfig As ConfigIO
    UserConfig As ConfigIO
    ProjectConfig As ConfigIO
End Type

Private this As TConfig

Private Sub Class_Initialize()
    Set this.GlobalConfig = Configs.GetGlobalConfig()
    Set this.UserConfig = Configs.GetUserConfig()
    Set this.ProjectConfig = Configs.GetProjectConfig()
End Sub

#If DEV Then
  Public Property Set Definitions(ByVal RHS As Dictionary)
#Else
  Public Property Set Definitions(ByVal RHS As Object)
#End If
    Set this.Definitions = RHS
End Property

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Public Function GetValue(ByVal DefinitionName As String) As Variant
    If Not this.Definitions.Exists(DefinitionName) Then
        Err.Raise 9, Information.TypeName(Me), "Can't find definition: " & DefinitionName
    End If

    Dim Definition As Definition: Set Definition = this.Definitions(DefinitionName)
    If IsTrue(this.Tokens) Then
        If this.Tokens.IncludeDefinition(Definition) Then
            GetValue = GetValueFromDefinition(Definition)
            Exit Function
        End If
    End If
    Dim Value As String
  #If DEV Then
    Dim Data As Dictionary
  #Else
    Dim Data As Object
  #End If
    Set Data = this.ProjectConfig.Read()
    If Data.Exists(Definition.Key) Then
        GetValue = Utils.ConvertToType(Data(Definition.Key), Definition.DataType)
        Exit Function
    End If

    Set Data = this.UserConfig.Read()
    If Data.Exists(Definition.Key) Then
        GetValue = Utils.ConvertToType(Data(Definition.Key), Definition.DataType)
        Exit Function
    End If

    Set Data = this.GlobalConfig.Read()
    If Data.Exists(Definition.Key) Then
        GetValue = Utils.ConvertToType(Data(Definition.Key), Definition.DataType)
        Exit Function
    End If

    If Not Information.IsMissing(Definition.Default) Then
        GetValue = Utils.ConvertToType(Definition.Default, Definition.DataType)
        Exit Function
    End If
End Function

Private Function GetValueFromDefinition(ByRef Definition As Definition) As Variant
    If Definition.DataType = vbBoolean Then
        GetValueFromDefinition = True
    Else
        GetValueFromDefinition = Utils.ConvertToType( _
            this.Tokens.GetTokenByDefinition(Definition).Text, Definition.DataType _
        )
    End If
End Function

Private Function GetValueFromFile(ByRef LocationDef As Definition, ByRef Definition As Definition) As Variant
    Dim Location As String
    Location = Strings.LCase(this.Tokens.GetTokenByDefinition(LocationDef).Text)
    Dim Config As ConfigIO
    Select Case Location
        Case "project": Set Config = this.ProjectConfig
        Case "user": Set Config = this.UserConfig
        Case "global": Set Config = this.GlobalConfig
    End Select
    GetValueFromFile = Config.Read()(Definition.Key)
End Function


