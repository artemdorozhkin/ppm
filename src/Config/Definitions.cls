VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Definitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PearPMProject.src.Config"
Option Explicit

#If DEV Then
  Public Property Get Items() As Dictionary
Attribute Items.VB_UserMemId = 0
'@DefaultMember
#Else
  Public Property Get Items() As Object
#End If
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If
    Dim Definition As Definition

    Set Definition = NewDefinition( _
        Key:="_after-dialog", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="author-name", _
        DataType:=vbString _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="author-url", _
        DataType:=vbString _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="dev-dep", _
        Short:="d", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="encoding", _
        Short:="e", _
        Default:="UTF-8" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="global", _
        Short:="g", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="help", _
        Short:="h", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="registry", _
        Default:=Constants.LocalRegistryPath _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="language", _
        Default:="eng" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="local", _
        Short:="l", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="location", _
        Default:="project" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="name", _
        Default:="VBAProject" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="no-clear", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="path", _
        Short:="p", _
        Default:="" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="save-struct", _
        Short:="s", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="version", _
        Short:="v", _
        Default:="1.0.0" _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="yes", _
        Short:="y", _
        Default:=False _
    )
    Set Buffer(Definition.Key) = Definition

    Set Items = Buffer
End Property
