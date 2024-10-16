VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Lang"
Option Explicit

Private Type TLang
    LangPath As String
  #If DEV Then
    XML As DOMDocument60
    Current As IXMLDOMNode
  #Else
    XML As Object
    Current As Object
  #End If
End Type

Private this As TLang

Public Sub SetLang(ByVal Language As String)
    ' Path name is ISO 639 language codes
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim Root As String
    Root = FSO.BuildPath(FSO.GetParentFolderName(Application.VBE.ActiveVBProject.FileName), "languages")
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
    Dim Root As String: Root = FSO.BuildPath(Constants.GlobalPPMPath, "languages")
  #End If

    this.LangPath = FSO.BuildPath(Root, PStrings.FString("{0}.xml", Language))

    Set this.XML = CreateObject("MSXML2.DOMDocument.6.0")
    this.XML.Load this.LangPath
End Sub

Public Function SetBlock(ByVal Path As String) As Lang
    Set this.Current = this.XML.SelectSingleNode(Path)
    Set SetBlock = Me
End Function

Public Function GetValues(ByVal Block As String, Optional ByVal Attr As String) As Variant
    If IsFalse(this.Current) Then Exit Function

  #If DEV Then
    Dim Nodes As IXMLDOMNodeList: Set Nodes = this.Current.SelectNodes(Block)
  #Else
    Dim Nodes As Object: Set Nodes = this.Current.SelectNodes(Block)
  #End If
    If Strings.Len(Attr) > 0 Then
        Dim Name As String: Name = Strings.Split(Attr, "=")(0)
        Dim Value As String: Value = Strings.Split(Attr, "=")(1)
        Dim Query As String: Query = PStrings.FString("{0}[@{1}='{2}']", Block, Name, Value)
        Set Nodes = this.Current.SelectNodes(Query)
    End If
    If Nodes.Length = 0 Then Exit Function

    Dim Items As Collection: Set Items = New Collection
  #If DEV Then
    Dim Child As IXMLDOMElement
  #Else
    Dim Child As Object
  #End If
    For Each Child In Nodes(0).ChildNodes
        Items.Add Child.Text
    Next
    GetValues = PCollection.ToArray(Items)
End Function

Public Function GetValue(ByVal Block As String, Optional ByVal Attr As String) As Variant
    If IsFalse(this.Current) Then Exit Function

  #If DEV Then
    Dim Nodes As IXMLDOMNodeList: Set Nodes = this.Current.SelectNodes(Block)
  #Else
    Dim Nodes As Object: Set Nodes = this.Current.SelectNodes(Block)
  #End If
    If Strings.Len(Attr) > 0 Then
        Dim Name As String: Name = Strings.Split(Attr, "=")(0)
        Dim Value As String: Value = Strings.Split(Attr, "=")(1)
        Dim Query As String: Query = PStrings.FString("{0}[@{1}='{2}']", Block, Name, Value)
        Set Nodes = this.Current.SelectNodes(Query)
    End If

    If Nodes.Length = 0 Then Exit Function
    GetValue = Nodes(0).Text
End Function
