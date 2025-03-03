VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyncCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Sync"
Option Explicit

Implements ICommand

Private Type TSyncCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TSyncCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/sync"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "sync"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If Not SelectedProject.IsSaved Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/saveRequired")
        End
    End If

    Dim ms As Double: ms = DateTime.Timer

    Immediate.WriteLine CLI.Lang.GetValue("messages/synchronisation")

    Dim ActivePaneName As String
    ActivePaneName = Application.VBE.ActiveCodePane.CodeModule.Parent.Name
  #If DEV Then
    Dim SyncPaths As Dictionary: Set SyncPaths = NewDictionary()
    Dim SyncNames As Dictionary: Set SyncNames = NewDictionary()
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim SyncPaths As Object: Set SyncPaths = NewDictionary()
    Dim SyncNames As Object: Set SyncNames = NewDictionary()
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Me.CollectPaths SelectedProject.Folder, SyncPaths

    Dim Module As Module
    Dim Name As Variant
    For Each Name In SyncPaths
        Dim Path As String: Path = SyncPaths(Name)
        Dim Ext As String: Ext = FSO.GetExtensionName(Path)
        Set Module = NewModule(SelectedProject.GetComponent(Name))
        Dim Code As String: Code = Me.ReadCode(Path)

        Dim CodePage As String: CodePage = Locale.GetAnsiCodePage()
        Dim LocalePath As String
        LocalePath = FSO.BuildPath( _
            FSO.GetParentFolderName(Path), _
            PStrings.FString("{0}{1}.{2}", Name, CodePage, Ext) _
        )
        PFileSystem.SaveToFile _
            LocalePath, _
            PFileSystem.ReadFile(Path), _
            PStrings.FString("Windows-{0}", CodePage)

        If IsFalse(Module.Item) And Not PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            SelectedProject.ImportComponent LocalePath
        ElseIf IsFalse(Module.Item) And PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            ' TODO: need logic for creation the sheet and paste the code
            Information.Err.Raise 5, Information.TypeName(Me), "Non implemented"
        ElseIf IsFalse(Module.Item) Then
            SelectedProject.ImportComponent LocalePath
        ElseIf IsTrue(Module.Item) And PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            Module.Code = Code
        Else
            If Module.Kind = ModuleType.FormModule Then
                SelectedProject.RemoveComponent Name
                SelectedProject.ImportComponent LocalePath
            Else
                Dim DelName As String: DelName = "PPMToDelete_" & SyncNames.Count
                Module.Name = DelName
                SelectedProject.ImportComponent LocalePath
                SelectedProject.RemoveComponent DelName
            End If
        End If
        SyncNames.Add Name, True
        FileSystem.Kill LocalePath
    Next

  #If DEV Then
    Dim Component As VBComponent
    Dim Components As VBComponents: Set Components = SelectedProject.Components
  #Else
    Dim Component As Object
    Dim Components As Object: Set Components = SelectedProject.Components
  #End If
    For Each Component In Components
        If Not SyncNames.Exists(Component.Name) _
        Or PStrings.StartsWith(Component.Name, "PPMToDelete_") Then
            Components.Remove Component
        End If
    Next

    ms = DateTime.Timer - ms
    ' Throws a message when an Ambiguous name detected.
    ' This is normal behaviour if PPMProject is set as the active project.
    ' The error should not occur in other projects.
    Immediate.WriteLine PStrings.FString( _
        CLI.Lang.GetValue("messages/synchronised"), _
        SyncPaths.Count, _
        Utils.ConvertTime(ms) _
    )

    On Error Resume Next ' if this component was deleted
    Components(ActivePaneName).Activate
    On Error GoTo 0
End Sub

#If DEV Then
  Public Sub CollectPaths(ByVal FolderPath As String, ByRef Paths As Dictionary)
#Else
  Public Sub CollectPaths(ByVal FolderPath As String, ByRef Paths As Object)
#End If
  #If DEV Then
    Dim Folder As Scripting.Folder: Set Folder = NewFolder(FolderPath)
    Dim SubFolder As Scripting.Folder
  #Else
    Dim Folder As Object: Set Folder = NewFolder(FolderPath)
    Dim SubFolder As Object
  #End If
    For Each SubFolder In Folder.SubFolders
        CollectPaths SubFolder.Path, Paths
    Next

  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim File As File
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
    Dim File As Object
  #End If
    For Each File In Folder.Files
        Dim Ext As String: Ext = Strings.LCase(FSO.GetExtensionName(File.Path))
        Dim Name As String: Name = FSO.GetBaseName(File.Path)
        If PArray.IncludesAny(Array("bas", "cls", "doccls", "frm"), Ext) Then
            Paths(Name) = File.Path ' Because Name is a unique value, we can't have two identical modules
        End If
    Next
End Sub

Public Function ReadCode(ByVal Path As String) As String
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim Ext As String: Ext = Strings.LCase(FSO.GetExtensionName(Path))
    Dim VBAttr As String

    Select Case Ext
        Case "bas": VBAttr = "Attribute VB_Name = "
        Case Else: VBAttr = "Attribute VB_Exposed = "
    End Select

    Dim Code As String: Code = PFileSystem.ReadFile(Path)
    Dim AttrPosition As String: AttrPosition = Strings.InStr(1, Code, VBAttr, vbTextCompare)
    Code = Strings.Right(Code, Strings.Len(Code) - AttrPosition)

    Dim NewLinePosition As String: NewLinePosition = Strings.InStr(1, Code, vbNewLine)
    Code = Strings.Right(Code, Strings.Len(Code) - NewLinePosition - 1)

    ReadCode = Code
End Function
