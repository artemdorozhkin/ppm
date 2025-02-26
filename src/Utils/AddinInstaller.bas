Attribute VB_Name = "AddinInstaller"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Sub SwitchAddin()
  #If DEV Then
    Dim VBProject As VBProject: Set VBProject = Application.VBE.ActiveVBProject
  #Else
    Dim VBProject As Object: Set VBProject = Application.VBE.ActiveVBProject
  #End If
    If VBProject.Name = "PearPMProject" Then
        Interaction.MsgBox "ppm can't enabled for itself", VbMsgBoxStyle.vbExclamation, "ppm project selected"
        Exit Sub
    End If

    Dim ppmEnabled As Boolean
  #If DEV Then
    Dim Ref As Reference
  #Else
    Dim Ref As Object
  #End If
    For Each Ref In VBProject.References
        ppmEnabled = Ref.Name = "PearPMProject"
        If ppmEnabled Then Exit For
    Next

    Dim Msg As String
    If ppmEnabled Then
        VBProject.References.Remove Ref
        Msg = "ppm: disabled"
    Else
      #If DEV Then
        Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
      #Else
        Dim FSO As Object: Set FSO = NewFileSystemObject()
      #End If
        Dim ppmAddinPath As String
        ppmAddinPath = FSO.BuildPath(Interaction.Environ("APPDATA"), "Microsoft")
        ppmAddinPath = FSO.BuildPath(ppmAddinPath, "AddIns")
        ppmAddinPath = FSO.BuildPath(ppmAddinPath, "ppm.xlam")
        If Not FSO.FileExists(ppmAddinPath) Then
            Interaction.MsgBox "Can't find ppm.xlam:" & vbNewLine & ppmAddinPath, VbMsgBoxStyle.vbExclamation, "ppm not found"
            Exit Sub
        End If
        VBProject.References.AddFromFile ppmAddinPath
        Msg = "ppm: enabled" & vbNewLine & "for more information run 'ppm' command in immediate window"
    End If

    Debug.Print Msg
    Interaction.MsgBox Msg, VbMsgBoxStyle.vbInformation, "Success"
End Sub
