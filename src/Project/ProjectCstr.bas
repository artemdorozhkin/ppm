Attribute VB_Name = "ProjectCstr"
'@Folder "PearPMProject.src.Project"
Option Explicit

#If DEV Then
  Public Function NewProject(ByVal Project As VBProject) As Project
#Else
  Public Function NewProject(ByVal Project As Object) As Project
#End If
    Set NewProject = New Project
    Set NewProject.Project = Project
End Function
