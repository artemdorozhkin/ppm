Attribute VB_Name = "ProjectExporterCstr"
'@Folder "PearPMProject.src.ProjectExporter"
Option Explicit

Public Function NewProjectExporter( _
    Optional ByVal Destination As String, _
    Optional ByVal SaveStruct As Boolean, _
    Optional ByVal RewriteLastExport As Boolean = True _
) As ProjectExporter
    Set NewProjectExporter = New ProjectExporter
    NewProjectExporter.Destination = Destination
    NewProjectExporter.SaveStruct = SaveStruct
    NewProjectExporter.RewriteLastExport = RewriteLastExport
End Function
