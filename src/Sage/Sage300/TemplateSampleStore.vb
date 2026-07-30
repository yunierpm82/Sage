Imports System
Imports System.IO

Namespace Sage300

    ''' Keeps a copy of the sample Excel file for each column-mapping template in a "Plantillas"
    ''' folder next to the application executable, so the column links can be shown again after
    ''' restarting the app without needing the original file to still exist in its original location.
    ''' An optional "category" (e.g. "AP") keeps each module's samples in their own subfolder; the
    ''' default (Nothing) keeps the original GL Booking location ("Plantillas" itself).
    Public Class TemplateSampleStore

        Private Shared Function GetFolder(Optional category As String = Nothing) As String
            Dim baseFolder = Path.Combine(AppContext.BaseDirectory, "Plantillas")
            Dim folder = If(String.IsNullOrEmpty(category), baseFolder, Path.Combine(baseFolder, category))
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function GetSamplePath(templateName As String, Optional category As String = Nothing) As String
            Return Path.Combine(GetFolder(category), templateName & ".xlsx")
        End Function

        Public Shared Function TryGetExistingSamplePath(templateName As String, Optional category As String = Nothing) As String
            Dim filePath = GetSamplePath(templateName, category)
            Return If(File.Exists(filePath), filePath, Nothing)
        End Function

        Public Shared Sub SaveSample(templateName As String, sourceFilePath As String, Optional category As String = Nothing)
            File.Copy(sourceFilePath, GetSamplePath(templateName, category), overwrite:=True)
        End Sub

        Public Shared Sub DeleteSample(templateName As String, Optional category As String = Nothing)
            Dim filePath = GetSamplePath(templateName, category)
            If File.Exists(filePath) Then File.Delete(filePath)
        End Sub

    End Class

End Namespace
