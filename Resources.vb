Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Xml

Friend Class App
    Public Shared ReadOnly Property [Assembly]() As System.Reflection.Assembly
        Get
            Return ClassType.Assembly
        End Get
    End Property
    Public Shared ReadOnly Property ClassType() As Type
        Get
            Return GetType(App)
        End Get
    End Property

    Public Shared Function GetEmbeddedResource(ByVal resName As String) As Image
        resName = resName.Replace("/", ".").Replace("\\", ".")
        If resName.IndexOf(App.ClassType.Namespace) <> 0 Then
            resName = App.ClassType.Namespace + "." + resName
        End If

        Dim s As Stream = App.Assembly.GetManifestResourceStream(resName)
        If Not IsNothing(s) Then
            Dim ext As String = resName.Substring(resName.LastIndexOf(".") + 1)
            Return New System.Drawing.Bitmap(s)
        End If
        Return Nothing
    End Function
End Class

