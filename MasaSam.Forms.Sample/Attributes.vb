Public Class Attributes
    Private Choices() As String = {"Rate", "Height", "Width", "Date"}
    Private Sub Attributes_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.Shown
    End Sub


    Public Sub UpdateLabel(filename As String)
        Exit Sub
        Dim AttributeName As String = ""
        Dim AttList As List(Of KeyValuePair(Of String, String)) = FileMetaData.GetMetaText(filename, "Q:\exiftool.exe")
        For Each f In AttList
            If ChosenString(f.Key) Then
                AttributeName += f.Key & ": " & f.Value & vbCrLf
            End If
        Next
        TextBox1.Text = AttributeName
    End Sub

    Private Function ChosenString(s As String) As Boolean
        Dim Flag As Boolean = False
        For i = 0 To 3
            If Not Flag Then Flag = s.Contains(Choices(i))

        Next
        Return Flag
    End Function
End Class


Public NotInheritable Class FileMetaData
    Public Delegate Function GetMetaTextDelegate(ByVal filePath As String, ByVal exifToolsExePath As String) As List(Of KeyValuePair(Of String, String))

    Public Shared Function GetMetaText(ByVal filePath As String,
                    ByVal exifToolsExePath As String) As List(Of KeyValuePair(Of String, String))

        Dim retVal As List(Of KeyValuePair(Of String, String)) = Nothing

        Try
            If String.IsNullOrWhiteSpace(exifToolsExePath) Then
                Throw New ArgumentException("The file path of the EXIF Tools executable" & vbCrLf &
                                            "cannot be null or empty.")

            ElseIf Not My.Computer.FileSystem.FileExists(exifToolsExePath) Then
                Throw New IO.FileNotFoundException("The file path of the EXIF Tools executable" & vbCrLf &
                                                   "could not be located.")

            ElseIf String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentException("The file path of the file to be examined" & vbCrLf &
                                            "cannot be null or empty.")

            ElseIf Not My.Computer.FileSystem.FileExists(filePath) Then
                Throw New IO.FileNotFoundException("The file path of the file to be examined" & vbCrLf &
                                                   "could not be located.")

            Else
                Using tool As New Process
                    Dim sb As New System.Text.StringBuilder

                    With tool
                        With .StartInfo
                            .FileName = exifToolsExePath
                            .Arguments = String.Format(" -s  {0}{1}{0}", Chr(34), filePath)
                            .UseShellExecute = False
                            .RedirectStandardOutput = True
                            .CreateNoWindow = True
                        End With

                        .Start()

                        sb.Append(tool.StandardOutput.ReadToEnd)
                        sb.Remove(sb.Length - 2, 2)

                        Dim pairs() As String = sb.ToString.Split(CChar(vbCrLf))
                        Dim delimiters() As String = New String() {" : "}

                        If retVal Is Nothing Then
                            retVal = New List(Of KeyValuePair(Of String, String))
                        End If

                        For Each element As String In pairs
                            Dim kv() As String = element.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)

                            If kv.Length = 2 Then
                                retVal.Add(New KeyValuePair(Of String, String)(kv(0).Trim, kv(1).Trim))
                            End If
                        Next
                    End With
                End Using
            End If

        Catch ex As Exception
            Throw
        End Try

        Return retVal

    End Function

End Class