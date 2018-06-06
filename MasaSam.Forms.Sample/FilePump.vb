Imports System.Threading
Public Class FilePump
    Dim t As Thread

    Private mListbox As ListBox
    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
        End Set
    End Property

    Private mPumpList As List(Of String)
    Public Property PumpList() As List(Of String)
        Get
            Return mPumpList
        End Get
        Set(ByVal value As List(Of String))
            mPumpList = value
        End Set
    End Property
    Private Sub MoveFilesThread()

        't = New Thread(New ThreadStart(Sub() FilePump()))
        't.IsBackground = True
        't.Start()
    End Sub
End Class
