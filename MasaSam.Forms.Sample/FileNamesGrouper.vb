Public Class FileNamesGrouper
    Public Event WordsParsed()

    Private mWordlist As New SortedList(Of String, Integer)
    Public Property WordList() As SortedList(Of String, Integer)
        Get
            Return mWordlist
        End Get
        Set(ByVal value As SortedList(Of String, Integer))
            mWordlist = value
        End Set
    End Property

#Region "Properties"
    Private mFilenames As New List(Of String)
    Public Property Filenames() As List(Of String)
        Get
            Return mFilenames
        End Get
        Set(ByVal value As List(Of String))
            mFilenames = value
            mGroups.Clear()
            mGroupNames.Clear()
            mWordlist.Clear()
            ParseNames()
        End Set
    End Property
    Private mGroups As New List(Of List(Of String))
    Public Property Groups As List(Of List(Of String))
        Get
            Return mGroups
        End Get
        Set(ByVal value As List(Of List(Of String)))
            mGroups = value
        End Set
    End Property
    Private mGroupNames As New List(Of String)
    Public Property GroupNames() As List(Of String)
        Get
            Return mGroupNames
        End Get
        Set(ByVal value As List(Of String))
            mGroupNames = value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Sub Clear()
        mFilenames.Clear()
        mGroupNames.Clear()
        mGroups.Clear()
    End Sub
    Private Sub ParseNames()
        For Each f In mFilenames
            StringParser(f)
        Next
        UpdateWordlist(5)

        GetGroups()
            If mWordlist.Count <> 0 Then RaiseEvent WordsParsed()
    End Sub
    ''' <summary>
    ''' Takes a string, and uses a regular expression to strip out all the words. 
    ''' Adds those words to a sorted list maintaining a count of the global frequency 
    ''' </summary>
    ''' <param name="str"></param>

    Private Sub StringParser(str As String)
        Dim r As New System.Text.RegularExpressions.Regex("([A-Z|a-z]+[AEIOUY|aeiouy][A-Z|a-z]+[_]*)+") '("([A-Z]^[0-9])\w+")
        If r.Matches(str).Count > 0 Then
            Dim localwords As New List(Of String)
            For Each s In r.Matches(str)
                If mWordlist.ContainsKey(s.ToString) Then
                    If localwords.Contains(s.ToString) Then
                    Else
                        Dim k As Integer = mWordlist.Item(s.ToString) 'currently counts multiple occurrences in same str
                        k = k + 1
                        mWordlist.Remove(s.ToString)
                        mWordlist.Add(s.ToString, k)
                        localwords.Add(s.ToString)
                    End If
                Else
                    mWordlist.Add(s.ToString, 1)
                    localwords.Add(s.ToString)

                End If

            Next
        End If
    End Sub

    ''' <summary>
    ''' For each filename, adds to a lists which pair each name with a target group and its associated count
    ''' </summary>
    Private Sub GetGroups()
        mGroups.Clear()
        Dim Group As New SortedList(Of String, Integer) 'Filenames, Count
        Dim TargetGroup As New SortedList(Of String, String) 'Filenames, Target
        For Each f In mFilenames
            Group.Add(f, 0)
            Console.WriteLine(Group.Last)

            TargetGroup.Add(f, "")
            Console.WriteLine(TargetGroup.Last)

            For Each w In mWordlist 'Go through wordlist and allocate files which contain it

                If InStr(f, w.Key) <> 0 And Group(f) < w.Value Then
                    'If the file contains the word, and the occurrence of the word is greater than the currently assigned group, reassign them. 
                    'The problem is that this allocates each and every file to the group which has the largest frequency overall, irrespective of whether 
                    Console.WriteLine(w.Value & ", " & w.Key)
                    Console.WriteLine()

                    If w.Value > 0 Then
                        Group(f) = w.Value
                        TargetGroup(f) = w.Key

                    End If
                End If

            Next
        Next
        Dim targets As New List(Of String)

        For Each f In TargetGroup
            If targets.Contains(f.Value) Or f.Value = "" Then
            Else
                targets.Add(f.Value)
            End If
        Next
        For Each tgt In targets
            Dim flst As New List(Of String)
            For Each f In TargetGroup
                If f.Value = tgt Then
                    flst.Add(f.Key)
                End If
            Next
            mGroups.Add(flst)
            mGroupNames.Add(tgt)
        Next
    End Sub

    ''' <summary>
    ''' Goes through the Wordlist and doesn't copy any entries which are singletons, or universal
    ''' </summary>
    ''' 
    Private Sub UpdateWordlist(Count As Byte)
        Dim x As New SortedList(Of String, Integer)

        For Each m In mWordlist
            For i = 100 To Count Step -1

                If m.Value = mFilenames.Count Or m.Value < i Then
                Else
                    If x.Keys.Contains(m.Key) Then
                    Else
                        x.Add(m.Key, m.Value)

                    End If
                End If
            Next
        Next
        mWordlist = x
    End Sub
#End Region

End Class
