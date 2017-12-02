''' <summary>
''' 
''' </summary>
Module Subsets
    Public Enum parameter
        name
        Dte
        size

    End Enum
    ''' <summary>
    ''' Returns a list consisting of all elements of a list whose files have a parameter between two values (inclusive at lower limit)
    ''' </summary>
    Public Function subset(list As List(Of String), para1 As Date, para2 As Date) As List(Of String)
        Dim dt As Date
        Dim rl As New List(Of String)
        For Each s In list
            Dim finfo As New IO.FileInfo(s)
            dt = finfo.CreationTime
            If dt < para2 And dt >= para1 Then
                rl.Add(s)
            End If
        Next
        Return rl

    End Function
    Public Function subset(list As List(Of String), para1 As Long, para2 As Long) As List(Of String)
        Dim l As Long
        Dim rl As New List(Of String)
        For Each s In list
            Dim finfo As New IO.FileInfo(s)
            l = finfo.Length
            If l < para2 AndAlso l >= para1 Then
                rl.Add(s)
            End If
        Next
        Return rl
    End Function

    Public Function subset(list As List(Of String), para1 As String, para2 As String) As List(Of String)
        Dim l As String
        Dim rl As New List(Of String)
        For Each s In list
            Dim finfo As New IO.FileInfo(s)
            l = finfo.Name
            If l < para2 AndAlso l >= para1 Then
                rl.Add(s)
            End If
        Next
        Return rl
    End Function



End Module
