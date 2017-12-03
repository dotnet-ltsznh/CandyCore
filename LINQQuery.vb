Public Class LINQQuery

    ''' <summary>
    ''' 打印筛选的数字
    ''' </summary>
    Public Sub printQueryInteger()
        Dim scores(10) As Integer
        scores = {97, 92, 81, 60}
        Dim scoreQuery As IEnumerable(Of Integer) = From score In scores
                                                    Where score > 80
                                                    Select score

        For Each score As Integer In scoreQuery
            Console.WriteLine(score)
        Next
    End Sub

End Class
