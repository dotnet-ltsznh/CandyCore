Public Class TestStopwatch

    Private unitStopwatch As Stopwatch '单元耗时
    Private totalStopwatch As Stopwatch '总共耗时

    ''' <summary>
    ''' 开启测试
    ''' </summary>
    ''' <param name="restart">是否重新开始</param>
    Public Sub startWatch(ByVal restart As Boolean)
        If restart Then
            totalStopwatch = New Stopwatch
            unitStopwatch = New Stopwatch

            totalStopwatch.Start()
            unitStopwatch.Start()
        Else
            unitStopwatch = New Stopwatch
            unitStopwatch.Start()
        End If
    End Sub

    ''' <summary>
    ''' 停止测试
    ''' </summary>
    ''' <param name="isFinished">是否结束测试</param>
    ''' <param name="message">测试消息</param>
    Public Sub stopWatch(ByVal isFinished As Boolean, ByVal message As String)
        If isFinished Then
            unitStopwatch.Stop()
            totalStopwatch.Stop()

            Console.WriteLine(Date.Now() & vbTab & message & "-耗时(ms)：" & unitStopwatch.ElapsedMilliseconds)
            Console.WriteLine(Date.Now() & vbTab & message & "-总共耗时(ms)：" & totalStopwatch.ElapsedMilliseconds)
        Else
            unitStopwatch.Stop()
            Console.WriteLine(Date.Now() & vbTab & message & "-耗时(ms)：" & unitStopwatch.ElapsedMilliseconds)
        End If
    End Sub

    ''' <summary>
    ''' 不暂停，临时显示耗时
    ''' </summary>
    ''' <param name="message">测试消息</param>
    Public Sub showWatch(ByVal message As String)
        Console.WriteLine(Date.Now() & vbTab & message & "-临时显示耗时(ms)：" & unitStopwatch.ElapsedMilliseconds)
        Console.WriteLine(Date.Now() & vbTab & message & "-临时显示总共耗时(ms)：" & totalStopwatch.ElapsedMilliseconds)
    End Sub

End Class
