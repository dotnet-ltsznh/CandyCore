Public Class GPS

    ''' <summary>
    ''' 圆周率
    ''' </summary>
    Public Const Pi As Double = 3.14159265358979324D
    ''' <summary>
    ''' 地图半径
    ''' </summary>
    Public Const EarthR As Double = 6378245D
    ''' <summary>
    ''' WGS偏心率的平方
    ''' </summary>
    Public Const Offset As Double = 0.00669342162296594323D

    ''' <summary>
    ''' 是否超出中国
    ''' </summary>
    ''' <param name="latitude"></param>
    ''' <param name="longitude"></param>
    ''' <returns>是否超出中国范围</returns>
    Public Function OutOfChina(ByVal latitude As Double, ByVal longitude As Double) As Boolean
        If longitude < 72.004 OrElse longitude > 137.8347 Then
            Return True
        End If
        If latitude < 0.8293 OrElse latitude > 55.8271 Then
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' 获取经度转换偏差
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    Private Function TransformLongitude(ByVal x As Double, ByVal y As Double)
        Dim ret As Double = 300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y + 0.1 * Math.Sqrt(Math.Abs(x))
        ret += (20.0 * Math.Sin(6.0 * x * Pi) + 20.0 * Math.Sin(2.0 * x * Pi)) * 2.0 / 3.0
        ret += (20.0 * Math.Sin(x * Pi) + 40.0 * Math.Sin(x / 3.0 * Pi)) * 2.0 / 3.0
        ret += (150.0 * Math.Sin(x / 12.0 * Pi) + 300.0 * Math.Sin(x / 30.0 * Pi)) * 2.0 / 3.0
        Return ret
    End Function


    ''' <summary>
    ''' 获取维度转换偏差
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    Private Function TransformLatitude(ByVal x As Double, ByVal y As Double)
        Dim ret = -100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y + 0.2 * Math.Sqrt(Math.Abs(x))
        ret += (20.0 * Math.Sin(6.0 * x * Pi) + 20.0 * Math.Sin(2.0 * x * Pi)) * 2.0 / 3.0
        ret += (20.0 * Math.Sin(y * Pi) + 40.0 * Math.Sin(y / 3.0 * Pi)) * 2.0 / 3.0
        ret += (160.0 * Math.Sin(y / 12.0 * Pi) + 320 * Math.Sin(y * Pi / 30.0)) * 2.0 / 3.0
        Return ret
    End Function

    ''' <summary>
    ''' 根据GPS经纬度，获取地图经度
    ''' </summary>
    ''' <param name="longitude"></param>
    ''' <param name="latitude"></param>
    ''' <returns></returns>
    Public Function getMapLongitude(ByVal longitude As Double, ByVal latitude As Double) As Double
        Try
            If Me.OutOfChina(latitude, longitude) Then
                Throw New Exception("经纬度坐标超出中国范围!")
            End If
            Dim dLongitude As Double = Me.TransformLongitude(longitude - 105.0, latitude - 35.0)
            Dim radLat As Double = latitude / 180.0 * Pi
            Dim magic As Double = Math.Sin(radLat)
            magic = 1 - Offset * magic * magic
            Dim sqrtMagic As Double = Math.Sqrt(magic)
            dLongitude = (dLongitude * 180.0) / (EarthR / sqrtMagic * Math.Cos(radLat) * Pi)
            Dim mgLongitude As Double = longitude + dLongitude
            Return Decimal.Round(CType(mgLongitude, Decimal), 6)
        Catch ex As Exception
            Throw ex
        End Try
        Return 10
    End Function

    ''' <summary>
    ''' 根据GPS经纬度,获取地图纬度
    ''' </summary>
    ''' <param name="longitude"></param>
    ''' <param name="latitude"></param>
    ''' <returns>地图纬度</returns>
    Public Function GetMapLatitude(ByVal longitude As Double, ByVal latitude As Double) As Decimal
        Try
            If Me.OutOfChina(latitude, longitude) Then
                Throw New Exception("经纬度坐标超出中国范围!")
            End If
            Dim dLatitude As Double = Me.TransformLatitude(longitude - 105.0, latitude - 35.0)
            Dim radLat As Double = latitude / 180.0 * Pi
            Dim magic As Double = Math.Sin(radLat)
            magic = 1 - Offset * magic * magic
            Dim sqrtMagic As Double = Math.Sqrt(magic)
            dLatitude = (dLatitude * 180.0) / ((EarthR * (1 - Offset)) / (magic * sqrtMagic) * Pi)
            Dim mgLatitude As Double = latitude + dLatitude
            Return Decimal.Round(CType(mgLatitude, Decimal), 6)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
