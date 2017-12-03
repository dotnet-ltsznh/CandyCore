Imports System.IO
Imports NPOI.XSSF.UserModel

Public Class ExcelRW

    Public Sub ExcelRW()

    End Sub


    Public Function ExcelReadAndWrite() As String
        Dim fs As FileStream
        Try
            fs = New FileStream("D:\123.xlsx", FileMode.Open, FileAccess.Read)
        Catch ex As Exception
            fs = Nothing
        End Try

        Dim book As XSSFWorkbook

        If fs IsNot Nothing Then
            Try
                book = New XSSFWorkbook(fs)
            Catch ex As Exception
                book = New XSSFWorkbook()
            End Try
            fs.Close()
            fs.Dispose()
        Else
            book = New XSSFWorkbook()
        End If

        Dim sheet As XSSFSheet
        If book.GetSheet("Excel Sheet") Is Nothing Then
            sheet = book.CreateSheet("Excel Sheet")
        Else
            sheet = book.GetSheet("Excel Sheet")
        End If


        '第一行
        Dim row0 As XSSFRow
        If sheet.GetRow(0) IsNot Nothing Then
            row0 = sheet.GetRow(0)
        Else
            row0 = sheet.CreateRow(0)
        End If


        If row0.GetCell(0) Is Nothing Then
            row0.CreateCell(0).SetCellValue("0,0")
        Else
            row0.GetCell(0).SetCellValue("0,0")
        End If

        '第二行
        Dim row1 As XSSFRow
        If sheet.GetRow(1) IsNot Nothing Then
            row1 = sheet.GetRow(1)
        Else
            row1 = sheet.CreateRow(1)
        End If

        If row1.GetCell(0) Is Nothing Then
            row1.CreateCell(0).SetCellValue("1,0")
        Else
            row1.GetCell(0).SetCellValue("1,0")
        End If



        fs = New FileStream("D:\123.xlsx", FileMode.OpenOrCreate, FileAccess.Write)
        book.Write(fs)
        fs.Close()
        fs.Dispose()


        '获取字节流
        'Dim ms As New IO.MemoryStream()
        'book.Write(ms)
        'Dim data As Byte() = ms.ToArray

        '写入客户端
        'Dim ms As New IO.MemoryStream()
        'book.Write(ms)

        '// 写入到客户端  
        'System.IO.MemoryStream MS = New System.IO.MemoryStream();
        'book.Write(MS);
        'Response.AddHeader("Content-Disposition", String.Format("attachment; filename={0}.xls", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
        'Response.BinaryWrite(MS.ToArray());
        'book = null;
        'MS.Close();
        'MS.Dispose();

        Return ""
    End Function

End Class
