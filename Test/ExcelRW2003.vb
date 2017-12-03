Imports System.IO
Imports System.Text
Imports NPOI.HSSF.UserModel
Imports NPOI.POIFS.FileSystem

Public Class ExcelRW2003

    Public Sub ExcelRW2003()

    End Sub


    Public Function ExcelReadAndWrite() As String
        Dim fs As New FileStream("D:\\123.xls", FileMode.Open, FileAccess.Read)
        'Dim data As Byte()
        'fs.Read(data, 0, fs.Length) '一次读完（到缓冲数组中）  
        ''TextBox1.Text = Encoding.ASCII.GetString(btArray)
        'fs.Close()
        'fs.Dispose()


        Dim poiFs As New POIFSFileSystem(fs)
        fs.Close()
        fs.Dispose()

        Dim book As New HSSFWorkbook(poiFs)
        Dim sheet As HSSFSheet = book.GetSheet("Excel Sheet")

        '第一行
        Dim row0 As HSSFRow = sheet.CreateRow(0)
        row0.CreateCell(0).SetCellValue("0,0")

        '第二行
        Dim row1 As HSSFRow = sheet.CreateRow(1)
        row1.CreateCell(0).SetCellValue("1,0")
        row1.CreateCell(1).SetCellValue("文档结束")
        row1.CreateCell(2).SetCellValue("文档结束文档结束")


        row1.CreateCell(3).SetCellValue("文档被修改了")


        fs = New FileStream("D:\\123.xls", FileMode.OpenOrCreate, FileAccess.Write)
        book.Write(fs)
        fs.Close()
        fs.Dispose()

        '写入本地
        'Dim fs As FileStream
        'fs = File.OpenWrite("D:\\123.xls")
        'For i = 0 To UBound(data)
        '    fs.WriteByte(data(i))
        'Next
        'fs.Close()
        'fs.Dispose()



        '写入文件 - 按块
        'fs = New FileStream("D:\\123.xls", FileMode.OpenOrCreate, FileAccess.Write)
        'fs.Write(data, 0, data.Length) '写入开始，结尾
        'fs.Close()
        'fs.Dispose()


        '写入文件 - 按块
        'Dim fs As New FileStream("D:\11.txt", FileMode.Append, FileAccess.Write)
        'Dim str As String = "This is my data block"
        'Dim btArray As Byte()
        'btArray = Encoding.ASCII.GetBytes(str)
        'fs.Write(btArray, 1, 3) '写入开始，结尾
        'fs.Close()



        '转换成字节流
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


    Function ExcelWrite() As Boolean
        Dim book As New HSSFWorkbook()
        Dim sheet As HSSFSheet = book.CreateSheet("Excel Sheet")

        '第一行
        Dim row0 As HSSFRow = sheet.CreateRow(0)
        row0.CreateCell(0).SetCellValue("0,0")

        '第二行
        Dim row1 As HSSFRow = sheet.CreateRow(1)
        row1.CreateCell(0).SetCellValue("1,0")
        row1.CreateCell(1).SetCellValue("文档结束")
        row1.CreateCell(2).SetCellValue("文档结束文档结束")

        '写入本地
        Dim ms As New IO.MemoryStream()
        book.Write(ms)
        Dim data As Byte() = ms.ToArray

        '写入本地
        'Dim fs As FileStream
        'fs = File.OpenWrite("D:\\123.xls")
        'For i = 0 To UBound(data)
        '    fs.WriteByte(data(i))
        'Next
        'fs.Close()
        'fs.Dispose()

        Dim fs As New FileStream("D:\\123.xls", FileMode.OpenOrCreate, FileAccess.Write)
        fs.Write(data, 0, data.Length) '写入开始，结尾
        fs.Close()
        fs.Dispose()

        '写入文件 - 按块
        'Dim fs As New FileStream("D:\11.txt", FileMode.Append, FileAccess.Write)
        'Dim str As String = "This is my data block"
        'Dim btArray As Byte()

        'btArray = Encoding.ASCII.GetBytes(str)
        'fs.Write(btArray, 1, 3) '写入开始，结尾
        'fs.Close()





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

        Return True
    End Function

End Class
