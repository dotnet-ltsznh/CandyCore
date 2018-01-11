Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO

''' <summary>
''' 打印机，目前仅可打印文本
''' </summary>
Public Class CandyPrinter
    Private printName As String = "Microsoft XPS Document Writer" ' 打印机名称
    Friend stringToPrint As String ' 打印内容
    Friend printDoc As New PrintDocument
    Private printFilePath As String = "d:\1.txt"

    ''' <summary>
    ''' 打印参数，字体、格式等
    ''' 未设置分页，打印大小等信息。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    Private Sub PrintPageHandler(ByVal sender As Object, ByVal args As PrintPageEventArgs)
        Dim myFont As New Font("Microsoft San Serif", 10)
        args.Graphics.DrawString(stringToPrint, New Font(myFont, FontStyle.Regular), Brushes.Black, 50, 50)
    End Sub


    ''' <summary>
    ''' 打印参数，字体格式等
    ''' 根据内容自动换行，自动分页
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    Private Sub printDoc_PrintPage(ByVal sender As Object, ByVal args As PrintPageEventArgs)
        Dim charactersOnPage As Int16 = 0
        Dim linesPerPage As Int16 = 0
        Dim myFont As New Font("Microsoft San Serif", 10)
        ' Sets the value of charactersOnPage to the number of characters
        ' of stringToPrint that will fit within the bounds of the page.
        args.Graphics.MeasureString(stringToPrint, myFont,
                                    args.MarginBounds.Size, StringFormat.GenericTypographic,
                                     charactersOnPage, linesPerPage)

        ' Draws the string within the bounds of the page
        args.Graphics.DrawString(stringToPrint, myFont,
                                 Brushes.Black, args.MarginBounds, StringFormat.GenericTypographic)

        ' Remove the portion of the string that has been printed.
        stringToPrint = stringToPrint.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed
        args.HasMorePages = (stringToPrint.Length > 0)

    End Sub

    ''' <summary>
    ''' 所有参数设置后，开始打印
    ''' </summary>
    Private Sub Print()
        If stringToPrint Is Nothing OrElse stringToPrint.Length <= 0 Then Return ' 如果内容为空，无需打印，则不打印
        Using (printDoc)
            printDoc.PrinterSettings.PrinterName = printName
            AddHandler printDoc.PrintPage, AddressOf Me.printDoc_PrintPage
            printDoc.Print()
            RemoveHandler printDoc.PrintPage, AddressOf Me.printDoc_PrintPage
        End Using
        'printDoc.Print()
    End Sub

    ''' <summary>
    ''' 设置打印文件路径
    ''' </summary>
    ''' <param name="filePath"></param>
    Private Sub setPrintFilePath(ByVal filePath As String)
        printFilePath = filePath
    End Sub

    ''' <summary>
    ''' 读取要打印的文件内容
    ''' </summary>
    Private Sub ReadFile()
        printDoc.DocumentName = "1.txt"
        Dim stream As New FileStream(printFilePath, FileMode.Open)
        Try
            Dim reader As New StreamReader(stream)
            Try
                stringToPrint = reader.ReadToEnd
            Finally
                reader.Dispose()
            End Try
        Finally
            stream.Dispose()
        End Try

    End Sub

    ''' <summary>
    ''' 按文件路径进行打印
    ''' </summary>
    ''' <param name="filePath"></param>
    Public Sub PrintFileByFilePath(ByVal filePath As String)
        setPrintFilePath(filePath)
        ReadFile()
        Print()
    End Sub

    ''' <summary>
    ''' 直接打印文本
    ''' </summary>
    ''' <param name="text"></param>
    Public Sub Print(ByVal text As String)
        stringToPrint = text
        Print()
    End Sub
End Class
