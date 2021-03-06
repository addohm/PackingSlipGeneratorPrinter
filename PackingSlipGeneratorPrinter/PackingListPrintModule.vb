﻿Imports System.Data.SqlClient 'needed for DB interactions
Imports System.Drawing
Imports System.IO 'needed for BLOB
Imports Word = Microsoft.Office.Interop.Word 'needed for COM object interaction with MS Word

Module PackingListPrintModule

    Dim appPath As String = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
    Dim appRootFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Integra Optics")
    Dim userAppFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Integra Optics", "Packing Slip")
    Dim userDesktopFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Packing Slips")

    ''' <summary>
    ''' Clean everything out of a text string the follows a \
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="separator"></param>
    ''' <returns>String</returns>
    Public Function cleanedText(text As String, separator As String) As String

        Dim startIndex As Int16 = 0

        startIndex = InStr(1, text, separator)
        cleanedText = text.Remove(0, startIndex)

    End Function

    ''' <summary>
    ''' Reads an image blob from the database and stores it as a bmp file in the local temp directory
    ''' </summary>
    ''' <returns></returns>

    Private Function WriteImageFromDb() As String
        'Set the full file path for the image
        Dim imagePath = Path.GetTempPath & "IntegraPLabel.png"

        ' The bytes returned from GetBytes.
        Dim retval As Long
        ' The starting position in the BLOB output.
        Dim startIndex As Long

        Using connection = New SqlConnection()
            connection.ConnectionString = My.Settings.ConnStr
            connection.Open() 'open connection
            'sql command specific for the label type
            Using cmd As New SqlCommand With {
                    .Connection = connection,
                    .CommandText = "select company_logo from aof_order_queue"
                    }

                Using cmdReader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)
                    If Not cmdReader.Read() = False Then

                        'if the file already exists, delete and create a new one
                        If File.Exists(imagePath) Then
                            File.Delete(imagePath)
                        End If

                        ' Create a file to hold the output.
                        Using stream = New FileStream(imagePath, FileMode.OpenOrCreate, FileAccess.Write)
                            Using writer = New BinaryWriter(stream)
                                ' The size of the BLOB buffer.
                                Dim bufferSize = 100
                                ' The BLOB byte() buffer to be filled by GetBytes.
                                Dim outByte(bufferSize - 1) As Byte
                                ' Reset the starting byte for a new BLOB.
                                startIndex = 0

                                ' Read bytes into outByte() and retain the number of bytes returned.
                                retval = cmdReader.GetBytes(0, startIndex, outByte, 0, bufferSize)

                                ' Continue while there are bytes beyond the size of the buffer.
                                Do While retval = bufferSize
                                    writer.Write(outByte)
                                    writer.Flush()

                                    ' Reposition start index to end of the last buffer and fill buffer.
                                    startIndex += bufferSize
                                    retval = cmdReader.GetBytes(0, startIndex, outByte, 0, bufferSize)
                                Loop
                                ' Write the remaining buffer.
                                writer.Write(outByte, 0, retval)
                                writer.Flush()
                                writer.Close()

                            End Using
                            stream.Close()
                        End Using
                    Else
                        MsgBox("Label image data query returned no results!")
                        GC.Collect()
                        Environment.Exit(0)
                    End If
                    WriteImageFromDb = imagePath
                End Using
            End Using
        End Using
    End Function

    'form load subroutine
    Sub Main()
        Dim boxID As Integer = 0
        Dim version As Integer = 0
        Dim noprint As Boolean = True
        Dim templateName As String = String.Empty
        Dim args() As String = Environment.GetCommandLineArgs()
        args = args.Skip(1).ToArray

        If args.Length >= 1 Then
            If HelpRequired(args(0)) Then DisplayHelp()

            'Parse all the command line arguments
            For Each c In args
                Dim argVal As String = String.Empty
                'Return the argument name
                Dim arg As String = c.Split("=")(0).Replace("/", "").ToLower
                'return the argument value
                If c.Contains("=") Then
                    argVal = c.Split("=")(1).ToLower
                End If
                Select Case arg
                    Case "boxid"
                        boxID = Convert.ToInt32(argVal)
                    Case "version"
                        version = Convert.ToInt32(argVal)
                    Case "noprint"
                        noprint = False
                End Select
            Next

            If version = 1 Then
                templateName = "ITO_PackListTemplate.dotm"
            ElseIf version = 2 Then
                templateName = "PackListTemplate.dotm"
            End If

            If Not boxID = 0 And version > 0 Then
                Dim objWordApp As New Word.Application
                Dim objDoc As Word.Document
                Dim objTable As Word.Table
                'Run application in foreground or background.
                'If in background (false), be sure to add objDoc.close() and objWordApp.Quit()
                objWordApp.Visible = False

                'Declarations
                Dim localDateTimeString As String = DateTime.Now.ToString
                Dim localDateTimeFileName As String = DateTime.Now.ToString("yyyyMMddhhmm")
                Dim soNumber As String = "No_SO_Selected"
                Console.WriteLine("Box ID set to: " & boxID)
                'My.Application.Log.WriteEntry("Box ID set to: " & boxID)
                'Dim remaining As Integer = argFromCommandLine("remaining")
                Dim saveString As String 'file name format
                Dim pages As Integer = 0
                Dim cb As New SqlConnectionStringBuilder With {
                .DataSource = My.Settings.sqlServer,
                .InitialCatalog = My.Settings.sqlDBName,
                .UserID = My.Settings.sqlUsername,
                .Password = My.Settings.sqlPassword
            }

                With My.Settings
                    .ConnStr = cb.ToString()
                    .Save()
                End With

                If (Not Directory.Exists(userDesktopFolder)) Then
                    Directory.CreateDirectory(userDesktopFolder)
                End If

                Try
                    'Open the template
                    objDoc = objWordApp.Documents.Open(Path.Combine(appPath, templateName), [ReadOnly]:=True)
                    'set word document as active
                    objDoc = objWordApp.ActiveDocument

                    Using conn = New SqlConnection(My.Settings.ConnStr)
                        conn.Open()

                        Using cmd As New SqlCommand()
                            cmd.Connection = conn
                            cmd.CommandType = CommandType.Text

                            With objDoc

                                Dim filePath As String = WriteImageFromDb()
                                Dim img As Image = Image.FromFile(filePath)
                                Dim imgX As Integer = img.Width
                                Dim imgXLim As Integer = 350
                                Dim imgY As Integer = img.Height
                                Dim imgYLim As Integer = 100
                                Dim imgXdiff As Integer = imgX - imgXLim
                                Dim imgYdiff As Integer = imgY - imgYLim
                                Dim growFactor As Double
                                Dim shrinkFactor As Double
                                .PageSetup.DifferentFirstPageHeaderFooter = 0
                                With .Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Shapes
                                    'Check to see if the image is too big
                                    If imgX > imgXLim Or imgY > imgYLim Then
                                        If imgXdiff > imgYdiff Then
                                            shrinkFactor = (imgXLim / imgX) - ((imgXLim / imgX) Mod 0.1) 'Decimal Floor
                                        ElseIf imgYdiff > imgXdiff Then
                                            shrinkFactor = (imgYLim / imgY) - ((imgYLim / imgY) Mod 0.1) 'Decimal Floor
                                        End If
                                        imgX = imgX * shrinkFactor
                                        imgY = imgY * shrinkFactor
                                        'If the image isn't too big, check to see if it's too small
                                    ElseIf imgX < imgXLim Or imgY < imgYLim Then
                                        If imgXdiff > imgYdiff Then
                                            growFactor = (imgXLim / imgX) - ((imgXLim Mod imgX) / imgX) 'Floor
                                        ElseIf imgYdiff > imgXdiff Then
                                            growFactor = (imgYLim / imgY) - ((imgYLim Mod imgY) / imgY) 'Floor
                                        End If
                                        imgX = imgX * growFactor
                                        imgY = imgY * growFactor
                                    End If
                                    With .AddPicture(filePath)
                                        .Width = imgX
                                        .Height = imgY
                                    End With
                                End With
                                If version = 1 Then
                                    cmd.CommandText = "SELECT [SALES_ORDER_NUMBER], ISNULL([CUSTOMER_PO], ''), ISNULL([SHIP_VIA], '')," _
                                        & "ISNULL([DELIVERY_COMPANY], ''), ISNULL([DELIVERY_ATTN_TO], ''), ISNULL([DELIVERY_STREET], ''), " _
                                        & "ISNULL([DELIVERY_CITY], ''), ISNULL([DELIVERY_STATE], ''), ISNULL([DELIVERY_POSTAL_CODE], ''), " _
                                        & "ISNULL([DELIVERY_COUNTRY], ''), ISNULL([BILLING_COMPANY], ''), ISNULL([BILLING_STREET], ''), " _
                                        & "ISNULL([BILLING_CITY], ''), ISNULL([BILLING_STATE], ''), ISNULL([BILLING_POSTAL_CODE], ''),  " _
                                        & "ISNULL([BILLING_COUNTRY], ''), ISNULL([SALES_ORDER_NOTES], '') FROM [AOF_ORDER_QUEUE]"
                                ElseIf version = 2 Then
                                    cmd.CommandText = "SELECT [SALES_ORDER_NUMBER], ISNULL([CUSTOMER_PO], ''), ISNULL([SHIP_VIA], '')," _
                                         & "ISNULL([DELIVERY_COMPANY], ''), ISNULL([DELIVERY_ATTN_TO], ''), ISNULL([DELIVERY_STREET], ''), " _
                                         & "ISNULL([DELIVERY_CITY], ''), ISNULL([DELIVERY_STATE], ''), ISNULL([DELIVERY_POSTAL_CODE], ''), " _
                                         & "ISNULL([DELIVERY_COUNTRY], ''), ISNULL([BILLING_COMPANY], ''), ISNULL([BILLING_STREET], ''), " _
                                         & "ISNULL([BILLING_CITY], ''), ISNULL([BILLING_STATE], ''), ISNULL([BILLING_POSTAL_CODE], ''),  " _
                                         & "ISNULL([BILLING_COUNTRY], ''), ISNULL([SALES_ORDER_NOTES], ''), ISNULL([COMPANY_NAME], '')," _
                                         & "ISNULL([COMPANY_STREET], ''), ISNULL([COMPANY_CITY], ''), ISNULL([COMPANY_STATE], '')," _
                                         & "ISNULL([COMPANY_POSTAL_CODE], ''), ISNULL([COMPANY_COUNTRY], ''), ISNULL([COMPANY_EMAIL], ''), " _
                                         & "ISNULL([COMPANY_PHONE], '') FROM [AOF_ORDER_QUEUE]"
                                End If
                                Dim readerOrderQueue As SqlDataReader = cmd.ExecuteReader()
                                readerOrderQueue.Read()
                                'Table 1 company information and sales order number
                                objTable = .Tables.Item(1) 'select table 1

                                If version = 1 Then
                                    With objTable
                                        'Insert Text into table 1
                                        'Company Address
                                        .Cell(1, 1).Range.Text = "Integra Optics, Inc."
                                        .Cell(2, 1).Range.Text = "745 Albany Shaker Rd" & vbCrLf & "Latham, NY 12110-1417" &
                                              vbCrLf & "Phone: (877) 402-3850" & vbCrLf & "FAX: (866) 847-5219" & vbCrLf &
                                              "Email: info@integraoptics.com"
                                        'Packing List & SO
                                        soNumber = cleanedText(readerOrderQueue(0), "/")
                                        .Cell(2, 2).Range.Text = soNumber
                                    End With
                                ElseIf version = 2 Then
                                    With objTable
                                        'Insert Text into table 1
                                        'Company Address
                                        .Cell(1, 1).Range.Text = readerOrderQueue(17)
                                        .Cell(2, 1).Range.Text = readerOrderQueue(18) & vbCrLf _
                                        & readerOrderQueue(19) & ", " & readerOrderQueue(20) & " " & readerOrderQueue(21) & vbCrLf _
                                        & readerOrderQueue(22) & vbCrLf _
                                        & "Phone: " & readerOrderQueue(23) & vbCrLf _
                                        & "Email: " & readerOrderQueue(24)
                                        'Packing List & SO
                                        soNumber = cleanedText(readerOrderQueue(0), "/")
                                        .Cell(2, 2).Range.Text = soNumber
                                    End With
                                End If

                                'Table 2 shipping information and billing information
                                objTable = .Tables.Item(2)
                                With objTable
                                    'Insert Text into table 2
                                    'Ship to
                                    .Cell(2, 1).Range.Text = readerOrderQueue(3) 'Company name
                                    If Not readerOrderQueue.IsDBNull(4) Then
                                        .Cell(2, 1).Range.Text = .Cell(2, 1).Range.Text & readerOrderQueue(4) 'Attn to if exists
                                    End If
                                    .Cell(2, 1).Range.Text = .Cell(2, 1).Range.Text & readerOrderQueue(5) & vbCrLf & readerOrderQueue(6) _
                        & ", " & readerOrderQueue(7) & " " & readerOrderQueue(8) & vbCrLf & readerOrderQueue(9)
                                    'Bill to
                                    .Cell(2, 2).Range.Text = readerOrderQueue(10) & vbCrLf & readerOrderQueue(11) & vbCrLf & readerOrderQueue(12) _
                        & " " & readerOrderQueue(13) & ", " & readerOrderQueue(14) & vbCrLf & readerOrderQueue(15)
                                    'Notes
                                    .Cell(3, 1).Range.Text = .Cell(3, 1).Range.Text & " " & readerOrderQueue(1)
                                    If Not readerOrderQueue.IsDBNull(16) Then
                                        .Cell(3, 1).Range.Text = .Cell(3, 1).Range.Text & " " & readerOrderQueue(16)
                                    End If
                                End With

                                'Table 3 sales order, customer PO, and shipping method
                                objTable = .Tables.Item(3)

                                With objTable
                                    'Insert Text into table 3
                                    'SO Info
                                    .Cell(2, 1).Range.Text = readerOrderQueue(1)
                                    .Cell(2, 2).Range.Text = readerOrderQueue(2)
                                End With

                                readerOrderQueue.Close() 'close the order queue reader

                                If Not readerOrderQueue Is Nothing Then readerOrderQueue = Nothing

                                'Line item information
                                cmd.CommandText = "SELECT bL.[FINISHED_PART_NUMBER], ISNULL(bL.[SERIAL_NUMBERS], ''), " _
                    & "bL.[QUANTITY], bL.[SO_LINE_NUMBER], aol.[QUANTITY_NEEDED], ISNULL(aoL.[DESCRIPTION], ''), " _
                    & "ISNULL(aoL.[CUSTOMER_PRODUCT_NUMBER], '') " _
                    & "FROM [AOF_BOXES_LINES] AS bL " _
                    & "LEFT JOIN [AOF_ALL_ORDER_LINES] AS aoL " _
                    & "On aoL.[SO_LINE_NUMBER] = bL.[SO_LINE_NUMBER] " _
                    & "LEFT JOIN [AOF_ORDER_LINE_QUEUE] AS lQ " _
                    & "On aoL.[SO_LINE_NUMBER] = lQ.[SO_LINE_NUMBER] " _
                    & "LEFT JOIN [AOF_ORDER_QUEUE] AS oQ " _
                    & "On aoL.[SALES_ORDER_NUMBER] = oQ.[SALES_ORDER_NUMBER] " _
                    & "WHERE oQ.[SELECTED] = 'True' AND bL.[AOF_BOXES_ID] = " & boxID
                                Dim readerOrderLines As SqlDataReader = cmd.ExecuteReader()
                                Dim rstLoop As Integer = 0

                                'Line items
                                objTable = .Tables.Item(4)

                                'Insert text into table 4
                                With objTable
                                    Do While readerOrderLines.Read()
                                        .Cell(rstLoop + 2, 1).Range.Text = readerOrderLines(0) 'Item
                                        .Cell(rstLoop + 2, 2).Range.Text = readerOrderLines(6) & vbCrLf _
                            & readerOrderLines(5) & vbCrLf _
                            & readerOrderLines(1) 'Description
                                        .Cell(rstLoop + 2, 3).Range.Text = readerOrderLines(4) 'Quantity Needed (Ordered)
                                        .Cell(rstLoop + 2, 4).Range.Text = readerOrderLines(2) 'Quantity Packed (This Box)
                                        .Rows.Add()
                                        rstLoop = rstLoop + 1
                                    Loop
                                    .Rows.Last.Cells.Delete() 'Remove bottom empty row
                                End With

                                readerOrderLines.Close() 'close the order lines reader
                                If Not readerOrderLines Is Nothing Then readerOrderLines = Nothing

                                'close com objects on parent system
                                If Not objTable Is Nothing Then
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objTable)
                                End If

                                'clear objTable object
                                If Not objTable Is Nothing Then objTable = Nothing

                                'Get page count from document
                                pages = .Application.ActiveDocument.Range.Information(Word.WdInformation.wdNumberOfPagesInDocument)

                                If version = 1 Then
                                    'Determine the pack method used for this order
                                    cmd.CommandText = "rt_sp_aof_packMode"
                                    cmd.CommandType = CommandType.StoredProcedure
                                    cmd.Parameters.Add("@manualPack", SqlDbType.Bit)
                                    cmd.Parameters("@manualPack").Direction = ParameterDirection.Output
                                    cmd.ExecuteNonQuery()

                                    Dim packMode As Boolean = cmd.Parameters("@manualPack").Value

                                    If packMode = True Then
                                        objWordApp.ActivePrinter = "Manual Pack Printer"
                                    Else
                                        objWordApp.ActivePrinter = "Auto Pack Printer"
                                    End If
                                ElseIf version = 2 Then
                                    objWordApp.ActivePrinter = "Manual Pack Printer"
                                End If


                                'disable alerts
                                .Application.DisplayAlerts = False

                                'Set save path string
                                saveString = userDesktopFolder & "\Integra Packing Lists - " & soNumber & " Box " & boxID & ".pdf"

                                'Save document and set recommendation read only
                                .SaveAs2(saveString, Word.WdSaveFormat.wdFormatPDF, AddToRecentFiles:=True, ReadOnlyRecommended:=True)

                                If noprint = False Then
                                    'Print document to default printer
                                    .PrintOut()
                                End If

                                'close without saving
                                .Close(False)

                                If Not objTable Is Nothing Then objTable = Nothing

                            End With

                            'clear objDoc object
                            If Not objDoc Is Nothing Then objDoc = Nothing

                            'Open a filestream
                            Dim saveStream As New FileStream(saveString, FileMode.Open, FileAccess.Read)

                            'Dimension a variable defined by using the file size (-1 in VB)
                            Dim bytes(saveStream.Length() - 1) As Byte

                            'Read the file into the variable
                            saveStream.Read(bytes, 0, bytes.Length)

                            'Close the filestream
                            saveStream.Close()
                            If Not saveStream Is Nothing Then saveStream = Nothing

                            'Write SQL command variable
                            'Set the PDF Blob & set box as unselected
                            cmd.CommandType = CommandType.Text
                            cmd.CommandText = "UPDATE [AOF_BOXES] Set " _
                    & "[PACKING_LIST_PDF] = @PACKING_LIST_PDF, [SELECTED] = 'False', " _
                    & "[PAGES] = @PAGES WHERE [ID] = " & boxID

                            'Write SQL parameter variable for the PDF Blob
                            Dim param As New SqlParameter("@PACKING_LIST_PDF", SqlDbType.VarBinary, bytes.Length,
                    ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytes)
                            cmd.Parameters.Add(param)
                            cmd.Parameters.AddWithValue("@PAGES", pages)

                            'write parameter and execute command
                            cmd.ExecuteNonQuery()

                            If Not param Is Nothing Then param = Nothing

                        End Using
                    End Using
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    Console.ReadLine()
                Finally

                    'quit msWord
                    objWordApp.Quit()

                    'clear objWord object
                    If Not objWordApp Is Nothing Then objWordApp = Nothing

                    'close com objects on parent system
                    If Not objDoc Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objDoc)
                    End If

                    If Not objWordApp Is Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objWordApp)
                    End If

                    'If Not objTable Is Nothing Then objTable = Nothing
                    If Not objDoc Is Nothing Then objDoc = Nothing
                    If Not objWordApp Is Nothing Then objWordApp = Nothing
                    'exit application with exit code 0 (successful)
                    ' Environment.Exit(0)
                    GC.Collect()
                End Try
            Else
                Console.WriteLine("A BoxID and version must be specified to run this application")
                Console.ReadLine()
            End If
        End If
    End Sub

    Private Function HelpRequired(param As String)
        If param = "-h" Or param = "--help" Or param = "/?" Then
            Return True
        End If
        Return False
    End Function

    Private Sub DisplayHelp()
        Console.WriteLine("======================================================================")
        Console.WriteLine("Robotics Packing Slip Generator and Printer")
        Console.WriteLine("Written by: Adam S. Leven of Automation Integrity")
        Console.WriteLine("http://automationintegrity.net")
        Console.WriteLine("======================================================================")
        Console.WriteLine("Arguments:")
        Console.WriteLine("boxid - A number that represents the box this packign slip is associated to")
        Console.WriteLine("noprint - Set to to skip printing")
        Console.WriteLine("version - A number that represents version of this software to run")
        Console.WriteLine("        > Version 1: uses the ITO_PackListTemplate.dotm template")
        Console.WriteLine("        > Version 2: uses the PackListTemplate.dotm template")
        Console.WriteLine("        >            expects the COMPANY_<field> fields in AOF_ORDER_QUEUE")
        Console.WriteLine("Example:")
        Console.WriteLine("packinglistprint.exe /boxid=234 /version=1")
        Console.WriteLine("> Would produce and print a packing list for box id 234 at the appropriate printer")
        Console.WriteLine("packinglistprint.exe /boxid=56 /version=2 /noprint")
        Console.WriteLine("> Would produce and print a packing list with the appropriate company logo in the")
        Console.WriteLine("> header, and company name and contact information in the top left table cell")
        Console.WriteLine("> This version will always print at the printer named 'Manual Pack Printer'")
        Console.WriteLine("======================================================================")
        Console.WriteLine("> All packing slips will be saved as a PDF formatted file in a desktop folder.")
        Console.WriteLine("> In '%USERPROFILE%\Desktop\Packing Slip' you can view the past packing slips")
        Environment.Exit(0)
    End Sub

End Module