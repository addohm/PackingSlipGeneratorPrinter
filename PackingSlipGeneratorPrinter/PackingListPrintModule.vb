Imports System.Data.SqlClient 'needed for DB interactions
Imports System.IO 'needed for BLOB
Imports Word = Microsoft.Office.Interop.Word 'needed for COM object interaction with MS Word

Module PackingListPrintModule

    'Establish application path, replace appPath on deployment
    Dim appPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

    Dim appRootFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Integra Optics")
    Dim userAppFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Integra Optics", "Packing Slip")
    Dim userDesktopFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Packing Slips")
    Public Property CustomLocation As String
    Dim logFile As New FileLogTraceListener
    Dim logFileLoc As LogFileLocation

    'set form level declarations
    Dim objWordApp As New Word.Application

    Dim objDoc As Word.Document
    Dim objTable As Word.Table
    Dim errorPosition As String
    Dim boxID As Integer = Convert.ToInt32(argFromCommandLine("boxID"))
    Dim verboseLogging As Boolean = Convert.ToBoolean(argFromCommandLine("logging"))

    '*********** NOT IMPLEMENTED DUE TO PERMISSIONS FAILURE*******
    '*************************************************************
    'NAME:          WriteToEventLog
    'PURPOSE:       Write to Event Log
    'PARAMETERS:    Entry - Value to Write
    '               AppName - Name of Client Application. Needed
    '               because before writing to event log, you must
    '               have a named EventLog source.
    '               EventType - Entry Type, from EventLogEntryType
    '               Structure e.g., EventLogEntryType.Warning,
    '               EventLogEntryType.Error
    '               LogNam1e: Name of Log (System, Application;
    '               Security is read-only) If you
    '               specify a non-existent log, the log will be
    '               created
    'RETURNS:       True if successful
    '*************************************************************
    Public Function WriteToEventLog(ByVal entry As String,
        Optional ByVal appName As String = "Integra Optics",
        Optional ByVal eventType As _
        EventLogEntryType = EventLogEntryType.Information,
        Optional ByVal logName As String = "Robotics") As Boolean

        Dim objEventLog As New EventLog

        Try

            'Register the Application as an Event Source
            If Not EventLog.SourceExists(appName) Then
                EventLog.CreateEventSource(appName, logName)
            End If

            'log the entry
            objEventLog.Source = appName
            objEventLog.WriteEntry(entry, eventType)

            Return True
        Catch Ex As Exception

            Return False

        End Try

    End Function

    '*************************************************************
    'NAME:          WriteToErrorLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    msg - message to be written to error file
    '               stkTrace - stack trace from error message
    '               title - title of the error file entry
    'RETURNS:       Nothing
    '*************************************************************
    Public Sub WriteToErrorLog(ByVal level As Int32, ByVal msg As String,
           Optional ByVal stkTrace As String = "", Optional ByVal title As String = "")

        If Not Directory.Exists(userAppFolder & "\Errors\") Then
            Directory.CreateDirectory(userAppFolder & "\Errors\")
        End If

        Dim errLogPath As String = userAppFolder & "\Errors\" & boxID & " - ErrorLog.txt"

        'check the file
        Dim fs As FileStream = New FileStream(errLogPath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        If Not s Is Nothing Then s = Nothing
        If Not fs Is Nothing Then fs = Nothing

        'log it
        Dim fs1 As FileStream = New FileStream(errLogPath, FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        If level = 1 Then
            s1.Write("Title: " & title & vbCrLf)
            s1.Write("Message: " & msg & vbCrLf)
            s1.Write("StackTrace: " & stkTrace & vbCrLf)
            s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
            s1.Write("================================================" & vbCrLf)
        ElseIf level = 2 Then
            s1.Write("Message: " & msg & vbCrLf)
            s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
            s1.Write("================================================" & vbCrLf)
        End If

        s1.Close()
        fs1.Close()

        If Not s1 Is Nothing Then s1 = Nothing
        If Not fs1 Is Nothing Then fs1 = Nothing

    End Sub

    '*************************************************************
    'NAME:          if verboseLogging = True then WriteToMessageLog
    'PURPOSE:       Open or create an message log and submit general message
    'PARAMETERS:    msg - message to be written to error file
    'RETURNS:       Nothing
    '*************************************************************
    Public Sub WriteToMessageLog(ByVal msg As String)

        'Check for and\or create the logs directory
        If Not Directory.Exists(userAppFolder & "\Logs\") Then
            Directory.CreateDirectory(userAppFolder & "\Logs\")
        End If

        'Set full log path
        Dim logPath As String = userAppFolder & "\Logs\" & boxID & " - RunLog.txt"

        'Delete files after 90 days
        Dim orderedFiles = New DirectoryInfo(userAppFolder & "\Logs\").GetFiles().OrderBy(Function(x) x.CreationTime)
        For Each f As FileInfo In orderedFiles
            If (Now - f.CreationTime).Days > 90 Then f.Delete()
        Next

        'check the file
        Dim logCheckFileStream As FileStream = New FileStream(logPath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim logCheckStreamWriter As StreamWriter = New StreamWriter(logCheckFileStream)

        logCheckStreamWriter.Close()
        logCheckFileStream.Close()

        If Not logCheckFileStream Is Nothing Then logCheckFileStream = Nothing
        If Not logCheckStreamWriter Is Nothing Then logCheckStreamWriter = Nothing

        'log it
        Dim logFileStream As FileStream = New FileStream(logPath, FileMode.Append, FileAccess.Write)
        Dim logStreamWriter As StreamWriter = New StreamWriter(logFileStream)

        logStreamWriter.Write(DateTime.Now.ToString() & " - Message: " & msg & vbCrLf)

        logStreamWriter.Close()
        logFileStream.Close()

        If Not logStreamWriter Is Nothing Then logStreamWriter = Nothing
        If Not logFileStream Is Nothing Then logFileStream = Nothing

    End Sub

    '*************************************************************
    'NAME:          argFromCommandLine
    'PURPOSE:       Parse command line arguments
    'PARAMETERS:    argName - Argument preceeded by a / and followed by a =
    'RETURNS:       Integer
    '*************************************************************

    Function argFromCommandLine(argName As String) As Integer

        Dim inputArg As String = "/" & argName & "="
        Dim inputVal As String = ""

        For Each s As String In My.Application.CommandLineArgs
            If s.ToLower.StartsWith(inputArg.ToLower) Then
                inputVal = s.Remove(0, inputArg.Length)
            End If

        Next

        If inputVal = "" Then
            inputVal = "0"
        End If

        argFromCommandLine = inputVal

    End Function

    '*************************************************************
    'NAME:          cleanedText
    'PURPOSE:       Clean everything out of a text string the follows a \
    'PARAMETERS:    text - String to clean
    '               separator - character to start cleaning from
    'RETURNS:       String
    '*************************************************************

    Public Function cleanedText(text As String, separator As String) As String

        Dim startIndex As Int16 = 0

        startIndex = InStr(1, text, separator)
        cleanedText = text.Remove(0, startIndex)

    End Function

    '*************************************************************
    'NAME:          getSettings
    'PURPOSE:       Read settings in from a text file
    'PARAMETERS:    path - Path to the file containing the settings
    '               fileName - Name of the file containing the settings
    'RETURNS:       Object array
    '*************************************************************

    Public Function getSettings(ByVal path As String, ByVal fileName As String) As Object
        'get config information from ini file

        Dim txt As String
        Dim i As Int16 = 0
        Dim x As Int16 = 0

        Dim sqlSettings() As Object = Nothing

        Try
            Dim fso As StreamReader = My.Computer.FileSystem.OpenTextFileReader(path & "\" & fileName)
            'ignore comments (begind with hyphen)
            'accept setting as variable in array (begins after trailing space after colon)
            Do Until fso.EndOfStream
                txt = fso.ReadLine
                ReDim Preserve sqlSettings(i)
                For x = 1 To Len(txt)
                    If Mid(txt, x, 1) = "-" Then
                        Exit For
                    End If
                    If Mid(txt, x, 1) = ":" Then
                        sqlSettings(i) = sqlSettings(i) + Mid(txt, x + 2, Len(txt))
                        Exit For
                    End If
                Next
                i = i + 1
            Loop
        Catch ex As Exception

            WriteToErrorLog(1, "Error", ex.Message, "Catch All - Main")
        Finally

            getSettings = sqlSettings

        End Try

    End Function

    Function setCustomLogLocation() As String
        logFile.CustomLocation = appPath
        setCustomLogLocation = appPath
    End Function

    'form load subroutine
    Sub Main()

        Try

            'Run application in foreground or background.
            'If in background (false), be sure to add objDoc.close() and objWordApp.Quit()
            objWordApp.Visible = False

            'Declarations
            Dim localDateTimeString As String = DateTime.Now.ToString
            Dim localDateTimeFileName As String = DateTime.Now.ToString("yyyyMMddhhmm")
            Dim soNumber As String = "No_SO_Selected"
            Dim sqlServer, sqlDBName, sqlUserName, sqlPassword As String
            Console.WriteLine("Box ID set to: " & boxID)
            'My.Application.Log.WriteEntry("Box ID set to: " & boxID)
            'Dim remaining As Integer = argFromCommandLine("remaining")
            Dim saveString As String 'file name format
            Dim sqlSettings As Object = New Object
            Dim connection As SqlConnection = New SqlConnection() 'set SQL server connection string
            Dim pages As Integer = 0

            'Add save folder if it doesn't exist
            Console.WriteLine("Checking that save location exists")
            If verboseLogging = True Then WriteToMessageLog("Checking that save location exists")

            If (Not Directory.Exists(userDesktopFolder)) Then
                Directory.CreateDirectory(userDesktopFolder)
            End If

            'SQL Server Connection
            Console.WriteLine("Connecting to DB")
            If verboseLogging = True Then WriteToMessageLog("Connecting to DB")

            'My.Application.Log.WriteEntry("Connecting to DB")

            sqlSettings = getSettings("C:\RT Engineering", "dbSettings.ini")
            'sqlSettings = getSettings("C:\RT Engineering", "dbSettings_local.ini")

            sqlServer = sqlSettings(0)
            sqlDBName = sqlSettings(1)
            sqlUserName = sqlSettings(2)
            sqlPassword = sqlSettings(3)

            connection.ConnectionString = "Data Source=" & sqlServer _
            & ";Initial Catalog=" & sqlDBName _
            & ";user id=" & sqlUserName _
            & ";password=" & sqlPassword

            If verboseLogging = True Then WriteToMessageLog("Read in settings as: " _
                & sqlSettings(0) & ", " & sqlSettings(1) & ", " & sqlSettings(2) & ", " & sqlSettings(3))

            connection.Open() 'open connection

            'Open an existing document.
            Console.WriteLine("Opening Template")
            If verboseLogging = True Then WriteToMessageLog("Opening Template" & appPath & "\ITO_PackListTemplate.dotm")
            'My.Application.Log.WriteEntry("Opening Template")

            'Open the template
            objDoc = objWordApp.Documents.Open(appPath & "\ITO_PackListTemplate.dotm", [ReadOnly]:=True)

            'set word document as active
            objDoc = objWordApp.ActiveDocument
            With objDoc

                'table manipulation
                Console.WriteLine("Populating tables...")
                If verboseLogging = True Then WriteToMessageLog("Populating tables...")
                'My.Application.Log.WriteEntry("Populating tables")

                'Sales order information
                Dim cmd As SqlCommand = New SqlCommand("SELECT [SALES_ORDER_NUMBER], ISNULL([CUSTOMER_PO], ''), ISNULL([SHIP_VIA], ''), " _
                & "ISNULL([DELIVERY_COMPANY], ''), ISNULL([DELIVERY_ATTN_TO], ''), ISNULL([DELIVERY_STREET], ''), ISNULL([DELIVERY_CITY], ''), " _
                & "ISNULL([DELIVERY_STATE], ''), ISNULL([DELIVERY_POSTAL_CODE], ''), ISNULL([DELIVERY_COUNTRY], ''), ISNULL([BILLING_COMPANY], ''), " _
                & "ISNULL([BILLING_STREET], ''), ISNULL([BILLING_CITY], ''), ISNULL([BILLING_STATE], ''), ISNULL([BILLING_POSTAL_CODE], ''), " _
                & "ISNULL([BILLING_COUNTRY], ''), ISNULL([SALES_ORDER_NOTES], '') FROM [AOF_ORDER_QUEUE] " _
                & "WHERE [SELECTED] = 'True' ", connection)
                Dim readerOrderQueue As SqlDataReader = cmd.ExecuteReader()

                If Not readerOrderQueue.Read() = False Then 'check data exists in the reader

                    'Table 1 company information and sales order number
                    Console.WriteLine("Populating table 1")
                    If verboseLogging = True Then WriteToMessageLog("Populating table 1")
                    'My.Application.Log.WriteEntry("Populating table 1")

                    objTable = .Tables.Item(1) 'select table 1

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

                    'Table 2 shipping information and billing information
                    Console.WriteLine("Populating table 2")
                    If verboseLogging = True Then WriteToMessageLog("Populating table 2")
                    'My.Application.Log.WriteEntry("Populating table 2")

                    objTable = .Tables.Item(2)
                    With objTable
                        'Insert Text into table 2
                        'Ship to
                        .Cell(2, 1).Range.Text = readerOrderQueue(3) 'Company name
                        If Not readerOrderQueue.IsDBNull(4) Then
                            .Cell(2, 1).Range.Text = .Cell(2, 1).Range.Text & readerOrderQueue(4) 'Attn to if exists
                        End If
                        .Cell(2, 1).Range.Text = .Cell(2, 1).Range.Text & readerOrderQueue(5) & 'The rest
                        vbCrLf & readerOrderQueue(6) & ", " & readerOrderQueue(7) & " " & readerOrderQueue(8) &
                        vbCrLf & readerOrderQueue(9)
                        'Bill to
                        .Cell(2, 2).Range.Text = readerOrderQueue(10) & vbCrLf & readerOrderQueue(11) &
                    vbCrLf & readerOrderQueue(12) & " " & readerOrderQueue(13) & ", " & readerOrderQueue(14) &
                    vbCrLf & readerOrderQueue(15)
                        'Notes
                        .Cell(3, 1).Range.Text = .Cell(3, 1).Range.Text & " " & readerOrderQueue(1)
                        If Not readerOrderQueue.IsDBNull(16) Then
                            .Cell(3, 1).Range.Text = .Cell(3, 1).Range.Text & " " & readerOrderQueue(16)
                        End If
                    End With

                    'Table 3 sales order, customer PO, and shipping method
                    Console.WriteLine("Populating table 3")
                    If verboseLogging = True Then WriteToMessageLog("Populating table 3")
                    'My.Application.Log.WriteEntry("Populating table 3")

                    objTable = .Tables.Item(3)
                    With objTable
                        'Insert Text into table 3
                        'SO Info
                        .Cell(2, 1).Range.Text = readerOrderQueue(1)
                        .Cell(2, 2).Range.Text = readerOrderQueue(2)
                    End With
                Else
                    WriteToErrorLog(1, "No order is selected in the AOF_ORDER_QUEUE (readerOrderQueue)", , "Initial order query failure")
                End If

                readerOrderQueue.Close() 'close the order queue reader

                If Not readerOrderQueue Is Nothing Then readerOrderQueue = Nothing
                If Not cmd Is Nothing Then cmd = Nothing

                'Line item information
                Dim cmdOrderLines As SqlCommand = New SqlCommand("SELECT bL.[FINISHED_PART_NUMBER], ISNULL(bL.[SERIAL_NUMBERS], ''), " _
                & "bL.[QUANTITY], bL.[SO_LINE_NUMBER], aol.[QUANTITY_NEEDED], ISNULL(aoL.[DESCRIPTION], ''), " _
                & "ISNULL(aoL.[CUSTOMER_PRODUCT_NUMBER], '') " _
                & "FROM [AOF_BOXES_LINES] AS bL " _
                & "LEFT JOIN [AOF_ALL_ORDER_LINES] AS aoL " _
                & "On aoL.[SO_LINE_NUMBER] = bL.[SO_LINE_NUMBER] " _
                & "LEFT JOIN [AOF_ORDER_LINE_QUEUE] AS lQ " _
                & "On aoL.[SO_LINE_NUMBER] = lQ.[SO_LINE_NUMBER] " _
                & "LEFT JOIN [AOF_ORDER_QUEUE] AS oQ " _
                & "On aoL.[SALES_ORDER_NUMBER] = oQ.[SALES_ORDER_NUMBER] " _
                & "WHERE oQ.[SELECTED] = 'True' AND bL.[AOF_BOXES_ID] = " & boxID, connection)
                Dim readerOrderLines As SqlDataReader = cmdOrderLines.ExecuteReader()
                Dim rstLoop As Integer = 0

                'Line items
                Console.WriteLine("Populating table 4")
                If verboseLogging = True Then WriteToMessageLog("Populating table 4")
                'My.Application.Log.WriteEntry("Populating table 4")

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
                If Not cmdOrderLines Is Nothing Then cmdOrderLines = Nothing

                'clear objTable object
                If Not objTable Is Nothing Then objTable = Nothing

                'Determine the pack method used for this order

                Dim cmdTotalQuantity As SqlCommand = New SqlCommand("rt_sp_aof_packMode", connection)
                cmdTotalQuantity.CommandType = CommandType.StoredProcedure
                cmdTotalQuantity.Parameters.Add("@manualPack", SqlDbType.Bit)
                cmdTotalQuantity.Parameters("@manualPack").Direction = ParameterDirection.Output
                Try
                    cmdTotalQuantity.ExecuteNonQuery()
                Catch ex As Exception
                    'MsgBox(ex.Message)
                    WriteToErrorLog(1, "The query that determines the quantity of clamshells being packed failed", , "Clamshell quantity query failure")
                End Try

                Dim packMode As Boolean = cmdTotalQuantity.Parameters("@manualPack").Value

                If packMode = True Then
                    objWordApp.ActivePrinter = "Manual Pack Printer"
                Else
                    objWordApp.ActivePrinter = "Auto Pack Printer"
                End If

                If Not cmdTotalQuantity Is Nothing Then cmdTotalQuantity = Nothing

                ''Run application macro that sets [PAGES] value in [BOXES] table
                'Console.WriteLine("Updating [PAGES] value In DB")
                '.Application.Run("updatePages", boxID) 'runs updatePages macro with the boxID as it's parameter

                'Get page count from document

                pages = .Application.ActiveDocument.Range.Information(Word.WdInformation.wdNumberOfPagesInDocument)

                'disable alerts
                .Application.DisplayAlerts = False

                'Set save path string

                saveString = userDesktopFolder & "\Integra Packing Lists - " & soNumber & " Box " & boxID & ".pdf"
                Console.WriteLine("Saving document" & saveString)
                If verboseLogging = True Then WriteToMessageLog("Saving document" & saveString)
                'My.Application.Log.WriteEntry("Saving document" & saveString)

                'Save document and set recommendation read only

                .SaveAs2(saveString, Word.WdSaveFormat.wdFormatPDF, AddToRecentFiles:=True, ReadOnlyRecommended:=True)

                'Print document to default printer
                Console.WriteLine("Printing " & pages & " pages")
                If verboseLogging = True Then WriteToMessageLog("Printing " & pages & " pages")
                'My.Application.Log.WriteEntry("Printing " & pages & " pages")

                .PrintOut()

                'close without saving
                Console.WriteLine("Closing word document")
                If verboseLogging = True Then WriteToMessageLog("Closing word document")
                'My.Application.Log.WriteEntry("Closing word document")

                .Close(False)

                If Not objTable Is Nothing Then objTable = Nothing

            End With

            'clear objDoc object
            If Not objDoc Is Nothing Then objDoc = Nothing

            'quit msWord
            Console.WriteLine("Quitting MS Word")
            If verboseLogging = True Then WriteToMessageLog("Quitting MS Word")
            'My.Application.Log.WriteEntry("Quitting MS Word")

            objWordApp.Quit()

            'clear objWord object
            If Not objWordApp Is Nothing Then objWordApp = Nothing

            Console.WriteLine("Writing PDF Data To DB")
            If verboseLogging = True Then WriteToMessageLog("Writing PDF Data To DB")
            'My.Application.Log.WriteEntry("Writing PDF Data To DB")

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
            Dim cmdStoreBlob As New SqlCommand("UPDATE [AOF_BOXES] Set " _
            & "[PACKING_LIST_PDF] = @PACKING_LIST_PDF, [SELECTED] = 'False', " _
            & "[PAGES] = @PAGES WHERE [ID] = " & boxID, connection)

            'Write SQL parameter variable for the PDF Blob
            Dim param1 As New SqlParameter("@PACKING_LIST_PDF", SqlDbType.VarBinary, bytes.Length,
                ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytes)
            'Write SQL parameter variable for the Page Count
            Dim param2 As New SqlParameter("@PAGES", SqlDbType.Int)
            param2.Value = pages

            'write parameter and execute command
            cmdStoreBlob.Parameters.Add(param1)
            cmdStoreBlob.Parameters.Add(param2)
            cmdStoreBlob.ExecuteNonQuery()

            If Not param1 Is Nothing Then param1 = Nothing
            If Not param2 Is Nothing Then param2 = Nothing
            If Not cmdStoreBlob Is Nothing Then cmdStoreBlob = Nothing

            'close SQL server connection
            connection.Close()
            If Not connection Is Nothing Then connection = Nothing

            'close binary reader and file stream
            Console.WriteLine("Cleaning Up")
            If verboseLogging = True Then WriteToMessageLog("Cleaning Up...")
            'My.Application.Log.WriteEntry("Cleaning Up")

            Console.WriteLine("Releasing COM objects")
            If verboseLogging = True Then WriteToMessageLog("Releasing COM objects")
            'My.Application.Log.WriteEntry("Releasing COM objects")
        Catch ex As Exception

            WriteToErrorLog(1, "Error", ex.Message, "Catch All - Main")
        Finally

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
            Environment.Exit(0)

        End Try
    End Sub

End Module