Imports System.Drawing
Namespace Print

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Print.PrintSupport
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Support  for printers
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	16/11/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class PrintSupport

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Find what port a printer (defined by logical name) prints too
        ''' </summary>
        ''' <param name="strPrintername"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/11/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function GetPrinterPort(ByRef strPrintername As String) As String
            Dim strPortName As String = ""

            Try  ' Protecting
                Dim wmiClass As New System.Management.SelectQuery("WIN32_PRINTER")
                Dim wmiSearcher As New Management.ManagementObjectSearcher(wmiClass)

                Dim Printer As Management.ManagementBaseObject
                For Each Printer In wmiSearcher.Get
                    'Retrieve the Device ID for the chosen printer 
                    Dim strPrinter As String = Printer.Item("DeviceID")
                    If InStr(strPrinter, strPrintername) > 0 Then       'See if we have a match with the requested printer
                        strPortName = Printer.Item("PortName")          'If we have a match set the return variable
                        'Debug.WriteLine(Printer.ToString & Printer.Item("PortName"))
                        Exit For
                    End If
                Next

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "PrintSupport-GetPrinterPort-35")
            Finally
            End Try

            Return strPortName
        End Function

        Public Shared Function PrinterStatus(ByRef strPrintername As String) As Integer
            Dim intStatus As Integer = -1

            Try  ' Protecting
                Dim wmiClass As New System.Management.SelectQuery("WIN32_PRINTER")
                Dim wmiSearcher As New Management.ManagementObjectSearcher(wmiClass)

                Dim Printer As Management.ManagementBaseObject
                For Each Printer In wmiSearcher.Get
                    'Retrieve the Device ID for the chosen printer 

                    Dim strPrinter As String = Printer.Item("DeviceID")
                    If InStr(strPrinter, strPrintername) > 0 Then       'See if we have a match with the requested printer
                        intStatus = Convert.ToInt32(Printer.Item("PrinterStatus"))          'If we have a match set the return variable
                        'Debug.WriteLine(Printer.ToString & Printer.Item("PortName"))
                        Exit For
                    End If
                Next


            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "PrintSupport-PrinterAvailability-59")
            Finally

            End Try

            Return intStatus
        End Function

        Public Shared Function PrinterStatusString(ByRef strPrintername As String) As String
            Dim strReturn As String = ""
            Dim intStatValue As Integer = PrinterStatus(strPrintername)
            Select Case intStatValue
                Case 0, -1
                    strReturn = "Code failure"
                Case 1
                    strReturn = "Other"
                Case 2
                    strReturn = "Unknown"
                Case 3
                    strReturn = "Idle"
                Case 4
                    strReturn = "Printing"
                Case 5
                    strReturn = "Warming Up"
                Case 6
                    strReturn = "Stopped"
                Case 7
                    strReturn = "Offline"

            End Select
            Return strReturn
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a OS based logical name for a printer given a Sample Manager Device NAme
        ''' </summary>
        ''' <param name="strShortName"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/11/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function FindPrinter(ByVal strShortName As String) As String
            Dim PrinterName As String = ""

            Dim oRead As OleDb.OleDbDataReader = Nothing
            Dim oCmd As New OleDb.OleDbCommand()
            Dim strLogicalName As String = ""                          'Return variable
            Try ' Protecting 
                oCmd.CommandText = "Select logical_name from printer where device_name = ?"
                oCmd.Parameters.Add(MedscreenLib.CConnection.StringParameter("DeviceID", strShortName, 10))
                oCmd.Connection = MedscreenLib.CConnection.DbConnection
                If MedscreenLib.CConnection.ConnOpen Then            'Attempt to open reader
                    oRead = oCmd.ExecuteReader          'Get Data
                    While oRead.Read                    'Loop through data
                        If Not oRead.IsDBNull(0) Then
                            strLogicalName = oRead.GetValue(0)
                            Exit While
                        End If
                    End While
                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "Report-FindPrinter-118")
            Finally
                MedscreenLib.CConnection.SetConnClosed()             'Close connection
                If Not oRead Is Nothing Then            'Try and close reader
                    If Not oRead.IsClosed Then oRead.Close()
                End If
            End Try

            Dim blnInstalled As Boolean = False
            'Dim objPrinter As System.Drawing.Printing.PrintDocument
            Try  ' Protecting
                Dim objPrinterName As String
                For Each objPrinterName In System.Drawing.Printing.PrinterSettings.InstalledPrinters
                    If InStr(objPrinterName, strLogicalName) = 1 Then
                        blnInstalled = True
                        PrinterName = strLogicalName
                        Exit For
                    End If
                Next

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "Report-FindPrinter-145")
            Finally
            End Try

            'Dim nameEnd As Short
            Return PrinterName
        End Function

    End Class
End Namespace
