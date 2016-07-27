
''' <summary>
''' Library of routines to support Crystal Reports
''' </summary>
''' <remarks></remarks>
''' <revisionHistory></revisionHistory>
''' <author></author>
Public Class CrystalSupport
    ''' <developer></developer>
    ''' <summary>
    ''' Log into Crystal and get report
    ''' </summary>
    ''' <param name="ReportSource"></param>
    ''' <param name="ReportDirectory"></param>
    ''' <param name="UserName"></param>
    ''' <param name="Password"></param>
    ''' <param name="instance"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    <CLSCompliant(False)> _
Public Shared Function LogInCrystal(ByVal ReportSource As String, ByVal ReportDirectory As String, Optional ByVal UserName As String = "", Optional ByVal Password As String = "", Optional ByVal instance As String = "", Optional ByVal noLogin As Boolean = False) As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim objTab As CrystalDecisions.CrystalReports.Engine.Table
        Dim Cr As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing
        Dim objLogInf As CrystalDecisions.Shared.TableLogOnInfo = Nothing

        Try
            Dim ReportPath As String = ReportSource & ReportDirectory
            Medscreen.LogAction("Creating report " & ReportPath)
            Cr = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            Medscreen.LogAction("Loading report " & ReportPath)

            Cr.Load(ReportPath)
            If Not noLogin Then
                Medscreen.LogAction(" Setting log on ")
                For Each objTab In Cr.Database.Tables
                    objLogInf = objTab.LogOnInfo
                    Debug.Print(objLogInf.ConnectionInfo.DatabaseName & " - " & objLogInf.ConnectionInfo.Password)
                    If instance.Trim.Length > 0 Then
                        objLogInf.ConnectionInfo.ServerName = instance
                    Else
                        objLogInf.ConnectionInfo.ServerName = CConnection.DBDatabase
                    End If
                    If UserName.Trim.Length > 0 Then
                        If UserName = "NULL" Then
                            objLogInf.ConnectionInfo.UserID = ""

                        Else
                            objLogInf.ConnectionInfo.UserID = UserName

                        End If
                    Else
                        objLogInf.ConnectionInfo.UserID = CConnection.DBUserName
                    End If
                    If Password.Trim.Length > 0 Then
                        If Password = "NULL" Then
                            objLogInf.ConnectionInfo.Password = ""

                        Else
                            objLogInf.ConnectionInfo.Password = Password
                        End If
                    Else
                        objLogInf.ConnectionInfo.Password = MedConnection.Instance.SID
                    End If
                    objTab.ApplyLogOnInfo(objLogInf)
                Next
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , objLogInf.ToString)
            'Throw ex                'Rethrow exception
        End Try

        Return Cr
    End Function

    ''' <developer></developer>
    ''' <summary>
    ''' Convert a date or metadate to date
    ''' </summary>
    ''' <param name="inputDate"></param>
    ''' <param name="NullDateDefault"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><modified>03-Apr-2012 15:10</modified><Author>CONCATENO\taylor</Author>Last working day and first working day added</revision></revisionHistory>
    Public Shared Function DateParameter(ByVal inputDate As String, Optional ByVal NullDateDefault As String = "PREVMONTHSTART", Optional ByVal BaseDate As String = "") As Date
        Dim tmpDate As Date = Today
        If BaseDate.Trim.Length > 0 Then
            Try
                tmpDate = Date.Parse(BaseDate)
            Catch ex As Exception

            End Try
        End If
        Dim inDate As String = inputDate.ToUpper
        If inDate.Trim.Length = 0 Then
            If Char.IsNumber(NullDateDefault(0)) OrElse NullDateDefault(0) = "-" OrElse NullDateDefault(0) = "+" Then       'Deal with passed change in date
                Dim multiplier As Integer = CInt(NullDateDefault)
                inDate = Now.AddDays(multiplier).ToString("dd-MMM-yyyy")
            Else
                inDate = NullDateDefault
            End If
        End If




        If inDate = "YEARSTART" OrElse inDate = "STARTOFYEAR" Then
            tmpDate = DateSerial(tmpDate.Year, 1, 1)
        ElseIf inDate = "YEAREND" Then
            tmpDate = DateSerial(tmpDate.Year, 12, 31)
        ElseIf inDate = "PREVMONTH" Then
            tmpDate = tmpDate.AddMonths(-1)
        ElseIf inDate = "PREVMONTHSTART" Then
            Dim td As Date = tmpDate.AddMonths(-1)
            tmpDate = DateSerial(td.Year, td.Month, 1)
        ElseIf inDate = "NEXTMONTH" Then
            tmpDate = tmpDate.AddMonths(+1)
        ElseIf inDate = "MONTHEND" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(1).AddDays(-1)
        ElseIf inDate = "MONTHSTART" OrElse inDate = "STARTOFMONTH" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1)
        ElseIf inDate = "LASTWEEK" Then
            tmpDate = tmpDate.AddDays(-7)
        ElseIf inDate = "NEXTWEEK" Then
            tmpDate = tmpDate.AddDays(7)
        ElseIf inDate = "PREVFORTNT" Then
            tmpDate = tmpDate.AddDays(-14)
        ElseIf inDate = "NEXTFORTNT" Then
            tmpDate = tmpDate.AddDays(14)
        ElseIf inDate = "TODAY" Then
            tmpDate = Today
        ElseIf inDate = "NOW" Then
            tmpDate = Now
        ElseIf inDate = "YESTERDAY" Then
            tmpDate = Today.AddDays(-1)
        ElseIf inDate = "STARTLASTY" Then
            tmpDate = DateSerial(tmpDate.Year - 1, 1, 1)
        ElseIf inDate = "STARTLASTM" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(-1)
        ElseIf inDate = "ENDLASTMON" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddDays(-1)
        ElseIf inDate = "ENDNEXTMON" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(2).AddDays(-1)
        ElseIf inDate = "NXTMNWRKDY" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(2).AddDays(-1)
            While Medscreen.IsAHoliday(tmpDate) Or tmpDate.DayOfWeek = DayOfWeek.Saturday Or DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(-1)
            End While
        ElseIf inDate = "LASTWRKDAY" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(1).AddDays(-1)
            While Medscreen.IsAHoliday(tmpDate) Or tmpDate.DayOfWeek = DayOfWeek.Saturday Or DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(-1)
            End While
        ElseIf inDate = "FRSTWRKDAY" Then
            tmpDate = DateSerial(tmpDate.Year, tmpDate.Month, 1).AddMonths(1)
            While Medscreen.IsAHoliday(tmpDate) Or tmpDate.DayOfWeek = DayOfWeek.Saturday Or DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(1)
            End While
        ElseIf inDate = "NEXTWRKDAY" Then
            tmpDate = Today.AddDays(1)
            While Medscreen.IsAHoliday(tmpDate) Or tmpDate.DayOfWeek = DayOfWeek.Saturday Or tmpDate.DayOfWeek = DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(1)
            End While
        ElseIf inDate = "NXTWEEKDAY" Then
            tmpDate = Today.AddDays(1)
            While tmpDate.DayOfWeek = DayOfWeek.Saturday Or tmpDate.DayOfWeek = DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(1)
            End While

        Else
            Try
                tmpDate = Date.Parse(inDate)
            Catch ex As Exception
                tmpDate = Now.AddDays(-30)
            End Try
        End If

        Return tmpDate
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Export as crystal report to appropriate file type
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <param name="SendMethod"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Shared Function ExportReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal SendMethod As Constants.SendMethod) As String
        Dim tmpFileName As String = ""
        Try
            If SendMethod = Constants.SendMethod.Email Then
                tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "DOC")
                ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.WordForWindows, tmpFileName)
            ElseIf SendMethod = Constants.SendMethod.PDF Then
                tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "PDF")
                ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, tmpFileName)
            ElseIf SendMethod = Constants.SendMethod.Excel Then
                tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "XLS")
                ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.ExcelRecord, tmpFileName)
            ElseIf SendMethod = Constants.SendMethod.RTF Then
                tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "RTF")
                ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.RichText, tmpFileName)
            ElseIf SendMethod = Constants.SendMethod.HTML Then
                tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "HTM")
                ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpFileName)
            End If
        Catch
        End Try
        Return tmpFileName


    End Function
    <CLSCompliant(False)> _
    Public Shared Sub ExportToDisk(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal opt As CrystalDecisions.Shared.ExportFormatType, ByVal destination As String)
        Try
            Dim expOpt As CrystalDecisions.Shared.ExportOptions = cr.ExportOptions
            expOpt.ExportFormatType = opt
            If opt = CrystalDecisions.Shared.ExportFormatType.ExcelRecord Or opt = CrystalDecisions.Shared.ExportFormatType.Excel Then
                'Dim ExclOptions As CrystalDecisions.Shared.ExcelDataOnlyFormatOptions = CrystalDecisions.Shared.ExportOptions.CreateDataOnlyExcelFormatOptions
                'ExclOptions.ExportPageHeaderAndPageFooter = True
                'ExclOptions.SimplifyPageHeaders = True
                'expOpt.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.ExcelRecord
                'expOpt.ExportFormatOptions = ExclOptions

            End If
            expOpt.ExportDestinationType = CrystalDecisions.[Shared].ExportDestinationType.DiskFile
            Dim diskexport As New CrystalDecisions.Shared.DiskFileDestinationOptions()
            diskexport.DiskFileName = destination
            expOpt.DestinationOptions = diskexport
            cr.ExportToDisk(opt, destination)
        Catch ex As Exception
            Medscreen.EmailError(ex)
            'Medscreen.LogError(ex, , "Exporting -" & destination & "-" & cr.FilePath & "-")
        End Try
    End Sub

End Class
