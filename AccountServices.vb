Imports MedscreenLib
Public Class AccountServices
    Private Shared Acco As AccountTypeServices

    Public Shared Function AccountTypeServicesLt(ByVal AccType As String) As AccountTypeServices.ServiceRow()
        Dim acRow As AccountTypeServices.AccTypeRow
        Dim myRet As AccountTypeServices.ServiceRow()
        For Each acRow In AccountTypeServiceList.AccType
            If acRow.name.ToUpper = AccType.ToUpper Then
                Dim aSerRow As AccountTypeServices.ServicesRow
                For Each aSerRow In acRow.GetServicesRows
                    Dim aSerrRow As AccountTypeServices.ServiceRow
                    myRet = aSerRow.GetServiceRows
                Next
                Exit For
            End If
        Next
        Return myRet
    End Function

    Public Shared Property AccountTypeServiceList() As AccountTypeServices
        Get
            If Acco Is Nothing Then
                Acco = New AccountTypeServices()
                Acco.ReadXml(Medscreen.LiveRoot & Medscreen.IniFiles & "AccTypeService.xml")
            End If
            Return Acco
        End Get
        Set(ByVal Value As AccountTypeServices)
            Acco = Value
        End Set
    End Property
End Class
