'$Revision: 1.7 $
'$Author: taylor $
'$Date: 2005-06-16 16:55:20+01 $
'$Log: clsJobs.vb,v $
Imports Intranet.intranet.sample
Imports Intranet.Intranet.jobs
Imports Intranet.intranet.support
'Imports Intranet.Customer
Public Class CCToolAppMessages
    Private Shared myCollectionMessages As Collections.Specialized.NameValueCollection
    Private Shared myRoutinePayments As Collections.Specialized.NameValueCollection
    Private Shared myPanelServices As Collections.Specialized.NameValueCollection
    Private Shared myPanelDisItemCodes As Collections.Specialized.NameValueCollection
    Private Shared myPanelStdItemCodes As Collections.Specialized.NameValueCollection
    Private Shared myPanelVisibility As Collections.Specialized.NameValueCollection
    Private Shared myTestPrices As Collections.Specialized.NameValueCollection
    Private Shared myPanelLabel As Collections.Specialized.NameValueCollection
    Private Shared myTreeIcons As Collections.Specialized.NameValueCollection
    Public Shared ReadOnly Property CollectionMessages() As Specialized.NameValueCollection
        Get
            If myCollectionMessages Is Nothing Then
                myCollectionMessages = System.Configuration.ConfigurationSettings.GetConfig("CCMMessages/CollectionMessages")
            End If
            Return myCollectionMessages
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Routine Line Items used for routine payments
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property RoutineLineItems() As Specialized.NameValueCollection
        Get
            If myRoutinePayments Is Nothing Then
                myRoutinePayments = System.Configuration.ConfigurationSettings.GetConfig("RoutinePayments/LineItems")
            End If
            Return myRoutinePayments
        End Get
    End Property

    Public Shared ReadOnly Property PanelServices() As Specialized.NameValueCollection
        Get
            If myPanelServices Is Nothing Then
                myPanelServices = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/Services")
            End If
            Return myPanelServices
        End Get
    End Property

    Public Shared ReadOnly Property PanelLabel() As Specialized.NameValueCollection
        Get
            If myPanelLabel Is Nothing Then
                myPanelLabel = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/Labels")
            End If
            Return myPanelLabel
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List of tree icons
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	23/06/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property TreeIcons() As Specialized.NameValueCollection
        Get
            If myTreeIcons Is Nothing Then
                myTreeIcons = System.Configuration.ConfigurationSettings.GetConfig("Tree/TreeIcons")
            End If
            Return myTreeIcons
        End Get
    End Property


    Public Shared ReadOnly Property PanelVisibility() As Specialized.NameValueCollection
        Get
            If myPanelVisibility Is Nothing Then
                myPanelVisibility = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/PanelVisibility")
            End If
            Return myPanelVisibility
        End Get
    End Property

    Public Shared ReadOnly Property TestPrices() As Specialized.NameValueCollection
        Get
            If myTestPrices Is Nothing Then
                myTestPrices = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/TestPrices")
            End If
            Return myTestPrices
        End Get
    End Property

    Public Shared ReadOnly Property PanelDiscountLineItems() As Specialized.NameValueCollection
        Get
            If myPanelDisItemCodes Is Nothing Then
                myPanelDisItemCodes = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/ItemCodesDis")
            End If
            Return myPanelDisItemCodes
        End Get
    End Property

    Public Shared ReadOnly Property PanelStandardLineItems() As Specialized.NameValueCollection
        Get
            If myPanelStdItemCodes Is Nothing Then
                myPanelStdItemCodes = System.Configuration.ConfigurationSettings.GetConfig("DiscountPanels/ItemCodesStd")
            End If
            Return myPanelStdItemCodes
        End Get
    End Property


End Class
