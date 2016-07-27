''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : PriceBookSupport
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Routines to support price books
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Taylor]</Author><date> [06/05/2009]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class PriceBookSupport
    Private Shared myPriceDate As Date
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''Date on which the book should be displayed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Taylor]</Author><date> [06/05/2009]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property PriceDate() As Date
        Get
            Return myPriceDate
        End Get
        Set(ByVal Value As Date)
            myPriceDate = Value
        End Set
    End Property
End Class
