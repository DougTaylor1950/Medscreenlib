Imports System.ComponentModel
Imports System.Windows.Forms
Namespace Combos
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Combos.SMIDProfileCombo
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A combo box specifically for SMID Profiles
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <ToolboxItem(True)> Public Class SMIDProfileCombo
        Inherits ComboBox
        Private objSiblings As Intranet.intranet.customerns.SMIDProfileCollection
        Private objSelectedItem As Intranet.intranet.customerns.Client
        Private objSelectedProfile As Intranet.intranet.customerns.SMIDProfile
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' create a new combo box
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action>Force the value member and identity properties</Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [23/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()
            Me.ValueMember = "CustomerID"
            Me.DisplayMember = "Info"

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' set or read sibling collection
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        '''         ''' Usually
        '''  sets the sibling collection
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [23/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Siblings() As Intranet.intranet.customerns.SMIDProfileCollection
            Get
                Return objSiblings
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.SMIDProfileCollection)

                objSiblings = Value
                Me.SelectedIndex = -1
                'Me.ListItems = New Controls.ListItems()
                Me.DataSource = objSiblings
                If Me.objSiblings Is Nothing Then Exit Property

                Me.ValueMember = "CustomerID"
                Me.DisplayMember = "Info"

                If objSiblings.Count > 0 Then
                    Me.SelectedIndex = 0
                Else
                    Me.SelectedIndex = -1
                End If
                Me.ValueMember = "CustomerID"
                Me.DisplayMember = "Info"
            End Set
        End Property

        Private Sub SMIDProfileCombo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SelectedIndexChanged
            Me.objSelectedItem = Nothing
            If Me.SelectedIndex > -1 AndAlso (Not Me.objSiblings Is Nothing) AndAlso Me.SelectedIndex <= Me.objSiblings.Count - 1 Then
                objSelectedProfile = Me.objSiblings.Item(Me.SelectedIndex)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Client selected in control
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/11/2005]</date><Action>Bug fixed in positioning code</Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [23/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property SelectedProfile() As Intranet.intranet.customerns.Client
            Get
                Return Me.objSelectedItem
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.Client)
                If Not objSelectedItem Is Nothing AndAlso Not Value Is Nothing AndAlso Value.Identity = objSelectedItem.Identity Then
                    If Not objSelectedItem Is Nothing Then Me.Text = objSelectedItem.Info
                    Exit Property
                End If
                Me.objSelectedItem = Value
                If Me.objSiblings Is Nothing Then Exit Property
                Dim objClient As Intranet.intranet.customerns.Client
                Dim i As Integer = 0
                For Each objClient In Me.objSiblings
                    If objClient.Identity = objSelectedItem.Identity Then
                        'Me.SelectedItem = i
                        Me.SelectedIndex = i
                        Me.Text = objSelectedItem.Info
                        Exit For
                    End If
                    i += 1
                Next

            End Set
        End Property
    End Class
End Namespace