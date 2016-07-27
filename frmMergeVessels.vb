Imports MedscreenLib.Medscreen
Imports MedscreenLib
Imports Intranet.intranet
Imports Intranet.intranet.support

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmMergeVessels
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for merging two vessels together
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action>
''' $Revision: 1.0 $ <para/>
'''$Author: taylor $ <para/>
'''$Date: 2005-11-30 09:04:05+00 $ <para/>
'''$Log: frmMergeVessels.vb,v $
'''Revision 1.0  2005-11-30 09:04:05+00  taylor
'''Moved to medscreen common GUI
'''
'''Revision 1.0  2005-11-30 07:16:21+00  taylor
'''Commented
''' <para/>
''' </Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmMergeVessels
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdSwap As System.Windows.Forms.Button
    Friend WithEvents txtRtName As System.Windows.Forms.TextBox
    Friend WithEvents txtRtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtltCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtLtName As System.Windows.Forms.TextBox
    Friend WithEvents txtLtVessId As System.Windows.Forms.TextBox
    Friend WithEvents txtRightVessid As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMergeVessels))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdSwap = New System.Windows.Forms.Button()
        Me.txtRtName = New System.Windows.Forms.TextBox()
        Me.txtRtCustomer = New System.Windows.Forms.TextBox()
        Me.txtltCustomer = New System.Windows.Forms.TextBox()
        Me.txtLtName = New System.Windows.Forms.TextBox()
        Me.txtLtVessId = New System.Windows.Forms.TextBox()
        Me.txtRightVessid = New System.Windows.Forms.TextBox()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.cmdOk})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 245)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(624, 40)
        Me.Panel2.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(520, 9)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "&Cancel"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(432, 9)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "&Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Bitmap)
        Me.PictureBox1.Location = New System.Drawing.Point(240, 96)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(128, 50)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(168, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(288, 32)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Merge Left hand vessel data into right hand vessel and remove left hand vessel"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdSwap
        '
        Me.cmdSwap.Image = CType(resources.GetObject("cmdSwap.Image"), System.Drawing.Bitmap)
        Me.cmdSwap.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSwap.Location = New System.Drawing.Point(240, 48)
        Me.cmdSwap.Name = "cmdSwap"
        Me.cmdSwap.Size = New System.Drawing.Size(128, 23)
        Me.cmdSwap.TabIndex = 5
        Me.cmdSwap.Text = "Swap Vessels"
        '
        'txtRtName
        '
        Me.txtRtName.Location = New System.Drawing.Point(392, 48)
        Me.txtRtName.Name = "txtRtName"
        Me.txtRtName.Size = New System.Drawing.Size(216, 20)
        Me.txtRtName.TabIndex = 6
        Me.txtRtName.Text = ""
        '
        'txtRtCustomer
        '
        Me.txtRtCustomer.Location = New System.Drawing.Point(392, 80)
        Me.txtRtCustomer.Name = "txtRtCustomer"
        Me.txtRtCustomer.Size = New System.Drawing.Size(216, 20)
        Me.txtRtCustomer.TabIndex = 7
        Me.txtRtCustomer.Text = ""
        '
        'txtltCustomer
        '
        Me.txtltCustomer.Location = New System.Drawing.Point(16, 80)
        Me.txtltCustomer.Name = "txtltCustomer"
        Me.txtltCustomer.Size = New System.Drawing.Size(208, 20)
        Me.txtltCustomer.TabIndex = 9
        Me.txtltCustomer.Text = ""
        '
        'txtLtName
        '
        Me.txtLtName.Location = New System.Drawing.Point(16, 48)
        Me.txtLtName.Name = "txtLtName"
        Me.txtLtName.Size = New System.Drawing.Size(208, 20)
        Me.txtLtName.TabIndex = 8
        Me.txtLtName.Text = ""
        '
        'txtLtVessId
        '
        Me.txtLtVessId.Location = New System.Drawing.Point(16, 112)
        Me.txtLtVessId.Name = "txtLtVessId"
        Me.txtLtVessId.Size = New System.Drawing.Size(208, 20)
        Me.txtLtVessId.TabIndex = 10
        Me.txtLtVessId.Text = ""
        '
        'txtRightVessid
        '
        Me.txtRightVessid.Location = New System.Drawing.Point(392, 112)
        Me.txtRightVessid.Name = "txtRightVessid"
        Me.txtRightVessid.Size = New System.Drawing.Size(216, 20)
        Me.txtRightVessid.TabIndex = 11
        Me.txtRightVessid.Text = ""
        '
        'frmMergeVessels
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 285)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtRightVessid, Me.txtLtVessId, Me.txtltCustomer, Me.txtLtName, Me.txtRtCustomer, Me.txtRtName, Me.cmdSwap, Me.Label1, Me.PictureBox1, Me.Panel2})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMergeVessels"
        Me.Text = "Merge vessels "
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private myVessel1 As Intranet.intranet.customerns.Vessel
    Private myVessel2 As Intranet.intranet.customerns.Vessel

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Swap the position of the two vessels
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdSwap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSwap.Click
        Dim tmpVessel As Intranet.intranet.customerns.Vessel
        tmpVessel = Vessel1
        Vessel1 = Vessel2
        Vessel2 = tmpVessel

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' First vessel to merge
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Vessel1() As Intranet.intranet.customerns.Vessel
        Get
            Return myVessel1
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Vessel)
            myVessel1 = Value
            If Not myVessel1 Is Nothing Then
                'See if we can translate the vessel manager
                Dim strCustomer As String = CConnection.PackageStringList("lib_customer.GetSMIDProfile", myVessel1.CustomerID)
                If strCustomer.Trim.Length = 0 Then
                    Me.txtltCustomer.Text = myVessel1.CustomerID
                Else
                    Me.txtltCustomer.Text = strCustomer
                End If
                Me.txtLtName.Text = myVessel1.VesselName
                Me.txtLtVessId.Text = myVessel1.VesselID
                Me.txtLtVessId.Enabled = Not MedscreenLib.Medscreen.IsNumber(myVessel1.VesselID)
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Second vessel to merge
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Vessel2() As Intranet.intranet.customerns.Vessel
        Get
            Return myVessel2
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Vessel)
            myVessel2 = Value
            If Not myVessel2 Is Nothing Then
                Dim strCustomer As String = CConnection.PackageStringList("lib_customer.GetSMIDProfile", myVessel2.CustomerID)
                If strCustomer.Trim.Length = 0 Then
                    Me.txtRtCustomer.Text = myVessel2.CustomerID
                Else
                    Me.txtRtCustomer.Text = strCustomer
                End If
                Me.txtRtName.Text = myVessel2.VesselName
                Me.txtRightVessid.Text = myVessel2.VesselID
                Me.txtRightVessid.Enabled = Not MedscreenLib.Medscreen.IsNumber(myVessel2.VesselID)
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Do the actual merge 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim oCmd As New OleDb.OleDbCommand()
        Dim strAction As String

        'Sort out collections all collections with the old name need to be changed to the new name
        strAction = "update collection set vessel_name = '" & Me.Vessel2.VesselID.Trim & _
            "' where vessel_name = '" & Me.Vessel1.VesselID.Trim & "'"

        Try
            oCmd.Connection = MedscreenLib.CConnection.DbConnection
            oCmd.CommandText = strAction

            MedscreenLib.CConnection.SetConnOpen()
            Dim intRet As Integer = oCmd.ExecuteNonQuery
        Catch ex As Exception
            LogError(ex, , "correcting vessel ids, merge")
        Finally
            MedscreenLib.CConnection.SetConnClosed()
        End Try

        'Sort out collection_history 
        strAction = "update collection_history set vessel_name = '" & Me.Vessel2.VesselID.Trim & _
            "' where vessel_name = '" & Me.Vessel1.VesselID.Trim & "'"

        Try
            oCmd.Connection = MedscreenLib.CConnection.DbConnection
            oCmd.CommandText = strAction

            MedscreenLib.CConnection.SetConnOpen()
            Dim intRet As Integer = oCmd.ExecuteNonQuery
        Catch ex As Exception
            LogError(ex)
        Finally
            MedscreenLib.CConnection.SetConnClosed()
        End Try

        'Sort out vessel_invoicing 
        strAction = "delete from vesselinvoicing where vessel_id = '" & Me.Vessel1.VesselID.Trim & "'"

        Try
            oCmd.Connection = MedscreenLib.CConnection.DbConnection
            oCmd.CommandText = strAction

            MedscreenLib.CConnection.SetConnOpen()
            Dim intRet As Integer = oCmd.ExecuteNonQuery
        Catch ex As Exception
            LogError(ex)
        Finally
            MedscreenLib.CConnection.SetConnClosed()
        End Try

        Me.Vessel1.Removed = True
        Me.Vessel1.Update()

    End Sub
End Class
