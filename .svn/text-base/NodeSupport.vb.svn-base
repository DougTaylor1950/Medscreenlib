Imports Intranet.intranet.customerns
Imports System.Windows.Forms
Imports MedscreenLib
Public Class ContractNodeSupport
    Private Shared instance As ContractNodeSupport

    Private Sub New()
        MyBase.new()
    End Sub

    Public Shared Sub AddContact(ByVal Node As TreeNode, ByVal clientId As String, Optional ByVal ContactType As String = "")
        If instance Is Nothing Then instance = New ContractNodeSupport()

        Dim objClient As New Intranet.intranet.customerns.Client(clientId)
        If objClient Is Nothing Then Exit Sub

        If MsgBox("Do you want to add a contact to " & objClient.SMIDProfile & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
            Dim objContact As contacts.CustomerContact = objClient.CustomerContacts.CreateCustomerContact()
            If objContact Is Nothing Then Exit Sub
            Dim frmEdCont As New MedscreenCommonGui.frmEditContact()
            If ContactType.Trim.Length > 0 Then objContact.ContactType = ContactType
            frmEdCont.Contact = objContact
            If frmEdCont.ShowDialog() = Windows.Forms.DialogResult.OK Then
                objContact = frmEdCont.Contact
                If TypeOf Node.Parent Is Treenodes.CustomerNodes.ContactsNode Then
                    Dim pNode As Treenodes.CustomerNodes.ContactsNode = Node.Parent
                    If pnode.ContactType = objContact.ContactType Then
                        pNode.ContactList.Add(objContact)
                        Dim cNode As New Treenodes.CustomerNodes.ContactNode(objContact)
                        pNode.Nodes.Add(cNode)
                    Else
                        If TypeOf pnode.Parent Is Treenodes.CustomerNodes.ContactHeaderNode Then
                            Dim cNode As Treenodes.CustomerNodes.ContactHeaderNode = pnode.Parent
                            Dim contNode As Treenodes.CustomerNodes.ContactsNode = cnode.FindContactsNode(objContact.ContactType)
                            If Not contNode Is Nothing Then
                                contNode.Refresh()
                            End If

                        End If
                    End If
                ElseIf TypeOf Node Is Treenodes.CustomerNodes.ContactsNode Then
                    Dim pNode As Treenodes.CustomerNodes.ContactsNode = Node
                    If pnode.ContactType = objContact.ContactType Then

                        pNode.ContactList.Add(objContact)
                        Dim cNode As New Treenodes.CustomerNodes.ContactNode(objContact)
                        pNode.Nodes.Add(cNode)
                    Else
                        If TypeOf pnode.Parent Is Treenodes.CustomerNodes.ContactHeaderNode Then
                            Dim cNode As Treenodes.CustomerNodes.ContactHeaderNode = pnode.Parent
                            Dim contNode As Treenodes.CustomerNodes.ContactsNode = cnode.FindContactsNode(objContact.ContactType)
                            If Not contNode Is Nothing Then
                                contNode.Refresh()
                            End If

                        End If
                    End If

                    End If

                ElseIf TypeOf Node Is Treenodes.CustomerNodes.ContactHeaderNode Then
                    'we have to find the right contactsnode based on the contact type
                    Dim cNode As Treenodes.CustomerNodes.ContactHeaderNode = Node
                    Dim contNode As Treenodes.CustomerNodes.ContactsNode = cnode.FindContactsNode(objContact.ContactType)
                    If Not contNode Is Nothing Then
                        contNode.Refresh()
                    End If

                End If
            End If




    End Sub

    Public Shared Sub EditContact(ByVal contact As contacts.CustomerContact)
        Dim frmEdCont As New MedscreenCommonGui.frmEditContact()
        frmEdCont.Contact = contact
        frmEdCont.ShowDialog()

    End Sub
End Class
