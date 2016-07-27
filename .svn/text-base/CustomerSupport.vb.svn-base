Imports System.Text.RegularExpressions
''' <summary>
''' Class to provide various support utilities for customers
''' </summary>
''' <remarks></remarks>
''' <revisionHistory><revision><created>11-11-11</created></revision></revisionHistory>
''' <author>Taylor</author>
Public Class CustomerSupport
    ''' <developer></developer>
    ''' <summary>
    ''' Check the validity of a password against a mask, a default AlphaNumeric mask is supplied
    ''' </summary>
    ''' <param name="Password"></param>
    ''' <param name="mask"></param>
    ''' <returns></returns>
    ''' <remarks>If any changes are made to the default mask i.e. the addition of any special characters 
    ''' the SpellItOut function needs to have corresponding text added the text is added in the form "@| [AT],#| [HASH]" 
    ''' i.e. character | description <see cref="MedscreenLib.Medscreen.SpellItOut" /></remarks>
    ''' <revisionHistory><revision><created>11-11-11</created><Author>Taylor</Author></revision></revisionHistory>
    Public Shared Function ValidatePassword(ByVal Password As String, Optional ByVal mask As String = "") As Boolean
        Dim blnRet As Boolean = True
        'if no mask supplied get default mask from config file
        If mask.Trim.Length = 0 Then 'get default mask
            mask = MedscreenCommonGUIConfig.Misc("DefPasswordMask")
        End If
        Dim myregex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(mask)
        Dim mymatch As Match = myregex.Match(Password)
        If mymatch.Index <> 0 OrElse Not mymatch.Success Then
            blnRet = False
        End If
        Return blnRet
    End Function
End Class
