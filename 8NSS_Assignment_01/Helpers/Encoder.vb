Imports System.Security

Public NotInheritable Class Encoder

    Public Shared saltValueBytes As Byte() = System.Text.Encoding.Unicode.GetBytes("Go On")

    Public Shared Function EncryptString(ByVal data As SecureString) As String
        Dim encryptedData As Byte() = System.Security.Cryptography.ProtectedData.Protect(System.Text.Encoding.Unicode.GetBytes(ToInsecureString(data)),
                                                                                         saltValueBytes,
                                                                                         System.Security.Cryptography.DataProtectionScope.CurrentUser)
        Return Convert.ToBase64String(encryptedData)
    End Function

    Public Shared Function DecryptString(ByVal data As String) As SecureString

        Try
            Dim decryptedData As Byte() = System.Security.Cryptography.ProtectedData.Unprotect(Convert.FromBase64String(data),
                                                                                 saltValueBytes,
                                                                                 System.Security.Cryptography.DataProtectionScope.CurrentUser)
            Return ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData))
        Catch ex As Exception
            Return New SecureString()
        End Try
    End Function

    Public Shared Function ToSecureString(ByVal data As String) As SecureString
        Dim secure As SecureString = New SecureString()
        For Each c As Char In data
            secure.AppendChar(c)
        Next
        secure.MakeReadOnly()
        Return secure
    End Function

    Public Shared Function ToInsecureString(ByVal data As SecureString) As String
        Dim returnString As String = String.Empty
        Dim ptr As IntPtr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(data)
        Try
            returnString = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr)
        Finally
            System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr)
        End Try
        Return returnString
    End Function
End Class
