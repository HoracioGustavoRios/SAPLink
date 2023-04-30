Imports System.Security.Cryptography
Imports System.IO
Imports VB = Microsoft.VisualBasic
Imports System.Text
Imports System.Runtime.InteropServices
'Insert security functions below...
Namespace Hash
    Module Hash
        Function CreateHash(ByVal strSource As String) As String
            Dim bytHash As Byte()
            Dim uEncode As New UnicodeEncoding
            'Store the source string in a byte array
            Dim bytSource() As Byte = uEncode.GetBytes(strSource)
            Dim sha1 As New SHA1CryptoServiceProvider
            'Create the hash
            bytHash = sha1.ComputeHash(bytSource)
            'return as a base64 encoded string
            Return Convert.ToBase64String(bytHash)
        End Function
    End Module
End Namespace

Namespace PrivateKey
    Module PrivateKey
        Function Encrypt(ByVal strPlainText As String, _
                ByVal strKey16 As String) As String
            Dim crp As New TripleDESCryptoServiceProvider
            Dim uEncode As New UnicodeEncoding
            Dim aEncode As New ASCIIEncoding
            'Store plaintext as a byte array
            Dim bytPlainText() As Byte = uEncode.GetBytes(strPlainText)
            'Create a memory stream for holding encrypted text
            Dim stmCipherText As New MemoryStream
            'Private key
            crp.Key = aEncode.GetBytes(Left(strKey16, 16))
            'Initialization vector is the encryption seed 
            crp.IV = aEncode.GetBytes(VB.Right(strKey16, 8))
            'Create a crypto-writer to encrypt a bytearray
            'into a stream
            Dim csEncrypted As New CryptoStream(stmCipherText, _
                crp.CreateEncryptor(), CryptoStreamMode.Write)
            csEncrypted.Write(bytPlainText, 0, bytPlainText.Length)
            csEncrypted.FlushFinalBlock()
            'Return result as a Base64 encoded string
            Return Convert.ToBase64String(stmCipherText.ToArray())
        End Function

        Function Decrypt(ByVal strCipherText As String, _
                ByVal strKey16 As String) As String
            Dim crp As New TripleDESCryptoServiceProvider
            Dim uEncode As New UnicodeEncoding
            Dim aEncode As New ASCIIEncoding
            'Store cipher text as a byte array
            Dim bytCipherText() As Byte = _
            Convert.FromBase64String(strCipherText)
            Dim stmPlainText As New MemoryStream
            Dim stmCipherText As New MemoryStream(bytCipherText)
            'Private key
            crp.Key = aEncode.GetBytes(VB.Left(strKey16, 16))
            'Initialization vector
            crp.IV = aEncode.GetBytes(VB.Right(strKey16, 8))
            'Create a crypto stream decoder to decode
            'a cipher text stream into a plain text stream
            Dim csDecrypted As New CryptoStream(stmCipherText, _
                crp.CreateDecryptor(), CryptoStreamMode.Read)
            Dim sw As New StreamWriter(stmPlainText)
            Dim sr As New StreamReader(csDecrypted)
            sw.Write(sr.ReadToEnd)
            'Clean up afterwards
            sw.Flush()
            csDecrypted.Clear()
            crp.Clear()
            Return uEncode.GetString(stmPlainText.ToArray())
        End Function
    End Module
End Namespace

Namespace PublicKey
    Module PublicKey
        Function CreateKeyPair() As String
            'Create a new random key pair
            Dim rsa As New RSACryptoServiceProvider
            CreateKeyPair = rsa.ToXmlString(True)
            rsa.Clear()
        End Function
        Function GetPublicKey(ByVal strPrivateKey As String) As String
            'Extract the public key from the 
            'public/private key pair
            Dim rsa As New RSACryptoServiceProvider
            rsa.FromXmlString(strPrivateKey)
            Return rsa.ToXmlString(False)
        End Function
        Function Encrypt(ByVal strPlainText As String, _
                ByVal strPublicKey As String) As String
            'Encrypt a string using the private or public key
            Dim rsa As New RSACryptoServiceProvider
            Dim bytPlainText() As Byte
            Dim bytCipherText() As Byte
            Dim uEncode As New UnicodeEncoding
            rsa.FromXmlString(strPublicKey)
            bytPlainText = uEncode.GetBytes(strPlainText)
            bytCipherText = rsa.Encrypt(bytPlainText, False)
            Encrypt = Convert.ToBase64String(bytCipherText)
            rsa.Clear()
        End Function
        Function Decrypt(ByVal strCipherText As String, _
                ByVal strPrivateKey As String) As String
            'Decrypt a string using the private key
            Dim rsa As New RSACryptoServiceProvider
            Dim bytPlainText() As Byte
            Dim bytCipherText() As Byte
            Dim uEncode As New UnicodeEncoding
            rsa.FromXmlString(strPrivateKey)
            bytCipherText = Convert.FromBase64String(strCipherText)
            bytPlainText = rsa.Decrypt(bytCipherText, False)
            Decrypt = uEncode.GetString(bytPlainText)
            rsa.Clear()
        End Function
    End Module
End Namespace
Namespace Settings
    Module Settings
        <StructLayout(LayoutKind.Sequential)> Private Structure DATA_BLOB
            Dim cbData As Integer
            Dim pbData As IntPtr
        End Structure
        <StructLayout(LayoutKind.Sequential)> Private Structure CRYPTPROTECT_PROMPTSTRUCT
            Dim cbSize As Integer
            Dim dwPromptFlags As Integer
            Dim hwndApp As IntPtr
            Dim szPrompt As String
        End Structure
        Private Declare Function CryptProtectData Lib "Crypt32.dll" ( _
            ByRef pDataIn As DATA_BLOB, _
            ByVal szDataDescr As String, _
            ByRef pOptionalEntropy As DATA_BLOB, _
            ByVal pvReserved As IntPtr, _
            ByRef pPromptStruct As CRYPTPROTECT_PROMPTSTRUCT, _
            ByVal dwFlags As Integer, _
            ByRef pDataOut As DATA_BLOB) As Boolean
        Private Declare Function CryptUnprotectData Lib "Crypt32.dll" ( _
            ByRef pDataIn As DATA_BLOB, _
            ByVal szDataDescr As String, _
            ByRef pOptionalEntropy As DATA_BLOB, _
            ByVal pvReserved As IntPtr, _
            ByRef pPromptStruct As CRYPTPROTECT_PROMPTSTRUCT, _
            ByVal dwFlags As Integer, _
            ByRef pDataOut As DATA_BLOB) As Boolean
        Function SaveEncrypted(ByVal strSettingName As String, ByVal strValue As String) As Boolean
            Dim bytPlainText() As Byte
            Dim bytCipherText() As Byte
            Dim strCipherText As String
            Dim strFilename As String
            Dim uEncode As New UnicodeEncoding
            Dim blnSuccess As Boolean = False
            Dim bbPlainText, bbCipherText, bbEntropy As DATA_BLOB
            Dim pmt As CRYPTPROTECT_PROMPTSTRUCT
            Dim intFileNumber As Integer
            'Initialize the pmt structure
            pmt.cbSize = Marshal.SizeOf(pmt)
            pmt.hwndApp = IntPtr.Zero
            pmt.szPrompt = vbNullString
            'Convert the plaintext into a byte array, and copy
            'to global memory
            bytPlainText = uEncode.GetBytes(strValue)
            bbPlainText.pbData = Marshal.AllocHGlobal(bytPlainText.Length)
            If bbPlainText.pbData.ToInt32 = 0 Then
                MsgBox("Global Alloc failed", , "ProtectString")
                Return False
            End If
            bbPlainText.cbData = bytPlainText.Length
            Marshal.Copy(bytPlainText, 0, bbPlainText.pbData, bytPlainText.Length)
            'Call the windows Crypto API CryptProtectData to encrypt the plain text
            blnSuccess = CryptProtectData(bbPlainText, "", bbEntropy, IntPtr.Zero, pmt, 1, bbCipherText)
            If blnSuccess = False Then
                MsgBox("CryptProtect failed", , "ProtectString")
                Return False
            End If
            'the result is stored in a block of memory. Convert this block to a string
            ReDim bytCipherText(bbCipherText.cbData)
            Marshal.Copy(bbCipherText.pbData, bytCipherText, 0, bbCipherText.cbData)
            Marshal.FreeHGlobal(bbPlainText.pbData)
            strCipherText = Convert.ToBase64String(bytCipherText)
            'Save the encrypted setting to a file in the user's application folder
            intFileNumber = FreeFile()
            strFilename = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & strSettingName & ".txt"
            FileOpen(intFileNumber, strFilename, OpenMode.Output, OpenAccess.Write)
            PrintLine(intFileNumber, strCipherText)
            FileClose(intFileNumber)
            Return True
        End Function
        Function LoadEncrypted(ByVal StrSettingName As String) As String
            Dim strCipherText As String
            Dim uEncode As New UnicodeEncoding
            Dim bytCipherText() As Byte
            Dim blnSuccess As Boolean = False
            Dim bbPlainText, bbCipher, bbEntropy As DATA_BLOB
            Dim pmt As CRYPTPROTECT_PROMPTSTRUCT
            Dim strFilename As String
            Dim intFileNumber As Integer
            'Load the encrypted setting from the file in the user's
            'application folder
            intFileNumber = FreeFile()
            strFilename = System.Environment.GetFolderPath( _
              Environment.SpecialFolder.ApplicationData) & _
              "\" & StrSettingName & ".txt"
            FileOpen(intFileNumber, strFilename, _
              OpenMode.Input, OpenAccess.Read)
            strCipherText = LineInput(intFileNumber)
            FileClose(intFileNumber)
            'initialize the pmt structure
            pmt.cbSize = Marshal.SizeOf(pmt)
            pmt.hwndApp = IntPtr.Zero
            pmt.szPrompt = vbNullString
            'Copy the ciphertext into a byte array and store
            'it in global memory
            bytCipherText = Convert.FromBase64String(strCipherText)
            bbCipher.pbData = Marshal.AllocHGlobal(bytCipherText.Length)
            If bbCipher.pbData.ToInt32 = 0 Then
                MsgBox("Global Alloc failed", , "UnprotectString")
                Return ""
            End If
            bbCipher.cbData = bytCipherText.Length
            Marshal.Copy(bytCipherText, 0, bbCipher.pbData, bbCipher.cbData)
            'Call the Windows API CryptUnprotectData
            blnSuccess = CryptUnprotectData(bbCipher, vbNullString, bbEntropy, IntPtr.Zero, pmt, 1, bbPlainText)
            If blnSuccess = False Then
                MsgBox("CryptUnprotect failed", , "UnprotectString")
                Return ("")
            End If
            'the result is stored in a block of memory. Convert to a
            'string and return to the user
            Marshal.FreeHGlobal(bbCipher.pbData)
            Dim plainText(bbPlainText.cbData) As Byte
            Marshal.Copy(bbPlainText.pbData, plainText, 0, bbPlainText.cbData)
            Return uEncode.GetString(plainText)
            MsgBox(System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
        End Function
    End Module
End Namespace
Namespace EventLog
    Module EventLog
        Sub LogException(ByVal ex As Exception)
            MsgBox("Ha ocurrido un error y sera registrado. Contacte a su Administrador de Sistemas.", MsgBoxStyle.Critical, "Error en la aplicacion")
            'If this is an NT based operating system (Windows NT4, Windows2000,
            'Windows XP, Windows Server 2003) then add the exception to the
            'application event log.
            'If the operating system is Windows98 or WindowsME, then
            'append it to a <appname>.log file in the ApplicationData directory
            Dim strApplicationName As String
            Dim blnIsWin9X As Boolean
            Dim FileNumber As Integer = -1
            Try
                'Get name of assembly
                strApplicationName = _
                System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
                blnIsWin9X = (System.Environment.OSVersion.Platform <> _
                PlatformID.Win32NT)
                If blnIsWin9X Then
                    'Windows98 or WindowsME
                    Dim strTargetDirectory, strTargetPath As String
                    'Get Application Data directory, and create path
                    strTargetDirectory = System.Environment.GetFolderPath( _
                      Environment.SpecialFolder.ApplicationData)
                    strTargetPath = strTargetDirectory & "\" & _
                      strApplicationName & ".Log"
                    'Append to the end of the log (or create a new one 
                    'if it doesn't already exist)
                    FileNumber = FreeFile()
                    FileOpen(FileNumber, strTargetPath, OpenMode.Append)
                    PrintLine(FileNumber, Now)
                    PrintLine(FileNumber, ex.ToString)
                    FileClose(FileNumber)
                Else
                    'WinNT4, Win2K, WinXP, Windows.NET
                    System.Diagnostics.EventLog.WriteEntry(strApplicationName, _
                      ex.ToString, EventLogEntryType.Error)
                End If
            Finally
                If FileNumber > -1 Then FileClose(FileNumber)
            End Try
        End Sub
    End Module
End Namespace
