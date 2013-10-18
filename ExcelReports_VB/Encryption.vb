Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic.VbStrConv
Imports System.Globalization


Friend Class Encryption
    Private ReadOnly iv() As Byte = {8, 7, 6, 5, 4, 3, 2, 1}
    Private sbox(255)
    Private key(255)

    ' define the triple des provider
    Private m_des As New TripleDESCryptoServiceProvider

    ' define the string handler
    Private m_utf8 As New UTF8Encoding

    ' define the local property arrays
    Private m_key() As Byte
    Private m_iv() As Byte

    Public Sub New(ByVal key As String)
        Me.m_key = StrToByteArray(key)
    End Sub

    Public Sub New(ByVal key() As Byte, ByVal iv() As Byte)
        Me.m_key = key
        Me.m_iv = iv
    End Sub

    Public Sub New()
        'Usual constructor
    End Sub

    Public Function DESEncrypt(ByVal input() As Byte) As Byte()
        Return Transform(input, m_des.CreateEncryptor(m_key, m_iv))
    End Function

    Public Function DESDecrypt(ByVal input() As Byte) As Byte()
        Return Transform(input, m_des.CreateDecryptor(m_key, m_iv))
    End Function

    Public Function DESEncrypt(ByVal text As String) As String
        Dim input() As Byte = m_utf8.GetBytes(text)
        Dim output() As Byte = Transform(input, _
                        m_des.CreateEncryptor(m_key, m_iv))
        Return Convert.ToBase64String(output)
    End Function

    Public Function DESDecrypt(ByVal text As String) As String
        Dim input() As Byte = Convert.FromBase64String(text)
        Dim output() As Byte = Transform(input, _
                         m_des.CreateDecryptor(m_key, m_iv))
        Return m_utf8.GetString(output)
    End Function

    Private Function Transform(ByVal input() As Byte, _
        ByVal CryptoTransform As ICryptoTransform) As Byte()
        ' create the necessary streams
        Dim memStream As MemoryStream = New MemoryStream
        Dim cryptStream As CryptoStream = New  _
            CryptoStream(memStream, CryptoTransform, _
            CryptoStreamMode.Write)
        ' transform the bytes as requested
        cryptStream.Write(input, 0, input.Length)
        cryptStream.FlushFinalBlock()
        ' Read the memory stream and convert it back into byte array
        memStream.Position = 0
        Dim result(CType(memStream.Length - 1, System.Int32)) As Byte
        memStream.Read(result, 0, CType(result.Length, System.Int32))
        ' close and release the streams
        memStream.Close()
        cryptStream.Close()
        ' hand back the encrypted buffer
        Return result
    End Function

    Public Shared Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.ASCIIEncoding()
        Return encoding.GetBytes(str)
    End Function 'StrToByteArray

    Sub RC4Initialize(ByVal strPwd)
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        ':::  This routine called by EnDeCrypt function. Initializes the
        ':::  sbox and the key array)
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

        Dim tempSwap As Byte
        Dim a
        Dim b

        Array.Clear(sbox, 0, sbox.Length)
        Array.Clear(key, 0, key.Length)

        Dim intLength = Len(strPwd)
        For a = 0 To 255
            key(a) = Asc(Mid(strPwd, (a Mod intLength) + 1, 1))
            sbox(a) = a
        Next

        b = 0
        For a = 0 To 255
            b = (b + sbox(a) + key(a)) Mod 256
            tempSwap = sbox(a)
            sbox(a) = sbox(b)
            sbox(b) = tempSwap
        Next
    End Sub

    Function RC4EnDecrypt(ByVal StValue As String, ByVal psw As String) As String
        Dim ByteArray() As Byte
        ByteArray = Encoding.ASCII.GetBytes(StValue)

        ByteArray = Me.RC4EnDecrypt(ByteArray, psw)

        Return Encoding.ASCII.GetString(ByteArray)
    End Function

    Function RC4EnDecrypt(ByVal byteArray As Byte(), ByVal psw As String) As Byte()

        Dim temp As Byte
        Dim a As Int32
        Dim i As Int32
        Dim j As Int32
        Dim k As Byte
        Dim cipherby As Byte
        Dim cipher As Byte()
        ReDim cipher(byteArray.Length - 1)


        i = 0
        j = 0

        RC4Initialize(psw)

        For a = 0 To byteArray.Length - 1
            i = (i + 1) Mod 256
            j = (j + sbox(i)) Mod 256
            temp = sbox(i)
            sbox(i) = sbox(j)
            sbox(j) = temp

            k = sbox((sbox(i) + sbox(j)) Mod 256)

            cipherby = byteArray(a) Xor k
            cipher(a) = cipherby
        Next

        Return cipher

    End Function

    Function convertStringToHex(ByVal str As String) As String
        'Dim StringToHex As String = ""
        'Dim x As Long, hexchar As String
        'Dim bytes() As Byte
        'bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(str)
        'For x = 0 To bytes.Length - 1
        '    hexchar = Hex$(bytes(x))
        '    If Len(hexchar) = 1 Then hexchar = "0" & hexchar
        '    StringToHex = StringToHex & hexchar
        'Next x
        'Return StringToHex


        Dim byteArray() As Byte = Nothing
        Dim hexNumbers As System.Text.StringBuilder = New System.Text.StringBuilder

        'byteArray = System.Text.ASCIIEncoding.ASCII.GetBytes(str)
        'byteArray = UTF32Encoding.UTF8.GetBytes(str)
        Dim i As Integer
        For i = 0 To str.Length - 2
            'Dim hexVal As String = byteArray(i).ToString("X")
            'Dim toAdd As String = hexVal.PadLeft(4, "0")
            hexNumbers.Append(Asc(str(i)).ToString("X").PadLeft(2, "0") & "-")
        Next
        hexNumbers.Append(Asc(str(i)).ToString("X").PadLeft(2, "0"))
        Return hexNumbers.ToString

    End Function

    Function convertHexToString(ByVal str As String) As String
        Dim HexToString As String = ""
        Dim x As Long


        'bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(str)
        'For x = 0 To (str.Length / 2) - 1
        '    If x > bytes.Length Then
        '        ReDim bytes(bytes.Length + 2)
        '    End If
        '    Dim newByte As Byte = Convert.ToByte(Convert.ToInt32(str.Substring(x, 2), 16))
        '    HexToString = HexToString + ChrW(newByte)
        'Next x
        ''HexToString = BitConverter.ToString(bytes)
        Dim hexChars As String() = str.Split("-")
        Dim newByte As Byte
        For x = 0 To hexChars.Length - 1
            newByte = Convert.ToByte(hexChars(x), 16)
            HexToString = HexToString & ChrW(newByte)
        Next

        'HexToString = UTF32Encoding.UTF32.GetString(bytes)
        Return HexToString
    End Function

    Function hexEncrypt(ByVal plaintxt As String, ByVal psw As String) As String
        Return convertStringToHex(RC4EnDecrypt(plaintxt, psw))
    End Function

    Function hexDecrypt(ByVal hexa As String, ByVal psw As String)
        Return RC4EnDecrypt(convertHexToString(hexa), psw)
    End Function

    Public WriteOnly Property pKey() As String
        Set(ByVal value As String)
            Me.m_key = StrToByteArray(value)
        End Set
    End Property

#Region " Nademote "

    'Sub RC4Initialize(ByVal strPwd)
    '    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '    ':::  This routine called by EnDeCrypt function. Initializes the
    '    ':::  sbox and the key array)
    '    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    '    Dim tempSwap
    '    Dim a
    '    Dim b

    '    Dim intLength = Len(strPwd)
    '    For a = 0 To 255
    '        key(a) = asc(mid(strpwd, (a Mod intLength) + 1, 1))
    '        sbox(a) = a
    '    Next

    '    b = 0
    '    For a = 0 To 255
    '        b = (b + sbox(a) + key(a)) Mod 256
    '        tempSwap = sbox(a)
    '        sbox(a) = sbox(b)
    '        sbox(b) = tempSwap
    '    Next
    'End Sub

    'Function RC4EnDecrypt(ByVal plaintxt As String, ByVal psw As String)

    '    Dim temp As String
    '    Dim a As Int32
    '    Dim i As Int32
    '    Dim j As Int32
    '    Dim k As Int32
    '    Dim cipherby
    '    Dim cipher
    '    cipher = ""

    '    i = 0
    '    j = 0

    '    RC4Initialize(psw)

    '    For a = 1 To Len(plaintxt)
    '        i = (i + 1) Mod 256
    '        j = (j + sbox(i)) Mod 256
    '        temp = sbox(i)
    '        sbox(i) = sbox(j)
    '        sbox(j) = temp

    '        k = sbox((sbox(i) + sbox(j)) Mod 256)

    '        cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
    '        cipher = cipher & Chr(cipherby)
    '    Next

    '    Return cipher

    'End Function

    'Public Sub RC4EnDecrypt(ByVal pInStream As IO.Stream, ByVal pOuStream As IO.Stream, ByVal psw As String)


    'End Sub

#End Region

End Class

