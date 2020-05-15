Imports Base32
Imports NUnit.Framework

Namespace NUnitBase32Test

    Public Class Tests

        <Test>
        Public Sub Encoder()
            For I As Integer = 0 To 100
                Dim RandomString = CreateRandomPassword(I + 5, &H30 + I, &H7A + I)
                Dim MyEncoder As String = RFC4648.Encode(System.Text.Encoding.UTF8.GetBytes(RandomString))
                TestContext.WriteLine(RandomString & " : " & MyEncoder)
                Dim MyDecode As Byte() = RFC4648.Decode(MyEncoder)
                Assert.IsTrue(MyDecode.SequenceEqual(System.Text.Encoding.UTF8.GetBytes(RandomString)))
            Next
        End Sub

        Public Function CreateRandomPassword(Len As String, Optional FromCharCode As UInt32 = &H33, Optional ToCharCode As UInt32 = &H7E, Optional ExcludeChars As String = "<>") As String
            Dim Ret1 As New System.Text.StringBuilder
            While Ret1.Length < Len
                Dim RandomNum As Integer = FromCharCode + GetRandomInteger(ToCharCode - FromCharCode)
                Dim OneChar As Char = ChrW(RandomNum)
                If Not ExcludeChars.Contains(OneChar) Then
                    Ret1.Append(OneChar)
                End If
            End While
            Return Ret1.ToString
        End Function

        Public Function GetRandomInteger(MaxValue As Integer) As Integer
            Dim byte_count As Byte() = New Byte(3) {}
            Dim random_number As New System.Security.Cryptography.RNGCryptoServiceProvider()
            random_number.GetBytes(byte_count)
            Return (Math.Abs(BitConverter.ToInt32(byte_count, 0)) / Int32.MaxValue) * MaxValue
        End Function
    End Class

End Namespace