Imports System

Public Module RFC4648
    Private Const EncodedBitCount As Integer = 5
    Private Const ByteBitCount As Integer = 8
    Private Const EncodingChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
    Private Const PaddingCharacter As Char = "="c

    Public Function Encode(InputArr As Byte()) As String
        If InputArr Is Nothing Then
            Throw New ArgumentNullException("no data")
        End If
        Dim OutLen As Integer = CInt(Decimal.Ceiling(InputArr.Length / EncodedBitCount)) * ByteBitCount
        Dim RetArray As Char() = New Char(OutLen - 1) {}
        Dim OutB As Byte = 0
        Dim Shift As Short = EncodedBitCount
        Dim I As Integer = 0
        For J As Integer = 0 To InputArr.Length - 1
            Dim InputB As Byte = InputArr(J)
            OutB = CByte(CInt(OutB) Or InputB >> CInt((ByteBitCount - Shift)))
            RetArray(I) = EncodingChars.ToArray(CInt(OutB))
            I = I + 1
            If Shift <= ByteBitCount - EncodedBitCount Then
                OutB = CByte(InputB >> CInt((ByteBitCount - EncodedBitCount - Shift)) And 31)
                RetArray(I) = EncodingChars.ToArray(CInt(OutB))
                I = I + 1
                Shift += EncodedBitCount
            End If
            Shift -= ByteBitCount - EncodedBitCount
            OutB = CByte(CInt(InputB) << CInt(Shift) And 31)
        Next
        If I <> OutLen Then
            RetArray(I) = EncodingChars.ToArray(CInt(OutB))
            I = I + 1
        End If
        For j As Integer = I To OutLen - 1
            RetArray(j) = PaddingCharacter
        Next
        Return New String(RetArray)
    End Function

    Public Function Decode(Base32String As String) As Byte()
        If String.IsNullOrEmpty(Base32String) Then
            Throw New ArgumentNullException("no data")
        End If
        Dim ClearText As String = Base32String.ToUpperInvariant().TrimEnd(PaddingCharacter)
        For I As Integer = 0 To ClearText.Length() - 1
            Dim C As Char = ClearText.ToArray(I)
            If EncodingChars.IndexOf(C) < 0 Then
                Throw New ArgumentException("base32 contains illegal characters")
            End If
        Next
        Dim OutLen As Integer = ClearText.Length() * EncodedBitCount \ ByteBitCount
        Dim Array As Byte() = New Byte(OutLen - 1) {}
        Dim Summ As Byte = 0
        Dim Shift As Short = ByteBitCount
        Dim Ind As Integer = 0
        For J As Integer = 0 To ClearText.Length() - 1
            Dim C2 As Char = ClearText.ToArray(J)
            Dim PosIndex As Integer = EncodingChars.IndexOf(C2)
            If Shift > EncodedBitCount Then
                Dim Left As Integer = PosIndex << CInt(Shift - EncodedBitCount And 31)
                Summ = CByte(CInt(Summ) Or Left)
                Shift = CShort(Shift - EncodedBitCount)
            Else
                Dim Right As Integer = PosIndex >> CInt(EncodedBitCount - Shift And 31)
                Summ = CByte(CInt(Summ) Or Right)
                Array(Ind) = Summ
                Ind = Ind + 1
                Summ = BitConverter.GetBytes(PosIndex << CInt(ByteBitCount - EncodedBitCount + Shift And 31))(0)
                Shift = CShort(Shift + ByteBitCount - EncodedBitCount)
            End If
        Next
        Return Array
    End Function

End Module

