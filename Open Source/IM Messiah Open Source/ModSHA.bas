Attribute VB_Name = "ModSHA"
Option Explicit

Public Type Word
    B0 As Byte
    B1 As Byte
    B2 As Byte
    B3 As Byte
End Type

Public Function AndW(w1 As Word, w2 As Word) As Word
    Dim W As Word
    W.B0 = w1.B0 And w2.B0: W.B1 = w1.B1 And w2.B1
    W.B2 = w1.B2 And w2.B2: W.B3 = w1.B3 And w2.B3
    AndW = W
End Function

Public Function OrW(w1 As Word, w2 As Word) As Word
    Dim W As Word
    W.B0 = w1.B0 Or w2.B0: W.B1 = w1.B1 Or w2.B1
    W.B2 = w1.B2 Or w2.B2: W.B3 = w1.B3 Or w2.B3
    OrW = W
End Function

Public Function XorW(w1 As Word, w2 As Word) As Word
    Dim W As Word
    W.B0 = w1.B0 Xor w2.B0: W.B1 = w1.B1 Xor w2.B1
    W.B2 = w1.B2 Xor w2.B2: W.B3 = w1.B3 Xor w2.B3
    XorW = W
End Function

Public Function NotW(W As Word) As Word
    Dim w0 As Word
    w0.B0 = Not W.B0: w0.B1 = Not W.B1
    w0.B2 = Not W.B2: w0.B3 = Not W.B3
    NotW = w0
End Function

Public Function AddW(w1 As Word, w2 As Word) As Word
    Dim I As Integer, W As Word
    I = CInt(w1.B3) + w2.B3: W.B3 = I Mod 256
    I = CInt(w1.B2) + w2.B2 + (I \ 256): W.B2 = I Mod 256
    I = CInt(w1.B1) + w2.B1 + (I \ 256): W.B1 = I Mod 256
    I = CInt(w1.B0) + w2.B0 + (I \ 256): W.B0 = I Mod 256
    AddW = W
End Function

Public Function CircShiftLeftW(W As Word, n As Integer) As Word
    Dim d1 As Double, d2 As Double
    d1 = WordToDouble(W): d2 = d1
    d1 = d1 * (2 ^ n): d2 = d2 / (2 ^ (32 - n))
    CircShiftLeftW = OrW(DoubleToWord(d1), DoubleToWord(d2))
End Function

Public Function WordToHex(W As Word) As String
    WordToHex = Right$("0" & Hex$(W.B0), 2) & Right$("0" & Hex$(W.B1), 2) & _
                Right$("0" & Hex$(W.B2), 2) & Right$("0" & Hex$(W.B3), 2)
End Function

Public Function HexToWord(h As String) As Word
    HexToWord = DoubleToWord(Val("&H" & h & "#"))
End Function

Public Function DoubleToWord(n As Double) As Word
    Dim W As Word
    W.B0 = Int(DMod(n, 2 ^ 32) / (2 ^ 24)): W.B1 = Int(DMod(n, 2 ^ 24) / (2 ^ 16))
    W.B2 = Int(DMod(n, 2 ^ 16) / (2 ^ 8)): W.B3 = Int(DMod(n, 2 ^ 8))
    DoubleToWord = W
End Function

Public Function WordToDouble(W As Word) As Double
    WordToDouble = (W.B0 * (2 ^ 24)) + (W.B1 * (2 ^ 16)) + (W.B2 * (2 ^ 8)) + W.B3
End Function

Public Function DMod(Value As Double, divisor As Double) As Double
    Dim n As Double
    n = Value - (Int(Value / divisor) * divisor)
    If (n < 0) Then n = n + divisor
    DMod = n
End Function

Public Function F(T As Integer, B As Word, C As Word, D As Word) As Word
    Select Case T
        Case Is <= 19
            F = OrW(AndW(B, C), AndW((NotW(B)), D))
        Case Is <= 39
            F = XorW(XorW(B, C), D)
        Case Is <= 59
            F = OrW(OrW(AndW(B, C), AndW(B, D)), AndW(C, D))
        Case Else
            F = XorW(XorW(B, C), D)
    End Select
End Function

Public Function SHAHash(inMessage As String) As String
    Dim inLen As Long, inLenW As Word, padMessage As String
    Dim numBlocks As Long, W(0 To 79) As Word
    Dim blockText As String, wordText As String
    Dim I As Long, T As Integer, temp As Word
    
    Dim K(0 To 3) As Word
    Dim H0 As Word, H1 As Word, H2 As Word, H3 As Word, H4 As Word
    Dim a As Word, B As Word, C As Word, D As Word, E As Word
  
    inLen = Len(inMessage)
    inLenW = DoubleToWord(CDbl(inLen) * 8)
    padMessage = inMessage & Chr$(128) & String$((128 - (inLen Mod 64) - 9) Mod 64, Chr$(0)) & _
                 String$(4, Chr$(0)) & Chr$(inLenW.B0) & Chr$(inLenW.B1) & Chr$(inLenW.B2) & Chr$(inLenW.B3)
    numBlocks = Len(padMessage) / 64

    K(0) = HexToWord("5A827999"): K(1) = HexToWord("6ED9EBA1")
    K(2) = HexToWord("8F1BBCDC"): K(3) = HexToWord("CA62C1D6")
    
    H0 = HexToWord("67452301"): H1 = HexToWord("EFCDAB89")
    H2 = HexToWord("98BADCFE"): H3 = HexToWord("10325476")
    H4 = HexToWord("C3D2E1F0")
  
    For I = 0 To numBlocks - 1
        blockText = Mid$(padMessage, (I * 64) + 1, 64)
        For T = 0 To 15
            wordText = Mid$(blockText, (T * 4) + 1, 4)
            W(T).B0 = Asc(Mid$(wordText, 1, 1))
            W(T).B1 = Asc(Mid$(wordText, 2, 1))
            W(T).B2 = Asc(Mid$(wordText, 3, 1))
            W(T).B3 = Asc(Mid$(wordText, 4, 1))
        Next T
        
        For T = 16 To 79
            W(T) = CircShiftLeftW(XorW(XorW(XorW(W(T - 3), W(T - 8)), W(T - 14)), W(T - 16)), 1)
        Next T
        
        a = H0: B = H1: C = H2: D = H3: E = H4
        For T = 0 To 79
            temp = AddW(AddW(AddW(AddW(CircShiftLeftW(a, 5), F(T, B, C, D)), E), W(T)), K(T \ 20))
            E = D: D = C: C = CircShiftLeftW(B, 30): B = a: a = temp
        Next T
        H0 = AddW(H0, a): H1 = AddW(H1, B): H2 = AddW(H2, C): H3 = AddW(H3, D): H4 = AddW(H4, E)
    Next I
  
    SHAHash = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) & WordToHex(H3) & WordToHex(H4)
End Function

