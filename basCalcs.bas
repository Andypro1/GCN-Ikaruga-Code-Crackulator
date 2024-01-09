Attribute VB_Name = "basCalcs"
Option Explicit
Private CodeKey As Collection  '  Our secret character string key

'  This function assembles a zero code.  It can get the full code
'  if the first 4 characters of the main code have been input.  Otherwise
'  It just assembles what it can.
Public Function GetZeroCode(ByVal strCode As String, ByVal blArcadeMode As Boolean) As String
    Dim i As Integer
    Dim strTempText As String
    Dim strTempValue As String
    Dim BinNum As String
    Dim ResultBin As String
    Dim ResultDec As String
    strTempText = ""
    strTempValue = ""

    Select Case Len(strCode)
        Case 1
            strTempText = strTempText & Mid$(strCode, 1, 1)
        Case 2
            strTempText = strTempText & Mid$(strCode, 1, 1)
            strTempText = strTempText & Mid$(strCode, 2, 1)
        Case 3
            strTempText = strTempText & Mid$(strCode, 1, 1)
            strTempText = strTempText & Mid$(strCode, 2, 1)
            strTempText = strTempText & Mid$(strCode, 3, 1)
        Case Is > 3
            If Len(strCode) <= 12 Then
                strTempText = strTempText & Mid$(strCode, 1, 1) '1
                strTempText = strTempText & Mid$(strCode, 2, 1) '2
                strTempText = strTempText & Mid$(strCode, 3, 1) '3
                strTempText = strTempText & Mid$(strCode, 4, 1) '4
                strTempText = strTempText & Mid$(strCode, 1, 1) '5
                strTempText = strTempText & Mid$(strCode, 2, 1) '6

                BinNum = DecToBin(FindCharPosition(Mid$(strCode, 3, 1)))
                If frmMain.optNTSC.Value = True Then
                    ResultBin = XORString(BinNum, "000100")
                Else
                    ResultBin = XORString(BinNum, "001000")
                End If
                ResultDec = Trim(BinToDec(ResultBin))
                strTempText = strTempText & FindChar(CInt(ResultDec)) '7

                If blArcadeMode = True Then
                    strTempValue = Mid$(strCode, 4, 1) '8
                Else
                    BinNum = DecToBin(FindCharPosition(Mid$(strCode, 4, 1)))
                    ResultBin = XORString(BinNum, "000001")
                    ResultDec = Trim(BinToDec(ResultBin))
                    strTempValue = FindChar(CInt(ResultDec)) '8
                End If
                
                If frmMain.opt2P.Value = True Then
                    BinNum = DecToBin(FindCharPosition(strTempValue))
                    ResultBin = XORString(BinNum, "000010")
                    ResultDec = Trim(BinToDec(ResultBin))
                    strTempValue = FindChar(CInt(ResultDec)) '8
                End If

                strTempText = strTempText & strTempValue  '8

                BinNum = DecToBin(FindCharPosition(Mid$(strCode, 1, 1)))
                ResultBin = XORString(BinNum, "111111")
                ResultDec = Trim(BinToDec(ResultBin))
                strTempText = strTempText & FindChar(CInt(ResultDec)) '9

                BinNum = DecToBin(FindCharPosition(Mid$(strCode, 2, 1)))
                ResultBin = XORString(BinNum, "110000")
                ResultDec = Trim(BinToDec(ResultBin))
                strTempText = strTempText & FindChar(CInt(ResultDec)) '10

                If blArcadeMode = True Then
                    If frmMain.optNTSC.Value = True Then
                        strTempValue = Mid$(strCode, 3, 1) '11
                    Else
                        BinNum = DecToBin(FindCharPosition(Mid$(strCode, 3, 1)))
                        ResultBin = XORString(BinNum, "000100")
                        ResultDec = Trim(BinToDec(ResultBin))
                        strTempValue = FindChar(CInt(ResultDec)) '11
                    End If
                Else
                    If frmMain.optNTSC.Value = True Then
                        BinNum = DecToBin(FindCharPosition(Mid$(strCode, 3, 1)))
                        ResultBin = XORString(BinNum, "000100")
                        ResultDec = Trim(BinToDec(ResultBin))
                        strTempValue = FindChar(CInt(ResultDec)) '11
                    Else
                        BinNum = DecToBin(FindCharPosition(Mid$(strCode, 3, 1)))
                        ResultBin = XORString(BinNum, "001000")
                        ResultDec = Trim(BinToDec(ResultBin))
                        strTempValue = FindChar(CInt(ResultDec)) '11
                    End If
                End If
          
                If frmMain.opt2P.Value = True Then
                    BinNum = DecToBin(FindCharPosition(strTempValue))
                    ResultBin = XORString(BinNum, "001000")
                    ResultDec = Trim(BinToDec(ResultBin))
                    strTempValue = FindChar(CInt(ResultDec)) '11
                End If

                strTempText = strTempText & strTempValue '11

                strTempText = strTempText & Mid$(strCode, 4, 1) '12
            End If
    End Select

    GetZeroCode = strTempText
End Function

Public Sub Main()
    InitCodeKey

    frmMain.Show
End Sub

'  A collection is used as opposed to an array so that we have an
'  index associated with each character for search purposes (it's not
'  efficient, however it works).
Private Sub InitCodeKey()
    'D3AK2g7LuycxCZniRQPf#6J1FUTHGXmvEhWV@s%t?rq9MjYk8NboBzwap$4dSe!+
    Set CodeKey = New Collection
    CodeKey.Add "D"
    CodeKey.Add "3"
    CodeKey.Add "A"
    CodeKey.Add "K"
    CodeKey.Add "2"
    CodeKey.Add "g"
    CodeKey.Add "7"
    CodeKey.Add "L"
    CodeKey.Add "u"
    CodeKey.Add "y"
    CodeKey.Add "c"
    CodeKey.Add "x"
    CodeKey.Add "C"
    CodeKey.Add "Z"
    CodeKey.Add "n"
    CodeKey.Add "i"
    CodeKey.Add "R"
    CodeKey.Add "Q"
    CodeKey.Add "P"
    CodeKey.Add "f"
    CodeKey.Add "#"
    CodeKey.Add "6"
    CodeKey.Add "J"
    CodeKey.Add "1"
    CodeKey.Add "F"
    CodeKey.Add "U"
    CodeKey.Add "T"
    CodeKey.Add "H"
    CodeKey.Add "G"
    CodeKey.Add "X"
    CodeKey.Add "m"
    CodeKey.Add "v"
    CodeKey.Add "E"
    CodeKey.Add "h"
    CodeKey.Add "W"
    CodeKey.Add "V"
    CodeKey.Add "@"
    CodeKey.Add "s"
    CodeKey.Add "%"
    CodeKey.Add "t"
    CodeKey.Add "?"
    CodeKey.Add "r"
    CodeKey.Add "q"
    CodeKey.Add "9"
    CodeKey.Add "M"
    CodeKey.Add "j"
    CodeKey.Add "Y"
    CodeKey.Add "k"
    CodeKey.Add "8"
    CodeKey.Add "N"
    CodeKey.Add "b"
    CodeKey.Add "o"
    CodeKey.Add "B"
    CodeKey.Add "z"
    CodeKey.Add "w"
    CodeKey.Add "a"
    CodeKey.Add "p"
    CodeKey.Add "$"
    CodeKey.Add "4"
    CodeKey.Add "d"
    CodeKey.Add "S"
    CodeKey.Add "e"
    CodeKey.Add "!"
    CodeKey.Add "+"
End Sub

'  This function finds a character's position, and thus its value,
'  by searching through the collection.
Private Function FindCharPosition(ByVal strChar As String) As Integer
    Dim i As Integer
    For i = 0 To 63
        If CodeKey(i + 1) = Trim(strChar) Then
            FindCharPosition = i 'We don't add one here because we want the character's
                                 'actual value as opposed to it's collection position.
            Exit Function
        End If
    Next
End Function

'  This function retrieves a character from the collection
'  given a position.
Private Function FindChar(ByVal Position As Integer) As String
    FindChar = CodeKey(Position + 1)
End Function

'  This function was taken from the support.microsoft.com
'  knowledge base.  It converts numbers to binary (again, using
'  strings is inefficient, however it makes the coding easier).
Public Function DecToBin(DecNum As String) As String
   Dim BinNum As String
   Dim lDecNum As Long
   Dim i As Integer

   On Error GoTo ErrorHandler

'  Check the string for invalid characters
   For i = 1 To Len(DecNum)
      If Asc(Mid(DecNum, i, 1)) < 48 Or _
         Asc(Mid(DecNum, i, 1)) > 57 Then
         BinNum = ""
         Err.Raise 1010, "DecToBin", "Invalid Input"
      End If
   Next i

   i = 0
   lDecNum = Val(DecNum)

   Do
      If lDecNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lDecNum
'  Return BinNum as a String
   DecToBin = BinNum
ErrorHandler:
End Function

'  This function was taken from the support.microsoft.com
'  knowledge base.  It converts numbers to decimal (again, using
'  strings is inefficient, however it makes the coding easier).
Public Function BinToDec(BinNum As String) As String
   Dim i As Integer
   Dim DecNum As Long
   
   On Error GoTo ErrorHandler
   
'  Loop thru BinString
   For i = Len(BinNum) To 1 Step -1
'     Check the string for invalid characters
      If Asc(Mid(BinNum, i, 1)) < 48 Or _
         Asc(Mid(BinNum, i, 1)) > 49 Then
         DecNum = ""
         Err.Raise 1002, "BinToDec", "Invalid Input"
      End If
'     If bit is 1 then raise 2^LoopCount and add it to DecNum
      If Mid(BinNum, i, 1) And 1 Then
         DecNum = DecNum + 2 ^ (Len(BinNum) - i)
      End If
   Next i
'  Return DecNum as a String
   BinToDec = DecNum
ErrorHandler:
End Function

'  This function is used when generating the zero code.  It takes 6-bit
'  values and xor's them.
Private Function XORString(ByVal strOp1 As String, ByVal strOp2 As String) As String
    '  First assure that both operator strings are the proper length.
    If Len(strOp1) > 6 Then strOp1 = Right(strOp1, 6)
    If Len(strOp1) < 6 Then strOp1 = Mid$("000000", 1, 6 - Len(strOp1)) & strOp1
    If Len(strOp2) > 6 Then strOp2 = Right(strOp2, 6)
    If Len(strOp2) < 6 Then strOp2 = Mid$("000000", 1, 6 - Len(strOp2)) & strOp2

    Dim i As Integer
    Dim strResult As String
    strResult = ""
    For i = 1 To 6
        If Mid$(strOp1, i, 1) <> Mid$(strOp2, i, 1) Then
            strResult = strResult & "1"
        Else
            strResult = strResult & "0"
        End If
    Next
    XORString = strResult
End Function

Public Sub ExitProgram()
    Set CodeKey = Nothing
End Sub

'  This function generates the binary code.
Public Function GetCodeBinText(ByVal strCode As String) As String
    Dim i As Integer
    Dim tempText As String
    tempText = ""
    Dim tempBin As String

    For i = 1 To Len(strCode)
        tempBin = DecToBin(FindCharPosition(Mid$(strCode, i, 1)))
        tempText = tempText & Mid$("000000", 1, 6 - Len(tempBin)) & tempBin & "  "
    Next
    GetCodeBinText = tempText
End Function

'  This function generates the binary zero code.
Public Function GetZeroCodeBinText(ByVal strCode As String) As String
    Dim i As Integer
    Dim tempText As String
    tempText = ""
    Dim tempBin As String

    For i = 1 To Len(strCode)
        tempBin = DecToBin(FindCharPosition(Mid$(strCode, i, 1)))
        tempText = tempText & Mid$("000000", 1, 6 - Len(tempBin)) & tempBin & "  "
    Next
    GetZeroCodeBinText = tempText
End Function

'  This code XOR's the binary code and binary zero code
'  and returns the result.
Public Function GetResultBinText(ByVal binCode1 As String, ByVal binCode2 As String) As String
    Dim i As Integer
    Dim tempText As String
    tempText = ""

    For i = 1 To Len(binCode1)
        If Mid$(binCode1, i, 1) = " " Then
            tempText = tempText & " "
        ElseIf Mid$(binCode1, i, 1) <> Mid$(binCode2, i, 1) Then
            tempText = tempText & "1"
        Else
            tempText = tempText & "0"
        End If
    Next
    GetResultBinText = tempText
End Function

'  This function displays the binary result in 8 bit sections.
Public Function GetResult8BinText(ByVal resultText As String) As String
    Dim strTemp As String
    strTemp = Replace(resultText, " ", "")  '  Eliminate spaces.
    strTemp = Left(strTemp, 8) & "  " & Right(strTemp, (72 - 8))
    strTemp = Left(strTemp, (8 * 2) + 2) & "  " & Right(strTemp, (72 - (8 * 2)))
    strTemp = Left(strTemp, (8 * 3) + 4) & "  " & Right(strTemp, (72 - (8 * 3)))
    strTemp = Left(strTemp, (8 * 4) + 6) & "  " & Right(strTemp, (72 - (8 * 4)))
    strTemp = Left(strTemp, (8 * 5) + 8) & "  " & Right(strTemp, (72 - (8 * 5)))
    strTemp = Left(strTemp, (8 * 6) + 10) & "  " & Right(strTemp, (72 - (8 * 6)))
    strTemp = Left(strTemp, (8 * 7) + 12) & "  " & Right(strTemp, (72 - (8 * 7)))
    strTemp = Left(strTemp, (8 * 8) + 14) & "  " & Right(strTemp, (72 - (8 * 8)))

    GetResult8BinText = strTemp
End Function

'  This function converts the score to binary.
Public Function GetScoreBinText(ByVal strScore As String) As String
    GetScoreBinText = DecToBin(strScore)
End Function

'  This function re-orders the result binary value in order
'  to produce the score.
'  NOTE:  The 26th bit of the score is not included because it
'  is currently not known.  We need a score greater than 2^26
'  in order to place this final bit.
Public Sub AttemptScoreGuess()
    Dim strXOR As String
    Dim strScoreBin As String
    Dim strScore As String
    strScore = ""
    strXOR = ""
    strScoreBin = ""

    strXOR = Trim(Replace(frmMain.txtResultBin.Text, " ", ""))
    strXOR = Right(strXOR, 48)  '  We don't need the 4 salts

    strScoreBin = Mid$(strXOR, 8, 1) & _
                  Mid$(strXOR, 17, 1) & _
                  Mid$(strXOR, 18, 1) & _
                  Mid$(strXOR, 19, 1) & _
                  Mid$(strXOR, 20, 1) & _
                  Mid$(strXOR, 21, 1) & _
                  Mid$(strXOR, 22, 1) & _
                  Mid$(strXOR, 23, 1) & _
                  Mid$(strXOR, 24, 1) & _
                  Mid$(strXOR, 9, 1) & _
                  Mid$(strXOR, 10, 1) & _
                  Mid$(strXOR, 11, 1) & _
                  Mid$(strXOR, 12, 1) & _
                  Mid$(strXOR, 13, 1) & _
                  Mid$(strXOR, 14, 1) & _
                  Mid$(strXOR, 15, 1) & _
                  Mid$(strXOR, 16, 1) & _
                  Mid$(strXOR, 1, 1) & _
                  Mid$(strXOR, 2, 1) & _
                  Mid$(strXOR, 3, 1) & _
                  Mid$(strXOR, 4, 1) & _
                  Mid$(strXOR, 5, 1) & _
                  Mid$(strXOR, 6, 1) & _
                  Mid$(strXOR, 7, 1) & _
                  "0"

    strScore = BinToDec(strScoreBin)

    '  If the score is not divisible by 10, then we check to see
    '  if it is possibly a valid score above 33.55 million by
    '  flipping two of the bits and adding the 2^25 bit:
    If strScore Mod 10 <> 0 And frmMain.chkDisable = 0 Then
        strScoreBin = "1" & _
            Mid$(strXOR, 8, 1) & _
            Mid$(strXOR, 17, 1) & _
            Mid$(strXOR, 18, 1) & _
            Mid$(strXOR, 19, 1) & _
            Mid$(strXOR, 20, 1) & _
            Mid$(strXOR, 21, 1) & _
            Mid$(strXOR, 22, 1) & _
            Mid$(strXOR, 23, 1) & _
            Mid$(strXOR, 24, 1) & _
            Mid$(strXOR, 9, 1) & _
            Mid$(strXOR, 10, 1) & _
            Mid$(strXOR, 11, 1) & _
            Mid$(strXOR, 12, 1) & _
            Mid$(strXOR, 13, 1) & _
            Mid$(strXOR, 14, 1) & _
            Mid$(strXOR, 15, 1)

        strScoreBin = strScoreBin & _
            IIf(Mid$(strXOR, 16, 1) = "1", "0", "1") & _
            Mid$(strXOR, 1, 1) & _
            Mid$(strXOR, 2, 1) & _
            Mid$(strXOR, 3, 1) & _
            Mid$(strXOR, 4, 1) & _
            Mid$(strXOR, 5, 1) & _
            Mid$(strXOR, 6, 1) & _
            IIf(Mid$(strXOR, 7, 1) = "1", "0", "1") & _
            "0"

        strScore = BinToDec(strScoreBin)
    End If

    frmMain.txtScore.Text = Trim(strScore)
    '         111111111122222222223333333333444444444
    '123456789012345678901234567890123456789012345678
    '555555666666777777888888999999000000111111222222
End Sub

'  This function takes what we know about the extra code
'  positions and performs validity checks on the code.  It
'  returns a state variable to the caller indicating failed tests.
Public Function CheckCodeValidity(ByVal strResultBin As String, ByVal strScoreDec As String, ByVal strScoreBin As String) As Integer
    Dim intScoreIsValid As Integer  '  Our state variable.
    Dim tempResultBin As String
    intScoreIsValid = 0  '  Indicates that it's a valid code.
    tempResultBin = Right(strResultBin, 48)  '  Don't need the salts.
    
    If Mid$(tempResultBin, 25, 8) <> "00000000" Then
        intScoreIsValid = 1  '  Indicates that it fails the 0's test.
    End If

    If Mid$(tempResultBin, 41, 8) <> Mid$(tempResultBin, 17, 8) Then
        intScoreIsValid = 2  '  Indicates that it fails the duplicate test
                             '  (or multiple tests).
    End If

    If strScoreDec Mod 10 <> 0 Then
        intScoreIsValid = 3  '  Indicates that it fails the Mod 10 test.
                             '  (or multiple tests).
    End If

    If AllCharsValid(Trim(frmMain.txtCode.Text)) <> True Then
        intScoreIsValid = 4  '  Indicates that it fails the valid chars
                             '  test (or multiple tests).
    End If

    If CDbl(strScoreDec) > 38500000 Then
        intScoreIsValid = 5  '  Indicates that it fails the max score
                             '  test (or multiple tests).
    End If

    CheckCodeValidity = intScoreIsValid
End Function

'  This function takes the code string and determines if all of the
'  characters in the string are valid code characters.
Private Function AllCharsValid(ByVal strScoreCode As String) As Boolean
    Dim i As Integer
    Dim intPosition As Integer
    Dim blState As Boolean
    blState = True

    For i = 1 To Len(strScoreCode)
        intPosition = FindCharPosition(Mid$(strScoreCode, i, 1))
        Select Case intPosition
            Case 0
                If Mid$(strScoreCode, i, 1) <> "D" Then
                    blState = False
                End If
            Case Is < 63
                '  Do nothing
            Case 63
                If Mid$(strScoreCode, i, 1) <> "+" Then
                    blState = False
                End If
            Case Else
                blState = False
        End Select
    Next

    AllCharsValid = blState
End Function
