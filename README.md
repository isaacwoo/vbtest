Dim result
result = 0

Function GenerateArray(min, max)

    Dim arr()
    ReDim arr(max - min)
    Dim i

    For i = 0 To max - min

        arr(i) = min + i

    Next

    GenerateArray = arr

End Function

Function RemoveElement(arr, elem)

    Dim newArr()
    Dim i, j
    ReDim newArr(UBound(arr) - 1)

    j = 0

    For i = 0 To UBound(arr)

        If arr(i) <> elem Then

            newArr(j) = arr(i)
            j = j + 1

        End If

    Next

    RemoveElement = newArr

End Function

Function GetRandomElement(arr)

    Randomize
    Dim index

    index = Int((UBound(arr) + 1) * Rnd)

    GetRandomElement = arr(index)

End Function

Function GetData(seriesNo, QuestionSub, QuestionSel, QuestionKey, QuestionAws)
 
    Dim Qsub
    Dim Qselect
    Dim QKey
    Dim QAws

    Select Case seriesNo

        Case "1"

            Qsub = "問題1" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 1
            QAws = "問1答え"
        
        Case "2"

            Qsub = "問題2" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 2
            QAws = "問2答え"

        Case "3"

            Qsub = "問題3" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 3
            QAws = "問3答え"

        Case "4"

            Qsub = "問題4" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 4
            QAws = "問4答え"

        Case "5"

            Qsub = "問題5" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 1
            QAws = "問5答え"

        Case "6"

            Qsub = "問題6" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 1
            QAws = "問6答え"
        
        Case "7"

            Qsub = "問題7" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 2
            QAws = "問7答え"

        Case "8"

            Qsub = "問題8" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 3
            QAws = "問8答え"

        Case "9"

            Qsub = "問題9" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 4
            QAws = "問9答え"

        Case "10"

            Qsub = "問題10" & vbCrLf + vbCrLf& seriesNo
            Qselect = "1. A" & vbCrLf & "2. B" & vbCrLf & "3. C" & vbCrLf & "4. D"
            QKey = 1
            QAws = "問10答え"

    End Select

    QuestionSub = Qsub
    QuestionSel = Qselect
    QuestionKey = QKey
    QuestionAws = QAws

End Function

Function Question(QuestionNo,seriesNo)

  Dim inputType
  
  GetData seriesNo, QuestionSub, QuestionSel, QuestionKey, QuestionAws

  Do

    inputType = InputBox(QuestionSub & vbCrLf & vbCrLf &  QuestionSel, "テスト")

    If NOT IsNumeric(inputType) Then

       WScript.Echo "入力エラー"

    ElseIf inputType = "" Then

       If MsgBox( "テスト終了しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
         
          WScript.Echo "テストが終了されました。"
          WScript.Quit
    
       End If

    ElseIf CInt(inputType) = QuestionKey Then

       result = result + 1

       Exit Do

    ElseIf CInt(inputType) <> QuestionKey Then

       WScript.Echo "wrong!the answer is " & vbCrLf & vbCrLf & QuestionAws

       Exit Do

    End If

  Loop

End Function
 
Function Test()

    Dim count
    Dim randomNumber
    Dim randomArray
    Dim dataLen 
    Dim TestVol
    Dim TestStart
    Dim TestEnd
    Dim TestRange
    
    count = 0
    dataLen = 10
    TestVol = 5

    Do

      TestRange = InputBox("試験範囲を選んでください:" & vbCrLf & vbCrLf & "0. 全部" & vbCrLf & "1. 前半" & vbCrLf & "2. 後半", "試験範囲選択","0")

      If TestRange = "0" OR TestRange = "1" OR TestRange = "2" Then

        Exit Do

      Else

        WScript.Echo "正しく入力してでください"

      End If

    Loop

    Select Case TestRange

        Case "0"
            
             TestStart = 1
             TestEnd = dataLen
    
        Case "1"
            
             TestStart = 1
             TestEnd = 5

        Case "2"
            
             TestStart = 6
             TestEnd = dataLen

    End Select

    randomArray = GenerateArray(TestStart, TestEnd)

    Do While count < TestVol
    
        randomNumber = GetRandomElement(randomArray)

        Question count,randomNumber

        randomArray = RemoveElement(randomArray, randomNumber)

        count = count + 1

    Loop
    
    Test = result

End Function

Dim score
score = Test()

Dim Subject
Dim Message
Dim mailTo
Dim mailCc

Subject = "【作業完了連絡】" 
Message = "お疲れ様です。" + vbCrLF + vbCrLF + "作業実施しました。採点は" & score & "点です。" + vbCrLF + vbCrLF + "以上、よろしくお願い致します。"
mailTo = "P381852@intra.meijiyasuda.co.jp"
mailCc = "P381852@intra.meijiyasuda.co.jp"

Sub sendMail()

    Set olApp = CreateObject("Outlook.Application")
    Set objMailItem = olApp.CreateItem(0)

    Dim olApp
    Dim objMailItem
    
    objMailItem.Subject = Subject
    objMailItem.To = mailTo
    objMailItem.Body = Message
    objMailItem.Cc = mailCc

    objMailItem.Send 
    
    MsgBox "送信しました。 " 

    Set olApp = Nothing
    Set objMailItem = Nothing

End Sub


Sub Main()   

    Do While result < 3

       MsgBox "採点は3点以下ですので、再度チャレンジしてください " 

       Test()

    Loop
 
    Dim confirm
    confirm = MsgBox( "実施完了しますので、完了通知メールを送信しますか？", vbYesNo + vbQuestion, "confirm")

    If confirm = vbYes Then

       SendMail

    Else

       WScript.Echo "キャンセルしました"

    End If

End Sub

Call Main
