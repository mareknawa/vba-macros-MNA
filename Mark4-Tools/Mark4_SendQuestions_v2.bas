Attribute VB_Name = "Mark4_SendQuestions"

Sub CreateEmailSet()
    Dim mailBody As String
    Dim sender()
    Dim receiver()
    Dim questionsStr As String: questionsStr = ""
    Dim questionsCol As New Collection ' to assure unique list of question numbers
    Dim runOffset As Integer ' to assure that each time receiver get other sender for sending feedback
    
    Dim teamMembers As Collection
    Dim receiversList As New Collection
    
    Dim teamSize As Long: teamSize = GetLastUsedRow(Worksheets("Team")) - 1
   
    Dim questionsNr As Integer: questionsNr = Int(Worksheets("Variables").Cells(3, 2))
    Dim shouldSendEmail As Boolean: shouldSendEmail = Worksheets("Variables").Cells(2, 2)
    
    bodyTemplate = Worksheets("Email").Cells(2, 2)
        
    Randomize ' initialization of random number generator
      
    ' Create Collection of team members
    Set teamMembers = createTeamMembers()
    
    ' create emails and write reports
    
    For x = 1 To teamSize
        sender = teamMembers.Item(x)
        receiver = FindReceiver(sender(2).Value, receiversList)

        questionsStr = QuestionList(questionsNr, 6)
        mailBody = StringFormat(bodyTemplate, sender(0), receiver(2), questionsStr)
        Debug.Print mailBody
        Call sendEmail(sender(3).Value, mailBody, shouldSendEmail)
        WriteReport sender, receiver, questionsStr
    Next


    
End Sub

Public Function FindReceiver(senderFullName As String, alreadyReceivedList As Collection) As Variant
    Dim reportRng As Range
    Dim ws As Worksheet: Set ws = Worksheets("Report")
    
    Dim receiverList As Collection: Set receiverList = createTeamMembers()
        
    DropMemberFromTheList senderFullName, receiverList
    If alreadyReceivedList.Count < receiverList.Count Then
        For Each alreadyRecived In alreadyReceivedList
            DropMemberFromTheList alreadyRecived(2).Value, receiverList
        Next
    End If
    
    Dim rng As Range: Set rng = ws.UsedRange ' ws.Range("A1:D126")
    Dim lastRow As Long: lastRow = GetLastUsedRow(ws)
    Dim selectedReceiver As Variant
    
    Debug.Print ("test = " & rng.Cells(1, 1).Value)
    
    rng.AutoFilter Field:=2, Criteria1:=senderFullName, VisibleDropDown:=False
    
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).EntireRow.Hidden = False Then
            If (receiverList.Count < 2) Then
                Exit For
            End If
            Debug.Print "Wylosowane pytania = "; ws.Cells(i, 4).Value
            DropMemberFromTheList ws.Cells(i, 3).Value, receiverList
            
        End If
        
    Next i
    
    If (receiverList.Count < 2) Then
        selectedReceiver = receiverList(1)
    Else
        randomReceiver = GetRandomNumber(receiverList.Count)
        selectedReceiver = receiverList.Item(randomReceiver)
        
    End If
    
    ws.AutoFilterMode = False
    FindReceiver = selectedReceiver
    alreadyReceivedList.Add (selectedReceiver)
    
End Function
Function GetRandomNumber(endRange As Integer, Optional startRange As Integer = 1) As Integer
    GetRandomNumber = Int((endRange - startRange + 1) * Rnd + startRange)
End Function



Function GetLastUsedRow(ws As Worksheet) As Long
    GetLastUsedRow = ws.UsedRange.Rows.Count

End Function

Public Function DropMemberFromTheList(memberFullName As String, membersList As Collection)
    Dim listElem() As Variant
    For elemNr = membersList.Count To 1 Step -1
        listElem = membersList.Item(elemNr)
        If listElem(2).Value = memberFullName Then
            membersList.Remove (elemNr)
            Exit For
            
        End If
    Next
End Function

Public Function createTeamMembers() As Collection
    Set createTeamMembers = New Collection
    lastRow = Worksheets("Team").UsedRange.Rows.Count
    ' Create Collection of team members
    For x = 2 To lastRow
        createTeamMembers.Add (Array(Worksheets("Team").Cells(x, 1), Worksheets("Team").Cells(x, 2), Worksheets("Team").Cells(x, 3), Worksheets("Team").Cells(x, 4)))
    Next

End Function

Public Function WriteReport(sender As Variant, receiver As Variant, questions As String)
    Dim unusedRow As Long
    ' find first empty row number
    unusedRow = Worksheets("Report").Cells.SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row
    
    Worksheets("Report").Cells(unusedRow, 1) = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    Worksheets("Report").Cells(unusedRow, 2) = sender(2)
    Worksheets("Report").Cells(unusedRow, 3) = receiver(2)
    Worksheets("Report").Cells(unusedRow, 4) = questions

End Function

Public Function sendEmail(receiver As String, text As String, Optional sendMail As Boolean = False)
    Dim olApp As Outlook.Application
    Dim olEmail As Outlook.MailItem
    
    Set olApp = New Outlook.Application
    Set olEmail = olApp.CreateItem(olMailItem)
    
    With olEmail
    .BodyFormat = olFormatHTML
    .Display
    .HTMLBody = text
    .To = receiver
    .CC = ""
    .Subject = Worksheets("Email").Cells(1, 2).Value
    End With
    
    If (sendMail) Then
        olEmail.Send
    End If
   
End Function

Public Function QuestionList(questionsNr As Integer, listLenght As Integer) As String
    Dim questionsDic: Set questionsDic = CreateObject("Scripting.Dictionary")
    
    Do
        On Error Resume Next
        ' ((upperbound - lowerbound + 1) * Rnd + lowerbound) - start from 2nd question because 1st one is feedback receiver
        myValue = Int((questionsNr - 2 + 1) * Rnd + 2)
        questionsDic.Add Str(myValue), "1"
    Loop Until (questionsDic.Count = listLenght)
    
    QuestionList = Join(questionsDic.Keys, ",")
End Function


Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String
 
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
    Next
    StringFormat = mask
 
End Function




