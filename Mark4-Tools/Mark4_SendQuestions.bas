Attribute VB_Name = "Mark4_SendQuestions"


Sub CreateEmailSet()
    Dim mailBody As String
    Dim person()
    Dim receiver()
    Dim questionsStr As String: questionsStr = ""
    Dim questionsCol As New Collection ' to assure unique list of question numbers
    Dim runOffset As Integer ' to assure that each time receiver get other person for sending feedback
    
    Dim teamMembers As New Collection
    
    
    Dim teamSize As Integer: teamSize = Worksheets("Team").Cells.SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row - 2
    Dim questionsNr As Integer: questionsNr = Int(Worksheets("Variables").Cells(3, 2))
    Dim shouldSendEmail As Boolean: shouldSendEmail = Worksheets("Variables").Cells(2, 2)
    
    bodyTemplate = Worksheets("Email").Cells(2, 2)
    
    runOffset = Int(Worksheets("Variables").Cells(1, 2))
    If runOffset = (teamSize - 2) Then
        Worksheets("Variables").Cells(1, 2) = 0
    Else
        Worksheets("Variables").Cells(1, 2) = runOffset + 1
    End If
    
    Randomize
      
    ' Create Collection of team members
    For x = 1 To teamSize
        teamMembers.Add (Array(Worksheets("Team").Cells(x + 1, 1), Worksheets("Team").Cells(x + 1, 2), Worksheets("Team").Cells(x + 1, 3), Worksheets("Team").Cells(x + 1, 4)))
    Next
    
    ' create emails and write reports
    For x = 1 To teamSize
        receiver = teamMembers.Item(x)
        Dim receiverNr As Integer: receiverNr = 1 + (runOffset + x) Mod (teamSize)
        person = teamMembers.Item(receiverNr)
        
        questionsStr = QuestionList(questionsNr, 6)
        mailBody = StringFormat(bodyTemplate, receiver(0), person(2), questionsStr)
        Call sendEmail(receiver(3).Value, mailBody, shouldSendEmail)
        WriteReport receiver, person, questionsStr
    Next
    
End Sub

Public Function WriteReport(receiver As Variant, person As Variant, questions As String)
    Dim unusedRow As Long
    ' find first empty row number
    unusedRow = Worksheets("Report").Cells.SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row
    
    Worksheets("Report").Cells(unusedRow, 1) = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    Worksheets("Report").Cells(unusedRow, 2) = receiver(2)
    Worksheets("Report").Cells(unusedRow, 3) = person(2)
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
    Dim questionsStr As String: questionsStr = ""
    Dim questionsCol As New Collection
    
    myValue = Int((questionsNr * Rnd) + 1)
    questionsStr = myValue
    questionsCol.Add "1", Str(myValue)
    questionsStr = myValue
    Do
        On Error Resume Next
        myValue = Int((questionsNr * Rnd) + 1)
        questionsCol.Add "1", Str(myValue)
        questionsStr = questionsStr & ", " & myValue
    Loop Until (questionsCol.Count = listLenght)
    
    QuestionList = questionsStr
End Function

Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String
 
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
    Next
    StringFormat = mask
 
End Function



