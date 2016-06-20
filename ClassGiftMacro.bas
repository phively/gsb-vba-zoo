Attribute VB_Name = "ClassGiftMacro"
Option Explicit

Sub ClassGiftMacro()
' Cleaned up by Paul Hively on 3/15/2013
' Completely refactored the old Class Gift Macro to minimize or eliminate duplicate code
' as well as break the double loop.
' Created a variables section at the top; nothing is hard-coded anymore
' Updated on 4/24/2014 with new accounts and a Gift Type column

' Array declaration
Dim GivingLevels As Variant
Dim GivingAmounts As Variant
ReDim Accounts(1 To 100) As String
ReDim Purposes(1 To 100) As String

' ================ Variable values - Can safely be edited! ==============

' Keep Giving Levels and Giving Amounts in order from highest to lowest
GivingLevels = Array("Maroon", "Black", "Platinum", "Gold", "Silver", "Bronze", "None")
GivingAmounts = Array(5000, 2014, 1000, 500, 241, 150, 0)

' Keep Accounts and Purposes in order so there is a 1-1 correspondence
' Be sure to increase Accounts(#) sequentially; do not skip account numbers!
' A blank purpose indicates an account not being counted for class giving
Accounts(1) = "Chicago Booth Annual Fund/2014 Full-Time Class Gift: Annual Fund"
Purposes(1) = "Annual Fund"
Accounts(2) = "Chicago Booth Annual Fund/2014 Full-Time Class Gift: Global Visibility"
Purposes(2) = "Global Visibility"
Accounts(3) = "Chicago Booth Annual Fund/2014 Full-Time Class Gift: Student-Alumni Programming"
Purposes(3) = "Student/Alumni Programming"
' ======================== Do Not Edit Below Here =======================

' Clean up unused array values
Dim i As Long
Dim v As Variant
' Loop through the array until we find an unused value
    i = 0
    For Each v In Accounts
        If v = "" Then Exit For
        i = i + 1
    Next v
    ReDim Preserve Accounts(1 To i)
    ReDim Preserve Purposes(1 To i)

' Count of rows used in the data
' This assumes the Kintera data is pasted in the upper left (A1) of the DATA tab
Dim DataRows As Long
DataRows = Worksheets("DATA").Cells(Range("A:A").Rows.Count, 1).End(xlUp).row
Dim Resultsrows As Long
Resultsrows = Worksheets("Results").Cells(Range("A:A").Rows.Count, 1).End(xlUp).row

' Set up variables to be used for looping
Dim datarow As Long
Dim resultsrow As Long
Dim account As Variant
Dim student As Variant
Dim ArrayPos As Long
Dim givingamount As Variant
Dim acctmatch As Boolean

' Loop through each row in the data file
datarow = 2
Do While datarow <= DataRows
    ' Check each account
    For Each account In Accounts
    ' Check that we're looking at a gift to an approved allocation
    ' Column 8 in DATA is Fund/Designation
        If Worksheets("DATA").Cells(datarow, 8) = account Then
        ' If gift is to an account with blank purpose, also flag with check by hand
            ArrayPos = Application.Match(account, Accounts, False)
            If Purposes(ArrayPos) = "" Then
                Worksheets("DATA").Cells(datarow, 10) = "CHECK BY HAND-ALT"
                Exit For
            End If
        ' Otherwise, check that the First and Last names match
            resultsrow = 2
            ' Loop through the list of students to find a matching name
            Do While resultsrow <= Resultsrows
                ' DATA First/Last are in columns 6 & 7; Results First/Last are in columns 1 & 2
                If Worksheets("DATA").Cells(datarow, 6) = Worksheets("Results").Cells(resultsrow, 1) And Worksheets("DATA").Cells(datarow, 7) = Worksheets("Results").Cells(resultsrow, 2) Then
                    ' Clear the debug field on DATA
                    Worksheets("DATA").Cells(datarow, 10) = ""
                    ' If the person has already made a gift, this needs to be flagged check by hand
                    If Worksheets("Results").Cells(resultsrow, 3) >= 1 Then
                        Worksheets("DATA").Cells(datarow, 10) = "SECOND GIFT - Change By Hand"
                        Exit For
                    End If
                    ' If none of the previous exceptions occurred, fill in the gift count, purpose, amount, etc.
                    ArrayPos = Application.Match(account, Accounts, False)
                    ' Write the data
                    Worksheets("Results").Cells(resultsrow, 3) = 1 ' They made one gift
                    Worksheets("Results").Cells(resultsrow, 4) = Worksheets("DATA").Cells(datarow, 4) ' Pull the amount from the data
                    Worksheets("Results").Cells(resultsrow, 5) = Purposes(ArrayPos) ' Pull purpose associated with this account
                    Worksheets("Results").Cells(resultsrow, 6) = Worksheets("DATA").Cells(datarow, 1) ' Pull the date from the data
                    Worksheets("Results").Cells(resultsrow, 8) = Worksheets("DATA").Cells(datarow, 5) ' Pull the gift type from the data
                    ' Figure out which giving segment they belong to
                    For Each givingamount In GivingAmounts
                        ' When a match is found, write the corresponding level
                        If Worksheets("Results").Cells(resultsrow, 4) >= givingamount Then
                            ArrayPos = Application.Match(givingamount, GivingAmounts, False) - 1
                            Worksheets("Results").Cells(resultsrow, 7) = GivingLevels(ArrayPos) ' Write the giving level
                            Exit For
                        End If
                    Next givingamount
                    ' Since we found the person, get out of this loop
                    Exit For
                ' If we loop through the entire list without finding a matching student, flag check by hand
                ElseIf resultsrow = Resultsrows Then
                    Worksheets("DATA").Cells(datarow, 10) = "CHECK BY HAND - no student match"
                    Exit For
                End If
                resultsrow = resultsrow + 1
            Loop
            
        End If
    Next account
    datarow = datarow + 1
Loop

' Reverse lookup: does anyone on the DATA sheet have the same name as someone on the Results sheet? If so, flag CHECK BY HAND

datarow = 2
Do While datarow <= DataRows
        ' Look only at rows that do not match one of the accounts we're interested in
        acctmatch = False
        For Each account In Accounts
            If Worksheets("DATA").Cells(datarow, 8) = account Then
                acctmatch = True
                Exit For
            End If
        Next account
        'If there was no match, continue
        If acctmatch = False Then
            resultsrow = 2
                ' Loop through the list of students to find a matching name
                Do While resultsrow <= Resultsrows
                    ' DATA First/Last are in columns 6 & 7; Results First/Last are in columns 1 & 2
                    If Worksheets("DATA").Cells(datarow, 6) = Worksheets("Results").Cells(resultsrow, 1) And Worksheets("DATA").Cells(datarow, 7) = Worksheets("Results").Cells(resultsrow, 2) Then
                        ' Mark as Check By Hand
                        Worksheets("DATA").Cells(datarow, 10) = "CHECK BY HAND - non-CG account"
                        ' Get out of the loop since we found the person
                        Exit Do
                    End If
                    resultsrow = resultsrow + 1
                Loop
        End If
    datarow = datarow + 1
Loop

End Sub

' Original macro below, archived for posterity

'Sub ClassGift2012()
''
'' ClassGift2012 Macro
''
'    i = 2
'    j = 2
'    Do Until Worksheets("get_exportfile").Cells(i, 1) = ""
'        flag = 0
'        Do While j < 552
'            If Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/Annual Fund" Then
'                If Worksheets("get_exportfile").Cells(i, 6) = Worksheets("Sheet1").Cells(j, 1) And Worksheets("get_exportfile").Cells(i, 7) = Worksheets("Sheet1").Cells(j, 2) Then
'                    flag = 1
'                    Worksheets("get_exportfile").Cells(i, 10) = ""
'                    If Worksheets("Sheet1").Cells(j, 3) = 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "SECOND GIFT - Change By Hand"
'                    Else
'                        Worksheets("Sheet1").Cells(j, 3) = 1
'                        Worksheets("Sheet1").Cells(j, 4) = Worksheets("get_exportfile").Cells(i, 4)
'                        Worksheets("Sheet1").Cells(j, 5) = "Annual Fund"
'                        Worksheets("Sheet1").Cells(j, 6) = Worksheets("get_exportfile").Cells(i, 1)
'                        If Worksheets("get_exportfile").Cells(i, 4) >= 1000 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Platinum"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 500 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Gold"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 250 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Silver"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 100 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Bronze"
'                        Else
'                            Worksheets("Sheet1").Cells(j, 7) = "None"
'                        End If
'                    End If
'                Else
'                    If flag <> 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND"
'                    End If
'                End If
'            ElseIf Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/Give to the Full Time Student MBA Class Gift - Case Competition Fund" Then
'                If Worksheets("get_exportfile").Cells(i, 6) = Worksheets("Sheet1").Cells(j, 1) And Worksheets("get_exportfile").Cells(i, 7) = Worksheets("Sheet1").Cells(j, 2) Then
'                    flag = 1
'                    Worksheets("get_exportfile").Cells(i, 10) = ""
'                    If Worksheets("Sheet1").Cells(j, 3) = 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "SECOND GIFT - Change By Hand"
'                    Else
'                        Worksheets("Sheet1").Cells(j, 3) = 1
'                        Worksheets("Sheet1").Cells(j, 4) = Worksheets("get_exportfile").Cells(i, 4)
'                        Worksheets("Sheet1").Cells(j, 5) = "Case Competition"
'                        Worksheets("Sheet1").Cells(j, 6) = Worksheets("get_exportfile").Cells(i, 1)
'                        If Worksheets("get_exportfile").Cells(i, 4) >= 1000 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Platinum"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 500 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Gold"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 250 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Silver"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 100 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Bronze"
'                        Else
'                            Worksheets("Sheet1").Cells(j, 7) = "None"
'                        End If
'                    End If
'                Else
'                    If flag <> 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND"
'                    End If
'                End If
'            ElseIf Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/Polsky Center for Entrepreneurship" Then
'                If Worksheets("get_exportfile").Cells(i, 6) = Worksheets("Sheet1").Cells(j, 1) And Worksheets("get_exportfile").Cells(i, 7) = Worksheets("Sheet1").Cells(j, 2) Then
'                    flag = 1
'                    Worksheets("get_exportfile").Cells(i, 10) = ""
'                    If Worksheets("Sheet1").Cells(j, 3) = 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "SECOND GIFT - Change By Hand"
'                    Else
'                        Worksheets("Sheet1").Cells(j, 3) = 1
'                        Worksheets("Sheet1").Cells(j, 4) = Worksheets("get_exportfile").Cells(i, 4)
'                        Worksheets("Sheet1").Cells(j, 5) = "Polsky"
'                        Worksheets("Sheet1").Cells(j, 6) = Worksheets("get_exportfile").Cells(i, 1)
'                        If Worksheets("get_exportfile").Cells(i, 4) >= 1000 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Platinum"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 500 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Gold"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 250 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Silver"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 100 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Bronze"
'                        Else
'                            Worksheets("Sheet1").Cells(j, 7) = "None"
'                        End If
'                    End If
'                Else
'                    If flag <> 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND"
'                    End If
'                End If
'            ElseIf Worksheets("get_exportfile").Cells(i, 8) = " Chicago Booth Annual Fund/Executive MBA Class Gift - Global Visibility" Then
'                If Worksheets("get_exportfile").Cells(i, 6) = Worksheets("Sheet1").Cells(j, 1) And Worksheets("get_exportfile").Cells(i, 7) = Worksheets("Sheet1").Cells(j, 2) Then
'                    flag = 1
'                    Worksheets("get_exportfile").Cells(i, 10) = ""
'                    If Worksheets("Sheet1").Cells(j, 3) = 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "SECOND GIFT - Change By Hand"
'                    Else
'                        Worksheets("Sheet1").Cells(j, 3) = 1
'                        Worksheets("Sheet1").Cells(j, 4) = Worksheets("get_exportfile").Cells(i, 4)
'                        Worksheets("Sheet1").Cells(j, 5) = "Global Visibility"
'                        Worksheets("Sheet1").Cells(j, 6) = Worksheets("get_exportfile").Cells(i, 1)
'                        If Worksheets("get_exportfile").Cells(i, 4) >= 1000 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Platinum"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 500 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Gold"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 250 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Silver"
'                        ElseIf Worksheets("get_exportfile").Cells(i, 4) >= 100 Then
'                            Worksheets("Sheet1").Cells(j, 7) = "Bronze"
'                        Else
'                            Worksheets("Sheet1").Cells(j, 7) = "None"
'                        End If
'                    End If
'                Else
'                    If flag <> 1 Then
'                        Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND"
'                    End If
'                End If
'                 ElseIf Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/Kilts Center for Marketing Annual Fund" Then
'                    Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND-ALT"
'                 ElseIf Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/Reunion Gift" Then
'                    Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND-ALT"
'                 ElseIf Worksheets("get_exportfile").Cells(i, 8) = "Chicago Booth Annual Fund/The Distinguished Fellows Program" Then
'                    Worksheets("get_exportfile").Cells(i, 10) = "CHECK BY HAND-ALT"
'            End If
'            j = j + 1
'        Loop
'        j = 2
'        i = i + 1
'    Loop
