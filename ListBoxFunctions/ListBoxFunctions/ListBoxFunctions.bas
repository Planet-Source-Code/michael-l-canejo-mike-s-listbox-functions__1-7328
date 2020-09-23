Attribute VB_Name = "Functions"
'                  |4-15-00|
'
'           Mike's Listbox Functions
'
'Thank you for downloading my listbox functions. After Many goof-ups
'
'I finally figured out how to get this working properly.
'
'This is a little different from previous RemoveEmps and RemoveDoubles
'
'You've seen because all these functions use VB's Built-In Listbox
'
'Procedures. So just examine my coding. I didn't comment it because if you're
'
'An intermediate which this was deisnged for, You should know these simple lines of coding.
'
'AIM: mike3dd
'AOL: TheLeadX
'E-mail: Mike_3d@hotmail.com
'Website: Http://8op.com/leaderx

'   -Mike Canejo(Planet Source Code)







Public Function RemoveDoubleString(TheList As ListBox, TheString As String)
On Error Resume Next
    Dim x As Integer, LastPlaceFound As Integer, CountRemoved As Integer, CheckRemoved As Integer
        If TheList.ListCount = 0 Then Exit Function: If TheString = "" Then Exit Function
            HowLong = Timer
CheckDoubles:                  CheckRemoved = TheList.ListCount
            For x = 0 To TheList.ListCount: If TheList.List(x) = TheString Then TheList.RemoveItem (x): LastPlaceFound = x: CountRemoved = CountRemoved + 1: GoTo CheckDoubles
    Next x: If TheList.ListCount > 1 Then If Not CheckRemoved Like TheList.ListCount Then TheList.List(LastPlaceFound) = TheString
MsgBox "Removed " & CountRemoved - 1 & " " & """" & TheString & """" & " doubles in: " & Format(Timer - HowLong, "#.00") & " second(s)", vbSystemModal + vbInformation, "Finished"
'Removes a specified string found more then once in a listbox
End Function

Public Function RemoveDoubles(TheList As ListBox)
On Error Resume Next
    Dim x As Integer, a As Integer, ListTracker As Integer, CountRemoved As Integer, HowLong
        If TheList.ListCount = 0 Then Exit Function
            HowLong = Timer
CheckDoubles:        For x = 0 To TheList.ListCount: ListTracker = TheList.ListCount
            For a = x To TheList.ListCount: If TheList.List(x) = TheList.List(a + 1) Then TheList.RemoveItem (x): CountRemoved = CountRemoved + 1
        Next a: If Not ListTracker Like TheList.ListCount Then GoTo CheckDoubles
    Next x
MsgBox "Removed " & CountRemoved & " doubles in: " & Format(Timer - HowLong, "#.00") & " second(s)", vbSystemModal + vbInformation, "Finished"
'Removes all string found more than once in a listbox
End Function

Public Function RemoveEmptyItems(TheList As ListBox)
On Error Resume Next
    Dim x As Integer, LastPlaceFound As Integer, CountRemoved As Integer, CheckRemoved As Integer
        If TheList.ListCount = 0 Then Exit Function
           HowLong = Timer: CheckRemoved = TheList.ListCount
        For x = 0 To TheList.ListCount: If TheList.List(x) = " " Then TheList.RemoveItem (x): LastPlaceFound = x: CountRemoved = CountRemoved + 1: x = 0
    Next x: If TheList.ListCount > 1 Then If Not CheckRemoved Like TheList.ListCount Then TheList.List(LastPlaceFound) = " "
    MsgBox "Removed " & CountRemoved - 1 & " empty items in: " & Format(Timer - HowLong, "#.00") & " second(s)", vbSystemModal + vbInformation, "Finished"
'Example on how to remove all empty items in a listbox
End Function

