Option Explicit

' Move old mail back into Inbox

Sub MoveAgedMail()
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objRootFolder As Outlook.Folder

    Dim lngMovedItems As Long

    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objRootFolder = objNamespace.GetDefaultFolder(olFolderInbox).Parent

    Call MoveAgedMailHelper( _
        objRootFolder.Folders("GTD-incubate"), _
        objRootFolder.Folders("cleanup").Folders("inbox"), _
        lngMovedItems)
    Call MoveAgedMailHelper( _
        objNamespace.GetDefaultFolder(olFolderSentMail), _
        objRootFolder.Folders("cleanup").Folders("sent"), _
        lngMovedItems)

    ' Display the number of items that were moved.
    MsgBox "Moved " & lngMovedItems & " messages(s)."

End Sub

Sub MoveAgedMailHelper(objSourceFolder As Outlook.MAPIFolder, _
    objDestFolder As Outlook.MAPIFolder, _
    lngMovedItems As Long)

    Dim intCount As Integer
    Dim objVariant As Variant
    Dim intDateDiff As Integer
    Dim itemDate As Variant

    For intCount = objSourceFolder.Items.Count To 1 Step -1
        Set objVariant = objSourceFolder.Items.Item(intCount)
        DoEvents ' Process any pending GUI events here

        On Error Resume Next
        itemDate = objVariant.SentOn
        If Not IsDate(itemDate) Then itemDate = objVariant.CreationTime
        If Not IsDate(itemDate) Then itemDate = Now
        On Error GoTo 0

'        If objVariant.Class = olMail Then
'            itemDate = objVariant.SentOn
'        ElseIf objVariant.Class = olReport Then
'            itemDate = objVariant.CreationTime
'        ElseIf objVariant.Class = olMeeting Then ' almost got it now
'            itemDate = objVariant.SentOn
'        Else
'            itemDate = Now
'        End If

        intDateDiff = DateDiff("d", itemDate, Now)

        ' Anything older than 40 days will go
        If intDateDiff > 40 Then
            objVariant.Move objDestFolder

            'count the # of items moved
            lngMovedItems = lngMovedItems + 1
        End If
    Next
End Sub


