#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms as Long)
#End If
Sub AutoOpen()
'
' AutoOpen Macro
'
Dim Default As String

' Priority Mail
Dim pMessage As String, pTitle As String, pNumCopies As Long, pCt As Range, pTl As Range, pcbk As String, ptbk As String
' Group Mail
Dim gMessage As String, gTitle As String, gNumCogies As Long, gCt As Range, gTl As Range, gcbk As String, gtbk As String
' Bulk Mail
Dim bMessage As String, bTitle As String, bNumCobies As Long, bCt As Range, bTl As Range, bcbk As String, btbk As String
' Subscriber Mail
Dim sMessage As String, sTitle As String, sNumCosies As Long, sCt As Range, sTl As Range, scbk As String, stbk As String


' Set Bookmarks
pcbk = "priorityct"
gcbk = "groupct"
bcbk = "bulkct"
scbk = "subct"

ptbk = "prioritytotal"
gtbk = "grouptotal"
btbk = "bulktotal"
stbk = "subtotal"


' Set Prompt
pMessage = "Enter the number of Priority tray covers that you want to print"
gMessage = "Enter the number of Group tray covers that you want to print"
bMessage = "Enter the number of Bulk tray covers that you want to print"
sMessage = "Enter the number of Subscriber tray covers that you want to print"


' Set Title
pTitle = "Priority"
gTitle = "Group"
bTitle = "Bulk"
sTitle = "Subscriber"

' Set Default
Default = "1"

' Display messages
pNumCopies = Val(InputBox(pMessage, pTitle, Default))
gNumCopies = Val(InputBox(xMessage, gTitle, Default))
bNumCopies = Val(InputBox(bMessage, bTitle, Default))
sNumCopies = Val(InputBox(sMessage, sTitle, Default))


' Retrieve Bookmarks
Set pCt = ActiveDocument.Bookmarks(pcbk).Range
Set gCt = ActiveDocument.Bookmarks(gcbk).Range
Set bCt = ActiveDocument.Bookmarks(bcbk).Range
Set sCt = ActiveDocument.Bookmarks(scbk).Range

Set pTl = ActiveDocument.Bookmarks(ptbk).Range
Set gTl = ActiveDocument.Bookmarks(gtbk).Range
Set bTl = ActiveDocument.Bookmarks(btbk).Range
Set sTl = ActiveDocument.Bookmarks(stbk).Range

' Set Stuff
pTl.Text = pNumCopies
gTl.Text = gNumCopies
bTl.Text = bNumCopies
sTl.Text = sNumCopies


'Print sheets

'Priority
Counter = 0
While Counter < pNumCopies
        pCt.Delete
        pCt.Text = Default + Counter
        ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="4"
        Sleep 3000
        ' Considering using PrintOut's Background:=False and foregoing the sleep.
        ' With the current configuration, Window's print spooler causes the first 9 pages to be spooled and printed last
        Counter = Counter + 1
Wend
pCt.Text = "count"
pTl.Text = "total"

'Group
Counter = 0
While Counter < gNumCopies
        gCt.Delete
        gCt.Text = Default + Counter
        ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="3"
        Sleep 3000
        Counter = Counter + 1
Wend
gCt.Text = "count"
gTl.Text = "total"

'Bulk
Counter = 0
While Counter < bNumCopies
        bCt.Delete
        bCt.Text = Default + Counter
        ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="2"
        Sleep 3000
        Counter = Counter + 1
Wend
bCt.Text = "count"
bTl.Text = "total"

'Subscriber
Counter = 0
While Counter < sNumCopies
        sCt.Delete
        sCt.Text = Default + Counter
        ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="1"
        Sleep 3000
        Counter = Counter + 1
Wend
sCt.Text = "count"
sTl.Text = "total"

' Recreate bookmarks for future use
With ActiveDocument.Bookmarks
        .Add Name:=pcbk Range:=pCt
        .Add Name:=gcbk Range:=gCt
        .Add Name:=bcbk Range:=bCt
        .Add Name:=scbk Range:=sCt
        .Add Name:=ptbk Range:=pTl
        .Add Name:=gtbk Range:=gTl
        .Add Name:=btbk Range:=bTl
        .Add Name:=stbk Range:=sTl
End With

ActiveDocument.Save
End Sub
