Sub main()

Dim StartTime As Double
Dim SecondsElapsed As Double


'Remember time when macro starts
  StartTime = Timer

sLocation = "D:\Book1.xls"
sTxtLocation ="D:\test.txt"
Set ObjExl = CreateObject("Excel.Application")
Set ObjWrkBk = ObjExl.Workbooks.Open(sLocation)
Set ObjWrkSht = ObjWrkBk.workSheets("Sheet1")
'ObjExl.Visible = True
'Set FSO = CreateObject("Scripting.FileSystemObject")
'Set FSOFile = FSO.CreateTextFile (sTxtLocation)
'sRowCnt = ObjWrkSht.usedRange.Rows.Count
'sColCnt = ObjWrkSht.usedRange.Columns.Count
'For iLoop = 1 To 10
 ' For jLoop = 1 To 10
 For i=0 To 1000


 ObjWrkSht.range("A1").offset(i,0).value= i

 Next
'oExcel.Cell(1,2).value
    'FSOFile.Write( ObjExl.Cells(iLoop,jLoop).value)
  'Next
'Next

Set ObjWrkBk = Nothing
Set ObjWrkSht = Nothing
Set ObjExl = Nothing
'Set FSO = Nothing
'Set FSOFile = Nothing

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub
