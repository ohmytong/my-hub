# my-hub
Sub test()
Dim Axls As String
Dim PsDoc As Workbook
Dim CurPath, Newpath As String
'On Error Resume Next
ChDrive "D"
ChDir "D:\test"
Axls = Dir("*.xls")
Application.ScreenUpdating = False
Do While Axls <> ""
CurPath = CurDir("D")
Set PsDoc = Workbooks.Open(Axls)
Workbooks(Axls).Activate

ChDrive "D"
ChDir "D:\test\结果"


    Sheets("易感者").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("死亡者").Select
    Range("C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
     
 


ActiveWindow.Close SaveChanges:=True

ChDrive "D"
ChDir "D:\test"
Axls = Dir()
Loop
Application.ScreenUpdating = True
End Sub

