Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'if you double click a cell in a column with "*source" as the header in the first row
Cancel = True
Dim tCol As Integer
tCol = Target.Column
'check for "*source" header in 2nd row and row number is atleast 4 before calling sub
If Cells(2, tCol).Value = "*source" And Target.Row >= 4 Then Call selectFile(Target)

End Sub

' ** this sub will make it work with a single click
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' 'if you select any cell in a column with "*source" as the header in the first row
' Dim tCol As Integer
' tCol = Target.Column
' 'check for "*source" header in 2nd row and row number is atleast 4 before calling sub
' If Cells(2, tCol).Value = "*source" And Target.Row >= 4 Then Call selectFile(Target)
' End Sub

 Sub selectFile(Target As Range)

'create and set dialogue box
Dim dialogBox As FileDialog
Set dialogBox = Application.FileDialog(msoFileDialogOpen)

With dialogBox
'Do not allow multiple files to be selected
.AllowMultiSelect = False

'set the title of the dialog box tab (the title that appears on top tab of finder)
.Title = "select a file"

'set the folder to open. Would this work on mac too?
.InitialFileName = "C:\Users\"

'clear dialog box filters that already exist
.Filters.Clear

'allow only these file extensions. "Excel Workbooks" is the description I am using to describe the kinds of files I want to allow
.Filters.Add "Excel Workbooks", "*.xlsx;*.xls;*.xlsm"

'output the full file *source
.Show
If .SelectedItems.Count = 0 Then
Exit Sub
Else
Target.Value = .SelectedItems(1)
End If

End With

End Sub

