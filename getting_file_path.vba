Sub selectFile()

'create and set dialogue box
Dim dialogBox As FileDialog
Set dialogBox = Application.FileDialog(msoFileDialogOpen)

'Do not allow multiple files to be selected
dialogBox.AllowMultiSelect = False

'set the title of the dialog box tab (the title that appears on top tab of finder)
dialogBox.Title = "select a file"

'set the folder to open
dialogBox.InitialFileName = "C:\Users\"

'clear dialog box filters that already exist
dialogBox.Filters.Clear

'allow only these file extensions. "Excel Workbooks" is the description I am using to describe the kinds of files I want to allow
dialogBox.Filters.Add "Excel Workbooks", "*.xlsx;*.xls;*.xlsm"

'output the full file path
If dialogBox.Show = -1 Then
    ActiveSheet.Range("filepath").Value = dialogBox.SelectedItems(1)
End If

End Sub
