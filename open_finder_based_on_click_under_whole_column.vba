Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'if you double click a cell in column A, change to your needs

Cancel = True

If Target.Column = 1 Then Call selectFile(Target)

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'if you select any cell in column A, change to your needs

'If Target.Column = 1 Then Call selectFile(Target)

End Sub

Sub selectFile(Target As Range)

'create and set dialogue box

Dim filepath As String

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

'output the full file path, added a check for no selected file so no errors

.Show

If .SelectedItems.Count = 0 Then Exit Sub Else filepath = .SelectedItems(1)

End With

End Sub
