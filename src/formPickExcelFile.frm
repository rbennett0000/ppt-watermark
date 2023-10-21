VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPickExcelFile 
   Caption         =   "UserForm1"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14655
   OleObjectBlob   =   "formPickExcelFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formPickExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComboEventDisabled As Boolean

Private Sub btnFileBrowse_Click()

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

ComboEventDisabled = True
cboTables.Clear
cboColumns.Clear
ComboEventDisabled = False

With fd
      .AllowMultiSelect = False
      .Title = "Please select the file."
      .Filters.Clear
      .Filters.Add "Excel Workbooks", "*.xlsx; *.xls", 1

      ' Show the dialog box. If the .Show method returns True, the user picked at least one file. If the .Show method returns False, the user clicked Cancel.
      If .Show = True Then
        formPickExcelFile.textboxFileName.Value = .SelectedItems(1)
      End If
End With

optionUseTable.Enabled = True
optionUseRange.Enabled = True
Set fd = Nothing
    
End Sub

Private Sub btnOK_Click()
    
    boolUseTable = optionUseTable.Value
    excelFilePath = formPickExcelFile.textboxFileName.Value
    excelTable = cboTables.Value
    excelColumn = cboColumns.Value
    userFirstRow = textboxFirstRow.Value
    userLastRow = textboxLastRow.Value
    userFirstCol = textboxFirstCol.Value
    userLastCol = textboxLastCol.Value
    excelSheet = textboxSheet.Value
    
    formPickExcelFile.Hide
    formPickPDFPath.Show
    
End Sub

Private Sub PopulateTableNames()

Dim xlApp As Object
Dim xlWorkbook As Object
Dim xlSheet As Object
Dim tbl As Object

On Error GoTo Cleanup
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(formPickExcelFile.textboxFileName.Value)
    For Each xlSheet In xlWorkbook.Sheets
        For Each tbl In xlSheet.ListObjects
            cboTables.AddItem tbl.Name
        Next tbl
    Next xlSheet
    
Cleanup:
    On Error Resume Next
    If Not Err.Number = 0 Then
        MsgBox "Error in PopulateTableNames() - cleaning up", vbCritical
    End If
    
    xlWorkbook.Close SaveChanges:=False
    xlApp.Quit
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    On Error GoTo 0

End Sub

Private Sub PopulateColumnNames()

Dim xlApp As Object
Dim xlWorkbook As Object
Dim xlSheet As Object
Dim tbl As Object
Dim column As Object
Dim tblName As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(formPickExcelFile.textboxFileName.Value)
    tblName = cboTables.Value
    ComboEventDisabled = True
    cboColumns.Clear
    ComboEventDisabled = False
    
    On Error Resume Next
    For Each xlSheet In xlWorkbook.Sheets
        Set tbl = xlSheet.ListObjects(tblName)
        If Not tbl Is Nothing Then
            For Each column In tbl.ListColumns
                cboColumns.AddItem column.Name
            Next column
            Exit For
        End If
    Next xlSheet
    
    excelWorkbook.Close SaveChanges:=False
    excelApp.Quit
    Set excelWorkbook = Nothing
    Set excelApp = Nothing

End Sub

Private Sub cboTables_Change()
    If Not ComboEventDisabled Then
        cboColumns.Clear
        PopulateColumnNames
        lblColumnNames.Enabled = True
    End If
End Sub


Private Sub optionUseTable_Click()

    PopulateTableNames
    
    ' Enable table detail fields
    lblTableNames.Enabled = True
    cboTables.Enabled = True
    cboColumns.Enabled = True
    
    cboTables.BackColor = &H80000005
    cboColumns.BackColor = &H80000005
    
    ' Disable range detail fields
    textboxFirstRow.Enabled = False
    textboxLastRow.Enabled = False
    textboxFirstCol.Enabled = False
    textboxLastCol.Enabled = False
    textboxSheet.Enabled = False
    lblFirstRow.Enabled = False
    lblLastRow.Enabled = False
    lblFirstCol.Enabled = False
    lblLastCol.Enabled = False
    lblSheet.Enabled = False
    
    textboxFirstRow.BackColor = &H80000004
    textboxLastRow.BackColor = &H80000004
    textboxFirstCol.BackColor = &H80000004
    textboxLastCol.BackColor = &H80000004
    textboxSheet.BackColor = &H80000004
    
End Sub

Private Sub optionUseRange_Click()

    ' Enable range detail fields
    textboxFirstRow.Enabled = True
    textboxLastRow.Enabled = True
    textboxFirstCol.Enabled = True
    textboxLastCol.Enabled = True
    textboxSheet.Enabled = True
    lblFirstRow.Enabled = True
    lblLastRow.Enabled = True
    lblFirstCol.Enabled = True
    lblLastCol.Enabled = True
    lblSheet.Enabled = True
    
    textboxFirstRow.BackColor = &H80000005
    textboxLastRow.BackColor = &H80000005
    textboxFirstCol.BackColor = &H80000005
    textboxLastCol.BackColor = &H80000005
    textboxSheet.BackColor = &H80000005
      
    ' Disable table detail fields
    cboTables.Enabled = False
    cboColumns.Enabled = False
    lblTableNames.Enabled = False
    lblColumnNames.Enabled = False
    
    cboTables.BackColor = &H80000004
    cboColumns.BackColor = &H80000004
    

End Sub

Private Sub TextBox2_Change()

End Sub
