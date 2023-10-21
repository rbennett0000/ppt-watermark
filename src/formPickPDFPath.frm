VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPickPDFPath 
   Caption         =   "UserForm2"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13335
   OleObjectBlob   =   "formPickPDFPath.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formPickPDFPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFileBrowse_Click()

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)

With fd
      .AllowMultiSelect = False
      .Title = "Please select the folder."
      .Filters.Clear

      ' Show the dialog box. If the .Show method returns True, the user picked at least one file. If the .Show method returns False, the user clicked Cancel.
      If .Show = True Then
        formPickPDFPath.textboxPath.Value = .SelectedItems(1)
      End If
End With

Set fd = Nothing

End Sub

Private Sub btnOK_Click()

    pdfFilePath = formPickPDFPath.textboxPath.Value & "\"
    formPickPDFPath.Hide
    AddWatermarks
    
End Sub

