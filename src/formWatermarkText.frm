VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formWatermarkText 
   Caption         =   "UserForm1"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12900
   OleObjectBlob   =   "formWatermarkText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formWatermarkText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
 
    textWatermark = TextBox1.Text
    formWatermarkText.Hide
    formPickExcelFile.Show
     
End Sub

Private Sub TextBox1_Change()

    Dim outputString As String
    
    outputString = Replace(Replace(TextBox1.Text, "{n}", textCompanyName), "{d}", Date)
    formWatermarkText.labelPreviewText.Caption = outputString & " - " & outputString & " - " & outputString & " - " & outputString
    
End Sub

Private Sub UserForm_Activate()
    Dim outputString As String
    
    textCompanyName = DEFAULT_COMPANYNAME
    TextBox1.Text = DEFAULT_WATERMARK
    outputString = Replace(Replace(TextBox1.Text, "{n}", textCompanyName), "{d}", Date)
    formWatermarkText.labelPreviewText.Caption = outputString & " - " & outputString & " - " & outputString & " - " & outputString
    
End Sub

