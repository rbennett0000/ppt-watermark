Attribute VB_Name = "Module1"
Option Explicit
Public textWatermark As String
Public textCompanyName As String
Public excelFilePath As String
Public pdfFilePath As String
Public excelSheet As String
Public excelTable As String
Public excelColumn As String
Public arrayNames As Variant
Public boolUseTable As Boolean
Public userFirstRow As Long
Public userLastRow As Long
Public userFirstCol As Long
Public userLastCol As Long

Public collectionNames As Collection
Public Const DEFAULT_COMPANYNAME As String = "Contoso Corp"
Public Const DEFAULT_WATERMARK As String = "Prepared for {n} on {d}"
Public Const DEFAULT_PDFPREFIX As String = "Watermarked_"

Sub ConvertPPTtoPDFWithTextReplace()

    formWatermarkText.Show
    
End Sub

Sub removeWatermarks()

    Dim pres As Presentation
    Dim slido As slide
    Dim shapo As shape

    Set pres = ActivePresentation
    For Each slido In pres.Slides
        For Each shapo In slido.Shapes
            If shapo.HasTextFrame = True Then
                If shapo.TextFrame.textRange.Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua." Then
                    shapo.Delete
                End If
            End If
        Next shapo
    Next slido
End Sub

Function GetDataFromExcel() As Collection

    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim xlRow As Object

    Dim tbl As Object
    Dim dataCollection As Collection
    Set dataCollection = New Collection
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(excelFilePath)
    
    If boolUseTable Then    ' If user selected a table
        On Error Resume Next
        For Each xlSheet In xlWorkbook.Sheets
            Set tbl = xlSheet.ListObjects(excelTable)
            If Not tbl Is Nothing Then
                Exit For
            End If
        Next xlSheet
        
        On Error GoTo 0
        If Not tbl Is Nothing Then
            Dim tmpRange As Object
            Set tmpRange = tbl.ListColumns(excelColumn).DataBodyRange
        
            For Each xlRow In tmpRange.Rows
                dataCollection.Add xlRow.Cells(1, 1).Value
            Next xlRow
   
        Else
            MsgBox "Table not found in workbook.", vbExclamation, "Table not found!"
            ' TODO where should we go after this error?
        End If
    Else    ' IF user selected a manual range
        
        Dim rowIndex, colIndex As Long
        Set xlSheet = xlWorkbook.Sheets(excelSheet)
        
        For rowIndex = userFirstRow To userLastRow
            For colIndex = userFirstCol To userLastCol
                dataCollection.Add xlSheet.Cells(rowIndex, colIndex).Value
            Next colIndex
        Next rowIndex
    End If
    
    ' Clean up
    xlWorkbook.Close SaveChanges:=False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    Set GetDataFromExcel = dataCollection
End Function

Sub AddWatermarks()

    Dim objPresentation As Presentation
    Dim shpWatermark As shape
    Dim slideIndex As Long
    
    Dim coll As New Collection
    Dim tmpName As Variant
    
    Set collectionNames = GetDataFromExcel()
    
    ' Create watermark shape on first slide
    On Error GoTo Cleanup
    Set objPresentation = ActivePresentation
    Set shpWatermark = objPresentation.Slides(1).Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=-200, Top:=75, Width:=2000, Height:=50)
    ' shpWatermark.TextFrame.textRange.Text = textFullWatermark
    shpWatermark.IncrementRotation (-30)
        
    With shpWatermark.TextFrame.textRange.Font
        .Size = 32
        .Name = "Segoe UI Semibold"
        .Color.RGB = RGB(89, 89, 89)
    End With
        
    With shpWatermark.TextFrame2.textRange.Font
        .Line.Visible = True
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Transparency = 0.56
        .Fill.Transparency = 0.84
    End With
            
    coll.Add shpWatermark
    
    ' Copy watermark shape and paste on each slide
    shpWatermark.Copy
    For slideIndex = 2 To objPresentation.Slides.Count
        coll.Add objPresentation.Slides(slideIndex).Shapes.Paste
    Next slideIndex
   
    For Each tmpName In collectionNames
        Dim item As Object
        Dim pdfFile As String
        Dim outputString As String
        Dim textFullWatermark As String
        
        outputString = Replace(Replace(textWatermark, "{n}", tmpName), "{d}", Date)
        textFullWatermark = Left(outputString & " - " & outputString & " - " & outputString & " - " & outputString, 75)
        
        ' Iterate through all the watermark shapes and change the text
        For Each item In coll
            item.TextFrame.textRange.Text = textFullWatermark
        Next item
           
        ' Export PDF version
        pdfFile = pdfFilePath & DEFAULT_PDFPREFIX & tmpName & ".pdf"
        objPresentation.ExportAsFixedFormat pdfFile, ppFixedFormatTypePDF
        
    Next tmpName
    
Cleanup:
    ' Clean up all the watermark shapes and message result
    For Each item In coll
        item.Delete
    Next item
    
    Set coll = Nothing
    Set item = Nothing
    Set objPresentation = Nothing
    Set shpWatermark = Nothing
    
    If Err.Number = 0 Then
        MsgBox "Watermarking and PDF conversion complete!", vbInformation
    Else
        MsgBox "Operation failed - error number " & Str(Err.Number), vbExclamation
    End If
End Sub

