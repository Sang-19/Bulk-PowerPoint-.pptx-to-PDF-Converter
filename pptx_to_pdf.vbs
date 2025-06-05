' Usage: cscript pptx_to_pdf.vbs "C:\path\to\file.pptx"
Dim objPPT, objPresentation
Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Dim pptxPath
pptxPath = WScript.Arguments.Item(0)

If LCase(Right(pptxPath, 5)) = ".pptx" Then
    Dim pdfPath
    pdfPath = Left(pptxPath, Len(pptxPath) - 5) & ".pdf"
    
    Set objPresentation = objPPT.Presentations.Open(pptxPath, False, False, False)
    objPresentation.SaveAs pdfPath, 32 ' 32 = ppSaveAsPDF
    objPresentation.Close
End If

objPPT.Quit
