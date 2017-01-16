Set ExcelObj = CreateObject("Excel.application")
ExcelObj.Application.Visible = True
    Dim fso
    Dim ts
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set ts = fso.OpenTextFile("### ENTER listfile.txt PATH###")
    
    Dim strLine
    
    Do While Not ts.AtEndOfStream
     strLine = ts.ReadLine()
        ExcelObj.Workbooks.Open ("###YOUR EXCEL FOLDER PATH###" & "\" & strLine)
    'ExcelObj.Worksheets(1).Activate
	
	######ADD YOUR CODE HERE########

	ExcelObj.ActiveWorkbook.SaveAs ("###YOUR EXCEL FOLDER PATH TO SAVE###" & "\" & strLine)

	ExcelObj.Workbooks.close
	
  Loop
ExcelObj.Application.Quit


Set ExcelObj= Nothing
