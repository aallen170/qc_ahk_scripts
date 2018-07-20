oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
oExcel.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)

oExcel.Range("A1").Value := 3 ; set cell A1 to 3
oExcel.Range("A2").Value := 7 ; set cell A2 to 7
oExcel.Range("A3").Formula := "=SUM(A1:A2)" ; set formula for cell A3 to SUM(A1:A2)

oExcel.Range("A1:A3").Interior.ColorIndex := 19 ; fill range of cells from A1 to A3 with color number 19
oExcel.Range("A3").Borders(8).LineStyle := 1 ; set top border line style for cell A3 (xlEdgeTop = 8, xlContinuous = 1)
oExcel.Range("A3").Borders(8).Weight := 2 ; set top border weight for cell A3 (xlThin = 2)
oExcel.Range("A3").Font.Bold := 1 ; set bold font for cell A3

A1 := oExcel.Range("A1").Value ; get value from cell A1, and store it in A1 variable
oExcel.Range("A4").Select ; select A4 cell
oExcel.Visible := 1 ; make Excel Application visible
MsgBox % A1 "`n" oExcel.Range("A2").Value ; check. You can use Round() function to round numbers to the nearest integer
ExitApp