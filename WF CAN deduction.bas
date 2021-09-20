Attribute VB_Name = "Module3"
Sub deduction()

    ''Decide Destination

    Dim Mypath As String
    Mypath = Application.ThisWorkbook.Path

    ''generate date , ACH #, total amount from excel file Name


    Dim excelname As String
    Dim tempdate As String
    Dim filedate As String
    Dim ach As String


    excelname = ActiveWorkbook.Name

    tempdate = Left(excelname, 6)
    filedate = Left(tempdate, 2) & "/" & Mid(tempdate, 3, 2) & "/" & Right(tempdate, 2)

    ach = Mid(excelname, 20, 7)


    ''generate tax csv

    Dim src As Worksheet
    Dim tgt As Worksheet
    Dim lastRow As Long


    Set src = ThisWorkbook.Sheets("deduction")
    src.AutoFilterMode = False

    Sheets.Add.Name = "tgt"
    Set tgt = ThisWorkbook.Sheets("tgt")
    
    lastRow = src.Range("A" & src.Rows.Count).End(xlUp).Row
    
    
    Dim header As Variant
    header = Array("External ID", "Credit #", "Customer", "Date", "Department", "Location", _
        "Currency", "Exchange Rate", "To Be Printed", "To Be E-mailed", "To Be Faxed", _
        "Memo", "PO #", "Item", "Quantity", "Price Level", "Rate", "Sale Amnt", "Description", _
        "Apply_Applied", "Apply_payment")
    tgt.Range("a1:u1") = header
    
    

    tgt.Range("a2:a" & lastRow).Value = "CR0001"
    tgt.Range("b2:b" & lastRow).Value = "21"
    tgt.Range("c2:c" & lastRow).Value = "Wayfair.com : Castlegate - CAN Toronto"
    tgt.Range("d2:d" & lastRow).Value = filedate
    tgt.Range("e2:e" & lastRow).Value = "Dot com"
    tgt.Range("f2:f" & lastRow).Value = "CG-CAN"
    tgt.Range("g2:g" & lastRow).Value = "USD"
    tgt.Range("h2:h" & lastRow).Value = "1"
    tgt.Range("i2:k" & lastRow).Value = "FALSE"
    tgt.Range("l2:l" & lastRow).Value = "Ref. ACH#" & ach
    tgt.Range("m2:m" & lastRow).Value = "Extra deductions (except 5%)"
    tgt.Range("n2:n" & lastRow).Value = src.Range("g2:g" & lastRow).Value
    tgt.Range("o2:o" & lastRow).Value = "1"
    tgt.Range("p2:p" & lastRow).Value = "Custom"
    tgt.Range("q2:q" & lastRow).Value = src.Range("h2:h" & lastRow).Value
    tgt.Range("r2:r" & lastRow).Value = src.Range("h2:h" & lastRow).Value
    tgt.Range("s2:s" & lastRow).Value = src.Range("b2:f" & lastRow).Value

    tgt.Range("q2:r" & lastRow).NumberFormat = "General"

    ''export
    tgt.Name = tempdate & "_WF Canada_deduction"
 

    tgt.Copy
    ActiveWorkbook.SaveAs Filename:=Mypath & "\" & tgt.Name & ".csv", _
    FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    ActiveWorkbook.Close
        
    Application.DisplayAlerts = False
    tgt.Delete
    Application.DisplayAlerts = True


End Sub

